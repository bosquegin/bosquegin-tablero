#!/usr/bin/env python3
"""
servidor_render.py — Servidor Flask para Render.com

Expone /actualizar (POST) que descarga datos de Google Drive,
procesa con actualizar_bosquegin.py y sube bosquegin_data.js a GitHub.

Variables de entorno requeridas (configurar en Render dashboard):
  GOOGLE_SERVICE_ACCOUNT_JSON  — JSON del service account (todo en una línea)
  DRIVE_ROOT_FOLDER_ID         — ID de la carpeta raíz en Drive
  GITHUB_TOKEN                 — Personal access token con repo scope

Variables opcionales:
  PORT                         — Puerto HTTP (Render lo setea automáticamente)
  CLOUD_COOLDOWN               — Segundos mínimos entre actualizaciones (default 180)
"""
import os, json, time, hmac, hashlib, threading
from flask import Flask, request, Response, stream_with_context
from flask_cors import CORS

app = Flask(__name__)
CORS(app, origins=["https://bosquegin.github.io"])

_lock          = threading.Lock()
_last_update   = 0.0
_COOLDOWN      = int(os.environ.get("CLOUD_COOLDOWN", "180"))
_running       = False

# URL pública del repo para obtener auth_static.js
_AUTH_URL = (
    "https://raw.githubusercontent.com/bosquegin/bosquegin-tablero/main/auth_static.js"
)


# ── Auth ─────────────────────────────────────────────────────────────────────

def _fetch_users():
    """Descarga auth_static.js de GitHub y parsea window.BG_AUTH."""
    import urllib.request, re
    try:
        with urllib.request.urlopen(_AUTH_URL, timeout=10) as r:
            raw = r.read().decode("utf-8")
        # Extraer el array JSON después de "window.BG_AUTH = "
        m = re.search(r'window\.BG_AUTH\s*=\s*(\[.*\])\s*;', raw, re.DOTALL)
        if m:
            return json.loads(m.group(1))
    except Exception as e:
        app.logger.warning(f"No se pudo obtener auth_static.js: {e}")
    return []


def _verify_token(username, token):
    """Verifica cloud_token contra la lista de usuarios en GitHub."""
    if not username or not token:
        return None
    users = _fetch_users()
    for u in users:
        if u.get("username") == username:
            stored = u.get("cloud_token", "")
            if stored and hmac.compare_digest(stored, token):
                return u
    return None


# ── Endpoints ─────────────────────────────────────────────────────────────────

@app.route("/health")
def health():
    return {"status": "ok", "time": time.time()}, 200


@app.route("/actualizar", methods=["POST"])
def actualizar():
    global _last_update, _running

    # 1. Auth
    try:
        body = request.get_json(force=True) or {}
    except Exception:
        return {"ok": False, "error": "JSON inválido"}, 400

    username = (body.get("username") or "").strip()
    token    = (body.get("cloud_token") or "").strip()
    user     = _verify_token(username, token)
    if not user:
        return {"ok": False, "error": "No autorizado"}, 401
    if user.get("role") not in ("admin", "editor"):
        return {"ok": False, "error": "Sin permiso"}, 403

    # 2. Rate limit
    with _lock:
        now = time.time()
        remaining = int(_COOLDOWN - (now - _last_update))
        if remaining > 0:
            return {"ok": False, "error": f"Espera {remaining}s antes de volver a actualizar"}, 429
        if _running:
            return {"ok": False, "error": "Ya hay una actualización en curso"}, 429
        _last_update = now
        _running = True

    # 3. Stream output
    def generate():
        global _running
        import io, sys, importlib.util, os

        # Capturar stdout con un writer que hace yield
        class StreamCapture(io.TextIOBase):
            def __init__(self):
                self.buf = []
            def write(self, s):
                if s:
                    self.buf.append(s)
                return len(s)
            def flush(self):
                pass

        cap = StreamCapture()
        old_stdout = sys.stdout
        sys.stdout = cap

        try:
            # Importar actualizar_cloud desde el mismo directorio
            here = os.path.dirname(os.path.abspath(__file__))
            spec = importlib.util.spec_from_file_location(
                "actualizar_cloud",
                os.path.join(here, "actualizar_cloud.py")
            )
            mod = importlib.util.module_from_spec(spec)
            spec.loader.exec_module(mod)
            mod.main()

            sys.stdout = old_stdout
            # Devolver todo lo capturado
            yield "".join(cap.buf)
            yield "\n✅ Actualización completada"

        except Exception as e:
            sys.stdout = old_stdout
            yield "".join(cap.buf)
            yield f"\n❌ Error: {e}"
        finally:
            sys.stdout = old_stdout
            _running = False

    return Response(
        stream_with_context(generate()),
        content_type="text/plain; charset=utf-8"
    )


# ── Main ──────────────────────────────────────────────────────────────────────

if __name__ == "__main__":
    port = int(os.environ.get("PORT", 8080))
    app.run(host="0.0.0.0", port=port)
