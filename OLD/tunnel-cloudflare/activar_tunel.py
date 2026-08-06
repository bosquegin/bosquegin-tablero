#!/usr/bin/env python3
"""
activar_tunel.py — Inicia el túnel Cloudflare y publica la URL en GitHub.

Uso:
  1. Asegurate de que el servidor (iniciar_tablero.bat) ya esté corriendo.
  2. Ejecutá este script (doble clic en activar_publico.bat).
  3. La URL del túnel se publica automáticamente en GitHub.
  4. Los editores ya pueden usar "Actualizar" desde el tablero.
  5. Para cerrar el túnel, cerrá esta ventana (Ctrl+C).
"""
import os, re, subprocess, sys, threading, time

import platform
if platform.system() == "Windows":
    BASE = r"C:\Users\SupplyDestileria\Documents\Bosque\Claude\Tablero operativo"
else:
    BASE = os.path.dirname(os.path.abspath(__file__))

CF_EXE          = os.path.join(BASE, "cloudflared.exe")
TUNNEL_URL_FILE = os.path.join(BASE, "tunnel_url.txt")
URL_RE          = re.compile(r'https://[a-z0-9-]+\.trycloudflare\.com')


def _git(cmds):
    for cmd in cmds:
        r = subprocess.run(cmd, cwd=BASE, capture_output=True, text=True)
        ok = r.returncode == 0 or "nothing to commit" in (r.stdout + r.stderr)
        print(f"  git {cmd[1]}: {'OK' if ok else r.stderr.strip()[:80]}")


def _push_url(url):
    print("\nPublicando URL en GitHub...")
    with open(TUNNEL_URL_FILE, "w", encoding="utf-8") as f:
        f.write(url)
    _git([
        ["git", "add", "tunnel_url.txt"],
        ["git", "commit", "-m", "tunnel: servidor publico activo"],
        ["git", "pull", "--rebase", "origin", "main"],
        ["git", "push", "origin", "main"],
    ])
    print(f"\n✅  Editores ya pueden usar Actualizar desde: https://bosquegin.github.io/bosquegin-tablero/")
    print(f"    (URL del túnel: {url})\n")


def _clear_url():
    print("\nCerrando túnel — limpiando URL en GitHub...")
    with open(TUNNEL_URL_FILE, "w", encoding="utf-8") as f:
        f.write("")
    _git([
        ["git", "add", "tunnel_url.txt"],
        ["git", "commit", "-m", "tunnel: servidor offline"],
        ["git", "pull", "--rebase", "origin", "main"],
        ["git", "push", "origin", "main"],
    ])
    print("  Hecho — tablero muestra 'servidor offline' a los editores.")


def main():
    if not os.path.exists(CF_EXE):
        print(f"ERROR: cloudflared.exe no encontrado en:\n  {CF_EXE}")
        input("Presioná Enter para cerrar...")
        sys.exit(1)

    print("=" * 60)
    print("  BOSQUE GIN — SERVIDOR PÚBLICO")
    print("=" * 60)
    print(f"\nIniciando túnel Cloudflare → http://127.0.0.1:7891")
    print("(Asegurate de que el servidor local ya esté corriendo)\n")

    proc = subprocess.Popen(
        [CF_EXE, "tunnel", "--url", "http://127.0.0.1:7891"],
        cwd=BASE,
        stdout=subprocess.PIPE,
        stderr=subprocess.STDOUT,
        text=True,
        bufsize=1,
        encoding="utf-8",
        errors="replace",
    )

    tunnel_url = None

    try:
        for line in proc.stdout:
            line = line.strip()
            if not line:
                continue
            # Mostrar solo líneas relevantes (sin spam de métricas)
            if any(k in line.lower() for k in ("error", "warn", "url", "tunnel", "visit", "https", "trycloudflare")):
                print(f"  {line[:120]}")

            m = URL_RE.search(line)
            if m and not tunnel_url:
                tunnel_url = m.group(0)
                # Publicar en hilo separado para no bloquear el stdout del proceso
                threading.Thread(target=_push_url, args=(tunnel_url,), daemon=True).start()

    except KeyboardInterrupt:
        pass
    finally:
        proc.terminate()
        try:
            proc.wait(timeout=5)
        except Exception:
            pass

    try:
        _clear_url()
    except Exception as e:
        print(f"  Advertencia al limpiar URL: {e}")


if __name__ == "__main__":
    try:
        main()
    except KeyboardInterrupt:
        print("\nCerrando...")
    input("\nPresioná Enter para cerrar esta ventana...")
