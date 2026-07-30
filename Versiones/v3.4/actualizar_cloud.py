#!/usr/bin/env python3
"""
actualizar_cloud.py — v2.0  Versión cloud sin dependencias externas de Google

Descarga archivos de Google Drive a un directorio temporal,
ejecuta la lógica de actualizar_bosquegin.py y sube bosquegin_data.js
a GitHub via API (sin necesidad de tener la PC encendida).

Variables de entorno requeridas:
  GOOGLE_OAUTH_CLIENT_ID       — OAuth2 client ID
  GOOGLE_OAUTH_CLIENT_SECRET   — OAuth2 client secret
  GOOGLE_OAUTH_REFRESH_TOKEN   — OAuth2 refresh token (obtenido con get_oauth_token.py)
  DRIVE_ROOT_FOLDER_ID         — ID de la carpeta raíz en Drive
                                  (la que contiene la carpeta "Data")
  GITHUB_TOKEN                 — Token de acceso personal a GitHub
"""
import os, sys, json, shutil, tempfile, importlib.util, urllib.request, urllib.parse, base64
from datetime import datetime, timedelta, timezone

_AR = timezone(timedelta(hours=-3))

GITHUB_REPO = "bosquegin/bosquegin-tablero"
GITHUB_FILE = "bosquegin_data.js"
HERE        = os.path.dirname(os.path.abspath(__file__))


# ═══════════════════════════════════════════════════════════════════════════════
#  GOOGLE OAUTH2 — token refresh sin dependencias externas
# ═══════════════════════════════════════════════════════════════════════════════

def _get_access_token():
    """Obtiene un access token OAuth2 via refresh token (stdlib puro)."""
    data = urllib.parse.urlencode({
        "client_id":     os.environ["GOOGLE_OAUTH_CLIENT_ID"],
        "client_secret": os.environ["GOOGLE_OAUTH_CLIENT_SECRET"],
        "refresh_token": os.environ["GOOGLE_OAUTH_REFRESH_TOKEN"],
        "grant_type":    "refresh_token",
    }).encode()
    req = urllib.request.Request(
        "https://oauth2.googleapis.com/token", data=data,
        headers={"Content-Type": "application/x-www-form-urlencoded"}
    )
    with urllib.request.urlopen(req, timeout=15) as r:
        resp = json.loads(r.read())
    if "access_token" not in resp:
        raise RuntimeError(f"Token refresh fallido: {resp}")
    return resp["access_token"]


# ═══════════════════════════════════════════════════════════════════════════════
#  GOOGLE DRIVE — llamadas directas a la REST API v3
# ═══════════════════════════════════════════════════════════════════════════════

def _drive_get(token, endpoint, params=None):
    """GET a Drive API v3."""
    url = f"https://www.googleapis.com/drive/v3/{endpoint}"
    if params:
        url += "?" + urllib.parse.urlencode(params)
    req = urllib.request.Request(url, headers={"Authorization": f"Bearer {token}"})
    with urllib.request.urlopen(req, timeout=30) as r:
        return json.loads(r.read())


def _find_folder(token, parent_id, name):
    """Busca una subcarpeta por nombre dentro de parent_id."""
    q = (f"'{parent_id}' in parents and name='{name}' "
         f"and mimeType='application/vnd.google-apps.folder' and trashed=false")
    r = _drive_get(token, "files", {"q": q, "fields": "files(id,name)", "pageSize": "50"})
    items = r.get("files", [])
    return items[0]["id"] if items else None


def _list_files(token, folder_id, name_contains=None):
    """Lista archivos (no carpetas) en una carpeta Drive."""
    q = (f"'{folder_id}' in parents "
         f"and mimeType!='application/vnd.google-apps.folder' and trashed=false")
    if name_contains:
        q += f" and name contains '{name_contains}'"
    r = _drive_get(token, "files",
                   {"q": q, "fields": "files(id,name)", "orderBy": "name", "pageSize": "1000"})
    return r.get("files", [])


def _download_file(token, file_id, dest_path):
    """Descarga un archivo de Drive a dest_path via streaming."""
    url = f"https://www.googleapis.com/drive/v3/files/{file_id}?alt=media"
    req = urllib.request.Request(url, headers={"Authorization": f"Bearer {token}"})
    os.makedirs(os.path.dirname(dest_path), exist_ok=True)
    with urllib.request.urlopen(req, timeout=120) as r, open(dest_path, "wb") as f:
        shutil.copyfileobj(r, f)


def _download_folder(token, folder_id, dest_dir, name_contains=None):
    """Descarga todos los archivos de una carpeta Drive a dest_dir."""
    os.makedirs(dest_dir, exist_ok=True)
    files = _list_files(token, folder_id, name_contains)
    for f in files:
        dest = os.path.join(dest_dir, f["name"])
        _download_file(token, f["id"], dest)
        print(f"  ↓ {f['name']}")
    return len(files)


def _download_named(token, folder_id, filename, dest_path):
    """Descarga un archivo específico por nombre."""
    files = _list_files(token, folder_id)
    match = next((f for f in files if f["name"] == filename), None)
    if not match:
        print(f"  ⚠ No encontrado en Drive: {filename}")
        return False
    _download_file(token, match["id"], dest_path)
    print(f"  ↓ {filename}")
    return True


def download_from_drive(tmpdir):
    """
    Descarga la estructura Data/ de Google Drive al directorio temporal.
    Estructura esperada en Drive:
      <DRIVE_ROOT_FOLDER_ID>/
        Data/
          Inventario/         ← Stock*.xlsx + Stock_consolidado*.xlsx
          Salidas/
            Bosque salidas.xlsx
            GC/               ← Remitos GC*.xlsx
          Costos y PVP/       ← Analisis de costos y PVP - COSTOS.csv
          Supply Chain/
            proyecciones/     ← proyecciones.xlsx (opcional)
    """
    root_id = os.environ.get("DRIVE_ROOT_FOLDER_ID", "").strip()
    if not root_id:
        raise ValueError("Falta variable de entorno: DRIVE_ROOT_FOLDER_ID")

    print("  Obteniendo token OAuth2...")
    token = _get_access_token()

    # Navegar a carpeta Data (puede ser el root mismo o una subcarpeta)
    data_id = _find_folder(token, root_id, "Data") or root_id
    print(f"  Drive: carpeta Data encontrada (id={data_id[:8]}...)")

    # Inventario (stock Excel + consolidado)
    print("  → Inventario...")
    inv_id = _find_folder(token, data_id, "Inventario")
    if inv_id:
        n = _download_folder(token, inv_id, os.path.join(tmpdir, "Data", "Inventario"))
        print(f"     {n} archivos descargados")
    else:
        print("  ⚠ Carpeta Inventario no encontrada")

    # Salidas (ventas Excel + GC remitos)
    print("  → Salidas...")
    sal_id = _find_folder(token, data_id, "Salidas")
    if sal_id:
        _download_named(token, sal_id, "Bosque salidas.xlsx",
                        os.path.join(tmpdir, "Data", "Salidas", "Bosque salidas.xlsx"))
        _download_named(token, sal_id, "Salidas_consolidado.xlsx",
                        os.path.join(tmpdir, "Data", "Salidas", "Salidas_consolidado.xlsx"))
        gc_id = _find_folder(token, sal_id, "GC")
        if gc_id:
            print("  → GC (remitos)...")
            n = _download_folder(token, gc_id,
                                 os.path.join(tmpdir, "Data", "Salidas", "GC"),
                                 name_contains="Remitos GC")
            print(f"     {n} remitos")

    # Costos CSV
    print("  → Costos...")
    cos_id = _find_folder(token, data_id, "Costos y PVP")
    if cos_id:
        _download_folder(token, cos_id, os.path.join(tmpdir, "Data", "Costos y PVP"))

    # Supply Chain / proyecciones
    sc_id = _find_folder(token, data_id, "Supply Chain")
    if sc_id:
        proy_id = _find_folder(token, sc_id, "proyecciones")
        if proy_id:
            print("  → Proyecciones...")
            _download_folder(token, proy_id,
                             os.path.join(tmpdir, "Data", "Supply Chain", "proyecciones"))

    # Productos (lookup rubro/subrubro)
    prod_id = _find_folder(token, data_id, "Productos")
    if prod_id:
        print("  → Productos...")
        _download_folder(token, prod_id, os.path.join(tmpdir, "Data", "Productos"))


# ═══════════════════════════════════════════════════════════════════════════════
#  GITHUB API
# ═══════════════════════════════════════════════════════════════════════════════

def push_to_github(content, token):
    """Sube bosquegin_data.js a GitHub via API."""
    api_url = f"https://api.github.com/repos/{GITHUB_REPO}/contents/{GITHUB_FILE}"
    hdrs = {
        "Authorization": f"token {token}",
        "Accept": "application/vnd.github.v3+json",
        "Content-Type": "application/json",
    }

    # Obtener SHA actual
    req = urllib.request.Request(api_url)
    for k, v in hdrs.items():
        req.add_header(k, v)
    with urllib.request.urlopen(req, timeout=30) as r:
        sha = json.loads(r.read())["sha"]

    # Push
    body = json.dumps({
        "message": f"data: actualizar cloud {datetime.now(_AR).strftime('%Y-%m-%d %H:%M')}",
        "content": base64.b64encode(content.encode("utf-8")).decode("ascii"),
        "sha": sha,
    }).encode("utf-8")
    req2 = urllib.request.Request(api_url, data=body, method="PUT")
    for k, v in hdrs.items():
        req2.add_header(k, v)
    with urllib.request.urlopen(req2, timeout=30) as r:
        result = json.loads(r.read())

    commit_url = result.get("commit", {}).get("html_url", "")
    print(f"  ✓ Publicado en GitHub: {commit_url}")


# ═══════════════════════════════════════════════════════════════════════════════
#  MAIN
# ═══════════════════════════════════════════════════════════════════════════════

def main():
    github_token = os.environ.get("GITHUB_TOKEN", "").strip()
    if not github_token:
        raise ValueError("Falta variable de entorno: GITHUB_TOKEN")

    tmpdir = tempfile.mkdtemp(prefix="bosquegin_cloud_")
    print(f"  Directorio temporal: {tmpdir}")

    try:
        # Crear estructura de directorios
        for d in ["Data/Inventario", "Data/Salidas/GC",
                  "Data/Costos y PVP", "Data/Insumos", "Data/Productos",
                  "Data/Supply Chain/proyecciones"]:
            os.makedirs(os.path.join(tmpdir, d), exist_ok=True)

        # ── [1] Descargar de Drive ────────────────────────────────────────────
        print("\n[1/4] Descargando archivos de Google Drive...")
        download_from_drive(tmpdir)

        # ── [2] Procesar con actualizar_bosquegin.py ──────────────────────────
        print("\n[2/4] Procesando datos...")
        spec = importlib.util.spec_from_file_location(
            "actualizar_bg", os.path.join(HERE, "actualizar_bosquegin.py"))
        mod = importlib.util.module_from_spec(spec)
        spec.loader.exec_module(mod)

        # Redirigir TODAS las rutas al directorio temporal
        mod.BASE           = tmpdir
        mod.DATA_DIR       = os.path.join(tmpdir, "Data")
        mod.INV_DIR        = os.path.join(tmpdir, "Data", "Inventario")
        mod.VENTAS_F       = os.path.join(tmpdir, "Data", "Salidas", "Bosque salidas.xlsx")
        mod.GC_SALIDAS_DIR = os.path.join(tmpdir, "Data", "Salidas", "GC")
        mod.COSTOS_CSV     = os.path.join(tmpdir, "Data", "Costos y PVP",
                                          "Analisis de costos y PVP - COSTOS.csv")
        mod.INSUMOS_CSV    = os.path.join(tmpdir, "Data", "Insumos", "Stock insumos.csv")
        mod.SALIDAS_CONS   = os.path.join(tmpdir, "Data", "Salidas", "Salidas_consolidado.xlsx")
        mod.CONS_FILE      = os.path.join(tmpdir, "Data", "Inventario",
                                          "Stock_consolidado_por_deposito_y_dia.xlsx")
        mod.PROD_F         = os.path.join(tmpdir, "Data", "Productos", "PRODUCTOS.xlsx")
        mod.PROY_DIR       = os.path.join(tmpdir, "Data", "Supply Chain", "proyecciones")
        mod.PROY_FILE      = os.path.join(tmpdir, "Data", "Supply Chain", "proyecciones",
                                          "proyecciones.xlsx")
        mod.OUT_JS         = os.path.join(tmpdir, "bosquegin_data.js")

        # En cloud no hay CDP (costos usa CSV de Drive)
        mod._download_costos_via_cdp = lambda url, port=9222: None

        # Saltar git push (lo hacemos nosotros via API)
        mod._SKIP_GIT_PUSH = True

        mod.main()

        # ── [3] Subir bosquegin_data.js a GitHub ──────────────────────────────
        print("\n[3/4] Publicando en GitHub via API...")
        with open(mod.OUT_JS, encoding="utf-8") as f:
            js_content = f.read()
        push_to_github(js_content, github_token)

        print("\n[4/4] ✅ Actualización cloud completada")
        print("      El tablero se actualiza en ~60 segundos")

    finally:
        shutil.rmtree(tmpdir, ignore_errors=True)


if __name__ == "__main__":
    main()
