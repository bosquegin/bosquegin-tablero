#!/usr/bin/env python3
"""
actualizar_cloud.py — Versión cloud de actualizar_bosquegin.py

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
import os, sys, json, shutil, tempfile, importlib.util, urllib.request, base64
from datetime import datetime

GITHUB_REPO = "bosquegin/bosquegin-tablero"
GITHUB_FILE = "bosquegin_data.js"
HERE        = os.path.dirname(os.path.abspath(__file__))


# ═══════════════════════════════════════════════════════════════════════════════
#  GOOGLE DRIVE
# ═══════════════════════════════════════════════════════════════════════════════

def _drive_service():
    from google.oauth2.credentials import Credentials
    from googleapiclient.discovery import build
    creds = Credentials(
        token=None,
        refresh_token=os.environ["GOOGLE_OAUTH_REFRESH_TOKEN"],
        client_id=os.environ["GOOGLE_OAUTH_CLIENT_ID"],
        client_secret=os.environ["GOOGLE_OAUTH_CLIENT_SECRET"],
        token_uri="https://oauth2.googleapis.com/token",
        scopes=["https://www.googleapis.com/auth/drive.readonly"],
    )
    return build("drive", "v3", credentials=creds, cache_discovery=False)


def _find_folder(svc, parent_id, name):
    """Busca una subcarpeta por nombre dentro de parent_id."""
    q = (f"'{parent_id}' in parents and name='{name}' "
         f"and mimeType='application/vnd.google-apps.folder' and trashed=false")
    r = svc.files().list(q=q, fields="files(id,name)").execute()
    items = r.get("files", [])
    return items[0]["id"] if items else None


def _list_files(svc, folder_id, name_contains=None):
    """Lista archivos (no carpetas) en una carpeta Drive."""
    q = (f"'{folder_id}' in parents "
         f"and mimeType!='application/vnd.google-apps.folder' and trashed=false")
    if name_contains:
        q += f" and name contains '{name_contains}'"
    r = svc.files().list(q=q, fields="files(id,name)", orderBy="name").execute()
    return r.get("files", [])


def _download_file(svc, file_id, dest_path):
    """Descarga un archivo de Drive a dest_path."""
    from googleapiclient.http import MediaIoBaseDownload
    import io
    os.makedirs(os.path.dirname(dest_path), exist_ok=True)
    req = svc.files().get_media(fileId=file_id)
    with open(dest_path, "wb") as f:
        dl = MediaIoBaseDownload(f, req)
        done = False
        while not done:
            _, done = dl.next_chunk()


def _download_folder(svc, folder_id, dest_dir, name_contains=None):
    """Descarga todos los archivos de una carpeta Drive a dest_dir."""
    os.makedirs(dest_dir, exist_ok=True)
    files = _list_files(svc, folder_id, name_contains)
    for f in files:
        dest = os.path.join(dest_dir, f["name"])
        _download_file(svc, f["id"], dest)
        print(f"  ↓ {f['name']}")
    return len(files)


def _download_named(svc, folder_id, filename, dest_path):
    """Descarga un archivo específico por nombre."""
    files = _list_files(svc, folder_id)
    match = next((f for f in files if f["name"] == filename), None)
    if not match:
        print(f"  ⚠ No encontrado en Drive: {filename}")
        return False
    _download_file(svc, match["id"], dest_path)
    print(f"  ↓ {filename}")
    return True


def download_from_drive(tmpdir):
    """
    Descarga la estructura Data/ de Google Drive al directorio temporal.
    Estructura esperada en Drive:
      <DRIVE_ROOT_FOLDER_ID>/
        Data/
          Inventario/         ← archivos Stock*.xlsx
          Salidas/
            Bosque salidas.xlsx
            Salidas_consolidado.xlsx (opcional)
            GC/               ← Remitos GC*.xlsx
          Costos y PVP/       ← Analisis de costos y PVP - COSTOS.csv
          Supply Chain/
            proyecciones/     ← proyecciones.xlsx (opcional)
    """
    root_id = os.environ.get("DRIVE_ROOT_FOLDER_ID", "").strip()
    if not root_id:
        raise ValueError("Falta variable de entorno: DRIVE_ROOT_FOLDER_ID")

    svc = _drive_service()

    # Navegar a carpeta Data (puede ser el root mismo o una subcarpeta)
    data_id = _find_folder(svc, root_id, "Data") or root_id
    print(f"  Drive: carpeta Data encontrada (id={data_id[:8]}...)")

    # Inventario (stock Excel)
    print("  → Inventario...")
    inv_id = _find_folder(svc, data_id, "Inventario")
    if inv_id:
        n = _download_folder(svc, inv_id, os.path.join(tmpdir, "Data", "Inventario"))
        print(f"     {n} archivos descargados")
    else:
        print("  ⚠ Carpeta Inventario no encontrada")

    # Salidas (ventas Excel + GC remitos)
    print("  → Salidas...")
    sal_id = _find_folder(svc, data_id, "Salidas")
    if sal_id:
        _download_named(svc, sal_id, "Bosque salidas.xlsx",
                        os.path.join(tmpdir, "Data", "Salidas", "Bosque salidas.xlsx"))
        _download_named(svc, sal_id, "Salidas_consolidado.xlsx",
                        os.path.join(tmpdir, "Data", "Salidas", "Salidas_consolidado.xlsx"))
        gc_id = _find_folder(svc, sal_id, "GC")
        if gc_id:
            print("  → GC (remitos)...")
            n = _download_folder(svc, gc_id,
                                 os.path.join(tmpdir, "Data", "Salidas", "GC"),
                                 name_contains="Remitos GC")
            print(f"     {n} remitos")

    # Costos CSV (backup del Sheet)
    print("  → Costos...")
    cos_id = _find_folder(svc, data_id, "Costos y PVP")
    if cos_id:
        _download_folder(svc, cos_id, os.path.join(tmpdir, "Data", "Costos y PVP"))

    # Supply Chain / proyecciones
    sc_id = _find_folder(svc, data_id, "Supply Chain")
    if sc_id:
        proy_id = _find_folder(svc, sc_id, "proyecciones")
        if proy_id:
            print("  → Proyecciones...")
            _download_folder(svc, proy_id,
                             os.path.join(tmpdir, "Data", "Supply Chain", "proyecciones"))


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
        "message": f"data: actualizar cloud {datetime.now().strftime('%Y-%m-%d %H:%M')}",
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
                  "Data/Costos y PVP", "Data/Insumos",
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

        # Redirigir todas las rutas al directorio temporal
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
        mod.OUT_JS         = os.path.join(tmpdir, "bosquegin_data.js")

        # En cloud no hay browser → deshabilitar CDP (costos usa CSV de Drive)
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
