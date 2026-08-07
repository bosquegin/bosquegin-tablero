#!/usr/bin/env python3
"""
actualizar_cloud.py — v2.1  Versión cloud sin dependencias externas de Google

Descarga archivos de Google Drive a un directorio temporal,
ejecuta la lógica de actualizar_bosquegin.py y sube los data_*.js
(los que realmente lee el tablero) a GitHub via API, uno por uno
(sin necesidad de tener la PC encendida ni de un working copy git).

Variables de entorno requeridas:
  GOOGLE_OAUTH_CLIENT_ID       — OAuth2 client ID
  GOOGLE_OAUTH_CLIENT_SECRET   — OAuth2 client secret
  GOOGLE_OAUTH_REFRESH_TOKEN   — OAuth2 refresh token (obtenido con get_oauth_token.py)
  DRIVE_ROOT_FOLDER_ID         — ID de la carpeta raíz en Drive
                                  (la que contiene la carpeta "Data")
  GITHUB_TOKEN                 — Token de acceso personal a GitHub
"""
import os, sys, json, shutil, tempfile, importlib.util, urllib.request, urllib.parse, urllib.error, base64
from datetime import datetime, timedelta, timezone

# Evita que un print con tildes/emojis crashee la corrida si la consola no
# está en UTF-8 (mismo fix que actualizar_bosquegin.py).
try:
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
    sys.stderr.reconfigure(encoding="utf-8", errors="replace")
except Exception:
    pass

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


def download_cervezas_cache_from_github(tmpdir, github_token):
    """
    Descarga Data/Costos y PVP/cervezas_meses/*.csv desde GitHub al tmpdir cloud.

    fetch_cervezas() arma las hojas mensuales históricas leyendo estos CSV de
    caché (o, si están "stale", intentando refrescarlos vía CDP/Brave local).
    En la nube no hay CDP, así que un refresco en vivo siempre falla — y como
    download_from_drive() no baja esta subcarpeta (Drive no la espeja, y de
    todos modos _download_folder no recorre subcarpetas), el tmpdir arrancaba
    con cervezas_meses/ vacía. Eso hacía fallar TODOS los meses históricos
    (solo quedaba el mes actual, vía el fallback CERV_CSV aparte) y esa
    corrida vacía se subía a GitHub pisando el historial bueno (bug real
    detectado 2026-08-06). La nube nunca escribe caché nuevo acá (el fetch en
    vivo siempre falla sin CDP), así que alcanza con traer lo ya commiteado.
    """
    folder_path = "Data/Costos y PVP/cervezas_meses"
    encoded = "/".join(urllib.parse.quote(seg) for seg in folder_path.split("/"))
    api_dir = f"https://api.github.com/repos/{GITHUB_REPO}/contents/{encoded}"
    hdrs = {"Authorization": f"token {github_token}", "Accept": "application/vnd.github.v3+json"}
    req = urllib.request.Request(api_dir, headers=hdrs)
    try:
        with urllib.request.urlopen(req, timeout=30) as r:
            entries = json.loads(r.read())
    except urllib.error.HTTPError as e:
        print(f"  ⚠ No se pudo listar caché de cervezas en GitHub: {e}")
        return
    dest_dir = os.path.join(tmpdir, "Data", "Costos y PVP", "cervezas_meses")
    os.makedirs(dest_dir, exist_ok=True)
    n = 0
    for entry in entries:
        name = entry.get("name", "")
        if not name.endswith(".csv"):
            continue
        raw_req = urllib.request.Request(
            entry["url"],
            headers={"Authorization": f"token {github_token}", "Accept": "application/vnd.github.v3.raw"})
        try:
            with urllib.request.urlopen(raw_req, timeout=30) as r:
                content = r.read()
        except Exception as e:
            print(f"  ⚠ {name}: {e}")
            continue
        with open(os.path.join(dest_dir, name), "wb") as f:
            f.write(content)
        n += 1
    print(f"  ↓ {n} cachés de cervezas descargados de GitHub")


# ═══════════════════════════════════════════════════════════════════════════════
#  GITHUB API
# ═══════════════════════════════════════════════════════════════════════════════

def push_file_to_github(repo_path, content, token, message):
    """Sube (crea o actualiza) un archivo puntual a GitHub via Contents API."""
    api_url = f"https://api.github.com/repos/{GITHUB_REPO}/contents/{repo_path}"
    hdrs = {
        "Authorization": f"token {token}",
        "Accept": "application/vnd.github.v3+json",
        "Content-Type": "application/json",
    }

    # Obtener SHA actual (si el archivo ya existe)
    sha = None
    req = urllib.request.Request(api_url)
    for k, v in hdrs.items():
        req.add_header(k, v)
    try:
        with urllib.request.urlopen(req, timeout=30) as r:
            sha = json.loads(r.read())["sha"]
    except urllib.error.HTTPError as e:
        if e.code != 404:
            raise

    body_dict = {
        "message": message,
        "content": base64.b64encode(content.encode("utf-8")).decode("ascii"),
    }
    if sha:
        body_dict["sha"] = sha
    body = json.dumps(body_dict).encode("utf-8")
    req2 = urllib.request.Request(api_url, data=body, method="PUT")
    for k, v in hdrs.items():
        req2.add_header(k, v)
    with urllib.request.urlopen(req2, timeout=30) as r:
        result = json.loads(r.read())

    commit_url = result.get("commit", {}).get("html_url", "")
    print(f"  ✓ {repo_path} publicado ({commit_url})")


# Archivos que el tablero realmente lee (bosquegin_data.js es un monolito
# viejo que el dashboard ya no carga — ver actualizar_bosquegin.py). Cada uno
# se sube por separado via Contents API porque el flujo cloud corre en un
# tmpdir sin git (no puede usar el git add/commit/push normal del flujo local).
SECTION_FILES = [
    "data_meta.js", "data_ventas.js", "data_clientes.js", "data_stock.js",
    "data_stock_cierre.js",
    "data_destileria.js", "data_costos.js", "data_insumos.js", "data_proyeccion.js",
]


def push_sections_to_github(tmpdir, token):
    """Sube los data_*.js reales (los que usa el tablero) a GitHub, uno por uno."""
    msg = f"data: actualizar cloud {datetime.now(_AR).strftime('%Y-%m-%d %H:%M')}"
    for fname in SECTION_FILES:
        fpath = os.path.join(tmpdir, fname)
        if not os.path.exists(fpath):
            print(f"  ⚠ {fname} no se generó — omitiendo")
            continue
        with open(fpath, encoding="utf-8") as f:
            content = f.read()
        push_file_to_github(fname, content, token, msg)


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
        download_cervezas_cache_from_github(tmpdir, github_token)

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

        # ── [3] Subir los data_*.js reales a GitHub (uno por uno via API) ─────
        # bosquegin_data.js NO se sube: es el monolito viejo que el dashboard
        # ya no lee (ver SECTION_FILES / actualizar_bosquegin.py).
        print("\n[3/4] Publicando en GitHub via API...")
        push_sections_to_github(tmpdir, github_token)

        print("\n[4/4] ✅ Actualización cloud completada")
        print("      El tablero se actualiza en ~60 segundos")

    finally:
        shutil.rmtree(tmpdir, ignore_errors=True)


if __name__ == "__main__":
    main()
