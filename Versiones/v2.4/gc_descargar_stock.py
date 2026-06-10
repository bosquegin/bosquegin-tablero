#!/usr/bin/env python3
"""
gc_descargar_stock.py
Descarga automáticamente el Excel de productos desde Gestión Cervecera
usando las cookies de sesión ya activas en Chrome.

Requiere: cryptography  (pip install cryptography)
"""
import os, json, base64, shutil, sqlite3, tempfile, ctypes, ctypes.wintypes
import urllib.request, urllib.parse
from datetime import datetime
import platform

# ── Rutas ─────────────────────────────────────────────────────────────────────
if platform.system() == "Windows":
    TABLERO = r"C:\Users\SupplyDestileria\Documents\Bosque\Claude\Tablero operativo"
    # Brave browser (Chromium-based, same format as Chrome)
    _LOCAL = os.path.expandvars("%LOCALAPPDATA%")
    CHROME_LOCAL = None
    for _candidate in [
        os.path.join(_LOCAL, "BraveSoftware", "Brave-Browser", "User Data"),
        os.path.join(_LOCAL, "Google",         "Chrome",        "User Data"),
        os.path.join(_LOCAL, "Microsoft",      "Edge",          "User Data"),
        os.path.join(_LOCAL, "Chromium",                        "User Data"),
    ]:
        if os.path.isdir(_candidate):
            CHROME_LOCAL = _candidate
            break
    if not CHROME_LOCAL:
        raise SystemExit("No se encontró carpeta de datos de Brave/Chrome/Edge.")
else:
    raise SystemExit("Este script solo corre en Windows (necesita Brave/Chrome + DPAPI)")

INV_DIR  = os.path.join(TABLERO, "Data", "Inventario")
GC_URL   = "https://www.gestioncervecera.com"

# ── DPAPI decrypt (Windows) ────────────────────────────────────────────────────
class _BLOB(ctypes.Structure):
    _fields_ = [("cbData", ctypes.wintypes.DWORD),
                ("pbData", ctypes.POINTER(ctypes.c_char))]

def _dpapi_decrypt(ciphertext: bytes) -> bytes:
    buf    = ctypes.create_string_buffer(ciphertext, len(ciphertext))
    inp    = _BLOB(ctypes.sizeof(buf), buf)
    out    = _BLOB()
    ok     = ctypes.windll.crypt32.CryptUnprotectData(
                 ctypes.byref(inp), None, None, None, None, 0, ctypes.byref(out))
    if not ok:
        raise ctypes.WinError()
    result = ctypes.string_at(out.pbData, out.cbData)
    ctypes.windll.kernel32.LocalFree(out.pbData)
    return result

# ── Chrome AES key ─────────────────────────────────────────────────────────────
def _get_chrome_aes_key() -> bytes:
    local_state_path = os.path.join(CHROME_LOCAL, "Local State")
    with open(local_state_path, "r", encoding="utf-8") as f:
        ls = json.load(f)
    enc_key = base64.b64decode(ls["os_crypt"]["encrypted_key"])
    # First 5 bytes are the literal prefix b'DPAPI'
    return _dpapi_decrypt(enc_key[5:])

# ── Decrypt single cookie value ────────────────────────────────────────────────
def _decrypt_cookie(aes_key: bytes, enc_val: bytes) -> str:
    if enc_val[:3] == b"v10" or enc_val[:3] == b"v11":
        from cryptography.hazmat.primitives.ciphers.aead import AESGCM
        nonce      = enc_val[3:15]
        ciphertext = enc_val[15:]
        return AESGCM(aes_key).decrypt(nonce, ciphertext, None).decode("utf-8")
    # Old-style DPAPI cookie (rare in modern Chrome)
    return _dpapi_decrypt(enc_val).decode("utf-8")

# ── Read gestioncervecera cookies from Chrome ──────────────────────────────────
def get_gc_cookies() -> dict:
    profile = os.path.join(CHROME_LOCAL, "Default")
    # Try Default profile first, then all profiles
    profiles = [profile]
    for entry in os.listdir(CHROME_LOCAL):
        if entry.startswith("Profile "):
            profiles.append(os.path.join(CHROME_LOCAL, entry))

    aes_key = _get_chrome_aes_key()
    cookies = {}

    for prof in profiles:
        db_path = os.path.join(prof, "Network", "Cookies")
        if not os.path.exists(db_path):
            db_path = os.path.join(prof, "Cookies")
        if not os.path.exists(db_path):
            continue
        # Copy because Chrome holds a lock
        tmp = tempfile.mktemp(suffix=".db")
        try:
            shutil.copy2(db_path, tmp)
            conn = sqlite3.connect(tmp)
            cur  = conn.execute(
                "SELECT name, encrypted_value FROM cookies "
                "WHERE host_key LIKE '%gestioncervecera%'")
            for name, enc in cur.fetchall():
                try:
                    cookies[name] = _decrypt_cookie(aes_key, bytes(enc))
                except Exception:
                    pass
            conn.close()
        except Exception as e:
            print(f"  [warn] profile {prof}: {e}")
        finally:
            try: os.unlink(tmp)
            except: pass
        if cookies:
            break

    return cookies

# ── HTTP helpers ───────────────────────────────────────────────────────────────
def _cookie_header(cookies: dict) -> str:
    return "; ".join(f"{k}={v}" for k, v in cookies.items())

def _post_json(url: str, data: dict, headers: dict) -> dict:
    body = urllib.parse.urlencode(data).encode()
    req  = urllib.request.Request(url, data=body, headers=headers, method="POST")
    with urllib.request.urlopen(req, timeout=30) as resp:
        return json.loads(resp.read().decode())

def _get_bytes(url: str, headers: dict) -> bytes:
    req = urllib.request.Request(url, headers=headers, method="GET")
    with urllib.request.urlopen(req, timeout=60) as resp:
        return resp.read()

# ── Main download ──────────────────────────────────────────────────────────────
def descargar_stock() -> str:
    print("Leyendo cookies de Chrome...")
    cookies = get_gc_cookies()
    if not cookies:
        raise SystemExit("No se encontraron cookies de gestioncervecera.com en Chrome. "
                         "Asegurate de estar logueado en el browser.")
    print(f"  {len(cookies)} cookies encontradas: {list(cookies.keys())}")

    hdrs = {
        "Cookie":           _cookie_header(cookies),
        "X-Requested-With": "XMLHttpRequest",
        "Content-Type":     "application/x-www-form-urlencoded",
        "User-Agent":       "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36",
        "Referer":          GC_URL + "/Producto",
        "Origin":           GC_URL,
    }

    # Step 1 – get FileGuid
    print("Solicitando exportación...")
    result = _post_json(GC_URL + "/Producto/ExportarProductosExcel",
                        {"ord": "", "dir": ""}, hdrs)
    if result.get("message") != "ok":
        raise SystemExit(f"Error al exportar: {result}")
    guid  = result["FileGuid"]
    fname = result["FileName"]
    print(f"  FileGuid: {guid}")
    print(f"  FileName: {fname}")

    # Step 2 – download file
    print("Descargando archivo...")
    hdrs_get = dict(hdrs)
    hdrs_get.pop("Content-Type", None)
    hdrs_get.pop("X-Requested-With", None)
    params   = urllib.parse.urlencode({"fileGuid": guid, "filename": fname})
    file_url = f"{GC_URL}/Producto/ArchivoExcel?{params}"
    data     = _get_bytes(file_url, hdrs_get)
    print(f"  Recibido: {len(data):,} bytes")

    # Validate it's a real xlsx (PK header)
    if data[:2] != b"PK":
        raise SystemExit(f"Respuesta inesperada (no es xlsx): {data[:200]}")

    # Save with today's date stamp (same naming as manual export)
    today    = datetime.now().strftime("%d-%m-%Y")
    out_name = f"gc_productos_{today}.xlsx"
    out_path = os.path.join(INV_DIR, out_name)
    os.makedirs(INV_DIR, exist_ok=True)
    with open(out_path, "wb") as f:
        f.write(data)
    print(f"  Guardado: {out_path}")
    return out_path

if __name__ == "__main__":
    path = descargar_stock()
    print(f"\n✓ Listo: {path}")
