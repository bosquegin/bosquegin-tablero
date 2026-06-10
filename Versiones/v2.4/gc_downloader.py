#!/usr/bin/env python3
"""
gc_downloader.py
Descarga desde Gestión Cervecera usando las cookies de sesión guardadas en
Chrome / Brave / Edge — sin abrir el browser.

Descarga:
  - Stock de Productos   (POST /Producto/StockExcel         → Stock productos D-M-YYYY.xlsx)
  - Stock Complementarios(POST /Producto/StockCompExcel     → Stock Productos Comp. D-M-YYYY.xlsx)
  - Remitos Detallados   (POST /Informes/InformeRemitosDetallados)

Requiere: cryptography
"""
import os, json, base64, shutil, sqlite3, tempfile, ctypes, ctypes.wintypes
import urllib.request, urllib.parse, urllib.error, re, platform
from datetime import datetime, date

if platform.system() != "Windows":
    raise SystemExit("gc_downloader solo corre en Windows (necesita DPAPI)")

BASE         = r"C:\Users\SupplyDestileria\Documents\Bosque\Claude\Tablero operativo"
INV_DIR      = os.path.join(BASE, "Data", "Inventario")
GC_SAL_DIR   = os.path.join(BASE, "Data", "Salidas", "GC")
GC_URL       = "https://www.gestioncervecera.com"
COOKIE_CACHE = os.path.join(BASE, "Data", "gc_session_cache.json")

# ── Copiar archivo bloqueado (para Cookies SQLite mientras el browser corre) ──
def _copy_locked(src: str, dst: str):
    GENERIC_READ          = 0x80000000
    FILE_SHARE_ALL        = 0x00000007
    OPEN_EXISTING         = 3
    FILE_ATTRIBUTE_NORMAL = 0x80
    INVALID_HANDLE_VALUE  = ctypes.c_size_t(-1).value

    k32 = ctypes.windll.kernel32
    k32.CreateFileW.restype = ctypes.c_size_t
    h = k32.CreateFileW(
        src, GENERIC_READ, FILE_SHARE_ALL,
        None, OPEN_EXISTING, FILE_ATTRIBUTE_NORMAL, None)
    if h == INVALID_HANDLE_VALUE:
        raise ctypes.WinError()
    try:
        buf = ctypes.create_string_buffer(65536)
        with open(dst, 'wb') as out:
            while True:
                read = ctypes.c_ulong(0)
                ok = k32.ReadFile(h, buf, len(buf), ctypes.byref(read), None)
                if not ok or read.value == 0:
                    break
                out.write(buf.raw[:read.value])
    finally:
        k32.CloseHandle(h)
    for ext in ("-wal", "-shm"):
        if os.path.exists(src + ext):
            try:
                shutil.copy2(src + ext, dst + ext)
            except Exception:
                pass


# ── Localizar perfil Chrome / Brave / Edge ────────────────────────────────────
_LOCAL = os.path.expandvars("%LOCALAPPDATA%")
CHROME_LOCAL = None
for _c in [
    os.path.join(_LOCAL, "BraveSoftware", "Brave-Browser", "User Data"),
    os.path.join(_LOCAL, "Google",        "Chrome",        "User Data"),
    os.path.join(_LOCAL, "Microsoft",     "Edge",          "User Data"),
    os.path.join(_LOCAL, "Chromium",                       "User Data"),
]:
    if os.path.isdir(_c):
        CHROME_LOCAL = _c
        break

# ── DPAPI / AES-GCM cookie decrypt ───────────────────────────────────────────
class _BLOB(ctypes.Structure):
    _fields_ = [("cbData", ctypes.wintypes.DWORD),
                ("pbData", ctypes.POINTER(ctypes.c_char))]

def _dpapi_decrypt(ct: bytes) -> bytes:
    buf = ctypes.create_string_buffer(ct, len(ct))
    inp = _BLOB(ctypes.sizeof(buf), buf)
    out = _BLOB()
    if not ctypes.windll.crypt32.CryptUnprotectData(
            ctypes.byref(inp), None, None, None, None, 0, ctypes.byref(out)):
        raise ctypes.WinError()
    r = ctypes.string_at(out.pbData, out.cbData)
    ctypes.windll.kernel32.LocalFree(out.pbData)
    return r

def _aes_key() -> bytes:
    ls_path = os.path.join(CHROME_LOCAL, "Local State")
    with open(ls_path, encoding="utf-8") as f:
        ls = json.load(f)
    return _dpapi_decrypt(base64.b64decode(ls["os_crypt"]["encrypted_key"])[5:])

def _decrypt_cookie(key: bytes, enc: bytes) -> str:
    if enc[:3] in (b"v10", b"v11"):
        from cryptography.hazmat.primitives.ciphers.aead import AESGCM
        return AESGCM(key).decrypt(enc[3:15], enc[15:], None).decode()
    return _dpapi_decrypt(enc).decode()

def _save_cookie_cache(cookies: dict):
    try:
        os.makedirs(os.path.dirname(COOKIE_CACHE), exist_ok=True)
        payload = {"ts": datetime.now().isoformat(), "cookies": cookies}
        with open(COOKIE_CACHE, "w", encoding="utf-8") as f:
            json.dump(payload, f)
    except Exception:
        pass

def _load_cookie_cache() -> dict:
    try:
        with open(COOKIE_CACHE, encoding="utf-8") as f:
            payload = json.load(f)
        ts = datetime.fromisoformat(payload["ts"])
        if (datetime.now() - ts).total_seconds() < 24 * 3600:
            return payload.get("cookies", {})
    except Exception:
        pass
    return {}

# ── CDP (Chrome DevTools Protocol) ────────────────────────────────────────────
def _cdp_get_gc_cookies(port: int = 9222) -> dict:
    """Lee cookies de GC via CDP. Requiere Brave con --remote-debugging-port=9222."""
    import socket as _sock, struct as _struct

    try:
        req = urllib.request.Request(f"http://localhost:{port}/json/list")
        with urllib.request.urlopen(req, timeout=2) as r:
            targets = json.loads(r.read())
    except Exception:
        return {}

    ws_url = next((t.get("webSocketDebuggerUrl") for t in targets
                   if t.get("type") == "page" and t.get("webSocketDebuggerUrl")), None)
    if not ws_url and targets:
        ws_url = targets[0].get("webSocketDebuggerUrl")
    if not ws_url:
        return {}

    def _ws_send(s, text):
        payload = text.encode()
        n = len(payload)
        mask = os.urandom(4)
        masked = bytes(b ^ mask[i % 4] for i, b in enumerate(payload))
        if n < 126:
            hdr = _struct.pack('BB', 0x81, 0x80 | n)
        elif n < 65536:
            hdr = _struct.pack('>BBH', 0x81, 0xFE, n)
        else:
            hdr = _struct.pack('>BBQ', 0x81, 0xFF, n)
        s.sendall(hdr + mask + masked)

    def _ws_recv(s):
        def _read(n):
            d = b""
            while len(d) < n:
                chunk = s.recv(n - len(d))
                if not chunk: raise ConnectionError("socket cerrado")
                d += chunk
            return d
        b1, b2 = _read(2)
        is_masked = (b2 & 0x80) != 0
        n = b2 & 0x7F
        if n == 126:   n = _struct.unpack('>H', _read(2))[0]
        elif n == 127: n = _struct.unpack('>Q', _read(8))[0]
        mask = _read(4) if is_masked else b""
        payload = _read(n)
        if is_masked:
            payload = bytes(b ^ mask[i % 4] for i, b in enumerate(payload))
        return payload.decode('utf-8', errors='replace')

    try:
        m = re.match(r'ws://([^/:]+)(?::(\d+))?(/.+)', ws_url)
        if not m: return {}
        host, ws_port_str, path = m.group(1), m.group(2), m.group(3)
        ws_port = int(ws_port_str) if ws_port_str else port

        s = _sock.create_connection((host, ws_port), timeout=5)
        key = base64.b64encode(os.urandom(16)).decode()
        s.sendall((
            f"GET {path} HTTP/1.1\r\n"
            f"Host: {host}:{ws_port}\r\n"
            f"Upgrade: websocket\r\nConnection: Upgrade\r\n"
            f"Sec-WebSocket-Key: {key}\r\nSec-WebSocket-Version: 13\r\n\r\n"
        ).encode())
        resp = b""
        while b"\r\n\r\n" not in resp:
            resp += s.recv(4096)
        if b"101" not in resp[:20]:
            s.close(); return {}

        _ws_send(s, json.dumps({
            "id": 1, "method": "Network.getCookies",
            "params": {"urls": ["https://www.gestioncervecera.com"]}
        }))

        cookies = {}
        for _ in range(20):
            try:
                result = json.loads(_ws_recv(s))
            except Exception:
                break
            if result.get("id") == 1:
                for c in result.get("result", {}).get("cookies", []):
                    cookies[c["name"]] = c["value"]
                break
        s.close()
        return cookies
    except Exception:
        return {}


BRAVE_EXE = r"C:\Program Files\BraveSoftware\Brave-Browser\Application\brave.exe"

def _brave_running() -> bool:
    import subprocess
    r = subprocess.run(
        ["tasklist", "/FI", "IMAGENAME eq brave.exe", "/NH"],
        capture_output=True, text=True,
        creationflags=subprocess.CREATE_NO_WINDOW
    )
    return "brave.exe" in r.stdout.lower()

def _restart_brave_with_cdp() -> bool:
    """Cierra Brave y lo vuelve a abrir con --remote-debugging-port=9222.
    Brave tiene restauración de sesión, los tabs vuelven solos."""
    import subprocess, time
    if not _brave_running():
        return False
    print("  [GC] Reiniciando Brave con CDP...")
    subprocess.run(
        ["taskkill", "/F", "/IM", "brave.exe"],
        capture_output=True,
        creationflags=subprocess.CREATE_NO_WINDOW
    )
    time.sleep(2)
    subprocess.Popen(
        [BRAVE_EXE, "--remote-debugging-port=9222"],
        creationflags=subprocess.CREATE_NO_WINDOW
    )
    for _ in range(15):          # esperar hasta 15 seg a que CDP esté listo
        time.sleep(1)
        cookies = _cdp_get_gc_cookies()
        if cookies:
            print("  [GC] Cookies leídas via CDP (tras reinicio)")
            _save_cookie_cache(cookies)
            return True
    return False

def _get_gc_cookies() -> dict:
    # 1. CDP: funciona con el browser abierto si tiene --remote-debugging-port=9222
    cookies = _cdp_get_gc_cookies()
    if cookies:
        print("  [GC] Cookies leídas via CDP")
        _save_cookie_cache(cookies)
        return cookies

    # 2. SQLite: solo funciona con el browser cerrado
    cookies = {}
    if CHROME_LOCAL:
        try:
            key      = _aes_key()
            profiles = [os.path.join(CHROME_LOCAL, "Default")]
            for e in os.listdir(CHROME_LOCAL):
                if e.startswith("Profile "):
                    profiles.append(os.path.join(CHROME_LOCAL, e))
            for prof in profiles:
                db = os.path.join(prof, "Network", "Cookies")
                if not os.path.exists(db):
                    db = os.path.join(prof, "Cookies")
                if not os.path.exists(db):
                    continue
                tmp = tempfile.mktemp(suffix=".db")
                try:
                    try:
                        _copy_locked(db, tmp)
                        conn = sqlite3.connect(tmp)
                    except (OSError, sqlite3.OperationalError):
                        db_uri = "///" + db.replace("\\", "/").replace(" ", "%20")
                        conn = sqlite3.connect(f"file:{db_uri}?mode=ro&immutable=1", uri=True)
                        tmp = None
                    for name, enc in conn.execute(
                            "SELECT name, encrypted_value FROM cookies "
                            "WHERE host_key LIKE '%gestioncervecera%'").fetchall():
                        try:
                            cookies[name] = _decrypt_cookie(key, bytes(enc))
                        except Exception:
                            pass
                    conn.close()
                except Exception as e:
                    print(f"  [warn] {os.path.basename(prof)}: {e}")
                finally:
                    if tmp:
                        for ext in ("", "-wal", "-shm"):
                            try: os.unlink(tmp + ext)
                            except: pass
                if cookies:
                    break
        except Exception as e:
            print(f"  [warn] Error leyendo cookies del browser: {e}")

    if cookies:
        _save_cookie_cache(cookies)
        return cookies

    # 3. Caché: válido 24h desde la última lectura exitosa
    cached = _load_cookie_cache()
    if cached:
        print("  [GC] Usando caché de cookies")
        return cached

    # 4. Último recurso: reiniciar Brave con CDP y reintentar
    if _brave_running():
        cookies = {}
        if _restart_brave_with_cdp():
            cookies = _cdp_get_gc_cookies()
            if cookies:
                _save_cookie_cache(cookies)
                return cookies

    return {}

# ── HTTP helpers ──────────────────────────────────────────────────────────────
_UA = ("Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
       "AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0 Safari/537.36")

def _base_headers(cookie_str: str, referer: str = "") -> dict:
    h = {
        "Cookie":     cookie_str,
        "User-Agent": _UA,
        "Referer":    referer or GC_URL + "/",
        "Origin":     GC_URL,
    }
    return h

def _post_json(url: str, data, hdrs: dict) -> dict:
    body = (urllib.parse.urlencode(data).encode() if isinstance(data, dict)
            else data.encode() if data else b"")
    req  = urllib.request.Request(
               url, data=body, headers={
                   **hdrs,
                   "Content-Type":     "application/x-www-form-urlencoded",
                   "X-Requested-With": "XMLHttpRequest",
               }, method="POST")
    try:
        with urllib.request.urlopen(req, timeout=60) as r:
            raw = r.read()
            try:
                return json.loads(raw.decode())
            except json.JSONDecodeError:
                raise RuntimeError(f"Respuesta no es JSON (status {r.status}): {raw[:200]!r}")
    except urllib.error.HTTPError as e:
        body_err = e.read()
        raise RuntimeError(f"HTTP {e.code} en {url}: {body_err[:200]!r}")

def _get_bytes(url: str, hdrs: dict) -> bytes:
    req = urllib.request.Request(url, headers=hdrs, method="GET")
    with urllib.request.urlopen(req, timeout=120) as r:
        return r.read()

def _download_excel(guid: str, fname: str, hdrs: dict) -> bytes:
    params = urllib.parse.urlencode({"fileGuid": guid, "filename": fname})
    data   = _get_bytes(f"{GC_URL}/Producto/ArchivoExcel?{params}", hdrs)
    if not data:
        raise RuntimeError("GC devolvió 0 bytes")
    if data[:2] != b"PK":
        raise RuntimeError(f"Respuesta no es xlsx: {data[:80]!r}")
    return data

def _get_csrf(cookie_str: str) -> str:
    hdrs = _base_headers(cookie_str)
    for url in [
        GC_URL + "/Informes/Ver?informe=RemitosDetallados",
        GC_URL + "/Informes/InformeRemitosDetallados",
    ]:
        try:
            html = _get_bytes(url, hdrs).decode("utf-8", errors="replace")
            m = re.search(r'name="__RequestVerificationToken"[^>]*value="([^"]+)"', html)
            if not m:
                m = re.search(r'value="([^"]+)"[^>]*name="__RequestVerificationToken"', html)
            if m:
                return m.group(1)
        except Exception:
            pass
    return ""


# ── Stock de Productos via API ────────────────────────────────────────────────
def _download_stock(cookie_str: str, endpoint: str, page_url: str,
                    out_path: str, verbose: bool = True) -> str:
    """
    Descarga un Excel de stock via POST endpoint → GET ArchivoExcel.
    Visita la página primero para inicializar el contexto de sesión en GC.
    """
    hdrs = _base_headers(cookie_str, referer=page_url)
    # Visitar la página para que GC inicialice el contexto del export
    _get_bytes(page_url, hdrs)
    j = _post_json(GC_URL + endpoint, {}, hdrs)
    if j.get("message") != "ok":
        raise RuntimeError(f"GC {endpoint}: {j}")
    data = _download_excel(j["FileGuid"], j["FileName"], hdrs)
    with open(out_path, "wb") as f:
        f.write(data)
    if verbose:
        print(f"  [GC] OK {os.path.basename(out_path)}  ({len(data):,} bytes)")
    return out_path


# ── Remitos ───────────────────────────────────────────────────────────────────
def _fetch_remitos(cookie_str: str, fecha_desde_date=None, verbose: bool = True) -> str | None:
    """Descarga Remitos Detallados desde fecha_desde_date hasta hoy.
    Si fecha_desde_date es None descarga desde el 1° del mes actual."""
    today = date.today()
    if fecha_desde_date is None:
        desde = date(today.year, today.month, 1)
    elif isinstance(fecha_desde_date, str):
        try:
            desde = date.fromisoformat(fecha_desde_date)
        except Exception:
            desde = date(today.year, today.month, 1)
    else:
        desde = fecha_desde_date
    fecha_desde = desde.strftime("%d/%m/%Y")
    fecha_hasta = today.strftime("%d/%m/%Y")
    hdrs        = _base_headers(cookie_str)

    if verbose: print("  [GC] Obteniendo token CSRF para Remitos...")
    csrf = _get_csrf(cookie_str)
    if verbose and csrf:
        print(f"  [GC] CSRF obtenido ({len(csrf)} chars)")
    elif verbose:
        print("  [GC] Sin CSRF — intentando sin token")

    if verbose: print(f"  [GC] Descargando Remitos {fecha_desde} - {fecha_hasta}...")

    form_data = {
        "filtro":       "",
        "tieneGrafico": "false",
        "tieneRanking": "false",
        "id":           "",
        "entidad":      "",
        "fechaDesde":   fecha_desde,
        "fechaHasta":   fecha_hasta,
    }
    if csrf:
        form_data["__RequestVerificationToken"] = csrf

    j = _post_json(GC_URL + "/Informes/InformeRemitosDetallados", form_data, hdrs)

    if j.get("message") != "ok":
        raise RuntimeError(f"Respuesta GC Remitos: {j}")

    data     = _download_excel(j["FileGuid"], j["FileName"], hdrs)
    rem_name = f"Remitos GC {fecha_desde.replace('/', '-')} {fecha_hasta.replace('/', '-')}.xlsx"
    out      = os.path.join(GC_SAL_DIR, rem_name)
    os.makedirs(GC_SAL_DIR, exist_ok=True)
    # Eliminar TODOS los remitos anteriores — solo debe existir uno a la vez
    for old in os.listdir(GC_SAL_DIR):
        if old.startswith("Remitos GC") and old.endswith(".xlsx") and old != rem_name:
            try:
                os.remove(os.path.join(GC_SAL_DIR, old))
                if verbose: print(f"  [GC] Eliminado remito anterior: {old}")
            except Exception:
                pass
    with open(out, "wb") as f:
        f.write(data)
    if verbose: print(f"  [GC] OK Remitos -> {rem_name}  ({len(data):,} bytes)")
    return out


# ── Descarga principal ────────────────────────────────────────────────────────
def descargar_todo(verbose: bool = True, fecha_desde_date=None) -> dict:
    """
    Descarga via API:
      - Stock de Productos     → Inventario/Stock productos D-M-YYYY.xlsx
      - Stock Complementarios  → Inventario/Stock Productos Comp. D-M-YYYY.xlsx
      - Remitos del mes        → Salidas/GC/Remitos GC ...xlsx
    """
    def log(msg):
        if verbose: print(msg)

    os.makedirs(INV_DIR,    exist_ok=True)
    os.makedirs(GC_SAL_DIR, exist_ok=True)

    today    = date.today()
    date_tag = f"{today.day}-{today.month}-{today.year}"

    log("  [GC] Leyendo cookies del browser...")
    cookies = _get_gc_cookies()
    if not cookies:
        raise RuntimeError(
            "No se pudieron leer las cookies de GC.\n"
            "  · Si nunca abriste GC en el browser: abrilo, iniciá sesión y hacé clic en Actualizar.\n"
            "  · Si el browser está abierto: cerralo y hacé clic en Actualizar (se guardará el caché).\n"
            "  · Luego de hacer eso una vez, el caché permitirá actualizar con el browser abierto.")
    cookie_str = "; ".join(f"{k}={v}" for k, v in cookies.items())
    log(f"  [GC] {len(cookies)} cookies encontradas")

    resultados = {}

    # ── 1. Stock de Productos ────────────────────────────────────────────────
    log("  [GC] Descargando Stock de Productos...")
    try:
        out = os.path.join(INV_DIR, f"Stock productos {date_tag}.xlsx")
        _download_stock(
            cookie_str,
            endpoint  = "/Producto/StockExcel",
            page_url  = GC_URL + "/Producto/Stock",
            out_path  = out,
            verbose   = verbose,
        )
        resultados["envases"] = out
    except Exception as e:
        log(f"  [GC] ERROR Stock: {e}")

    # ── 2. Stock Complementarios ─────────────────────────────────────────────
    log("  [GC] Descargando Stock Complementarios...")
    try:
        out = os.path.join(INV_DIR, f"Stock Productos Comp. {date_tag}.xlsx")
        _download_stock(
            cookie_str,
            endpoint  = "/Producto/StockCompExcel",
            page_url  = GC_URL + "/Producto/StockComp",
            out_path  = out,
            verbose   = verbose,
        )
        resultados["comp"] = out
    except Exception as e:
        log(f"  [GC] ERROR Comp: {e}")

    # ── 3. Remitos desde fecha_desde_date hasta hoy ──────────────────────────
    try:
        out = _fetch_remitos(cookie_str, fecha_desde_date=fecha_desde_date, verbose=verbose)
        if out:
            resultados["remitos"] = out
    except Exception as e:
        log(f"  [GC] ERROR Remitos: {e}")

    if not resultados:
        raise RuntimeError("No se pudo descargar ningún archivo de GC.")

    return resultados


if __name__ == "__main__":
    r = descargar_todo()
    print(f"\nListo: {len(r)} archivo(s) descargados")
    for k, v in r.items():
        print(f"  {k}: {v}")
