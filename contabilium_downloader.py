"""
contabilium_downloader.py
Descarga stock e historial de Contabilium usando CDP (Brave/Chrome con remote debugging).

Funciones principales:
  descargar_stock()      →  Data/Inventario/Contabilium/_tmp_Stock_16749_YYYYMMDD.xlsx
  descargar_historial()  →  Data/Salidas/Contabilium/_tmp_HistorialStock_16749_YYYYMMDD.xlsx
  descargar_todo()       →  ambas

Requiere Brave corriendo con --remote-debugging-port=9222 y sesión Contabilium activa.
"""
import os, json, re, time, glob, base64, struct, socket
import urllib.request
from datetime import date, timedelta

BASE                  = r"C:\Users\SupplyDestileria\Documents\Bosque\Claude\Tablero operativo"
CONTABILIUM_DIR       = os.path.join(BASE, "Data", "Salidas", "Contabilium")
CONTABILIUM_STOCK_DIR = os.path.join(BASE, "Data", "Inventario", "Contabilium")
CONTABILIUM_CUTOVER   = "2026-06-29"
CDP_PORT              = 9222

STOCK_URL    = "https://app.contabilium.com/modulos/reportes/stock.aspx"
HISTORIAL_URL = "https://app.contabilium.com/modulos/reportes/historicoStock.aspx"

# ── CDP mínimo ─────────────────────────────────────────────────────────────────

class _CDP:
    """Cliente CDP síncrono sobre WebSocket. Acumula eventos en cola."""

    def __init__(self, ws_url, timeout=90):
        self._timeout = timeout
        self._id = 0
        self._responses = {}   # msg_id -> mensaje completo
        self._events = []      # lista de (method, params)

        m = re.match(r'ws://([^/:]+)(?::(\d+))?(/.+)', ws_url)
        host, port_str, path = m.group(1), m.group(2), m.group(3)
        port = int(port_str) if port_str else CDP_PORT
        self._sock = socket.create_connection((host, port), timeout=timeout)

        key = base64.b64encode(os.urandom(16)).decode()
        self._sock.sendall((
            f"GET {path} HTTP/1.1\r\nHost: {host}:{port}\r\n"
            f"Upgrade: websocket\r\nConnection: Upgrade\r\n"
            f"Sec-WebSocket-Key: {key}\r\nSec-WebSocket-Version: 13\r\n\r\n"
        ).encode())
        resp = b""
        while b"\r\n\r\n" not in resp:
            resp += self._sock.recv(4096)
        if b"101" not in resp[:20]:
            raise RuntimeError("WebSocket handshake falló")

    def _ws_send(self, obj):
        text = json.dumps(obj).encode()
        n = len(text)
        mask = os.urandom(4)
        masked = bytes(b ^ mask[i % 4] for i, b in enumerate(text))
        hdr = (struct.pack('BB', 0x81, 0x80 | n) if n < 126
               else struct.pack('>BBH', 0x81, 0xFE, n) if n < 65536
               else struct.pack('>BBQ', 0x81, 0xFF, n))
        self._sock.sendall(hdr + mask + masked)

    def _ws_recv_one(self):
        def rd(n):
            d = b""
            while len(d) < n:
                chunk = self._sock.recv(n - len(d))
                if not chunk: raise ConnectionError("socket cerrado")
                d += chunk
            return d
        b1, b2 = rd(2)
        is_masked = (b2 & 0x80) != 0
        n = b2 & 0x7F
        if n == 126: n = struct.unpack('>H', rd(2))[0]
        elif n == 127: n = struct.unpack('>Q', rd(8))[0]
        mask = rd(4) if is_masked else b""
        payload = rd(n)
        if is_masked:
            payload = bytes(b ^ mask[i % 4] for i, b in enumerate(payload))
        return json.loads(payload.decode('utf-8', errors='replace'))

    def _drain(self, max_ms=50):
        """Lee mensajes disponibles sin bloquear mucho."""
        self._sock.settimeout(max_ms / 1000)
        for _ in range(200):
            try:
                msg = self._ws_recv_one()
                if "id" in msg:
                    self._responses[msg["id"]] = msg
                elif "method" in msg:
                    self._events.append((msg["method"], msg.get("params", {})))
            except socket.timeout:
                break

    def call(self, method, **params):
        self._id += 1
        msg_id = self._id
        self._ws_send({"id": msg_id, "method": method, "params": params})
        t0 = time.time()
        while time.time() - t0 < self._timeout:
            self._drain(max_ms=100)
            if msg_id in self._responses:
                resp = self._responses.pop(msg_id)
                if "error" in resp:
                    raise RuntimeError(f"CDP {method}: {resp['error']}")
                return resp.get("result", {})
            time.sleep(0.05)
        raise TimeoutError(f"CDP timeout esperando respuesta de {method}")

    def wait_event(self, event_name, timeout=None, filter_fn=None):
        """Espera un evento específico; retorna sus params."""
        t0 = time.time()
        tmax = timeout or self._timeout
        while time.time() - t0 < tmax:
            self._drain(max_ms=200)
            remaining = []
            found = None
            for name, params in self._events:
                if found is None and name == event_name and (filter_fn is None or filter_fn(params)):
                    found = params
                else:
                    remaining.append((name, params))
            self._events = remaining
            if found is not None:
                return found
            time.sleep(0.1)
        raise TimeoutError(f"CDP timeout esperando evento {event_name}")

    def close(self):
        try: self._sock.close()
        except: pass


# ── Helpers CDP ────────────────────────────────────────────────────────────────

def _cdp_port_open(port=CDP_PORT):
    try:
        urllib.request.urlopen(f"http://localhost:{port}/json/version", timeout=2)
        return True
    except Exception:
        return False


def _browser_ws_url(port=CDP_PORT):
    req = urllib.request.Request(f"http://localhost:{port}/json/version")
    with urllib.request.urlopen(req, timeout=3) as r:
        return json.loads(r.read()).get("webSocketDebuggerUrl")


def _new_tab_url(port=CDP_PORT):
    # Chrome/Brave recientes exigen PUT para /json/new (antes aceptaba GET)
    req = urllib.request.Request(
        f"http://localhost:{port}/json/new?about:blank", method="PUT")
    with urllib.request.urlopen(req, timeout=5) as r:
        info = json.loads(r.read())
    return info["webSocketDebuggerUrl"], info["id"]


def _close_tab(tab_id, port=CDP_PORT):
    try:
        urllib.request.urlopen(f"http://localhost:{port}/json/close/{tab_id}", timeout=3)
    except Exception:
        pass


# ── JS para clickear el botón de exportar Excel ───────────────────────────────

_EXPORT_JS = """
(function() {
    // Contabilium (ASP.NET) usa un flujo de 2 pasos:
    // 1) <a id="divIconoDescargar" href="javascript:exportar();">Exportar</a>
    //    dispara la generación del archivo en el servidor.
    // 2) <a id="lnkDownload" onclick="resetearExportacion();">Descargar</a>
    //    se habilita cuando el archivo está listo y dispara la descarga real.
    const btn = document.querySelector('#divIconoDescargar')
        || [...document.querySelectorAll('a[href], button, input[type="submit"], input[type="button"]')]
            .find(b => (b.textContent || b.value || '').trim().toLowerCase().includes('exportar'));
    if (!btn) {
        const all = [...document.querySelectorAll('a[href], button, input')];
        const ids = all.map(b => (b.id || b.value || '').slice(0, 20)).filter(Boolean);
        return 'NO ENCONTRADO boton Exportar. Elementos: ' + ids.join(', ');
    }
    btn.click();
    return 'click Exportar: ' + (btn.id || btn.textContent || '').trim().slice(0, 60);
})()
"""

_DOWNLOAD_JS = """
(function() {
    const btn = document.querySelector('#lnkDownload')
        || [...document.querySelectorAll('a[href], button, input[type="submit"], input[type="button"]')]
            .find(b => (b.textContent || b.value || '').trim().toLowerCase().includes('descargar'));
    if (!btn) {
        const all = [...document.querySelectorAll('a[href], button, input')];
        const ids = all.map(b => (b.id || b.value || '').slice(0, 20)).filter(Boolean);
        return 'NO ENCONTRADO boton Descargar. Elementos: ' + ids.join(', ');
    }
    btn.click();
    return 'click Descargar: ' + (btn.id || btn.textContent || '').trim().slice(0, 60);
})()
"""

def _ultima_fecha_historial():
    """
    Última fecha (columna Fecha) presente en los xlsx ya descargados de
    _tmp_HistorialStock_*.xlsx. El nombre del archivo es la fecha de
    DESCARGA, no la de los datos — hay que leer el contenido.
    """
    import openpyxl
    max_fecha = None
    for fpath in glob.glob(os.path.join(CONTABILIUM_DIR, "_tmp_HistorialStock_*.xlsx")):
        try:
            wb = openpyxl.load_workbook(fpath, read_only=True, data_only=True)
            ws = wb.active
            for row in ws.iter_rows(min_row=2, max_col=1, values_only=True):
                raw = row[0] if row else None
                if raw is None: continue
                d = raw.date() if hasattr(raw, "date") else None
                if d is None:
                    try:
                        d = date.fromisoformat(str(raw)[:10])
                    except Exception:
                        continue
                if max_fecha is None or d > max_fecha:
                    max_fecha = d
            wb.close()
        except Exception:
            continue
    return max_fecha


def _historial_setup_js(desde_date, hasta_date):
    """JS para setear el rango de fechas en historicoStock.aspx."""
    desde = desde_date.strftime("%d/%m/%Y")
    hasta = hasta_date.strftime("%d/%m/%Y")

    return f"""
(function() {{
    const desde = "{desde}";
    const hasta = "{hasta}";
    const inputs = [...document.querySelectorAll('input[type="text"], input[type="date"]')];
    const dateInputs = inputs.filter(i => {{
        const ctx = (i.id + ' ' + i.name + ' ' + (i.placeholder||'')).toLowerCase();
        return ctx.includes('fecha') || ctx.includes('date') || ctx.includes('desde') || ctx.includes('hasta');
    }});
    let res = [];
    if (dateInputs.length >= 2) {{
        dateInputs[0].value = desde;
        dateInputs[0].dispatchEvent(new Event('change', {{bubbles: true}}));
        dateInputs[1].value = hasta;
        dateInputs[1].dispatchEvent(new Event('change', {{bubbles: true}}));
        res = dateInputs.map(i => i.id + '=' + i.value);
    }} else {{
        // Fallback: primeros dos text inputs no hidden
        const vis = inputs.filter(i => i.type !== 'hidden').slice(0, 2);
        if (vis.length >= 2) {{
            vis[0].value = desde; vis[0].dispatchEvent(new Event('change', {{bubbles: true}}));
            vis[1].value = hasta; vis[1].dispatchEvent(new Event('change', {{bubbles: true}}));
            res = vis.map(i => i.id + '=' + i.value);
        }}
    }}
    return 'fechas setteadas: ' + (res.join(', ') || 'ninguna — inputs: ' + inputs.slice(0,5).map(i=>i.id).join(','));
}})()
"""


# ── Descarga genérica vía Browser.setDownloadBehavior ─────────────────────────

def _download_page(page_url, dest_dir, setup_js, verbose):
    """
    Abre una pestaña, navega a page_url, opcionalmente corre setup_js para setear
    fechas/filtros, y luego hace clic en el botón de exportar Excel.
    Usa Browser.setDownloadBehavior para que el archivo aterrice en dest_dir.
    Retorna la ruta del archivo descargado, o None si falló.
    """
    os.makedirs(dest_dir, exist_ok=True)

    # ── 1. Setear directorio de descarga a nivel browser ──────────────────────
    browser_ws = _browser_ws_url()
    if not browser_ws:
        raise RuntimeError("CDP no disponible — Brave no está corriendo con --remote-debugging-port=9222")

    browser_cdp = _CDP(browser_ws, timeout=30)
    try:
        browser_cdp.call("Browser.setDownloadBehavior",
                         behavior="allow",
                         downloadPath=dest_dir,
                         eventsEnabled=True)
    except Exception as e:
        browser_cdp.close()
        raise RuntimeError(f"Browser.setDownloadBehavior falló: {e}")

    # Anotar archivos preexistentes para detectar el nuevo
    before = set(glob.glob(os.path.join(dest_dir, "*.xlsx")))

    # ── 2. Abrir pestaña nueva ────────────────────────────────────────────────
    ws_url, tab_id = _new_tab_url()
    tab_cdp = _CDP(ws_url, timeout=90)

    try:
        tab_cdp.call("Page.enable")

        if verbose: print(f"    Navegando a {page_url}...")
        tab_cdp.call("Page.navigate", url=page_url)

        # Esperar carga (timeout generoso para ASP.NET)
        try:
            tab_cdp.wait_event("Page.loadEventFired", timeout=25)
        except TimeoutError:
            pass
        time.sleep(2)   # ASP.NET a veces renderiza después del load event

        # ── 3. Setup (fechas, filtros) ─────────────────────────────────────
        if setup_js:
            res = tab_cdp.call("Runtime.evaluate",
                               expression=setup_js,
                               awaitPromise=True,
                               returnByValue=True)
            val = res.get("result", {}).get("value", "")
            if verbose: print(f"    Setup: {val}")
            time.sleep(1)

        # ── 4. Click Exportar (dispara la generación del archivo en el servidor) ──
        if verbose: print(f"    Exportando...")
        res = tab_cdp.call("Runtime.evaluate",
                            expression=_EXPORT_JS,
                            awaitPromise=True,
                            returnByValue=True)
        val = res.get("result", {}).get("value", "")
        if verbose: print(f"    Export: {val}")

        if val and val.startswith("NO ENCONTRADO"):
            raise RuntimeError(f"Botón de exportar no encontrado: {val}")

        # ── 4b. Click Descargar (segundo paso, habilitado tras generar el archivo) ──
        time.sleep(4)   # esperar a que el servidor genere el archivo
        res = tab_cdp.call("Runtime.evaluate",
                            expression=_DOWNLOAD_JS,
                            awaitPromise=True,
                            returnByValue=True)
        val2 = res.get("result", {}).get("value", "")
        if verbose: print(f"    Download: {val2}")
        # Si no hay botón Descargar separado, asumimos que Exportar ya disparó la descarga

        # ── 5. Esperar el evento de descarga completada ───────────────────
        if verbose: print(f"    Esperando archivo...")

        def _is_complete(p):
            return p.get("state") == "completed"

        try:
            browser_cdp.wait_event("Browser.downloadProgress",
                                   timeout=60,
                                   filter_fn=_is_complete)
        except TimeoutError:
            # Fallback: monitorear el directorio directamente
            for _ in range(30):
                time.sleep(2)
                after = set(glob.glob(os.path.join(dest_dir, "*.xlsx")))
                new_files = after - before
                if new_files:
                    break

        # ── 6. Detectar el archivo nuevo ──────────────────────────────────
        time.sleep(1)
        after = set(glob.glob(os.path.join(dest_dir, "*.xlsx")))
        new_files = after - before
        if not new_files:
            # Puede que el archivo sobreescribió uno existente (mismo nombre)
            # Buscar el modificado más recientemente
            all_xlsx = sorted(glob.glob(os.path.join(dest_dir, "*.xlsx")),
                              key=os.path.getmtime, reverse=True)
            if all_xlsx and (time.time() - os.path.getmtime(all_xlsx[0])) < 120:
                new_files = {all_xlsx[0]}

        if not new_files:
            raise RuntimeError("No apareció ningún archivo xlsx en el directorio de descarga")

        fpath = sorted(new_files, key=os.path.getmtime)[-1]
        if verbose:
            print(f"    Descargado: {os.path.basename(fpath)} "
                  f"({os.path.getsize(fpath):,} bytes)")
        return fpath

    finally:
        tab_cdp.close()
        _close_tab(tab_id)
        browser_cdp.close()


# ── API pública ────────────────────────────────────────────────────────────────

def descargar_stock(verbose=True):
    """Descarga stock actual de Contabilium → Data/Inventario/Contabilium/."""
    if not _cdp_port_open():
        print("  [Contabilium stock] CDP no disponible, omitiendo descarga")
        return None
    if verbose: print("  [Contabilium] Descargando stock...")
    try:
        return _download_page(STOCK_URL, CONTABILIUM_STOCK_DIR, setup_js=None, verbose=verbose)
    except Exception as e:
        print(f"  [Contabilium stock] Error: {e}")
        return None


def descargar_historial(verbose=True, desde_date=None, hasta_date=None):
    """
    Descarga historial de stock (salidas) de Contabilium → Data/Salidas/Contabilium/.
    El archivo se nombra con el RANGO de fechas pedido (no la fecha de descarga),
    para que dos descargas el mismo día no se pisen entre sí.
    """
    if not _cdp_port_open():
        print("  [Contabilium historial] CDP no disponible, omitiendo descarga")
        return None

    if desde_date is None:
        ultima = _ultima_fecha_historial()
        desde_date = (ultima + timedelta(days=1)) if ultima else date.fromisoformat(CONTABILIUM_CUTOVER)
    if hasta_date is None:
        hasta_date = max(date.today(), desde_date)   # asegura desde <= hasta

    if desde_date > date.today():
        if verbose: print(f"  [Contabilium historial] Ya está al día (desde {desde_date} es futuro)")
        return None

    if verbose: print(f"  [Contabilium] Descargando historial de stock ({desde_date} a {hasta_date})...")
    try:
        fpath = _download_page(HISTORIAL_URL, CONTABILIUM_DIR,
                               setup_js=_historial_setup_js(desde_date, hasta_date),
                               verbose=verbose)
        if fpath:
            nuevo_nombre = os.path.join(
                CONTABILIUM_DIR,
                f"_tmp_HistorialStock_16749_{desde_date:%Y%m%d}_{hasta_date:%Y%m%d}.xlsx")
            if os.path.abspath(fpath) != os.path.abspath(nuevo_nombre):
                if os.path.exists(nuevo_nombre):
                    os.remove(nuevo_nombre)
                os.rename(fpath, nuevo_nombre)
                if verbose: print(f"    Renombrado a: {os.path.basename(nuevo_nombre)}")
                fpath = nuevo_nombre
        return fpath
    except Exception as e:
        print(f"  [Contabilium historial] Error: {e}")
        return None


def descargar_todo(verbose=True):
    """Descarga historial y stock de Contabilium."""
    descargar_historial(verbose=verbose)
    descargar_stock(verbose=verbose)
