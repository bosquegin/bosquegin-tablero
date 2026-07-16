"""
contabilium_api.py
Cliente de la API REST de Contabilium (https://rest.contabilium.com).

Auth: OAuth2 client_credentials. client_id = email de la cuenta,
client_secret = API Key (Mi cuenta > Configuración > API).
Token cacheado en contabilium_token_cache.json (expira a las ~24hs).

Uso principal: obtener stock en vivo por depósito, sin depender de
exports manuales de Contabilium.
"""
import os, json, time
import urllib.request, urllib.parse, urllib.error

BASE          = r"C:\Users\SupplyDestileria\Documents\Bosque\Claude\Tablero operativo"
CONFIG_FILE   = os.path.join(BASE, "contabilium_config.json")
TOKEN_CACHE   = os.path.join(BASE, "contabilium_token_cache.json")
TOKEN_URL     = "https://rest.contabilium.com/token"
API_BASE      = "https://rest.contabilium.com/api"
_UA           = ("Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
                  "(KHTML, like Gecko) Chrome/124.0.0.0 Safari/537.36")


def _load_config():
    with open(CONFIG_FILE, encoding="utf-8") as f:
        return json.load(f)


def _fetch_new_token():
    cfg = _load_config()
    body = urllib.parse.urlencode({
        "grant_type": "client_credentials",
        "client_id": cfg["client_id"],
        "client_secret": cfg["client_secret"],
    }).encode()
    req = urllib.request.Request(
        TOKEN_URL, data=body,
        headers={"Content-Type": "application/x-www-form-urlencoded",
                 "User-Agent": _UA},
        method="POST")
    with urllib.request.urlopen(req, timeout=20) as r:
        data = json.loads(r.read().decode())
    token = data["access_token"]
    expires_in = data.get("expires_in", 3600)
    cache = {"token": token, "expires_at": time.time() + expires_in - 60}
    with open(TOKEN_CACHE, "w", encoding="utf-8") as f:
        json.dump(cache, f)
    return token


def get_token():
    """Devuelve un bearer token válido, cacheado en disco."""
    if os.path.exists(TOKEN_CACHE):
        try:
            with open(TOKEN_CACHE, encoding="utf-8") as f:
                cache = json.load(f)
            if cache.get("expires_at", 0) > time.time():
                return cache["token"]
        except Exception:
            pass
    return _fetch_new_token()


def _api_get(path, params=None, retry_on_401=True):
    qs = ("?" + urllib.parse.urlencode(params)) if params else ""
    url = f"{API_BASE}{path}{qs}"
    token = get_token()
    req = urllib.request.Request(
        url, headers={"Authorization": f"Bearer {token}", "User-Agent": _UA},
        method="GET")
    try:
        with urllib.request.urlopen(req, timeout=30) as r:
            return json.loads(r.read().decode())
    except urllib.error.HTTPError as e:
        if e.code == 401 and retry_on_401:
            _fetch_new_token()
            return _api_get(path, params, retry_on_401=False)
        raise


def get_depositos():
    """Lista de depósitos: [{'Id': int, 'Nombre': str, ...}, ...]"""
    return _api_get("/inventarios/getDepositos")


def get_stock_by_deposito(dep_id):
    """
    Stock actual de TODOS los productos en un depósito.
    Pagina automáticamente (50 items por página).
    Devuelve dict {codigo: stock_actual}.
    """
    stock = {}
    page = 1
    while True:
        data = _api_get("/inventarios/getStockByDeposito",
                        {"id": dep_id, "page": page})
        items = data.get("Items", [])
        if not items:
            break
        for it in items:
            cod = str(it["Codigo"]).strip()
            stock[cod] = stock.get(cod, 0.0) + float(it.get("StockActual", 0))
        if len(items) < 50:
            break
        page += 1
    return stock


def get_stock_todos_depositos(nombres_filtro=None):
    """
    Stock en vivo para todos los depósitos (o solo los indicados en nombres_filtro,
    comparación case-insensitive contra el nombre de depósito).
    Devuelve dict {nombre_deposito_upper: {codigo: stock_actual}}.
    """
    deps = get_depositos()
    result = {}
    filtro = {n.upper() for n in nombres_filtro} if nombres_filtro else None
    for dep in deps:
        nombre = dep["Nombre"].strip().upper()
        if filtro and nombre not in filtro:
            continue
        result[nombre] = get_stock_by_deposito(dep["Id"])
    return result


def get_active_codigos():
    """
    Códigos de productos ACTIVOS en Contabilium (conceptos/search excluye
    inactivos por defecto). Devuelve set de strings de Codigo.
    """
    codigos = set()
    page = 1
    while True:
        data = _api_get("/conceptos/search", {"page": page})
        items = data.get("Items", [])
        if not items:
            break
        for it in items:
            cod = str(it.get("Codigo", "")).strip()
            if cod:
                codigos.add(cod)
        if len(items) < 50:
            break
        page += 1
    return codigos
