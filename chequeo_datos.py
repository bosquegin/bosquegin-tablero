#!/usr/bin/env python3
"""
chequeo_datos.py — chequeo periodico de frescura de datos publicados.

Corre solo (sin Contabilium, sin Drive, sin CDP): lee data_meta.js ya
publicado en GitHub y compara ventas_hasta/stock_hasta contra la fecha real.
Si estan desactualizados, abre (o mantiene) un Issue de GitHub asignado al
dueno del repo -- GitHub manda mail automatico por eso, sin necesidad de
configurar SMTP ni ningun secreto nuevo.

Pensado para correr por cron (ver .github/workflows/chequeo_datos.yml),
para enterarse de un problema aunque nadie abra el tablero ese dia.
"""
import json, os, re, subprocess, sys, urllib.request
from datetime import date, datetime

GITHUB_REPO    = "bosquegin/bosquegin-tablero"
META_URL       = f"https://raw.githubusercontent.com/{GITHUB_REPO}/main/data_meta.js"
ISSUE_MARCA    = "Datos desactualizados"  # se busca por texto en el titulo, sin depender de un label
ISSUE_ASSIGNEE = "bosquegin"

# Mismos umbrales que el banner del tablero (SupplyBosquegin.html): ventas
# se mide en meses (dato mensual), stock en dias (dato que se actualiza a
# diario cuando todo funciona).
UMBRAL_VENTAS_MESES = 2
UMBRAL_STOCK_DIAS   = 5


def _meses_atras(ym_str, hoy):
    y, m = (int(x) for x in ym_str.split("-"))
    return (hoy.year * 12 + hoy.month) - (y * 12 + m)


def _dias_atras(ymd_str, hoy):
    d = datetime.strptime(ymd_str[:10], "%Y-%m-%d").date()
    return (hoy - d).days


def _gh(*args, input_text=None):
    # encoding explicito: en Windows subprocess.run cae al codec de la
    # consola (cp1252) por defecto, que no puede decodificar el emoji del
    # titulo del Issue -- se rompe el thread lector y r.stdout queda None
    # (bug real detectado probando esto en local 2026-08-07).
    r = subprocess.run(["gh", *args], capture_output=True, text=True,
                        encoding="utf-8", errors="replace", input=input_text)
    if r.returncode != 0:
        raise RuntimeError(f"gh {' '.join(args)} fallo: {r.stderr.strip()}")
    return r.stdout.strip()


def main():
    hoy = date.today()
    try:
        with urllib.request.urlopen(META_URL, timeout=20) as r:
            raw = r.read().decode("utf-8")
    except Exception as e:
        print(f"No se pudo leer data_meta.js: {e}")
        sys.exit(1)

    m = re.search(r"window\.BOSQUE_DATA\.meta\s*=\s*(\{.*\});", raw)
    if not m:
        print("data_meta.js con formato inesperado, no se puede chequear.")
        sys.exit(1)
    meta = json.loads(m.group(1))

    problemas = []
    vh = meta.get("ventas_hasta")
    if vh and re.match(r"^\d{4}-\d{2}$", vh):
        atraso = _meses_atras(vh, hoy)
        if atraso >= UMBRAL_VENTAS_MESES:
            problemas.append(f"- **Salidas/Ventas**: último dato real es de {vh} ({atraso} meses atrás).")

    sh = meta.get("stock_hasta")
    if sh and re.match(r"^\d{4}-\d{2}-\d{2}", sh):
        atraso_d = _dias_atras(sh, hoy)
        if atraso_d >= UMBRAL_STOCK_DIAS:
            problemas.append(f"- **Stock**: último dato real es del {sh} ({atraso_d} días atrás).")

    todos_abiertos = json.loads(_gh("issue", "list", "--repo", GITHUB_REPO,
                                     "--state", "open", "--json", "number,title") or "[]")
    existentes = [it for it in todos_abiertos if ISSUE_MARCA in it["title"]]

    if problemas:
        cuerpo = (
            "El chequeo periódico automático detectó datos desactualizados en el tablero:\n\n"
            + "\n".join(problemas) +
            "\n\nEsto suele pasar cuando Contabilium devuelve 429 (demasiadas peticiones) "
            "durante Actualizar. El tablero ya muestra un banner rojo avisando esto, pero "
            "este chequeo corre solo (sin que nadie tenga que abrir el tablero) para que "
            "te enteres apenas pasa.\n\n"
            "_Este Issue se cierra solo cuando el chequeo vuelve a ver los datos al día._"
        )
        if existentes:
            print(f"Ya hay un Issue abierto (#{existentes[0]['number']}), no se crea uno nuevo.")
        else:
            titulo = f"⚠️ {ISSUE_MARCA} — {hoy.isoformat()}"
            url = _gh("issue", "create", "--repo", GITHUB_REPO,
                      "--title", titulo, "--body", cuerpo,
                      "--assignee", ISSUE_ASSIGNEE)
            print(f"Issue creado: {url}")
    else:
        print("Datos al día, sin problemas.")
        for it in existentes:
            _gh("issue", "close", str(it["number"]), "--repo", GITHUB_REPO,
                "--comment", "Los datos volvieron a estar al día — cierro este aviso automáticamente.")
            print(f"Issue #{it['number']} cerrado (datos recuperados).")


if __name__ == "__main__":
    main()
