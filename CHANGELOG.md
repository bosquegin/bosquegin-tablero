# Changelog — SupplyBosquegin

Historial de cambios del Tablero Operativo, reconstruido desde `git log` y mantenido a partir de ahora versión por versión. Orden: más reciente primero.

> **Cómo se actualiza este archivo:** cada vez que sumemos cambios en una sesión, se agregan como bullets bajo **[Sin versionar]**. Cuando se hace un bump de versión (`vX.Y`), esos bullets pasan a formar la sección de esa versión con su fecha.

---

## [Sin versionar]

_(sin cambios pendientes de versión todavía)_

---

## v3.41 — 2026-08-08

- **seguridad (arreglo final, no parche):** el login del sitio publicado dejó de exponer `password_hash`/`salt`/`cloud_token` de cada usuario en el repo público. El problema real no era el nombre del archivo (`auth_static.js`) sino que el repo entero es público — cualquier archivo en él, esté o no linkeado desde el sitio, se puede descargar vía la API de GitHub o `raw.githubusercontent.com`. Mover esos datos a otro archivo del mismo repo no solucionaba nada.
  - Los usuarios (hash+salt+cloud_token) ahora viven en **Cloudflare Workers KV** (`AUTH_KV`), fuera del repo git, accesible solo por el Worker.
  - Nuevo endpoint `POST /login` en el Worker: la contraseña se verifica server-side contra KV: el navegador ya no descarga ni compara hashes de nadie.
  - `auth_static.js` queda como un stub vacío (`window.BG_AUTH = []`) — nada lo lee más. El pipeline de Actualizar dejó de subirlo a git.
  - El panel de Administración (crear/editar/borrar usuarios) sigue funcionando igual en ambos lados — local escribe en `users.json` y espeja cada cambio al Worker (best-effort); publicado ya pegaba directo al Worker, ahora el Worker guarda en KV en vez del archivo.
  - Migración: se importaron los 10 usuarios existentes preservando sus contraseñas actuales (nadie tiene que cambiarla) pero **rotando todos los `cloud_token`** — los anteriores habían circulado en texto plano en el repo público, así que quedan invalidados de una.
  - **Nota:** esto cierra la exposición hacia adelante. El contenido viejo de `auth_static.js` (hasta v3.40) sigue existiendo en el historial de git de este repo público — no se puede "desexponer" sin reescribir el historial (una operación grande y disruptiva, no se hizo). Los `cloud_token` viejos ya están invalidados por la rotación; los `password_hash` viejos requerirían crackeo offline para ser útiles, mismo riesgo residual bajo que ya existía antes de este arreglo.
- **fix:** email de Loana corregido (`.con` → `.com`) — quedó pendiente de una revisión anterior.

---

## v3.40 — 2026-08-07

- **fix:** email mal cargado de Loana — era `loa.desalvo@bosquegin.con` (terminaba en `.con`), corregido a `.com`. Corregido en `users.json` local y regenerado `auth_static.js`.
- **fix:** Control Inventario comparaba TODOS los códigos que aparecían en Contabilium o Klozer, incluyendo cosas que no son productos de catálogo (ítems internos de Klozer, combos raros, etc.). Ahora se filtra a solo los códigos que existen en la hoja **Inventario Producto** (`PRODUCTOS.xlsx`).
- **feat:** Control Inventario suma columnas **Rubro** y **Subrubro** a la tabla.
- **feat:** el encabezado de la tabla de Control Inventario ahora es clickeable para ordenar por cualquier columna (Cód., Producto, Rubro, Subrubro, Contabilium, Klozer, Diferencia, % Diferencia), con flecha indicando la dirección. Por defecto ordena por |Diferencia| descendente, como antes.
- **pendiente:** sigue faltando mover la verificación de contraseña del login al Worker (para que `auth_static.js` deje de exponer hashes/cloud_token en el repo público) — es el próximo paso.

---

## v3.39 — 2026-08-07

- **feat:** nueva hoja **Control Inventario** — compara stock de KLOZER entre Contabilium (API) y la página propia de Klozer (klozer.co), por código de producto, ordenado por diferencia. Tiene su propio botón "Actualizar Klozer" (no corre con el Actualizar general), columnas de Contabilium/Klozer/Diferencia/% Diferencia, observaciones editables, checkbox de "ajustado", botón "Guardar cambios" e informe de la corrida. Klozer no tiene API — se lee la página vía CDP sobre una pestaña ya logueada en el Brave local (mismo mecanismo que ya usa cervezas), así que **solo funciona en el tablero local**; el sitio publicado muestra la última comparación guardada de solo lectura.
- **fix crítico (mientras se probaba):** el servidor local (`servidor_bosquegin.py`) era single-threaded — un solo pedido lento (ej. esperando hasta 30-40s una pestaña de Klozer trabada) dejaba el tablero entero sin responder a nada más, ni login ni navegación. Pasado a `ThreadingHTTPServer`; verificado con una prueba real (pedido lento + pedido rápido en simultáneo) que ya no se traba.
- **fix (durante las pruebas):** las cervezas en caja mostraban una diferencia falsa de ~83% entre Contabilium (cuenta latas sueltas) y Klozer (cuenta cajas) — se detecta el tamaño de caja desde el propio nombre del producto en Klozer (ej. "CAJA X6") y se multiplica antes de comparar.
- **fix (durante las pruebas):** Brave suspende pestañas en segundo plano y dejan de responder a CDP — se relanza con flags que evitan la suspensión, y el código ahora trae la pestaña al frente (`/json/activate`) antes de leerla, con reintento.
- **fix (durante las pruebas):** la página de Klozer no conserva la Unidad de negocio seleccionada entre reinicios del navegador — el código ahora selecciona "Temple Brewery" y dispara la búsqueda él solo, en vez de depender de que quede pre-configurada a mano.
- Verificado de punta a punta con un usuario de prueba (creado y borrado en la misma sesión): actualizar trae datos reales, guardar persiste observaciones/ajustado y genera el informe.
- **pendiente:** falta terminar de mover la verificación de contraseña del login al Worker (para que `auth_static.js` deje de exponer hashes/cloud_token en el repo público) — quedó acordado pero no implementado todavía.

---

## v3.38 — 2026-08-07

- **feat:** "Administración" (gestión de usuarios) ahora funciona también en el sitio publicado, no solo en local. El Cloudflare Worker suma `/admin_users` (listar/crear/editar/borrar), con hashing PBKDF2-SHA256/100.000 idéntico al del servidor local y al del login del sitio publicado, leyendo y escribiendo `auth_static.js` directo en GitHub (mismo archivo del que ya se validan las sesiones).
- **operativo:** `auth_static.js` (local) y `auth_static.js` en GitHub (publicado) son dos "fuentes de verdad" separadas para el lado cloud — no se sincronizan solas entre ediciones. Si se edita un usuario desde el tablero local y desde el publicado sin correr Actualizar entre medio, el último guardado pisa al otro (mismo comportamiento que ya tenían los demás edits publicados de este tablero).
- **seguridad — a tener en cuenta:** `auth_static.js` (necesario para que el login funcione sin backend en el sitio estático) contiene el hash+salt de la contraseña de cada usuario y su `cloud_token`, y el repositorio es público — cualquiera puede descargarlo. No es nuevo de esta versión, ya era así antes, pero vale la pena que se sepa: un `cloud_token` filtrado permite guardar cambios (Actualizar, Proyección, usuarios) sin conocer la contraseña, y los hashes son atacables offline si la contraseña es débil. No se tocó este diseño en esta versión — se avisa para decidir si conviene revisarlo (repo privado, o mover la verificación a un lugar no público).
- **pendiente:** falta pegar el `worker.js` actualizado en el dashboard de Cloudflare para que esto quede activo en el sitio publicado (paso manual, sin auto-deploy).

---

## v3.37 — 2026-08-07

- **feat:** en Proyección Producción, al desplegar un trimestre ahora aparece una explicación en criollo de los 7 indicadores del encabezado (En alerta, A comprar, Cumplim. objetivo, Forecast Accuracy/Bias/MAPE). Para el trimestre EN CURSO suma una advertencia: los tres indicadores de forecast comparan el objetivo del mes completo contra lo vendido en lo que va del mes, así que a principios de mes se ven mucho peor de lo que realmente son (se corrigen solos con el correr del mes) — evita que se lean como una alarma real cuando es solo un efecto de comparar un mes a medio transcurrir.

---

## v3.36 — 2026-08-07

- **fix:** en Proyección Producción, los meses ya cerrados de trimestres pasados (Q1 y Q2 2026) mostraban un saldo de cierre calculado por cascada (stock + abastecimiento − venta objetivo) en vez del stock real medido ese mes — llegaba a mostrar alarmas de "COMPRAR" falsas cuando el stock real era saludable (ej. Gin Bosque Nativo en junio: -1.092 calculado vs. 2.204 real). Ahora esos meses se anclan directo al cierre real de `STOCK_CIERRE_MES` — el mismo histórico que ya usa "Rotación mensual por rubro" en Inventario Productos, para que coincida en todo el tablero.
- **fix:** enero 2026 no tenía cierre real propio (el último export de depósitos es del 16, quince días antes de fin de mes). Se usa como proxy el stock real del 4 de febrero — primer día de febrero con un reporte de depósitos válido (el del 3/2 resultó ser de otro tipo de reporte, "stock en tanques"). Aprobado explícitamente después de comparar ambas opciones lado a lado.
- **verificado:** simulación mostrada y aprobada antes de tocar código; luego corrida real de Actualizar local de punta a punta (Contabilium respondió sin 429) confirmando que el dato publicado coincide exactamente con lo simulado — 8 alarmas de faltante falsas eliminadas, mayor corrección individual +3.296 unidades.

---

## v3.35 — 2026-08-07

- **fix:** en Lista de Precios, el primer cuadro ("Mes a mes") quedaba recortado con scroll interno a 400px de alto — ahora se ve completo, igual que el segundo cuadro ("Evolución mensual").
- **feat:** la tabla evolutiva de Cervezas (Costo Producción → Temple) ahora muestra quién elaboró cada lata **por mes** (Ortuzar/Filidoro/Bierhaus/Cmq/etc.), en una columna nueva "Elaborador" a la izquierda — antes solo se guardaba el elaborador del mes más reciente, así que un cambio de fasón a mitad de año no se veía reflejado en el historial.

---

## v3.34 — 2026-08-07

- **feat:** chequeo periódico automático de frescura de datos (`chequeo_datos.py` + `.github/workflows/chequeo_datos.yml`), corre solo 4 veces al día sin intervención y sin gastar cupo de Contabilium (solo lee `data_meta.js` ya publicado). Si Salidas o Stock quedan desactualizados, abre un Issue de GitHub asignado a la cuenta del tablero (dispara mail automático de GitHub, sin SMTP ni secretos nuevos que configurar) y lo cierra solo cuando los datos se recuperan. Objetivo: enterarse de un problema el mismo día, aunque nadie abra el tablero.

---

## v3.33 — 2026-08-07

- **fix:** en "Lista de Precios" las cervezas tenían un rubro inventado "CERVEZAS" que no existe en ningún otro lado del tablero — en Inventario Producto (y ahora también en Ventas/Salidas) esos mismos códigos son rubro **BEBIDAS**, sub-rubro CERVEZAS. Corregido en el pipeline y republicado.
- **fix:** "VASOS" (Lista de Precios) y "CRISTALERIA" (Inventario Producto) eran el mismo rubro real con nombre distinto en cada planilla — se normaliza al nombre de Inventario (CRISTALERIA), que es el que se toma como canónico.
- **fix:** el selector de rubro de la hoja Salidas tenía una lista fija a mano (BEBIDAS/BOLSAS/CERVEZA/COMBO/ESTUCHE/FUNDA/INDUMENTARIA/MOBILIARIO/SIN RUBRO/VASOS) que ya no coincidía con los rubros reales — "CERVEZA" y "VASOS" no existen en los datos (filtrar por esos nunca traía resultados) y faltaba "CRISTALERIA" (no se podía filtrar por él). Ahora el selector se arma solo, a partir de los rubros que realmente aparecen en los datos.
- **feat:** el selector de Rubro + buscador de artículo de la hoja Salidas se movió al encabezado de filtros (junto a Año/Mes/Cliente) — ya afectaba a toda la hoja (KPIs, gráfico mensual, tabla de rubros), pero estaba visualmente pegado solo a la tabla "Detalle por producto", lo que hacía parecer que filtraba nada más que esa tabla.

---

## v3.32 — 2026-08-06

- **fix crítico:** Salidas y Proyección Producción volvieron a mostrar solo hasta junio, sin julio ni agosto. Causa raíz confirmada en el log real de la corrida cloud: Contabilium devolvió `429 Too Many Requests` desde la primerísima llamada del run (probablemente el cupo del día ya gastado por varias corridas de prueba hoy), y el pipeline — aunque detectó y logueó el error (`❌ Consolidado salidas — desactualizado`) — igual publicó el consolidado viejo sin avisar nada visible en el tablero. Se restauraron los datos publicados con la corrida local (que sí tenía julio y agosto al día) mientras se aplican las mejoras de fondo.
- **feat:** nuevo banner rojo, visible en cualquier pestaña del tablero (no solo en el resumen), que se activa automáticamente cuando `ventas_hasta` queda 2+ meses detrás de la fecha real — algo que el punto verde de "actualizado hoy" nunca detectaba, porque ese punto solo mira cuándo corrió el Actualizar, no si Contabilium realmente trajo datos nuevos. Objetivo: que nunca más se tome una decisión con Salidas desactualizado sin saberlo.
- **fix:** `contabilium_api.py` reintenta automáticamente con espera creciente (10s/30s/60s) ante un `429`, en vez de rendirse en el primer intento — cubre el caso (ya confirmado antes, ver v3.30) de que el límite de Contabilium a veces es una ráfaga corta que se libera sola en minutos.
- **operativo:** la evidencia de hoy (429 incluso con más de 1h48m desde la corrida anterior, y fallando desde la primerísima llamada del run) apunta más fuerte a un tope **diario** de Contabilium, agotado por las varias corridas de prueba de hoy — no algo que un cooldown por hora pueda evitar. Pendiente: consultarle a Contabilium cuál es el límite real de su API para dimensionar esto con precisión en vez de a ciegas.

---

## v3.31 — 2026-08-06

- **fix:** en la hoja "Costo Producción → Cervezas/Temple" solo aparecía el mes actual, sin historial. Causa raíz: la corrida cloud arma su directorio de trabajo bajando archivos de Google Drive (no hace `git checkout`), y esa descarga no trae la subcarpeta de caché `Data/Costos y PVP/cervezas_meses/` (Drive no la espeja y la descarga tampoco recorre subcarpetas). Sumado a que la nube no tiene acceso a CDP/Chrome local para descubrir y leer las hojas mensuales en vivo, cada corrida cloud arrancaba con el caché vacío y todos los meses históricos fallaban — solo quedaba el mes actual, vía un fallback aparte. Ahora `actualizar_cloud.py` descarga ese caché ya commiteado desde GitHub antes de correr el pipeline, así los meses históricos se leen del caché igual que en una corrida local.
- **fix:** quedaron sin versionar dos cambios ya subidos en la sección "Costo Producción": se sacó una leyenda de fuente que quedaba pegada entre el encabezado y el selector de producto, se corrigió que los dos filtros quedaran "pegoteados" y se duplicaran visualmente al hacer scroll (tenían `position:sticky` heredado de una clase global, ahora `static` en esta sección), y se separó el cuadro combinado Bosque+Feriado+Temple en 3 pestañas independientes (Gin / Vermú / Cerveza) para que cada línea tenga su propia vista, sin mezclar productos de rubros distintos en el mismo selector.

---

## v3.30 — 2026-08-06

- **fix:** el número de "Proyección abastecimiento" solo se tachaba después de que volviera la respuesta del servidor (varias llamadas a la API de GitHub vía el Worker) — se sentía como que no tachaba solo. Ahora se tacha/destacha al toque, apenas se clickea el checkbox, y se revierte si el guardado falla.
- **operativo:** nueva evidencia sobre el límite de Contabilium — una corrida a las 20:10 (5 horas después de la anterior, muy por fuera del cooldown de 1 hora) también devolvió 429. Esto sugiere que el límite real puede ser un **tope diario** (probablemente gastado hoy a fuerza de tanto probar en el día), no solo "no correr dos veces seguidas". Una corrida posterior a las 23:21 sí funcionó bien. A seguir de cerca — si se repite, el cooldown va a necesitar ajustarse a un límite diario además del de 1 hora.
- Se restauraron los datos publicados al último estado bueno tras el 429 de las 20:10, y se les reaplicó el fix de consistencia de totales (v3.29) para no perder ese arreglo en la restauración manual.

---

## v3.29 — 2026-08-06

- **feat:** editar la proyección de abastecimiento y tildar "ya ingresó" en Proyección Producción ahora funciona también en el **sitio publicado** (antes solo en el tablero local). El Cloudflare Worker suma dos endpoints (`/guardar_proyeccion_abastecimiento`, `/guardar_proyeccion_ingreso`) que leen y escriben directo en GitHub via Contents API, con el mismo cálculo de `actualizar_bosquegin.py` portado a JS — mismo resultado en los dos lugares, verificado con una prueba real de punta a punta (ida y vuelta sobre un mes cerrado, sin tocar ningún número).
- **operativo:** el token de GitHub del Worker (`GH_TOKEN`) tenía permiso "Contents: Read-only" — alcanzaba para disparar Actualizar, pero no para guardar ediciones. Se subió a "Read and write".
- **fix:** `stock_total`/`total_objetivo_ventas`/`meses_stock` de un producto podían mostrar un valor distinto según si lo había tocado por última vez el Actualizar (sumaba TODOS los meses del trimestre, incluidos los cerrados) o una edición manual (excluye los cerrados, correcto) — detectado comparando el resultado real de la nueva función contra el del Actualizar. Ahora los dos caminos calculan igual.

---

## v3.28 — 2026-08-06

- **fix:** los cargos de servicio de Contabilium (ej. "ROYALTIES julio", `Codigo "000"`, `Tipo "S"`) se procesaban igual que un producto y aparecían como fila fantasma "código 0 / SIN RUBRO" en Salidas. Ahora solo se cuentan ítems con `Tipo "P"` (producto) — confirmado contra la API real que todo producto físico viene con `Tipo "P"`.
- **fix:** tildar "ya ingresó" o editar la proyección de abastecimiento en el **sitio publicado** (`bosquegin.github.io`, sin backend) tiraba un error críptico ("Unexpected token '<'... is not valid JSON") — esas funciones solo pueden guardar en el servidor local. Ahora avisan claro que hay que usar el tablero local en vez de intentar guardar donde no hay dónde.
- **verificado:** el reporte de "Actualizar no actualiza" / Salidas sin julio-agosto / checkbox no visible correspondía a una corrida anterior — la corrida de las 12:03 (ART) terminó bien (sin 429, 539 filas de Contabilium) y los datos publicados ya estaban correctos al momento de revisar (`generado: 2026-08-06T12:08:42-03:00`). El Service Worker de la PWA ya tenía manejo robusto de actualización (chequeo al abrir + cada 15 min + reload automático al detectar versión nueva, de v3.20/v3.21) — no se encontró una falla de fondo ahí.

---

## v3.27 — 2026-08-06

- **fix:** el mensaje de error largo del límite de 1 hora (con los minutos restantes) se mostraba cortado con "…" — el modal lo ponía en una línea de una sola fila (`text-overflow:ellipsis`) en vez del cuadro de error dedicado, que sí envuelve el texto completo. Se unificó también el manejo de excepciones inesperadas del flujo cloud, que tenía el mismo problema duplicado.
- **fix:** el cooldown de 1 hora del servidor local vivía solo en memoria — reiniciar el servidor lo reseteaba, justo lo que pasó hoy al reiniciar para aplicar el fix anterior. Ahora se persiste en disco (`Data/_ultima_actualizacion_local.txt`, no trackeado) y sobrevive a un reinicio. Se sembró con la hora de la última corrida real para que quede protegido también contra la corrida en la nube de hoy (local y nube no comparten el cooldown automáticamente, cada uno cuenta el suyo).

---

## v3.26 — 2026-08-06

- **feat:** límite de **1 corrida de Actualizar por hora**, sin importar quién lo dispare — para no volver a saturar la API de Contabilium (dos corridas seguidas alcanzan para gatillar el 429 del incidente de hoy). Implementado en el Cloudflare Worker (rechaza con 429 y minutos restantes) y en `servidor_bosquegin.py` (mismo límite para el Actualizar local).
- **fix:** el modal de Actualizar mostraba solo "HTTP 429"/"HTTP 503" genérico ante un error — ahora muestra el mensaje real del servidor (ej. "probá de nuevo en 42 min").

---

## v3.25 — 2026-08-06

- **fix:** tildar/destildar "ya ingresó" o corregir el abastecimiento de un mes **cerrado** todavía cambiaba `stock_total`, `total_objetivo_ventas`, `meses_stock` y `comprar` del producto (el fix anterior solo protegía `saldo_stock`). Ahora un mes cerrado no aporta nada a esos totales — tildarlo queda 100% informativo. Verificado con prueba aislada.
- **feat:** el checkbox de "ya ingresó" en Proyección Producción solo se mostraba para el mes en curso — en un mes ya cerrado no había forma de tildarlo desde el tablero. Ahora se muestra también ahí (con un tooltip distinto aclarando que es solo informativo).
- **operativo:** la API de Contabilium devolvió `429 Too Many Requests` dos veces en el día (primero por varias corridas seguidas de verificación, después por dos corridas separadas por ~5 minutos) — cada corrida completa hace ~200 llamadas a la API (stock por depósito + detalle de cada comprobante), así que dos corridas muy seguidas alcanzan para gatillar el límite. Cuando pasa, `Salidas` vuelve a cortar en `CONTABILIUM_CUTOVER` (2026-06-29) para esa corrida. No hay límite documentado públicamente por Contabilium. Mientras tanto se restauró `data_proyeccion.js` (local y publicado) al último estado bueno conocido, sin volver a llamar a la API.

---

## v3.24 — 2026-08-06

- **cambio de cálculo (intencional):** el saldo de stock del mes en curso en Proyección Producción ahora resta la `proyección_mensual` completa (antes solo `proyección_mensual − venta_actual`) — más conservador temprano en el mes, cuando la proyección lineal todavía tiene poca muestra. Aplicado en `aplicar_venta_real_mes_actual` (el cálculo del Actualizar).
- **fix:** ese mismo cambio de fórmula también se aplicó a `_proy_recalcular_derivados`, la función que recalcula el saldo al editar el abastecimiento o tildar "ya ingresó" a mano desde el tablero — para que quede consistente con lo que calcula el Actualizar.
- **fix:** tildar "ya ingresó" (o corregir el abastecimiento) de **cualquier** mes de un producto recalculaba en cascada **todos** los meses del trimestre, incluidos los ya cerrados — pisando el saldo real de stock de un mes pasado (ej. julio) con la fórmula genérica de mes futuro. Ahora los meses cerrados quedan excluidos de esa cascada: su saldo de stock quedó congelado en el cierre real, y tildar "ya ingresó" en un mes cerrado queda solo como dato informativo, sin cambiarle el saldo.

---

## v3.23 — 2026-08-06

- **infra:** el servidor Flask en Render (free tier) se reemplaza por **Cloudflare Worker + GitHub Actions** para el botón Actualizar — se dormía y no reaccionaba a tiempo. El Worker valida el `cloud_token` y dispara `workflow_dispatch`; `actualizar_cloud.py` corre como job de GitHub Actions (sin servidor siempre encendido, sin cold-start). Probado de punta a punta: corrida real completada en 3m29s.
- Se cargan los secretos de Contabilium, Google OAuth y `DRIVE_ROOT_FOLDER_ID` en GitHub Actions.
- Se desactiva el workflow "Keep Render awake" (ya sin uso).
- Se da de baja el servicio en Render.
- **chore:** repaso completo del tablero — todo lo que ya no se usa se mueve a `OLD/` (organizado por categoría) en vez de borrarse:
  - `OLD/render/`: `servidor_render.py`, `render.yaml`.
  - `OLD/tunnel-cloudflare/`: túnel Cloudflare del viejo Actualizar remoto (`activar_tunel.py`, `activar_publico.bat`, `instalar_tunel.bat`, `setup_tunnel.bat`, `tunnel_url.txt`) — ya no hace falta, el Worker le habla directo a GitHub Actions.
  - `OLD/gc-legacy/`: scripts de descarga de Gestión Cervecera, reemplazada por Contabilium desde 2026-06-29 (`gc_descargar_remitos.js`, `gc_descargar_via_browser.js`, `gc_descargar_stock.py`, `gc_write_batch.py`). `gc_relay_server.py` **no** se tocó — sigue en uso por el bookmarklet de descarga manual.
  - `OLD/legacy-data/`: `bosquegin_data.js`, el monolito de ~5MB que el dashboard ya no lee (reemplazado por los `data_*.js` con lazy loading desde v2.27).
  - `OLD/scratch/`: scripts y logs sueltos de debug (`read_temp.py`, `run_temp.bat`, `read_temp_out.txt`, `auth_url.txt`, `oauth_err.txt`, `oauth_out.txt`).
  - `OLD/data-temp/`: archivos temporales sin trackear que habían quedado sueltos en `Data/Salidas/Contabilium/`.
  - `OLD/windows-startup/`: `SupplyChain_Servidor.vbs`, que estaba en la carpeta de Inicio de Windows intentando levantar un proyecto "Supply Chain" que ya no existe (se fusionó a este tablero como la pestaña Compras en v2.12) — fallaba en silencio en cada arranque. También `iniciar_servidores.vbs`, un borrador viejo del lanzador ya superado por `iniciar_tablero_silencioso.vbs`.
- `Lib/site-packages` (155MB, dependencias de Python) sale de git tracking pero **queda en disco** — la usa directo el Python embebido (`python311._pth`), no era basura.
- `requirements.txt`: se sacan `flask`, `flask-cors`, `gunicorn` (solo las usaba `servidor_render.py`).
- **fix:** `index.html`, `manifest.json`, `service-worker.js` y `_redirects` quedaron sin pushear tras el rename a `SupplyBosquegin.html`, y el sitio publicado devolvía 404 (seguían apuntando a `bosquegin_dashboard.html`). Detectado y corregido el mismo día.
- **fix crítico:** `contabilium_api.py` tenía la ruta a la carpeta del proyecto hardcodeada a Windows (`C:\Users\...`). En GitHub Actions (runner Linux) esa ruta no existe, así que la autenticación con la API de Contabilium fallaba en silencio en **todas** las corridas en la nube desde que salió v3.22 — sin que se viera como error visible, `Salidas` se quedaba pegada en 2026-06-29 (el `CONTABILIUM_CUTOVER`) porque la única fuente de datos para julio/agosto en adelante es Contabilium. Se cambia `BASE` a una ruta relativa al propio script (`os.path.dirname(os.path.abspath(__file__))`), igual que ya hacía `actualizar_bosquegin.py` para el caso no-Windows. Verificado con una corrida real: pasó de 0 a 536 filas de Contabilium, julio y agosto ya aparecen en `Salidas`.
- **fix crítico:** `update_stock_cierre_mes()` (Proyección Producción → saldo de stock de meses cerrados) nunca corría en la nube — usaba la existencia de `SupplyBosquegin.html` como proxy de "modo local", pero en la nube `BASE` es un directorio temporal vacío que nunca tiene ese archivo. La función se cancelaba en silencio en cada corrida cloud y `data_stock_cierre.js` nunca se generaba ahí; sin ese dato, el saldo de stock de un mes ya cerrado (ej. julio) se seguía calculando como si fuera un mes futuro (stock actual − venta objetivo), dando saldos totalmente incorrectos (ej. código 100001: -576 en vez de 2944, el cierre real). Se saca la guarda vieja (no hacía falta para nada más en la función) y se agrega `data_stock_cierre.js` a `SECTION_FILES` de `actualizar_cloud.py` para que también se suba a GitHub desde la nube (antes solo lo pusheaba el flujo local). Verificado con una corrida real.
- **cambio de cálculo (intencional):** el saldo de stock del **mes en curso** en Proyección Producción restaba solo lo que falta vender (`proyección_mensual − venta_actual`) para no descontar dos veces la venta ya reflejada en el stock en vivo. Se cambia a restar la `proyección_mensual` completa a propósito — el saldo queda más conservador (más bajo) temprano en el mes, cuando la proyección lineal (venta real ÷ día del mes × días totales) todavía es poco confiable con pocos días de muestra. Aplicado tanto en el cálculo del Actualizar (`aplicar_venta_real_mes_actual`) como en el que se dispara al editar abastecimiento o tildar "ya ingresó" a mano desde el tablero (`_proy_recalcular_derivados`), para que no queden inconsistentes entre sí.

---

## v3.22 — 2026-08-05

- Las salidas se calculan ahora desde **comprobantes de Contabilium** (factura/cotización) en vez de CDP.

## v3.21 — 2026-08-05

- Chequeo periódico de actualización en la PWA instalada.

## v3.20 — 2026-08-05

- Fix: la PWA instalada no revisaba actualizaciones al abrir.

## v3.19 — 2026-08-04

- El modal de "Actualizar" pasa de mostrar un video a una animación cuadro por cuadro.

## v3.18 — 2026-08-04

- Pruebas del mecanismo de *cloud push* (crear/actualizar) antes de aplicar el fix.
- **Fix crítico:** el botón Actualizar (Render) no estaba actualizando los datos.

## v3.17 — 2026-08-03

- Corrige el orden de severidad al ordenar por Estado en Rotación por rubro y subrubro.

## v3.16 — 2026-08-03

- Agrega la opción de ordenar por Estado en Rotación por rubro y subrubro.

## v3.15 — 2026-08-03

- Permite ordenar por columna en Rotación por rubro y subrubro.

## v3.14 — 2026-08-03

- Corrige venta real y saldo de stock de meses cerrados en Proyección Producción.

## v3.13 — 2026-07-31

- Keep-alive de Render para el botón Actualizar.
- Amplía el margen de espera del botón Actualizar de 60s a 150s.

## v3.12 — 2026-07-31

- Fusiona el Desglose semanal dentro del Desglose mensual en Ventas.

## v3.11 — 2026-07-31

- La semana de calendario gregoriano ahora es continua y no reinicia al cambiar de mes.

## v3.10 — 2026-07-31

- Agrega la etiqueta "Semana N" arriba del rango de días en el desglose semanal.

## v3.9 — 2026-07-31

- Fix: el service worker servía la "cáscara" vieja de la app por caché HTTP.

## v3.8 — 2026-07-31

- El ícono de la PWA pasa a ser la botella Nativo en vez del pino.

## v3.7 — 2026-07-31

- El tablero se vuelve instalable como **PWA "SupplyBosquegin"** (manifest + service worker).

## v3.6 — 2026-07-31

- El desglose semanal se unifica al calendario gregoriano.

## v3.5 — 2026-07-31

- Checkbox "ya ingresó" en Proyección de abastecimiento.
- Fix de corrupción numpy/openpyxl.

## v3.4 — 2026-07-30

- Caché del cálculo de "última fecha de historial" de Contabilium.
- Fix de CORS del servidor cloud (faltaba `bosquegin.com`).
- Detección de sesión de Contabilium caducada con mensaje claro.
- Inventario Productos: filtro por producto + **desglose de stock por semana** (varias iteraciones: respeta Depósito y mes, tabla dinámica artículos×semanas, agrega Subrubro/Código y orden por encabezado, estima semanas faltantes, se reordena debajo de la tabla principal).
- Filtros superiores fijos (*sticky*) en todas las hojas que los usan; se corrige la franja vacía al hacer scroll.
- Se compacta la barra del título del header.
- Fix: el botón Actualizar podía quedar trabado; se corrige la barra de progreso.
- *Circuit breaker* para llamadas a CDP cuando no responden.
- Fixes en cervezas: no perder meses ya descargados si falla el refresco; traía mal el mes siguiente al cargar con anticipación.
- Modal de Actualizar con video real de botella sirviendo el vaso (con varios ajustes de recorte y scrub).
- Permitir ordenar por Estado en Inventario Productos.
- Se suman las cervezas (lata) a la Lista de Precios.

## v3.3 — 2026-07-21

- Proyección Producción: se quitan columnas de pedido sugerido de la grilla; nueva grilla resumen con detalle desplegable por producto; resumen de indicadores en una sola fila; proyección de abastecimiento editable; descripción de producto y venta promedio reales; stock igual al de Inventario Productos (Klozer+Oficina) en todos los trimestres; stock histórico correcto con promedios y tendencia de venta; venta real del mes en curso con tendencia graficada; saldo del mes en curso = stock actual − venta real; proyección mensual con saldo = stock − proyección.
- Ventas: "Salidas por rubro y subrubro" adopta el mismo formato que Detalle por artículo / Rotación por rubro.
- Compras: se quita KLOZER MKT de Alertas, Reposición sugerida y Stock completo; limpieza de hallazgos menores.
- Botón **Imprimir pedidos**: etiquetas y saldo del mes, descripción completa, encabezado por marca, todo en una sola hoja con dos cuadros compactos.
- Corrige el rubro Botánicos cayendo en Mobiliario por falta de tilde.
- Pipeline: cronómetro por paso con tabla de tiempos al final; más rápido cacheando Objetivo 2026 y sin pushear `bosquegin_data.js` en local.
- **Base histórica consolidada en SQLite** (`construir_base_historica.py`).
- Corrige el doble descuento de venta real en el saldo del mes en curso.
- Housekeeping de datos y fix del constructor de base histórica.
- Reorganización del **encabezado del tablero**: cascadas de productos + barril Feriado, layout robusto sin superposición.

## v3.2 — 2026-07-16

- Nueva pestaña **Proyección**: forecast trimestral de compra/venta.
- Unifica el formato de Rotación por rubro y subrubro.
- Rotación mensual por rubro: suma KLOZER MKT al stock (excepto Bebidas).
- Proyección: acordeón por trimestre, plegado de Objetivo cuando hay revisión; distribución de la cantidad a comprar por mes según venta objetivo; plantilla fija de 7 trimestres (2025 Q1–Q4, 2026 Q1–Q3); corrección del descubrimiento de trimestres en la hoja FORECAST.
- Proyección Producción: indicadores **Forecast Accuracy / Bias / MAPE**.

## v3.1 — 2026-07-13

- **Contabilium API en vivo.**
- Correcciones de salidas persistentes.
- Mejoras de rendimiento.
- Archivado en `Versiones/v3.1/`.

## v3.0 — 2026-06-29

- Fix: restaurar stock cierre 2025 en `data_stock_cierre.js`; preservar meses históricos 2025.
- La deduplicación pasa a aplicar solo a BOSQUE_SALIDAS, no a GC_REMITO — y luego se elimina del todo en `parse_ventas`: todas las filas del consolidado se procesan.
- Se rellenan nombres de productos sin descripción desde `PRODUCTOS.xlsx`.
- Tipografía "Heroes" (Cinzel Decorative) en el encabezado.
- **Cambio de arquitectura:** desglose semanal por producto, sin deduplicación.

## v2.29 — 2026-06-29

- Fix: filas duplicadas en Salidas (error de carga).
- Fix: costos y promedios con inicialización *lazy*.

## v2.28 — 2026-06-25

- Fix de lazy loading: renders correctos al cambiar de pestaña.

## v2.27 — 2026-06-25

- **Base de datos compactada por sección con lazy loading** — el tablero carga más rápido.

## v2.26 — 2026-06-25

- Gráficos en Lista de Precios: evolución, ranking, dispersión; fix del mes Dic'24.
- Fix: los gráficos ahora respetan los filtros activos.

## v2.25 — 2026-06-24

- Fix: costo de cervezas lata en inventario usa `costo_fab` directo ($/lata).
- Imágenes de productos en el header del dashboard (varias iteraciones de diseño: thumbnails cuadrados, layout con título centrado, recorte de botella portrait 2:3, fecha inline).
- Se agregan latas Temple Wolf IPA y Sin Alcohol al header.
- Fix: igualar tamaño de latas Temple.
- Rename de pestaña: Destilería → **"Costo Producción Bosque-Feriado-Temple"**.
- **Lista de Precios con vista mensual A+B+C** y reordenamiento de pestañas.

## v2.23 — 2026-06-23

- Se integra `update_stock_cierre_mes` en el flujo de actualización.
- Nueva hoja **Costo Destilería y Cervezas** desde RESUMEN PRODUCTOS X DESTILERIA, con cuadro evolutivo y Anualizado.
- Costos de cervezas (barril y lata) integrados en Destilería: card de latas, mapeo de códigos de inventario, cuadro evolutivo, detección dinámica de períodos, lectura de columna K en hojas mensuales vía descubrimiento de GIDs por gviz JSON.
- Fixes de robustez en cervezas: encabezado de tabla, búsqueda de columna "Costo Mes ACT" por nombre (no por índice fijo), primer match gana en lata, exploración de meses pre-GID para 2025.
- Se elimina el selector de Métrica en Destilería (estaba sin uso).

## v2.22 — 2026-06-17

- Gráfico de barras apiladas de meses de cobertura por rubro.

## v2.21 — 2026-06-17

- Reemplaza 3 gráficos por un área apilada de evolución de stock.

## v2.20 — 2026-06-17

- Gráficos de tendencias debajo de rotación mensual.

## v2.19 — 2026-06-17

- Modal de actualización con estética retro 80s.

## v2.18 — 2026-06-17

- Bebidas 2026 desde consolidado KLOZER+OFI + script de actualización.

## v2.15 – v2.17 — 2026-06-16

- Rotación por rubro: columna Subrubro inline (Cristalería, Combo, Mobiliario, luego Indumentaria); varias vueltas sobre la fila de encabezados al pie de cada grupo.
- Stock mensual: corrección del cálculo al cierre de cada mes; stock = Klozer + Oficina para todos los rubros; reemplazo de movimientos por remitos detallados.
- **v2.15:** stock real al cierre de mes desde snapshots diarios.
- **v2.16:** stock real por snapshot directo, sin estimación *backward*.
- **v2.17:** stock cierre de mes con datos reales de mayo/junio 2026.

## v2.14 — 2026-06-11

- Tabla de **rotación mensual por rubro**: meses de cobertura por mes histórico, con drill-down por rubro (luego simplificado a ventas + stock + meses), columnas ordenables con caché para re-ordenar, detalle en dos líneas por artículo.
- Stock mensual real: integración de movimientos Klozer 2025 y 2026 (ene–may).
- Rotación por rubro: pie de tabla aclarando depósitos por rubro.

## v2.13 — 2026-06-10

- "Supply Chain" se renombra a **Compras** + refactor de Reposición Sugerida; columna "En tránsito".
- Meses de stock con 1 decimal en Inventario Productos.
- Gran refactor de **Rotación por rubro y subrubro**: columnas Cód+Descripción combinadas, headers repetidos por grupo, orden A→Z, botón de impresión, páginas separadas por grupo, tipografía a 11px, stock solo Klozer para Bebidas, grupos separados para Vasos/Indumentaria/Combo, impresión A4 apaisada por grupo con thead repetido.

## v2.12 — 2026-06-08

- Tipografía general reducida a 9px (encabezados 14px+ sin cambios).
- **Costo como métrica de monto:** costUnit×un en KPIs, tabla mensual, semanal, gráfico y canales; labels Valor→Costo, Monto→Costo (con fix posterior de cuándo corresponde cada label).
- "OFI" → "OFICINA" en labels de display.
- Inventario: columna Costo unitario, se quita columna Fecha; reordenamiento y compactación de encabezados.
- Normalización "KLOZER_MKT" → "KLOZER MKT" en toda la sección de salidas.
- Fix de mínimos KLOZER MKT / Shop Gallery y botón Actualizar.
- Wake-up automático del servidor Render antes de llamar a `/actualizar`.
- Streaming real en servidor Render + timestamps con zona horaria AR.
- Fix: redondeo hacia arriba de mínimos y meses de stock en Inventario.
- Fix: descarga correcta de remitos GC hasta la fecha de hoy.
- **Sincronización automática de código con Google Drive.**
- "Inv. Intel" se fusiona en Inventario Productos.
- Modal de fecha+proveedor al aprobar compras.

## v2.10 — 2026-05-27

- Login modal client-side para GitHub Pages.
- Botón Actualizar visible para roles admin y editor.
- Actualización remota para editores vía **túnel Cloudflare**.
- Nace el **servidor Render** con integración a Google Drive.
- Migración de auth de Google Drive de service account a **OAuth2** refresh token.
- Deploy en Render.com; reemplazo de `googleapiclient` por HTTP stdlib.
- Cache-busting en `bosquegin_data.js` para siempre cargar datos frescos.
- "uds" → "un" en texto visible del dashboard.
- Multi-select de clientes con orden alfabético en Salidas.

## v1 — 2026-05-26

- **Deploy inicial** del Tablero Operativo Bosque Gin en GitHub Pages, con redirect de `index.html`.

---

### Nota sobre versiones sin entrada propia

`v2.1`–`v2.9`, `v2.11` y `v2.24` existen como snapshots en `Versiones/` pero no tienen un commit que identifique qué cambió puntualmente en cada una — sus cambios están repartidos entre las versiones vecinas de esta lista, agrupados por lo que sí quedó documentado en los mensajes de commit.

### Hoy (2026-08-05, sin version bump)

- `bosquegin_dashboard.html` renombrado a **`SupplyBosquegin.html`**, alineado con el nombre que ya usaban el manifest y el service worker. Actualizadas las referencias en `index.html`, `manifest.json`, `service-worker.js` (cache bump a v3), `_redirects`, `actualizar_bosquegin.py`, `servidor_bosquegin.py` e `iniciar_tablero_silencioso.vbs`.
