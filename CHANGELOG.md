# Changelog — SupplyBosquegin

Historial de cambios del Tablero Operativo, reconstruido desde `git log` y mantenido a partir de ahora versión por versión. Orden: más reciente primero.

> **Cómo se actualiza este archivo:** cada vez que sumemos cambios en una sesión, se agregan como bullets bajo **[Sin versionar]**. Cuando se hace un bump de versión (`vX.Y`), esos bullets pasan a formar la sección de esa versión con su fecha.

---

## [Sin versionar]

_(sin cambios pendientes de versión todavía)_

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
