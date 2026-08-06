// ─────────────────────────────────────────────────────────────────────────────
// gc_descargar_remitos.js
// Descarga Remitos Detallados desde Gestión Cervecera y los envía al relay local.
//
// Cómo usar:
//   1. Abrir cualquier página de GC en el browser (ej: la de Informes).
//   2. Abrir la consola del browser (F12 → Console).
//   3. Ajustar las fechas FECHA_DESDE / FECHA_HASTA si hace falta.
//   4. Pegar y ejecutar este script.
// ─────────────────────────────────────────────────────────────────────────────

// Fechas del mes actual (auto-calculadas; cambiar manualmente si es necesario)
const _hoy        = new Date();
const _dd         = String(_hoy.getDate()).padStart(2, '0');
const _mm         = String(_hoy.getMonth() + 1).padStart(2, '0');
const _yyyy       = _hoy.getFullYear();
const FECHA_DESDE = `01/${_mm}/${_yyyy}`;   // dd/mm/yyyy — primer día del mes actual
const FECHA_HASTA = `${_dd}/${_mm}/${_yyyy}`;   // dd/mm/yyyy — hoy
const RELAY_URL   = 'http://127.0.0.1:7893/save';
const SAVE_SUBDIR = 'Data/Salidas/GC';

(async () => {
  // ── 1. CSRF token ──────────────────────────────────────────────────────────
  const csrf = document.querySelector('[name="__RequestVerificationToken"]')?.value || '';
  if (!csrf) {
    console.error('[remitos] ✗ No se encontró __RequestVerificationToken en la página.');
    console.error('[remitos]   Asegurate de ejecutar este script desde una página de GC.');
    return;
  }

  // ── 2. POST para solicitar generación del Excel ────────────────────────────
  const params = new URLSearchParams({
    __RequestVerificationToken: csrf,
    formato:      'excel',
    filtro:       '',
    tieneGrafico: 'false',
    tieneRanking: 'false',
    id:           '',
    entidad:      '',
    fechaDesde:   FECHA_DESDE,
    fechaHasta:   FECHA_HASTA,
  });

  console.log(`[remitos] Solicitando informe ${FECHA_DESDE} → ${FECHA_HASTA}...`);

  let guid, fileName;
  try {
    const r1 = await fetch('/Informes/InformeRemitosDetallados', {
      method:  'POST',
      headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
      body:    params.toString(),
    });
    if (!r1.ok) throw new Error(`POST devolvió status ${r1.status}`);
    const j = await r1.json();
    if (j.message !== 'ok') throw new Error(`Respuesta inesperada: ${JSON.stringify(j)}`);
    guid     = j.FileGuid;
    fileName = j.FileName;
    console.log(`[remitos] ✓ Excel generado | GUID: ${guid} | Archivo: ${fileName}`);
  } catch (e) {
    console.error('[remitos] ✗ Error en POST de generación:', e.message);
    return;
  }

  // ── 3. GET para descargar el archivo Excel ─────────────────────────────────
  let arrayBuf;
  try {
    const url = `/Producto/ArchivoExcel?fileGuid=${encodeURIComponent(guid)}&filename=${encodeURIComponent(fileName)}`;
    const r2  = await fetch(url);
    if (!r2.ok) throw new Error(`GET devolvió status ${r2.status}`);
    arrayBuf = await r2.arrayBuffer();
    console.log(`[remitos] ✓ Descargado: ${arrayBuf.byteLength.toLocaleString('es-AR')} bytes`);
  } catch (e) {
    console.error('[remitos] ✗ Error descargando Excel:', e.message);
    return;
  }

  // ── 4. Convertir a base64 ──────────────────────────────────────────────────
  const bytes   = new Uint8Array(arrayBuf);
  let   binary  = '';
  const CHUNK   = 8192;
  for (let i = 0; i < bytes.length; i += CHUNK) {
    binary += String.fromCharCode(...bytes.subarray(i, i + CHUNK));
  }
  const b64 = btoa(binary);

  // ── 5. Nombre del archivo ──────────────────────────────────────────────────
  const dDesde   = FECHA_DESDE.replace(/\//g, '-');
  const dHasta   = FECHA_HASTA.replace(/\//g, '-');
  const saveName = `Remitos GC ${dDesde} ${dHasta}.xlsx`;

  // ── 6. Enviar al relay ─────────────────────────────────────────────────────
  console.log(`[remitos] Enviando a relay → ${SAVE_SUBDIR}/${saveName} ...`);
  try {
    const r3 = await fetch(RELAY_URL, {
      method:  'POST',
      headers: { 'Content-Type': 'application/json' },
      body:    JSON.stringify({ data: b64, name: saveName, subdir: SAVE_SUBDIR }),
    });
    const result = await r3.json();
    if (result.ok) {
      console.log(`[remitos] ✓ Guardado: ${result.path}  (${result.bytes.toLocaleString('es-AR')} bytes)`);
    } else {
      console.error('[remitos] ✗ Relay error:', result.error);
    }
  } catch (e) {
    console.error('[remitos] ✗ No se pudo conectar al relay (¿está corriendo?):', e.message);
    console.error('[remitos]   Iniciá el tablero con iniciar_tablero.bat y volvé a intentar.');
  }
})();
