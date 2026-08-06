/**
 * gc_descargar_via_browser.js
 * Pegar en la consola del browser desde cualquier página de GC
 * (https://www.gestioncervecera.com/Producto/Stock o /StockComp)
 *
 * Descarga 2 archivos de stock al relay server local (puerto 7893):
 *   1. Stock de Productos Simplificado*.xlsx  (envases por depósito)
 *   2. Stock Productos Comp.*.xlsx            (complementarios por depósito)
 *
 * Requiere que gc_relay_server.py esté corriendo.
 */
(async function() {
  const RELAY = 'http://127.0.0.1:7893/save';
  const log = (msg) => console.log('[GC-STOCK] ' + msg);

  async function exportAndSave(exportUrl, body, label) {
    log(`${label}: solicitando exportación...`);
    let ex;
    try {
      ex = await fetch(exportUrl, {
        method: 'POST',
        headers: { 'Content-Type': 'application/x-www-form-urlencoded', 'X-Requested-With': 'XMLHttpRequest' },
        body: body,
        credentials: 'include',
      }).then(r => r.json());
    } catch(e) { log(`${label}: ERROR exportación: ${e.message}`); return; }

    if (ex.message !== 'ok') { log(`${label}: ERROR: ${JSON.stringify(ex)}`); return; }
    log(`${label}: ${ex.FileName}`);

    let bytes;
    try {
      const p = new URLSearchParams({ fileGuid: ex.FileGuid, filename: ex.FileName });
      bytes = new Uint8Array(await fetch('/Producto/ArchivoExcel?' + p, { credentials: 'include' }).then(r => r.arrayBuffer()));
    } catch(e) { log(`${label}: ERROR descarga: ${e.message}`); return; }
    log(`${label}: ${bytes.length} bytes descargados`);

    let b64 = '';
    for (let i = 0; i < bytes.length; i += 8192)
      b64 += String.fromCharCode(...bytes.subarray(i, i + 8192));
    b64 = btoa(b64);

    try {
      const res = await fetch(RELAY, {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ data: b64, name: ex.FileName }),
      }).then(r => r.json());
      log(`${label}: ${res.ok ? '✓ guardado → ' + res.path : 'ERROR relay: ' + res.error}`);
    } catch(e) { log(`${label}: ERROR relay: ${e.message}`); }
  }

  // 1. Stock principal (envases por depósito)
  await exportAndSave(
    '/Producto/StockSimplificadoExcel',
    'idDeposito=&fechaEntregaEstimadaDesde=&fechaEntregaEstimadaHasta=&idDepositoEnvase=',
    'Stock'
  );

  // 2. Stock complementarios (insumos / merch por depósito)
  await exportAndSave(
    '/Producto/StockCompExcel',
    '',
    'StockComp'
  );

  log('Listo. Correr actualizar_bosquegin.py para actualizar el dashboard.');
})();
