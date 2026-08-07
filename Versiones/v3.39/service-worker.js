// SupplyBosquegin — cachea la "cáscara" de la app (HTML/manifest/íconos) para
// que instale como PWA y abra rápido; los datos en vivo (data_*.js) siempre
// se piden a la red primero y nunca se guardan en caché (deben estar frescos).
const CACHE = 'supplybosquegin-v3';
const PRECACHE_URLS = [
  'SupplyBosquegin.html',
  'manifest.json',
  'icons/icon-192.png',
  'icons/icon-512.png',
];

self.addEventListener('install', (event) => {
  event.waitUntil(
    caches.open(CACHE)
      .then((cache) => cache.addAll(
        PRECACHE_URLS.map((u) => new Request(u, { cache: 'reload' }))
      ))
      .then(() => self.skipWaiting())
  );
});

self.addEventListener('activate', (event) => {
  event.waitUntil(
    caches.keys()
      .then((keys) => Promise.all(keys.filter((k) => k !== CACHE).map((k) => caches.delete(k))))
      .then(() => self.clients.claim())
  );
});

const DATA_FILE_RE = /\/data_[a-z0-9_]+\.js$/i;

self.addEventListener('fetch', (event) => {
  const req = event.request;
  if (req.method !== 'GET') return;

  const url = new URL(req.url);
  if (url.origin !== location.origin) return;

  // "no-store" fuerza a ignorar la caché HTTP del navegador (GitHub Pages
  // manda Cache-Control: max-age=600 en el HTML) — sin esto, un simple
  // fetch() de red "primero" podia devolver igual una copia vieja servida
  // por la caché HTTP local, dejando el shell pegado en la version anterior
  // hasta que venciera ese max-age (bug real detectado 2026-07-31).
  const freshReq = new Request(req, { cache: 'no-store' });

  // data_*.js: siempre red primero, sin persistir en caché (evita servir
  // datos viejos y evita que la caché crezca sin límite por el cache-busting).
  if (DATA_FILE_RE.test(url.pathname)) {
    event.respondWith(fetch(freshReq).catch(() => caches.match(req)));
    return;
  }

  // Resto de archivos estáticos (shell, íconos, manifest): red primero,
  // guardando la última copia buena para poder servir offline.
  event.respondWith(
    fetch(freshReq)
      .then((res) => {
        const copy = res.clone();
        caches.open(CACHE).then((cache) => cache.put(req, copy));
        return res;
      })
      .catch(() => caches.match(req))
  );
});
