// Service worker: guarda la app en el celu para que funcione sin internet.
// Con internet trae siempre la última versión (y la guarda); sin internet usa la guardada.
const VERSION = 'pharmapp-v4';
const FILES = [
  './', 'index.html', 'styles.css', 'app.js', 'util.js', 'db.js', 'excel.js', 'drive.js',
  'manifest.webmanifest', 'vendor/exceljs.min.js', 'vendor/jszip.min.js',
  'templates/rendicion.xlsx', 'templates/logo.jpeg',
  'icons/icon-192.png', 'icons/icon-512.png', 'icons/maskable-512.png', 'icons/apple-touch-icon.png'
];
self.addEventListener('install', e => {
  // cache: 'reload' evita copiar archivos viejos que el navegador tenga guardados
  e.waitUntil(caches.open(VERSION).then(c => c.addAll(FILES.map(f => new Request(f, { cache: 'reload' })))).then(() => self.skipWaiting()));
});
self.addEventListener('activate', e => {
  e.waitUntil(caches.keys().then(keys => Promise.all(keys.filter(k => k !== VERSION).map(k => caches.delete(k))))
    .then(() => self.clients.claim()));
});
self.addEventListener('fetch', e => {
  const req = e.request;
  if (req.method !== 'GET' || new URL(req.url).origin !== location.origin) return;
  e.respondWith((async () => {
    const cache = await caches.open(VERSION);
    try {
      const ctrl = new AbortController();
      const t = setTimeout(() => ctrl.abort(), 4000); // con poca señal no esperamos de más
      const res = await fetch(req, { cache: 'no-cache', signal: ctrl.signal });
      clearTimeout(t);
      if (res.ok) cache.put(req, res.clone());
      return res;
    } catch (err) {
      return (await cache.match(req, { ignoreSearch: true })) || (await cache.match('index.html'));
    }
  })());
});
