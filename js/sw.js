// Service worker: la app abre sin conexión. Los datos se sincronizan al volver la red.
const VERSION = 'mdc-v2';
const SHELL = ['./', 'index.html', 'css/app.css', 'js/app.js', 'js/config.js', 'js/store.js', 'js/planner.js', 'js/parser.js', 'js/outlook.js', 'js/summary.js', 'js/util.js', 'js/flows.js', 'js/seguimiento.js', 'js/xlsx.js', 'manifest.webmanifest', 'icons/icon.svg', 'icons/icon-192.png', 'icons/icon-512.png'];

self.addEventListener('install', (e) => {
  e.waitUntil(caches.open(VERSION).then((c) => c.addAll(SHELL)).then(() => self.skipWaiting()));
});
self.addEventListener('activate', (e) => {
  e.waitUntil(caches.keys().then((keys) => Promise.all(keys.filter((k) => k !== VERSION).map((k) => caches.delete(k)))).then(() => self.clients.claim()));
});
self.addEventListener('fetch', (e) => {
  const url = new URL(e.request.url);
  if (e.request.method !== 'GET') return;
  // Nunca cachear APIs (Supabase, Microsoft)
  if (/supabase\.co|microsoftonline\.com|graph\.microsoft\.com/.test(url.host)) return;
  // Librería de Supabase desde CDN: caché primero
  if (['cdn.jsdelivr.net', 'esm.sh', 'esm.run'].includes(url.host)) {
    e.respondWith(caches.open(VERSION).then(async (c) => (await c.match(e.request)) || fetch(e.request).then((r) => { c.put(e.request, r.clone()); return r; })));
    return;
  }
  if (url.origin !== location.origin) return;
  // Archivos de la app: red primero (para recibir actualizaciones), caché si no hay conexión
  e.respondWith(fetch(e.request).then((r) => { const copy = r.clone(); caches.open(VERSION).then((c) => c.put(e.request, copy)); return r; })
    .catch(() => caches.match(e.request, { ignoreSearch: true }).then((r) => r || caches.match('index.html'))));
});
