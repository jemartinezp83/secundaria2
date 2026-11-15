/* Service worker: guarda la app en el iPad para usarla sin internet. Versión c2e9970872 */
const CACHE = 'sec2-c2e9970872';
const FILES = ['./', 'index.html', 'manifest.webmanifest', 'icon-180.png', 'icon-192.png', 'icon-512.png'];
self.addEventListener('install', e => { e.waitUntil(caches.open(CACHE).then(c => c.addAll(FILES)).then(() => self.skipWaiting())); });
self.addEventListener('activate', e => { e.waitUntil(caches.keys().then(ks => Promise.all(ks.filter(k => k.indexOf('sec2-') === 0 && k !== CACHE).map(k => caches.delete(k)))).then(() => self.clients.claim())); });
self.addEventListener('fetch', e => {
  if (e.request.method !== 'GET') return;
  /* La app de 3.° vive en /tercero/ con su propio service worker: no se intercepta aquí. */
  if (new URL(e.request.url).pathname.indexOf('/tercero/') >= 0) return;
  /* Red primero para la página (recibe actualizaciones si hay internet); si no hay red, usa la copia guardada. */
  if (e.request.mode === 'navigate') { e.respondWith(fetch(e.request).then(r => { const c = r.clone(); caches.open(CACHE).then(x => x.put('index.html', c)); return r; }).catch(() => caches.match('index.html'))); return; }
  e.respondWith(caches.match(e.request).then(r => r || fetch(e.request)));
});
