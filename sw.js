const CACHE = 'prev-v27';
const ASSETS = [
  './', './index.html', './styles.css', './app.js', './version.js', './manifest.webmanifest',
  './vendor/zxing.min.js', './vendor/xlsx.full.min.js', './vendor/JsBarcode.all.min.js',
  './vendor/exceljs.min.js', './data/equipos.json',
  './icons/icon-192.png', './icons/icon-512.png'
];

self.addEventListener('install', (e) => {
  e.waitUntil(caches.open(CACHE).then((c) => c.addAll(ASSETS)).then(() => self.skipWaiting()));
});
self.addEventListener('activate', (e) => {
  e.waitUntil(caches.keys()
    .then((keys) => Promise.all(keys.filter((k) => k !== CACHE).map((k) => caches.delete(k))))
    .then(() => self.clients.claim()));
});
self.addEventListener('fetch', (e) => {
  if (e.request.method !== 'GET' || new URL(e.request.url).origin !== self.location.origin) return;
  e.respondWith(caches.match(e.request).then((c) => c || fetch(e.request).then((r) => {
    const copia = r.clone(); caches.open(CACHE).then((cc) => cc.put(e.request, copia)); return r;
  }).catch(() => caches.match('./index.html'))));
});
