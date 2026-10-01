// Кеширует оболочку приложения. Запросы к Apps Script не трогает — ими управляет app.js.
// При любом изменении файлов оболочки поднять VERSION, иначе телефоны останутся на старой версии.
const VERSION = 'v1.1.2';
const CACHE = 'budget-' + VERSION;
const SHELL = ['./', 'index.html', 'styles.css', 'app.js', 'manifest.webmanifest',
  'icons/icon-192.png', 'icons/icon-512.png', 'icons/apple-touch-icon.png'];

self.addEventListener('install', (e) => {
  // cache: 'reload' — мимо HTTP-кеша браузера (GitHub Pages держит файлы 10 минут),
  // иначе новая версия может собраться из старых файлов.
  const fresh = SHELL.map((u) => new Request(u, { cache: 'reload' }));
  e.waitUntil(caches.open(CACHE).then((c) => c.addAll(fresh)).then(() => self.skipWaiting()));
});

self.addEventListener('activate', (e) => {
  e.waitUntil(
    caches.keys()
      .then((keys) => Promise.all(keys.filter((k) => k !== CACHE).map((k) => caches.delete(k))))
      .then(() => self.clients.claim())
  );
});

self.addEventListener('fetch', (e) => {
  const url = new URL(e.request.url);
  if (e.request.method !== 'GET' || url.origin !== self.location.origin) return;
  // Сначала кеш — мгновенный запуск; в фоне обновляем кеш из сети.
  e.respondWith(
    caches.open(CACHE).then(async (cache) => {
      const cached = await cache.match(e.request, { ignoreSearch: true });
      const network = fetch(new Request(e.request.url, { cache: 'no-cache' }))
        .then((r) => { if (r.ok) cache.put(e.request, r.clone()); return r; })
        .catch(() => cached);
      return cached || network;
    })
  );
});
