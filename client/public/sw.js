const CACHE_NAME = 'snaggi-v32';
const STATIC_ASSETS = [
  '/',
  '/manifest.json',
  '/icons/icon-192.png',
  '/icons/icon-512.png'
];

self.addEventListener('install', (event) => {
  event.waitUntil(
    caches.open(CACHE_NAME).then((cache) => cache.addAll(STATIC_ASSETS))
  );
  self.skipWaiting();
});

self.addEventListener('activate', (event) => {
  event.waitUntil(
    caches.keys().then((keys) =>
      Promise.all(keys.filter((k) => k !== CACHE_NAME).map((k) => caches.delete(k)))
    )
  );
  self.clients.claim();
});

self.addEventListener('fetch', (event) => {
  const { request } = event;
  // Skip API calls and non-GET requests
  if (request.url.includes('/api/') || request.method !== 'GET') return;

  const url = new URL(request.url);
  const isPage = request.mode === 'navigate' || url.pathname === '/' || url.pathname.endsWith('.html') || url.pathname === '/manifest.json';

  if (isPage) {
    // Network first: always load the newest app when online, so a deploy takes effect on
    // the next open/refresh. The cached copy is only used offline.
    event.respondWith(
      fetch(request).then((response) => {
        if (response.ok) {
          const clone = response.clone();
          caches.open(CACHE_NAME).then((cache) => cache.put(request, clone));
        }
        return response;
      }).catch(() => caches.match(request).then((c) => c || caches.match('/')))
    );
    return;
  }

  // Everything else (hashed /assets/ bundles, icons, fonts): cache first, they never change.
  event.respondWith(
    caches.match(request).then((cached) => cached || fetch(request).then((response) => {
      if (response.ok) {
        const clone = response.clone();
        caches.open(CACHE_NAME).then((cache) => cache.put(request, clone));
      }
      return response;
    }))
  );
});
