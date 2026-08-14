const CACHE = 'cellendar-v10';
const ASSETS = [
  './', 'index.html', 'styles.css', 'config.js', 'storage.js', 'sync.js',
  'vendor/msal-browser.min.js',
  'app.js', 'manifest.webmanifest', 'icons/cellendar-icon.png'
];

self.addEventListener('install', event => {
  event.waitUntil(caches.open(CACHE).then(cache => cache.addAll(ASSETS)));
  self.skipWaiting();
});

self.addEventListener('activate', event => {
  event.waitUntil(
    caches.keys().then(keys => Promise.all(keys.filter(key => key !== CACHE).map(key => caches.delete(key))))
  );
  self.clients.claim();
});

self.addEventListener('fetch', event => {
  if (event.request.method !== 'GET') return;
  const url = new URL(event.request.url);
  const local = url.origin === self.location.origin;
  const shouldRefresh = event.request.mode === 'navigate' || /\.(?:js|css|webmanifest)$/.test(url.pathname);

  if (local && shouldRefresh) {
    event.respondWith(
      fetch(event.request)
        .then(response => {
          caches.open(CACHE).then(cache => cache.put(event.request, response.clone()));
          return response;
        })
        .catch(() => caches.match(event.request))
    );
    return;
  }

  event.respondWith(
    caches.match(event.request).then(cached => cached || fetch(event.request).then(response => {
      if (local) caches.open(CACHE).then(cache => cache.put(event.request, response.clone()));
      return response;
    }))
  );
});
