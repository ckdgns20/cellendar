const CACHE = 'cellendar-v13';
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

self.addEventListener('message', event => {
  if (event.data?.type !== 'SHOW_NOTIFICATION') return;
  event.waitUntil(self.registration.showNotification(event.data.title || 'Cellendar', {
    body: event.data.body || '',
    icon: './icons/cellendar-icon.png',
    badge: './icons/cellendar-icon.png',
    tag: event.data.tag || 'cellendar-event',
    renotify: true,
    data: { url: './' },
  }));
});

self.addEventListener('notificationclick', event => {
  event.notification.close();
  event.waitUntil(clients.matchAll({ type: 'window', includeUncontrolled: true }).then(windows => {
    const existing = windows[0];
    return existing ? existing.focus() : clients.openWindow(event.notification.data?.url || './');
  }));
});
