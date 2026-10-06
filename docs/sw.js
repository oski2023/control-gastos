// Service Worker v75 para PWA Control de Gastos
self.addEventListener('install', event => {
  self.skipWaiting();
});

self.addEventListener('activate', event => {
  event.waitUntil(
    caches.keys().then(keys => Promise.all(
      keys.map(key => caches.delete(key))
    )).then(() => self.clients.claim())
  );
});

self.addEventListener('fetch', event => {
  // Para navegación y HTML, siempre intentar red fresca sin usar caché HTTP antiguo
  if (event.request.mode === 'navigate' || event.request.destination === 'document') {
    event.respondWith(
      fetch(event.request, { cache: 'no-store' }).catch(() => {
        return caches.match(event.request).then(cached => cached || new Response('Sin conexión'));
      })
    );
    return;
  }
  event.respondWith(fetch(event.request).catch(() => new Response('Sin conexión')));
});
