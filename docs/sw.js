// Service Worker mínimo para habilitar instalación PWA
self.addEventListener('install', event => {
  self.skipWaiting();
});

self.addEventListener('activate', event => {
  event.waitUntil(self.clients.claim());
});

self.addEventListener('fetch', event => {
  // Pasa las solicitudes a la red
  event.respondWith(fetch(event.request).catch(() => new Response('Sin conexión')));
});
