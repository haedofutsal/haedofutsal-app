// Versión del cache — incrementar para forzar actualización en todos los clientes
const CACHE_NAME = 'haedo-futsal-v10';

const ASSETS = [
  '/',
  '/Index.html',
  '/logo.png',
  '/manifest.json'
];

// Instala el nuevo service worker y pre-cachea los assets
self.addEventListener('install', event => {
  console.log('[SW] Instalando nueva versión:', CACHE_NAME);
  event.waitUntil(
    caches.open(CACHE_NAME).then(cache => {
      return cache.addAll(ASSETS);
    }).then(() => {
      // Activar inmediatamente sin esperar a que se cierren las pestañas viejas
      return self.skipWaiting();
    })
  );
});

// Al activarse, elimina todos los caches viejos
self.addEventListener('activate', event => {
  console.log('[SW] Activando y limpiando caches viejos...');
  event.waitUntil(
    caches.keys().then(cacheNames => {
      return Promise.all(
        cacheNames
          .filter(name => name !== CACHE_NAME)
          .map(name => {
            console.log('[SW] Borrando cache viejo:', name);
            return caches.delete(name);
          })
      );
    }).then(() => {
      // Tomar control de todas las pestañas abiertas inmediatamente
      return self.clients.claim();
    })
  );
});

// Estrategia: Network-first para Index.html, Cache-first para el resto
self.addEventListener('fetch', event => {
  const url = new URL(event.request.url);

  // Para Index.html, siempre intentar la red primero
  if (url.pathname === '/' || url.pathname === '/Index.html') {
    event.respondWith(
      fetch(event.request)
        .then(response => {
          // Si la respuesta es válida, actualizar el cache
          if (response && response.status === 200) {
            const responseClone = response.clone();
            caches.open(CACHE_NAME).then(cache => {
              cache.put(event.request, responseClone);
            });
          }
          return response;
        })
        .catch(() => {
          // Si no hay red, servir desde cache
          return caches.match(event.request);
        })
    );
    return;
  }

  // Para el resto: cache-first
  event.respondWith(
    caches.match(event.request).then(response => {
      return response || fetch(event.request);
    }).catch(() => {
      return caches.match('/Index.html');
    })
  );
});
