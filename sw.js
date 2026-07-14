const CACHE_NAME = 'haedo-futsal-v36';

self.addEventListener('install', event => {
  console.log('[SW] Instalado.');
  self.skipWaiting();
});

self.addEventListener('activate', event => {
  console.log('[SW] Activado.');
  event.waitUntil(
    caches.keys().then(cacheNames => {
      return Promise.all(
        cacheNames.map(cache => {
          if (cache !== CACHE_NAME && cache !== 'receipt-share') {
            console.log('[SW] Borrando caché vieja:', cache);
            return caches.delete(cache);
          }
        })
      );
    }).then(() => self.clients.claim())
  );
});

self.addEventListener('fetch', event => {
  const url = new URL(event.request.url);

  // Interceptar la acción de compartir comprobante (Web Share Target)
  if (url.pathname === '/share-receipt' && event.request.method === 'POST') {
    event.respondWith((async () => {
      try {
        const formData = await event.request.formData();
        const file = formData.get('receipt');
        if (file) {
          console.log('[SW] Comprobante compartido recibido:', file.name, file.size);
          // Guardar el archivo en Cache Storage para que la app lo recupere al cargar
          const cache = await caches.open('receipt-share');
          const responseHeaders = new Headers({
            'Content-Type': file.type,
            'Content-Length': file.size.toString(),
            'X-File-Name': encodeURIComponent(file.name)
          });
          const fileResponse = new Response(file, { headers: responseHeaders });
          await cache.put('/shared-receipt-file', fileResponse);
          
          // Redirigir al cliente a la aplicación con el parámetro indicando que hay un compartido
          return Response.redirect('/?shared=receipt', 303);
        }
      } catch (err) {
        console.error('[SW] Error procesando archivo compartido:', err);
      }
      return Response.redirect('/', 303);
    })());
    return;
  }

  // Paso directo a la red (no caché) para evitar código obsoleto
  event.respondWith(fetch(event.request));
});
