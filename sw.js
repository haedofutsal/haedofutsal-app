self.addEventListener('push', function(event) {
  if (event.data) {
    const data = event.data.json();
    const options = {
      body: data.body,
      icon: 'https://cdn-icons-png.flaticon.com/512/3075/3075908.png',
      badge: 'https://cdn-icons-png.flaticon.com/512/3075/3075908.png',
      vibrate: [100, 50, 100],
      data: {
        dateOfArrival: Date.now(),
        primaryKey: '2'
      }
    };
    event.waitUntil(
      self.registration.showNotification(data.title, options)
    );
  }
});

self.addEventListener('notificationclick', function(event) {
  event.notification.close();
  event.waitUntil(
    clients.matchAll({ type: 'window' }).then(windowClients => {
      // Check if there is already a window/tab open with the target URL
      for (var i = 0; i < windowClients.length; i++) {
        var client = windowClients[i];
        if (client.url === '/' && 'focus' in client) {
          return client.focus();
        }
      }
      if (clients.openWindow) {
        return clients.openWindow('/');
      }
    })
  );
});

// Interceptar Web Share Target (POST /share-receipt)
self.addEventListener('fetch', function(event) {
  const url = new URL(event.request.url);
  if (event.request.method === 'POST' && url.pathname === '/share-receipt') {
    event.respondWith((async () => {
      try {
        const formData = await event.request.formData();
        const file = formData.get('receipt');
        if (file) {
          const cache = await caches.open('receipt-share');
          const responseHeaders = new Headers({
            'Content-Type': file.type,
            'X-File-Name': encodeURIComponent(file.name || 'comprobante.png')
          });
          const fileResponse = new Response(file, { headers: responseHeaders });
          await cache.put('/shared-receipt-file', fileResponse);
        }
      } catch (err) {
        console.error('[SW] Error procesando share target:', err);
      }
      // Redirigir usando 303 (See Other) para que el navegador cambie a GET
      return Response.redirect('/?shared=1', 303);
    })());
  }
});

self.addEventListener('install', event => {
  self.skipWaiting();
});

self.addEventListener('activate', event => {
  event.waitUntil(self.clients.claim());
});


