/* KILL-SWITCH service worker.
   Eski keshni tozalaydi, o'zini o'chiradi va sahifani bir marta yangilaydi.
   Shundan keyin sahifa doim tarmoqdan (eng yangi versiya) yuklanadi — avtomatik reload/refresh bo'lmaydi. */
self.addEventListener('install', function () {
  self.skipWaiting();
});

self.addEventListener('activate', function (event) {
  event.waitUntil((async function () {
    try {
      var keys = await caches.keys();
      await Promise.all(keys.map(function (k) { return caches.delete(k); }));
    } catch (e) {}
    try { await self.registration.unregister(); } catch (e) {}
    try {
      var clients = await self.clients.matchAll({ type: 'window' });
      clients.forEach(function (c) { c.navigate(c.url); });
    } catch (e) {}
  })());
});

/* Hech narsani keshlamaymiz — barcha so'rovlar to'g'ridan-to'g'ri tarmoqqa ketadi */
self.addEventListener('fetch', function () { /* passthrough */ });
