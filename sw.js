const CACHE_NAME = 'pwa-cache-v2';

// مرحلة التثبيت والتخزين الآمن
self.addEventListener('install', (event) => {
  self.skipWaiting();
  event.waitUntil(
    caches.open(CACHE_NAME).then((cache) => {
      // استخدام addAll بحذر لتفادي توقف الخدمة إذا لم يجد ملفاً معيناً
      return cache.addAll(['./', './index.html', './manifest.json']).catch(err => {
        console.log('Cache Add Failed, bypassing initial cache error:', err);
      });
    })
  );
});

// مرحلة التفعيل وتنظيف التخزين القديم
self.addEventListener('activate', (event) => {
  event.waitUntil(
    caches.keys().then((cacheNames) => {
      return Promise.all(
        cacheNames.map((cache) => {
          if (cache !== CACHE_NAME) {
            return caches.delete(cache);
          }
        })
      );
    }).then(() => self.clients.claim())
  );
});

// التعامل مع طلبات الشبكة
self.addEventListener('fetch', (event) => {
  // تجنب تخزين طلبات البرمجة الخارجية مثل Google Apps Script بروابط POST
  if (event.request.method !== 'GET') return;

  event.respondWith(
    caches.match(event.request).then((cachedResponse) => {
      if (cachedResponse) {
        return cachedResponse;
      }
      return fetch(event.request).catch(() => {
        // العودة لصفحة index.html في حال انقطاع الاتصال
        return caches.match('./index.html');
      });
    })
  );
});
