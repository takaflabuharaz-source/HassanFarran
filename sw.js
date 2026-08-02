const CACHE_NAME = 'pwa-cache-v1';
const assetsToCache = [
  '/',
  '/index.html',
  // أضف هنا ملفات CSS أو JS الأساسية لديك
];

// مرحلة التثبيت والتخزين المؤقت
self.addEventListener('install', (event) => {
  event.waitUntil(
    caches.open(CACHE_NAME).then((cache) => {
      return cache.addAll(assetsToCache);
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
    })
  );
});

// التعامل مع طلبات الشبكة (مهم جداً لتفعيل الـ PWA)
self.addEventListener('fetch', (event) => {
  event.respondWith(
    caches.match(event.request).then((response) => {
      return response || fetch(event.request);
    })
  );
});
