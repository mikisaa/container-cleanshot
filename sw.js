const CACHE_NAME = 'container-cleanshot-v10';
const ASSETS_TO_CACHE = [
  './',
  './index.html',
  './manifest.json',
  './app-icon.jpg',
  './container_images/01_back_wall.png',
  './container_images/02_left_back.png',
  './container_images/03_right_back.png',
  './container_images/04_ceiling_back.png',
  './container_images/05_floor_back.png',
  './container_images/06_left_front.png',
  './container_images/07_right_front.png',
  './container_images/08_ceiling_front.png',
  './container_images/09_floor_front.png',
  './container_images/10_container_no.png',
  './container_images/11_before_full.png',
  './container_images/12_after_full.png'
];

// インストール時にアセットをキャッシュ
self.addEventListener('install', (event) => {
  event.waitUntil(
    caches.open(CACHE_NAME).then((cache) => {
      console.log('[Service Worker] アセットをキャッシュ中...');
      // 存在しないアセットがあった場合にもインストール全体を失敗させないよう、1つずつキャッシュ
      return Promise.allSettled(
        ASSETS_TO_CACHE.map((asset) => {
          return cache.add(asset)
            .then(() => console.log(`[Service Worker] キャッシュ成功: ${asset}`))
            .catch(err => console.error(`[Service Worker] キャッシュ失敗: ${asset}`, err));
        })
      );
    })
  );
  self.skipWaiting();
});

// 古いキャッシュのクリア
self.addEventListener('activate', (event) => {
  event.waitUntil(
    caches.keys().then((cacheNames) => {
      return Promise.all(
        cacheNames.map((cacheName) => {
          if (cacheName !== CACHE_NAME) {
            console.log('[Service Worker] 古いキャッシュを削除:', cacheName);
            return caches.delete(cacheName);
          }
        })
      );
    })
  );
  self.clients.claim();
});

// フェッチハンドラ (Cache-First, GAS APIは除外)
self.addEventListener('fetch', (event) => {
  // GAS APIへのPOSTリクエストや外部ドメインはキャッシュから除外
  if (event.request.method !== 'GET' || event.request.url.includes('script.google.com')) {
    return;
  }

  event.respondWith(
    caches.match(event.request).then((cachedResponse) => {
      if (cachedResponse) {
        return cachedResponse;
      }
      return fetch(event.request).then((networkResponse) => {
        // レスポンスが正常なGETリクエストなら動的にキャッシュに追加
        if (networkResponse && networkResponse.status === 200 && networkResponse.type === 'basic') {
          const responseToCache = networkResponse.clone();
          caches.open(CACHE_NAME).then((cache) => {
            cache.put(event.request, responseToCache);
          });
        }
        return networkResponse;
      }).catch((err) => {
        console.warn('[Service Worker] オフラインかつキャッシュにありません:', event.request.url);
        // 代替レスポンスなし
        return new Response('Offline', { status: 503, statusText: 'Offline' });
      });
    })
  );
});
