// 설치(PWA)용 최소 서비스워커 — 캐시 안 함(항상 최신 버전), 같은 주소 요청만 통과
self.addEventListener('install', () => self.skipWaiting());
self.addEventListener('activate', e => e.waitUntil(self.clients.claim()));
self.addEventListener('fetch', e => {
  if (new URL(e.request.url).origin !== self.location.origin) return;
  e.respondWith(fetch(e.request));
});
