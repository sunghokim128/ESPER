const CACHE_NAME = 'esper-app-v139';
const CACHEABLE_ORIGINS = [self.location.origin];

self.addEventListener('install', (event) => {
  event.waitUntil(self.skipWaiting());
});

self.addEventListener('activate', (event) => {
  event.waitUntil((async () => {
    const cachesNames = await caches.keys();
    await Promise.all(
      cachesNames
        .filter((name) => name !== CACHE_NAME)
        .map((name) => caches.delete(name))
    );
    await self.clients.claim();
  })());
});

self.addEventListener('fetch', (event) => {
  const request = event.request;
  if (request.method !== 'GET') return;

  const requestUrl = new URL(request.url);
  if (!CACHEABLE_ORIGINS.includes(requestUrl.origin)) return;

  const isAppResource = request.destination === 'document'
    || request.destination === 'script'
    || request.destination === 'style'
    || requestUrl.pathname.endsWith('.html')
    || requestUrl.pathname.endsWith('.js')
    || requestUrl.pathname.endsWith('.css');

  if (!isAppResource) return;

  event.respondWith((async () => {
    const cache = await caches.open(CACHE_NAME);
    try {
      const response = await fetch(request, { cache: 'no-store' });
      if (response.ok) await cache.put(request, response.clone());
      return response;
    } catch {
      const cached = await cache.match(request);
      if (cached) return cached;
      throw new Error('ESPER is unavailable and no cached copy exists.');
    }
  })());
});
