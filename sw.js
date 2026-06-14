/* FinanceOS Service Worker
   Strategy:
   - App shell (index.html, manifest, icons): cache-first
   - Google Sheet CSV fetches (docs.google.com): network-first, fall back to cache
   - Tells open pages when a fetch was served from cache so they can show
     the "Offline — showing last data" banner.
*/
const CACHE_NAME = 'financeos-v3'; // bump on app updates so old caches are dropped

const APP_SHELL = [
  './',
  './index.html',
  './manifest.json',
  './icon-192.svg',
  './icon-512.svg',
  // Photo-theme scenes (bundled so the theme works fully offline)
  './scene-peaks-day.svg', './scene-peaks-night.svg',
  './scene-shore-day.svg', './scene-shore-night.svg',
  './scene-pines-day.svg', './scene-pines-night.svg',
  './scene-skyline-day.svg', './scene-skyline-night.svg'
];

self.addEventListener('install', (event) => {
  event.waitUntil(
    caches.open(CACHE_NAME)
      .then((cache) => cache.addAll(APP_SHELL))
      .catch(() => { /* partial cache is fine — never block install */ })
      .then(() => self.skipWaiting())
  );
});

self.addEventListener('activate', (event) => {
  event.waitUntil(
    caches.keys()
      .then((keys) => Promise.all(
        keys.filter((k) => k !== CACHE_NAME).map((k) => caches.delete(k))
      ))
      .then(() => self.clients.claim())
  );
});

function isSheetRequest(url) {
  return url.hostname === 'docs.google.com' && url.pathname.includes('/gviz/');
}

async function notifyClients(message) {
  const clients = await self.clients.matchAll({ includeUncontrolled: true });
  clients.forEach((c) => c.postMessage(message));
}

self.addEventListener('fetch', (event) => {
  const url = new URL(event.request.url);
  if (event.request.method !== 'GET') return;

  // Network-first for Sheet CSV data
  if (isSheetRequest(url)) {
    event.respondWith(
      fetch(event.request)
        .then((response) => {
          if (response && response.ok) {
            const copy = response.clone();
            caches.open(CACHE_NAME).then((cache) => cache.put(event.request, copy));
          }
          return response;
        })
        .catch(async () => {
          const cached = await caches.match(event.request);
          if (cached) {
            notifyClients({ type: 'OFFLINE_DATA', message: 'Offline — showing last data' });
            return cached;
          }
          return new Response('', { status: 503, statusText: 'Offline' });
        })
    );
    return;
  }

  // Navigations (the app itself): network-first so updates arrive,
  // cached shell when offline
  if (event.request.mode === 'navigate') {
    event.respondWith(
      fetch(event.request)
        .then((response) => {
          if (response && response.ok) {
            const copy = response.clone();
            caches.open(CACHE_NAME).then((cache) => {
              cache.put(event.request, copy.clone());
              cache.put('./index.html', copy);
            });
          }
          return response;
        })
        .catch(async () => {
          const shell = (await caches.match(event.request)) || (await caches.match('./index.html'));
          if (shell) {
            notifyClients({ type: 'OFFLINE_DATA', message: 'Offline — showing last data' });
            return shell;
          }
          return new Response('', { status: 503, statusText: 'Offline' });
        })
    );
    return;
  }

  // Cache-first for static assets (icons, manifest, fonts)
  event.respondWith(
    caches.match(event.request).then((cached) => {
      if (cached) return cached;
      return fetch(event.request)
        .then((response) => {
          if (response && response.ok && url.origin === self.location.origin) {
            const copy = response.clone();
            caches.open(CACHE_NAME).then((cache) => cache.put(event.request, copy));
          }
          return response;
        })
        .catch(() => new Response('', { status: 503, statusText: 'Offline' }));
    })
  );
});

/* Background refresh of Sheet data.
   Uses Periodic Background Sync where supported (registered by the page);
   pages also run their own 30-min timer as a universal fallback. */
self.addEventListener('periodicsync', (event) => {
  if (event.tag === 'financeos-refresh') {
    event.waitUntil(refreshSheetCache());
  }
});

self.addEventListener('message', (event) => {
  if (event.data && event.data.type === 'REFRESH_SHEETS') {
    event.waitUntil ? null : null;
    refreshSheetCache();
  }
});

async function refreshSheetCache() {
  try {
    const cache = await caches.open(CACHE_NAME);
    const requests = await cache.keys();
    const sheetReqs = requests.filter((r) => isSheetRequest(new URL(r.url)));
    await Promise.all(sheetReqs.map(async (req) => {
      try {
        const fresh = await fetch(req);
        if (fresh && fresh.ok) await cache.put(req, fresh);
      } catch (e) { /* offline — keep old copy */ }
    }));
    notifyClients({ type: 'SHEETS_REFRESHED', at: Date.now() });
  } catch (e) { /* ignore */ }
}
