/* CheckappExcel - service worker
 * Strategia:
 *   - "cache-first" per risorse statiche del sito
 *   - "stale-while-revalidate" per le librerie CDN (xlsx, exceljs)
 *     -> usa sempre la cache se c'è, ma aggiorna in background
 *   - versionato tramite CACHE_NAME: aggiornando la versione il vecchio
 *     cache viene ripulito automaticamente
 */

const CACHE_VERSION = "v1.1.0";
const CACHE_STATIC  = `checkapp-static-${CACHE_VERSION}`;
const CACHE_CDN     = `checkapp-cdn-${CACHE_VERSION}`;

// File che vogliamo disponibili subito offline
const CORE_ASSETS = [
  "./",
  "./index.html",
  "./manifest.json",
  "./assets/favicon.svg",
  "./assets/apple-touch-icon.png",
  "./assets/icon-192.png",
  "./assets/icon-512.png",
  "./assets/og-image.png",
];

// CDN che consentiamo di cachare per l'uso offline
const CDN_HOSTS = new Set(["cdn.jsdelivr.net"]);

self.addEventListener("install", event => {
  event.waitUntil(
    caches.open(CACHE_STATIC).then(cache => cache.addAll(CORE_ASSETS))
      .then(() => self.skipWaiting())
  );
});

self.addEventListener("activate", event => {
  event.waitUntil(
    caches.keys().then(keys =>
      Promise.all(
        keys
          .filter(k => !k.endsWith(CACHE_VERSION))
          .map(k => caches.delete(k))
      )
    ).then(() => self.clients.claim())
  );
});

self.addEventListener("fetch", event => {
  const req = event.request;
  if (req.method !== "GET") return;

  const url = new URL(req.url);

  // CDN: stale-while-revalidate
  if (CDN_HOSTS.has(url.host)) {
    event.respondWith(staleWhileRevalidate(req, CACHE_CDN));
    return;
  }

  // Same-origin: cache-first, fallback network, poi cache della risposta
  if (url.origin === self.location.origin) {
    event.respondWith(cacheFirst(req, CACHE_STATIC));
    return;
  }
  // Tutto il resto passa direttamente alla rete
});

async function cacheFirst(request, cacheName) {
  const cache  = await caches.open(cacheName);
  const cached = await cache.match(request);
  if (cached) return cached;
  try {
    const response = await fetch(request);
    if (response && response.status === 200) {
      cache.put(request, response.clone());
    }
    return response;
  } catch (err) {
    // offline + nessuna cache -> prova a restituire la index per le navigazioni
    if (request.mode === "navigate") {
      const fallback = await cache.match("./index.html");
      if (fallback) return fallback;
    }
    throw err;
  }
}

async function staleWhileRevalidate(request, cacheName) {
  const cache  = await caches.open(cacheName);
  const cached = await cache.match(request);
  const networkPromise = fetch(request).then(response => {
    if (response && response.status === 200) {
      cache.put(request, response.clone());
    }
    return response;
  }).catch(() => cached);
  return cached || networkPromise;
}

// Permette a index.html di forzare l'aggiornamento immediato
self.addEventListener("message", event => {
  if (event.data === "SKIP_WAITING") self.skipWaiting();
});
