// sw.js — ChargeCap service worker
// Cache-first app shell so the app opens and works offline once installed.
// Bump CACHE_VERSION whenever app.js/styles.css/data.js change so clients
// pick up the new version instead of serving a stale cache forever.

const CACHE_VERSION = "chargecap-v2";
const APP_SHELL = [
  "./",
  "./index.html",
  "./styles.css",
  "./app.js",
  "./data.js",
  "./manifest.json",
  "./icons/icon-192.png",
  "./icons/icon-512.png",
];

// Third-party libraries loaded from CDN — cached opportunistically (see
// fetch handler below) so OCR/storage keep working offline after first load.
const CDN_HOSTS = ["unpkg.com"];

self.addEventListener("install", (event) => {
  event.waitUntil(
    caches.open(CACHE_VERSION).then((cache) => cache.addAll(APP_SHELL)).then(() => self.skipWaiting())
  );
});

self.addEventListener("activate", (event) => {
  event.waitUntil(
    caches.keys().then((keys) =>
      Promise.all(keys.filter((k) => k !== CACHE_VERSION).map((k) => caches.delete(k)))
    ).then(() => self.clients.claim())
  );
});

self.addEventListener("fetch", (event) => {
  const url = new URL(event.request.url);

  // Never cache Google APIs / auth — always go to network for these.
  if (url.hostname.includes("google")) return;

  // App shell + CDN libs: cache-first, falling back to network, and
  // updating the cache in the background (stale-while-revalidate).
  if (APP_SHELL.some((p) => event.request.url.endsWith(p.replace("./", ""))) || CDN_HOSTS.includes(url.hostname)) {
    event.respondWith(
      caches.match(event.request).then((cached) => {
        const fetchPromise = fetch(event.request)
          .then((networkResp) => {
            if (networkResp && networkResp.ok) {
              caches.open(CACHE_VERSION).then((cache) => cache.put(event.request, networkResp.clone()));
            }
            return networkResp;
          })
          .catch(() => cached);
        return cached || fetchPromise;
      })
    );
  }
});
