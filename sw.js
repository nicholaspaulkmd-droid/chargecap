// sw.js — ChargeCap service worker
// Cache-first app shell so the app opens and works offline once installed.
// Bump CACHE_VERSION whenever app.js/styles.css/data.js change so clients
// pick up the new version instead of serving a stale cache forever.

const CACHE_VERSION = "chargecap-v26";
const APP_SHELL = [
  "./",
  "./index.html",
  "./styles.css",
  "./app.js",
  "./data.js",
  "./msal-browser.min.js",
  "./manifest.json",
  "./icons/icon-192.png",
  "./icons/icon-512.png",
];
const SHELL_PATHS = new Set(APP_SHELL.map((p) => new URL(p, self.location).pathname));

// Third-party libraries loaded from CDN — cached opportunistically so
// OCR/storage keep working offline after first load.
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
  const req = event.request;
  // v2.0: only GETs of our own app files and the CDN libraries are
  // cached. Everything else — Microsoft sign-in, Graph API calls, the
  // workbook template — always goes straight to the network. (v1.x
  // matched every URL by mistake: "./" became "", and every URL ends
  // with "", so API reads could have been answered from cache.)
  if (req.method !== "GET") return;
  const url = new URL(req.url);
  const isShell = url.origin === self.location.origin && SHELL_PATHS.has(url.pathname);
  if (!isShell && !CDN_HOSTS.includes(url.hostname)) return;

  event.respondWith(
    caches.match(req, { ignoreSearch: isShell }).then((cached) => {
      const fetchPromise = fetch(req)
        .then((networkResp) => {
          if (networkResp && networkResp.ok) {
            const copy = networkResp.clone();
            caches.open(CACHE_VERSION).then((cache) => cache.put(req, copy));
          }
          return networkResp;
        })
        .catch(() => cached);
      return cached || fetchPromise;
    })
  );
});
