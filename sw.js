// Homebase service worker.
// Strategy:
//   - Navigation requests (HTML): network-first, cache fallback for offline.
//   - Same-origin assets (JSX, CSS, icons): stale-while-revalidate.
//   - Cross-origin (CDNs, APIs): stale-while-revalidate in a separate cache.

const VERSION = "v2";
const SHELL_CACHE = `homebase-shell-${VERSION}`;
const RUNTIME_CACHE = `homebase-runtime-${VERSION}`;

const SHELL_ASSETS = [
  "/",
  "/index.html",
  "/dashboard.css",
  "/dashboard.jsx",
  "/manifest.webmanifest",
  "/icons/icon.svg"
];

self.addEventListener("install", (event) => {
  event.waitUntil(
    caches.open(SHELL_CACHE).then((cache) =>
      cache.addAll(SHELL_ASSETS).catch(() => {})
    ).then(() => self.skipWaiting())
  );
});

self.addEventListener("activate", (event) => {
  event.waitUntil(
    caches.keys().then((keys) =>
      Promise.all(
        keys.filter((k) => k !== SHELL_CACHE && k !== RUNTIME_CACHE)
            .map((k) => caches.delete(k))
      )
    ).then(() => self.clients.claim())
  );
});

self.addEventListener("fetch", (event) => {
  const req = event.request;
  if (req.method !== "GET") return;
  const url = new URL(req.url);

  // Never cache API calls — they need to be fresh.
  if (
    url.hostname.endsWith("graph.microsoft.com") ||
    url.hostname.endsWith("login.microsoftonline.com") ||
    url.hostname.endsWith("api.anthropic.com") ||
    url.hostname.endsWith("newsapi.org") ||
    url.hostname.endsWith("open-meteo.com") ||
    url.hostname.endsWith("espn.com") ||
    url.hostname.endsWith("firebaseio.com")
  ) {
    return; // let it pass through
  }

  const isNavigation =
    req.mode === "navigate" ||
    (req.destination === "" && (req.headers.get("accept") || "").includes("text/html"));

  if (isNavigation) {
    event.respondWith(networkFirst(req, SHELL_CACHE));
    return;
  }

  if (url.origin === self.location.origin) {
    event.respondWith(staleWhileRevalidate(req, SHELL_CACHE));
    return;
  }

  event.respondWith(staleWhileRevalidate(req, RUNTIME_CACHE));
});

async function networkFirst(req, cacheName) {
  const cache = await caches.open(cacheName);
  try {
    const fresh = await fetch(req);
    if (fresh && fresh.status === 200) cache.put(req, fresh.clone());
    return fresh;
  } catch (e) {
    const cached = await cache.match(req);
    if (cached) return cached;
    return (await cache.match("/index.html")) || (await cache.match("/")) || Response.error();
  }
}

async function staleWhileRevalidate(req, cacheName) {
  const cache = await caches.open(cacheName);
  const cached = await cache.match(req);
  const network = fetch(req)
    .then((res) => {
      if (res && res.status === 200 && (res.type === "basic" || res.type === "cors")) {
        cache.put(req, res.clone());
      }
      return res;
    })
    .catch(() => cached);
  return cached || network;
}

self.addEventListener("message", (event) => {
  if (event.data === "skipWaiting") self.skipWaiting();
});
