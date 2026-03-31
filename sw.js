/**
 * 家計簿 PWA — 静的ファイルを先読みキャッシュ（オフライン時はキャッシュ優先）
 * 利用条件: https または http://localhost（file:// では登録不可）
 */
const CACHE_NAME = "household-budget-pwa-v39";

const PRECACHE_FILES = [
  "index.html",
  "styles.css",
  "app.js",
  "star-lines.svg",
  "manifest.webmanifest",
  "icon-192.png",
  "icon-512.png",
  "apple-touch-icon.png",
  "icon.svg",
];

const PRECACHE_URLS = PRECACHE_FILES.map((path) => new URL(path, self.location).href);

self.addEventListener("install", (event) => {
  event.waitUntil(
    caches
      .open(CACHE_NAME)
      .then((cache) => cache.addAll(PRECACHE_URLS))
      .then(() => self.skipWaiting())
      .catch(() => {})
  );
});

self.addEventListener("activate", (event) => {
  event.waitUntil(
    caches
      .keys()
      .then((keys) =>
        Promise.all(
          keys.filter((k) => k !== CACHE_NAME).map((k) => caches.delete(k))
        )
      )
      .then(() => self.clients.claim())
  );
});

self.addEventListener("fetch", (event) => {
  if (event.request.method !== "GET") return;

  const url = new URL(event.request.url);
  if (url.origin !== self.location.origin) return;

  const indexHref = new URL("index.html", self.location).href;

  if (event.request.mode === "navigate") {
    event.respondWith(
      fetch(event.request).catch(() => caches.match(indexHref))
    );
    return;
  }

  event.respondWith(
    caches.match(event.request).then((cached) => {
      if (cached) return cached;
      return fetch(event.request).then((res) => {
        if (!res || res.status !== 200 || res.type !== "basic") return res;
        const copy = res.clone();
        caches.open(CACHE_NAME).then((cache) => cache.put(event.request, copy));
        return res;
      });
    })
  );
});
