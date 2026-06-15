const CACHE_NAME = "expense-tracker-v2";
const STATIC_ASSETS = ["/static/styles.css", "/static/app.js", "/manifest.json"];

self.addEventListener("install", (event) => {
  event.waitUntil(caches.open(CACHE_NAME).then((cache) => cache.addAll(STATIC_ASSETS)).then(() => self.skipWaiting()));
});

self.addEventListener("activate", (event) => {
  event.waitUntil(caches.keys().then((keys) => Promise.all(keys.filter((key) => key !== CACHE_NAME).map((key) => caches.delete(key)))));
});

self.addEventListener("fetch", (event) => {
  const { request } = event;
  // For HTML navigation requests always go to the network so Flask serves fresh content.
  if (request.mode === "navigate" || request.headers.get("Accept")?.includes("text/html")) {
    event.respondWith(fetch(request));
    return;
  }
  // For static assets use cache-first.
  event.respondWith(caches.match(request).then((cached) => cached || fetch(request)));
});

