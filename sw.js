/* Offline-Cache für die App-Dateien (Daten kommen immer frisch aus OneDrive). */
const CACHE = "holzwirtschaft-iphone-1.0.0";
const SQLJS = "https://cdn.jsdelivr.net/npm/sql.js@1.10.3/dist/";
const FILES = ["./", "./index.html", "./app.js", "./store.js", "./config.js", "./manifest.webmanifest",
  "./icon-180.png", "./icon-192.png", "./icon-512.png", SQLJS + "sql-wasm.js", SQLJS + "sql-wasm.wasm"];
self.addEventListener("install", e => e.waitUntil(caches.open(CACHE).then(c => c.addAll(FILES)).then(() => self.skipWaiting())));
self.addEventListener("activate", e => e.waitUntil(caches.keys()
  .then(keys => Promise.all(keys.filter(k => k !== CACHE).map(k => caches.delete(k)))).then(() => self.clients.claim())));
self.addEventListener("fetch", e => {
  const url = e.request.url;
  if (e.request.method !== "GET") return;
  if (!(url.startsWith(self.registration.scope) || url.startsWith(SQLJS))) return;   // Login/Graph nie cachen
  // Erst Netz (Updates), sonst Cache
  e.respondWith(fetch(e.request).then(r => {
    if (r.ok && !url.includes("?code=")) { const copy = r.clone(); caches.open(CACHE).then(c => c.put(e.request, copy)); }
    return r;
  }).catch(() => caches.match(e.request, { ignoreSearch: true })));
});
