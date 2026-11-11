/* Só o shell deste componente: nunca guarda páginas do Drive ou credenciais. */
const CACHE = "ps-visitas-shell-v1";
const FILES = ["./index.html", "./app.js", "./app.css"];
self.addEventListener("install", event => {
  event.waitUntil(caches.open(CACHE).then(cache => cache.addAll(FILES)).then(() => self.skipWaiting()));
});
self.addEventListener("activate", event => {
  event.waitUntil(caches.keys().then(keys => Promise.all(
    keys.filter(key => key.startsWith("ps-visitas-shell-") && key !== CACHE).map(key => caches.delete(key))
  )).then(() => self.clients.claim()));
});
self.addEventListener("fetch", event => {
  const url = new URL(event.request.url);
  if (event.request.method !== "GET" || !url.href.startsWith(self.registration.scope)) return;
  const filename = url.pathname.split("/").pop();
  if (!["", "index.html", "app.js", "app.css"].includes(filename)) return;
  const asset = new URL(filename || "index.html", self.registration.scope).href;
  event.respondWith(fetch(event.request).then(response => {
    if (response.ok) {
      const copy = response.clone();
      event.waitUntil(caches.open(CACHE).then(cache => cache.put(asset, copy)));
      return response;
    }
    return caches.match(asset).then(cached => cached || response);
  }).catch(() => caches.match(asset)));
});
