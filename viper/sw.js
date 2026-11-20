// The app moved to ../sidewinder/. This worker replaces the old offline copy of Viper: it deletes the old
// cached files and serves nothing from cache, so old Home Screen icons show the "moved" page.
self.addEventListener("install", () => self.skipWaiting());

self.addEventListener("activate", e => {
  e.waitUntil(caches.keys()
    .then(keys => Promise.all(keys.filter(k => k.startsWith("viper-")).map(k => caches.delete(k))))
    .then(() => self.clients.claim()));
});
