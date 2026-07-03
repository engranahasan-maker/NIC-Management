// NIC Management Software — minimal service worker
// This exists mainly to satisfy PWA installability requirements
// (camera/GPS/Supabase all keep working normally through the network — nothing is cached that would go stale).

const CACHE_NAME = "nic-shell-v1";

self.addEventListener("install", (event) => {
  self.skipWaiting();
});

self.addEventListener("activate", (event) => {
  event.waitUntil(
    caches.keys().then((keys) =>
      Promise.all(keys.filter((k) => k !== CACHE_NAME).map((k) => caches.delete(k)))
    )
  );
  self.clients.claim();
});

// Network-first: always try the network (so data is always fresh), only
// fall back to a very small shell cache if the device is fully offline.
self.addEventListener("fetch", (event) => {
  if (event.request.method !== "GET") return;
  event.respondWith(
    fetch(event.request).catch(() => caches.match(event.request))
  );
});
