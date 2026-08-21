/* ─────────────────────────────────────────────────────────────────────────
   PACE Room Tracker — minimal app-shell service worker.

   Scope is intentionally narrow: only caches this app's own static shell
   files (HTML/CSS/JS/manifest/logo) so the kiosk still opens if the iPad's
   WiFi drops momentarily. It never intercepts MSAL or Graph requests —
   those are cross-origin and simply pass through untouched, so sign-in and
   SharePoint sync always hit the real network (and correctly surface the
   "Unable to sync" retry screen when offline, per the app's design).
   ───────────────────────────────────────────────────────────────────────── */

const CACHE_NAME = "pace-tracker-shell-v5"; // bumped: PATCH 003 temporary diagnostic added
const SHELL_FILES = [
  "./",
  "./index.html",
  "./styles.css",
  "./config.js",
  "./demo-data.js",
  "./auth.js",
  "./graph.js",
  "./roster.js",
  "./pace-data.js",
  "./diagnostic.js", // TEMPORARY PATCH 003 DIAGNOSTIC — remove this line when retracted
  "./app.js",
  "./manifest.webmanifest",
  "./assets/iu29-logo.png"
];

self.addEventListener("install", event => {
  event.waitUntil(caches.open(CACHE_NAME).then(cache => cache.addAll(SHELL_FILES)));
  self.skipWaiting();
});

self.addEventListener("activate", event => {
  event.waitUntil(
    caches.keys().then(keys => Promise.all(keys.filter(k => k !== CACHE_NAME).map(k => caches.delete(k))))
  );
  self.clients.claim();
});

self.addEventListener("fetch", event => {
  const url = new URL(event.request.url);
  // Only ever handle this app's own same-origin GET requests.
  if (event.request.method !== "GET" || url.origin !== self.location.origin) return;

  event.respondWith(
    fetch(event.request)
      .then(resp => {
        const copy = resp.clone();
        caches.open(CACHE_NAME).then(cache => cache.put(event.request, copy));
        return resp;
      })
      .catch(() => caches.match(event.request))
  );
});
