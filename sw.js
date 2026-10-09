///////////////////////////////////////////////////////////////////////////////////////////////
// Author: R-E Miller & Greg Macer
// Creation Date: August 20, 2026
// Filename: sw.js
// Purpose: Minimal network-first service worker that caches the PACE Room Tracker kiosk's
//          static app shell (HTML/CSS/JS/manifest/logo) after successful network reads, so
//          the kiosk still opens if the iPad's WiFi drops momentarily. Cross-origin MSAL and
//          Graph requests are intentionally left untouched, so sign-in and SharePoint sync
//          always hit the real network. Cache entries let the shell load offline, but
//          production data still requires the network — a cached UI must never be mistaken
//          for synchronized SharePoint data.
///////////////////////////////////////////////////////////////////////////////////////////////

const CACHE_NAME = "pace-tracker-shell-v18"; // STAFF CORRECTIONS PATCH: new visit-corrections.js + index.html/app.js/graph.js/pace-data.js/demo-data.js/styles.css changed
const SHELL_FILES = [
  "./",
  "./index.html",
  "./styles.css",
  "./config.js",
  "./demo-data.js",
  "./auth.js",
  // REVIEW: iep-app-users.js is loaded by index.html but is not currently in
  // this shell list; add it when the application-permission gate is enabled.
  "./graph.js",
  "./roster.js",
  "./recent-activity.js",
  "./visit-workflow.js",
  "./visit-corrections.js",
  "./student-select.js",
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
