// Service worker: tiene in cache solo i file dell'app (mai i dati dei pazienti).
const CACHE = "gp-shell-v5";
const SHELL = ["./", "index.html", "styles.css", "app.js", "auth.js", "drive.js", "indexer.js", "docgen.js", "bmcdoc.js", "config.js",
  "vendor/msal-browser.min.js", "vendor/docx.umd.js", "manifest.webmanifest"];

self.addEventListener("install", (e) => { e.waitUntil(caches.open(CACHE).then((c) => c.addAll(SHELL)).then(() => self.skipWaiting())); });
self.addEventListener("activate", (e) => {
  e.waitUntil(caches.keys().then((ks) => Promise.all(ks.filter((k) => k !== CACHE).map((k) => caches.delete(k)))).then(() => self.clients.claim()));
});
// Stale-while-revalidate, solo per richieste della stessa origine
self.addEventListener("fetch", (e) => {
  const u = new URL(e.request.url);
  if (e.request.method !== "GET" || u.origin !== location.origin) return;
  e.respondWith(caches.open(CACHE).then(async (c) => {
    const hit = await c.match(e.request);
    const net = fetch(e.request).then((r) => { if (r.ok) c.put(e.request, r.clone()); return r; }).catch(() => hit);
    return hit || net;
  }));
});
