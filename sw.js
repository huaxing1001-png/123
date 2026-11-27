const V = 'jinhe-v2', SHELL = ['./', 'index.html', 'manifest.webmanifest', 'icons/icon-192.png', 'icons/icon-512.png', 'icons/icon-maskable-512.png', 'icons/apple-touch-icon.png'];
self.addEventListener('install', e => { e.waitUntil(caches.open(V).then(c => c.addAll(SHELL)).then(() => self.skipWaiting())); });
self.addEventListener('activate', e => { e.waitUntil(caches.keys().then(k => Promise.all(k.filter(x => x !== V).map(x => caches.delete(x)))).then(() => self.clients.claim())); });
self.addEventListener('fetch', e => {
  const r = e.request; if (r.method !== 'GET' || !r.url.startsWith('http')) return;
  const same = new URL(r.url).origin === location.origin;
  if (r.mode === 'navigate' || (same && r.url.endsWith('index.html'))) {      // pages: network first, offline fallback
    e.respondWith(fetch(r).then(res => { const c = res.clone(); caches.open(V).then(x => x.put('index.html', c)); return res; }).catch(() => caches.match('index.html')));
    return;
  }
  e.respondWith(caches.match(r).then(hit => {                                 // assets & CDN libs/images: cache, refresh in background
    const net = fetch(r).then(res => { if (res.ok || res.type === 'opaque') { const c = res.clone(); caches.open(V).then(x => x.put(r, c)); } return res; }).catch(() => hit);
    return hit || net;
  }));
});
