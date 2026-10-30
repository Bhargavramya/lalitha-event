// Lalitha Quotation Maker – offline cache. Change VERSION on every release.
const VERSION = 'led-beta-13.5';
importScripts('https://cdn.onesignal.com/sdks/web/v16/OneSignalSDK.sw.js');
const CORE = ['./', 'index.html', 'c.html', 'manifest.json', 'icon-180.png', 'icon-192.png', 'icon-512.png'];
self.addEventListener('install', e => { e.waitUntil(caches.open(VERSION).then(c => c.addAll(CORE)).catch(() => {}).then(() => self.skipWaiting())); });
self.addEventListener('activate', e => { e.waitUntil(caches.keys().then(ks => Promise.all(ks.filter(k => k !== VERSION).map(k => caches.delete(k)))).then(() => self.clients.claim())); });
self.addEventListener('fetch', e => {
  const r = e.request, u = new URL(r.url);
  if (r.method !== 'GET') return;                                   // sync POSTs go straight to Google
  if (u.hostname.endsWith('google.com') || u.hostname.endsWith('googleusercontent.com')) return; // never cache Sheet data
  const fonts = /fonts\.(googleapis|gstatic)\.com|cdnjs\.cloudflare\.com/.test(u.hostname);
  if (u.origin !== location.origin && !fonts) return;
  e.respondWith(caches.open(VERSION).then(async c => {
    if (r.mode === 'navigate' || /\.html$|\/$|\/c$/.test(u.pathname)) {          // pages: newest first, cache only when offline
      try { const res = await fetch(r, { cache: 'no-store' }); if (res && res.ok) c.put(r, res.clone()); return res; }
      catch (x) { return (await c.match(r, { ignoreSearch: true })) || (await c.match('index.html')) || new Response('Offline', { status: 503 }); }
    }
    const hit = await c.match(r, { ignoreSearch: u.origin === location.origin });
    const net = fetch(r).then(res => { if (res && (res.ok || res.type === 'opaque')) c.put(r, res.clone()); return res; }).catch(() => null);
    if (hit) { e.waitUntil(net); return hit; }
    return (await net) || (r.mode === 'navigate' && !/\/c(\.html)?$/.test(u.pathname) ? await c.match('index.html') : null) || new Response('Offline', { status: 503 });
  }));
});
