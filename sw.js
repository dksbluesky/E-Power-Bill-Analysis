const CACHE_NAME = 'e-power-bills-shell-v1';
const APP_SHELL = ['./', './index.html', './ac_dashboard.html', './manifest-electricity.webmanifest', './manifest-air-conditioner.webmanifest', './icons/electricity-192.png', './icons/electricity-512.png', './icons/air-conditioner-192.png', './icons/air-conditioner-512.png'];
self.addEventListener('install', (event) => event.waitUntil(caches.open(CACHE_NAME).then((cache) => cache.addAll(APP_SHELL)).then(() => self.skipWaiting())));
self.addEventListener('activate', (event) => event.waitUntil(caches.keys().then((keys) => Promise.all(keys.filter((key) => key.startsWith('e-power-bills-') && key !== CACHE_NAME).map((key) => caches.delete(key)))).then(() => self.clients.claim())));
self.addEventListener('fetch', (event) => {
  const request = event.request;
  if (request.method !== 'GET' || new URL(request.url).origin !== self.location.origin) return;
  event.respondWith(caches.match(request).then((cached) => cached || fetch(request)));
});
