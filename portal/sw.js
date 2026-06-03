const V = 'dbc-v4';  // 版號升級，強制清除舊快取
const SHELL = ['/dbc-base/portal/', '/dbc-base/portal/index.html'];

self.addEventListener('install', e => {
  e.waitUntil(caches.open(V).then(c => c.addAll(SHELL)));
  self.skipWaiting();
});

self.addEventListener('activate', e => {
  e.waitUntil(
    caches.keys().then(ks =>
      Promise.all(ks.filter(k => k !== V).map(k => caches.delete(k)))
    )
  );
  self.clients.claim();
});

self.addEventListener('fetch', e => {
  const url = e.request.url;
  if (url.includes('script.google.com') ||
      url.includes('googleapis.com') ||
      url.includes('accounts.google.com') ||
      url.includes('gstatic.com')) return;

  if (url.includes('/dbc-base/portal/')) {
    // Network first：永遠優先抓最新版，失敗才用快取
    e.respondWith(
      fetch(e.request)
        .then(res => {
          if (res.ok) caches.open(V).then(c => c.put(e.request, res.clone()));
          return res;
        })
        .catch(() => caches.match(e.request))
    );
  }
});
