var CACHE = 'ft-v1';
var CORE = ['./', './index.html', './manifest.json', './icon.svg'];

self.addEventListener('install', function(e) {
  e.waitUntil(
    caches.open(CACHE).then(function(cache) {
      return cache.addAll(CORE);
    })
  );
  self.skipWaiting();
});

self.addEventListener('activate', function(e) {
  e.waitUntil(
    caches.keys().then(function(keys) {
      return Promise.all(
        keys.filter(function(k) { return k !== CACHE; }).map(function(k) { return caches.delete(k); })
      );
    })
  );
  self.clients.claim();
});

self.addEventListener('fetch', function(e) {
  if (e.request.method !== 'GET') return;

  var url = new URL(e.request.url);

  // Don't cache Google Apps Script API calls
  if (url.hostname.indexOf('script.google.com') !== -1 ||
      url.hostname.indexOf('script.googleusercontent.com') !== -1) {
    return;
  }

  e.respondWith(
    caches.match(e.request).then(function(cached) {
      var networkFetch = fetch(e.request).then(function(response) {
        if (response.ok) {
          var clone = response.clone();
          caches.open(CACHE).then(function(c) { c.put(e.request, clone); });
        }
        return response;
      }).catch(function() {
        if (e.request.mode === 'navigate') {
          return caches.match('./index.html');
        }
        return cached;
      });
      return cached || networkFetch;
    })
  );
});
