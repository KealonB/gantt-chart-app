/* Gantt app-shell service worker. Bump CACHE when shipping a new version. */
var CACHE = "gantt-shell-v3";
var SHELL = [
  "./",
  "./index.html",
  "./styles.css",
  "./app.js",
  "./schedule.js",
  "./excel.js",
  "./vendor/xlsx.bundle.js",
  "./manifest.webmanifest",
  "./icons/icon-192.png",
  "./icons/icon-512.png",
  "./icons/icon-maskable-512.png",
  "./icons/apple-touch-icon.png",
];

self.addEventListener("install", function (event) {
  event.waitUntil(
    caches.open(CACHE).then(function (cache) {
      return cache.addAll(SHELL);
    })
  );
});

self.addEventListener("activate", function (event) {
  event.waitUntil(
    caches.keys().then(function (keys) {
      return Promise.all(
        keys
          .filter(function (k) {
            return k !== CACHE;
          })
          .map(function (k) {
            return caches.delete(k);
          })
      );
    }).then(function () {
      return self.clients.claim();
    })
  );
});

self.addEventListener("message", function (event) {
  if (event.data && event.data.type === "SKIP_WAITING") {
    self.skipWaiting();
  }
});

function isFontRequest(url) {
  return /fonts\.(googleapis|gstatic)\.com$/.test(url.hostname);
}

function offlineFallback(url) {
  if (isFontRequest(url) && /\.css($|\?)/.test(url.pathname + url.search)) {
    return new Response("/* offline font fallback */", {
      headers: { "Content-Type": "text/css; charset=utf-8" },
    });
  }
  return undefined;
}

self.addEventListener("fetch", function (event) {
  var request = event.request;
  if (request.method !== "GET") return;
  var url = new URL(request.url);
  var sameOrigin = url.origin === self.location.origin;
  if (!sameOrigin && !isFontRequest(url)) return;

  event.respondWith(
    caches.match(request).then(function (cached) {
      var fetched = fetch(request)
        .then(function (response) {
          if (response && response.ok) {
            var copy = response.clone();
            caches.open(CACHE).then(function (cache) {
              cache.put(request, copy);
            });
          }
          return response;
        })
        .catch(function () {
          return cached || offlineFallback(url);
        });
      // App shell: prefer cache so offline starts instantly; HTML/JS updates
      // arrive when a new service worker precaches a new CACHE name.
      if (sameOrigin) return cached || fetched;
      return fetched.then(function (response) {
        return response || cached || offlineFallback(url);
      });
    })
  );
});
