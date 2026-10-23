// غيّر رقم النسخة عند أي تعديل في ملفات التطبيق عشان الموبايل يحدّث
var VERSION = "majd-v3";
var FILES = ["./", "index.html", "style.css", "app.js", "manifest.webmanifest",
  "lib/xlsx.mini.min.js", "icons/icon-192.png", "icons/icon-512.png", "icons/apple-touch-icon.png", "icons/icon-maskable-512.png", "icons/logo.png"];

self.addEventListener("install", function (e) {
  e.waitUntil(caches.open(VERSION).then(function (c) { return c.addAll(FILES); }).then(function () { return self.skipWaiting(); }));
});
self.addEventListener("activate", function (e) {
  e.waitUntil(caches.keys().then(function (ks) {
    return Promise.all(ks.filter(function (k) { return k !== VERSION; }).map(function (k) { return caches.delete(k); }));
  }).then(function () { return self.clients.claim(); }));
});
self.addEventListener("fetch", function (e) {
  var req = e.request;
  if (req.method !== "GET") return;
  var url = new URL(req.url);
  if (url.origin !== location.origin || url.pathname.endsWith("data.xlsx")) return; // ملف البيانات دايماً من النت
  e.respondWith(
    fetch(req).then(function (res) {
      var copy = res.clone();
      caches.open(VERSION).then(function (c) { c.put(req, copy); });
      return res;
    }).catch(function () { return caches.match(req).then(function (r) { return r || caches.match("index.html"); }); })
  );
});
