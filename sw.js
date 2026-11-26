/* Service worker do KP CHECKLIST: deixa o app abrir sem internet
   e envia a fila em segundo plano quando o sinal voltar. */
importScripts("config.js", "sync.js");
const CACHE = "kp-checklist-v7";
const ARQUIVOS = ["./", "index.html", "config.js", "sync.js", "manifest.webmanifest", "icon-192.png", "icon-512.png", "logo.jpg", "fundo.jpg"];

self.addEventListener("install", e => {
  e.waitUntil(caches.open(CACHE).then(c => c.addAll(ARQUIVOS)).then(() => self.skipWaiting()));
});
self.addEventListener("activate", e => {
  e.waitUntil(caches.keys()
    .then(ks => Promise.all(ks.filter(k => k !== CACHE).map(k => caches.delete(k))))
    .then(() => self.clients.claim()));
});
self.addEventListener("fetch", e => {
  const u = new URL(e.request.url);
  if (e.request.method !== "GET" || u.origin !== location.origin) return;
  // tenta a rede (para pegar atualizações) por até 4 s; sem sinal, usa a cópia guardada
  e.respondWith((async () => {
    const cache = await caches.open(CACHE);
    try {
      const r = await Promise.race([fetch(e.request), new Promise((_, no) => setTimeout(() => no(), 4000))]);
      if (r && r.ok) cache.put(e.request, r.clone());
      return r;
    } catch (_) {
      return (await cache.match(e.request, { ignoreSearch: true })) || (await cache.match("index.html")) || Response.error();
    }
  })());
});
self.addEventListener("sync", e => {
  if (e.tag === "kp-enviar") {
    e.waitUntil(KPFila.enviarTudo().then(r => { if (r.pendentes && r.erro === "rede") throw new Error("tentar de novo"); }));
  }
});
self.addEventListener("message", e => {
  if (e.data === "enviar") e.waitUntil(KPFila.enviarTudo());
});
