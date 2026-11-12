/* Service worker — guarda só os ARQUIVOS do app para abrir rápido.
   NUNCA guarda respostas do servidor (dados, login): elas sempre vão à internet.
   A cada mudança no app, aumente a VERSAO abaixo (ex.: v2 → v3). */
const VERSAO = 'hist-v1';
const ARQUIVOS = ['./', 'index.html', 'manifest.webmanifest', 'icons/icon-192.png', 'icons/icon-512.png', 'icons/apple-touch-icon.png'];
const CDN = ['https://cdnjs.cloudflare.com/ajax/libs/Chart.js/4.4.1/chart.umd.min.js', 'https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js'];

self.addEventListener('install', e => {
  e.waitUntil(caches.open(VERSAO).then(c => c.addAll(ARQUIVOS).then(() => Promise.all(CDN.map(u => c.add(u).catch(() => {}))))).then(() => self.skipWaiting()));
});
self.addEventListener('activate', e => {
  e.waitUntil(caches.keys().then(ks => Promise.all(ks.filter(k => k !== VERSAO).map(k => caches.delete(k)))).then(() => self.clients.claim()));
});
self.addEventListener('fetch', e => {
  const req = e.request; const url = new URL(req.url);
  // API e qualquer coisa que não seja GET: direto para a internet, sem cache
  if (req.method !== 'GET' || /script\.google(usercontent)?\.com$/.test(url.hostname) || url.hostname.endsWith('googleusercontent.com') || url.pathname.endsWith('/api')) return;
  // config.js: sempre da internet (é onde fica o endereço do servidor); sem internet, usa a cópia
  if (url.origin === location.origin && url.pathname.endsWith('/config.js')) {
    e.respondWith(fetch(req).then(r => { const c = r.clone(); caches.open(VERSAO).then(x => x.put(req, c)); return r; }).catch(() => caches.match(req)));
    return;
  }
  const meuSite = url.origin === location.origin; const cdn = CDN.includes(req.url);
  if (!meuSite && !cdn) return;
  // página: internet primeiro (pega a versão nova), cópia se estiver sem internet
  if (req.mode === 'navigate') {
    e.respondWith(fetch(req).then(r => { const c = r.clone(); caches.open(VERSAO).then(x => x.put('index.html', c)); return r; }).catch(() => caches.match('index.html')));
    return;
  }
  // demais arquivos do app e bibliotecas: cópia primeiro, internet se não tiver
  e.respondWith(caches.match(req).then(h => h || fetch(req).then(r => { if (r.ok) { const c = r.clone(); caches.open(VERSAO).then(x => x.put(req, c)); } return r; })));
});
