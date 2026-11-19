// 📶 Appli chauffeurs (roll.html) : s'ouvre même sans réseau.
// Le téléphone garde une copie de la page et des 2 bibliothèques (Supabase, lecteur de QR).
// Avec du réseau : toujours la DERNIÈRE version (réseau d'abord). Sans réseau : la copie gardée.
// Seule roll.html est concernée : les autres pages (bureau, GMAO, linge…) ne sont pas touchées.
const CACHE = 'roll-bdl-v1';
const LIBS = [
  'https://cdn.jsdelivr.net/npm/@supabase/supabase-js@2',
  'https://cdn.jsdelivr.net/npm/@zxing/library@0.21.3/umd/index.min.js'
];

self.addEventListener('install', e => {
  e.waitUntil(caches.open(CACHE).then(c => Promise.all(
    ['roll.html'].concat(LIBS).map(u => c.add(new Request(u, { cache: 'reload' })).catch(() => {}))
  )).then(() => self.skipWaiting()));
});

self.addEventListener('activate', e => {
  e.waitUntil(caches.keys().then(ks => Promise.all(ks.filter(k => k.startsWith('roll-bdl-') && k !== CACHE).map(k => caches.delete(k))))
    .then(() => self.clients.claim()));
});

self.addEventListener('fetch', e => {
  const req = e.request;
  if (req.method !== 'GET') return;
  const url = new URL(req.url);
  // la page des chauffeurs (avec ou sans ?r=R-xxx / ?c=client) : réseau d'abord, copie gardée sinon
  if (url.origin === self.location.origin && url.pathname.endsWith('/roll.html')) {
    e.respondWith(
      fetch(req).then(rep => {
        if (rep && rep.ok) { const copie = rep.clone(); caches.open(CACHE).then(c => c.put(new URL('roll.html', self.registration.scope).href, copie)); }
        return rep;
      }).catch(() => caches.match(new URL('roll.html', self.registration.scope).href))
    );
    return;
  }
  // bibliothèques : copie gardée d'abord (elles ne changent pas), réseau sinon
  if (LIBS.includes(req.url)) {
    e.respondWith(caches.match(req.url).then(m => m || fetch(req).then(rep => {
      if (rep && rep.ok) { const copie = rep.clone(); caches.open(CACHE).then(c => c.put(req.url, copie)); }
      return rep;
    })));
  }
  // tout le reste (serveur Supabase, Apps Script, autres pages) : on ne touche à rien
});
