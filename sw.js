/* Plantes mères — Amaltis : l'écran s'ouvre même sans réseau.
   Réseau d'abord (toujours la dernière version publiée) ; après 4 s ou en cas de coupure, la copie gardée sur le téléphone.
   Les données ne passent jamais par ici : elles viennent du script Google à chaque ouverture. */
const CACHE = 'meres-v3';
const FICHIERS = ['./', './index.html', './manifest.webmanifest', './icone-180.png', './icone-192.png', './icone-512.png'];

self.addEventListener('install', e => {
  e.waitUntil(caches.open(CACHE).then(c => c.addAll(FICHIERS)).then(() => self.skipWaiting()));
});
self.addEventListener('activate', e => {
  e.waitUntil(caches.keys().then(ks => Promise.all(ks.filter(k => k.startsWith('meres-') && k !== CACHE).map(k => caches.delete(k)))).then(() => self.clients.claim()));
});
self.addEventListener('fetch', e => {
  const u = new URL(e.request.url);
  if (e.request.method !== 'GET' || u.origin !== self.location.origin) return;  // le script Google n'est jamais intercepté
  const copie = () => caches.match(e.request, { ignoreSearch: true }).then(r => r || caches.match('./index.html'));
  const reseau = fetch(e.request).then(r => {
    if (r.ok) { const c = r.clone(); caches.open(CACHE).then(ca => ca.put(e.request, c)); }
    return r;
  });
  const delai = new Promise(ok => setTimeout(ok, 4000)).then(copie);
  e.respondWith(Promise.race([reseau.catch(copie), delai.then(r => r || reseau)]));
});
