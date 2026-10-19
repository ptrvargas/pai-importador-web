const CACHE='pai-importador-v1.1.0';
const FILES=['./','./index.html','./style.css','./app.js','./core.js','./storage.js','./install.js','./vendor/jszip.min.js','./manifest.webmanifest','./icon.svg','./icon-192.png','./icon-512.png'];
const URLS=new Set(FILES.map(path=>new URL(path,self.registration.scope).href));
self.addEventListener('install',event=>event.waitUntil(caches.open(CACHE).then(cache=>cache.addAll(FILES.map(path=>new Request(new URL(path,self.registration.scope),{cache:'reload'})))).then(()=>self.skipWaiting())));
self.addEventListener('activate',event=>event.waitUntil(caches.keys().then(keys=>Promise.all(keys.filter(k=>k.startsWith('pai-importador-')&&k!==CACHE).map(k=>caches.delete(k)))).then(()=>self.clients.claim())));
self.addEventListener('fetch',event=>{if(event.request.method!=='GET'||!URLS.has(event.request.url))return;event.respondWith(caches.open(CACHE).then(async cache=>(await cache.match(event.request))||fetch(event.request)));});
