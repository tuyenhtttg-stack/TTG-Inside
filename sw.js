/* Service worker – Không Gian Số Nội Bộ (TTG Inside)
   Mỗi lần cập nhật file trên GitHub: tăng VERSION để máy cán bộ tải bản mới. */
const VERSION='2.1.0';
const CACHE='ttg-inside-'+VERSION;
const ASSETS=['./','index.html','manifest.webmanifest','lib/qrcode.js','lib/jsQR.js',
  'icons/icon-192.png','icons/icon-512.png','icons/icon-maskable-512.png','icons/apple-touch-icon.png','icons/favicon.png'];

self.addEventListener('install',e=>{e.waitUntil(caches.open(CACHE).then(c=>c.addAll(ASSETS)).then(()=>self.skipWaiting()))});
self.addEventListener('activate',e=>{e.waitUntil(caches.keys().then(ks=>Promise.all(ks.filter(k=>k.startsWith('ttg-inside-')&&k!==CACHE).map(k=>caches.delete(k)))).then(()=>self.clients.claim()))});
self.addEventListener('fetch',e=>{
  const req=e.request;if(req.method!=='GET'||new URL(req.url).origin!==location.origin)return;
  if(req.mode==='navigate'){ // trang chính: ưu tiên mạng để luôn có bản mới, mất mạng thì dùng bản đã lưu
    e.respondWith(fetch(req).then(r=>{const cp=r.clone();caches.open(CACHE).then(c=>c.put('index.html',cp));return r}).catch(()=>caches.match('index.html')));return;
  }
  e.respondWith(caches.match(req).then(hit=>hit||fetch(req).then(r=>{if(r.ok){const cp=r.clone();caches.open(CACHE).then(c=>c.put(req,cp))}return r})));
});

/* ===== Thông báo đẩy =====
   Tin đẩy không mang nội dung. Khi được đánh thức, service worker dùng "khóa máy" (lưu lúc cán bộ bật thông báo)
   hỏi máy chủ lấy nội dung rồi hiện thông báo. Không lấy được thì hiện câu chung. */
function kvGet(k){return new Promise(res=>{try{const q=indexedDB.open('ttg-inside',1);q.onupgradeneeded=()=>q.result.createObjectStore('kv');
  q.onsuccess=()=>{const r=q.result.transaction('kv').objectStore('kv').get(k);r.onsuccess=()=>res(r.result||null);r.onerror=()=>res(null)};q.onerror=()=>res(null)}catch(e){res(null)}})}
self.addEventListener('push',e=>{e.waitUntil((async()=>{
  let items=[];
  try{const c=await kvGet('push');
    if(c&&c.api&&c.key){const r=await fetch(c.api,{method:'POST',headers:{'Content-Type':'text/plain;charset=utf-8'},body:JSON.stringify({action:'tb_pull',key:c.key})});items=(await r.json()).items||[]}}catch(err){}
  if(!items.length)items=[{ma:'ttg-chung',tieuDe:'Không Gian Số Nội Bộ',ngan:'Có thông báo mới. Mở TTG Inside để xem.',view:''}];
  await Promise.all(items.map(x=>self.registration.showNotification(x.tieuDe,{body:x.ngan,tag:x.ma,icon:'icons/icon-192.png',data:{view:x.view||''}})));
})())});
self.addEventListener('notificationclick',e=>{
  const view=(e.notification.data&&e.notification.data.view)||'';e.notification.close();
  e.waitUntil(self.clients.matchAll({type:'window',includeUncontrolled:true}).then(cs=>{
    const c=cs.find(x=>new URL(x.url).pathname.startsWith(new URL(self.registration.scope).pathname));
    if(c){c.postMessage({ttgGo:view});return c.focus()}
    return self.clients.openWindow('./'+(view?'#'+view:''));
  }));
});
