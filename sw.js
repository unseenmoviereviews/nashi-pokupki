// Service worker «Наших покупочек»: хранит оформление сайта, чтобы приложение
// быстро открывалось и без интернета показывало экран, а не ошибку браузера.
// Данные (покупки, фото) идут через API Google и здесь не кэшируются.
// Ещё он показывает пуш-уведомления.
const CACHE = 'pokupochki-v2';
// Тот же адрес API, что в index.html.
const API_URL = 'https://script.google.com/macros/s/AKfycbyODiIKQtRk3uT9ESsm1jVkgl5u2OUj9X_Y9kZE1eU7JlPYA4t26ema6eV7tRu5Rl0/exec';
const SHELL = ['./', './index.html', './manifest.webmanifest', './icons/icon-192.png', './icons/icon-512.png', './icons/apple-touch-icon.png'];

self.addEventListener('install', e => {
  e.waitUntil(caches.open(CACHE).then(c => c.addAll(SHELL)).then(() => self.skipWaiting()));
});

self.addEventListener('activate', e => {
  e.waitUntil(caches.keys()
    .then(keys => Promise.all(keys.filter(k => k !== CACHE).map(k => caches.delete(k))))
    .then(() => self.clients.claim()));
});

self.addEventListener('fetch', e => {
  const req = e.request;
  if (req.method !== 'GET') return;
  const url = new URL(req.url);

  // Шрифты Google почти не меняются — берём из кэша, если уже есть.
  if (url.hostname === 'fonts.googleapis.com' || url.hostname === 'fonts.gstatic.com') {
    e.respondWith(caches.match(req).then(hit => hit || fetch(req).then(res => {
      const copy = res.clone();
      caches.open(CACHE).then(c => c.put(req, copy));
      return res;
    })));
    return;
  }

  // Файлы самого сайта — сначала из сети (чтобы обновления приходили сразу), без сети — из кэша.
  if (url.origin === location.origin) {
    e.respondWith(fetch(req).then(res => {
      if (res.ok) {
        const copy = res.clone();
        caches.open(CACHE).then(c => c.put(req, copy));
      }
      return res;
    }).catch(() => caches.match(req, { ignoreSearch: true }).then(hit => hit || caches.match('./index.html'))));
  }
});

/* ---------- Пуш-уведомления ---------- */

// Пуш приходит пустым, поэтому спрашиваем у API, что произошло. Ключ входа сайт хранит в IndexedDB.
function savedKey() {
  return new Promise(done => {
    const open = indexedDB.open('pokupochki', 1);
    open.onupgradeneeded = () => open.result.createObjectStore('kv');
    open.onerror = () => done('');
    open.onsuccess = () => {
      const q = open.result.transaction('kv', 'readonly').objectStore('kv').get('key');
      q.onsuccess = () => done(q.result || '');
      q.onerror = () => done('');
    };
  });
}

self.addEventListener('push', e => {
  e.waitUntil((async () => {
    // iPhone отзывает разрешение, если на пуш не показать уведомление, — поэтому запасной текст есть всегда.
    let n = { title: 'Покупочки', body: 'Есть что-то новое', tag: 'news', url: './' };
    try {
      const key = await savedKey();
      if (key) {
        const res = await fetch(API_URL, { method: 'POST', headers: { 'Content-Type': 'text/plain;charset=utf-8' }, body: JSON.stringify({ key, fn: 'getNews', args: [] }) });
        const j = await res.json();
        if (j.ok && j.data) n = j.data;
      }
    } catch (err) { /* покажем запасной текст */ }
    await self.registration.showNotification(n.title, {
      body: n.body, tag: n.tag, icon: 'icons/icon-192.png', badge: 'icons/icon-192.png', data: { url: n.url },
    });
  })());
});

// Нажатие на уведомление открывает приложение сразу на нужной покупке.
self.addEventListener('notificationclick', e => {
  e.notification.close();
  const url = new URL((e.notification.data && e.notification.data.url) || './', self.registration.scope).href;
  e.waitUntil(self.clients.matchAll({ type: 'window', includeUncontrolled: true }).then(list => {
    const win = list.find(c => 'focus' in c);
    if (win) return win.navigate(url).then(c => (c || win).focus());
    return self.clients.openWindow(url);
  }));
});
