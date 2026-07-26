// Punch List Service Worker — V1 (fork từ FieldSnap V42)
// Mục tiêu: mở app KHÔNG CẦN MẠNG (cache app shell) + LUÔN lấy bản mới khi có mạng (network-first).
// HTML/sw luôn bỏ qua cache trình duyệt → cập nhật tức thì, không kẹt bản cũ.
// API backend là POST (không cache) — outbox trong app đã lo phần gửi lại.

const CACHE = 'punchlist-v1';
const ASSETS = ['./', './index.html', './manifest.json', './icon-192.png', './icon-512.png'];

self.addEventListener('install', (e) => {
  e.waitUntil(
    caches.open(CACHE)
      .then((c) => c.addAll(ASSETS))
      .then(() => self.skipWaiting())
  );
});

self.addEventListener('activate', (e) => {
  e.waitUntil(
    caches.keys()
      .then((keys) => Promise.all(keys.filter((k) => k !== CACHE).map((k) => caches.delete(k))))
      .then(() => self.clients.claim())
  );
});

self.addEventListener('fetch', (e) => {
  const req = e.request;
  const url = new URL(req.url);
  // Chỉ xử lý GET cùng nguồn (app shell). POST API + ảnh Drive đi thẳng ra mạng.
  if (req.method !== 'GET' || url.origin !== location.origin) return;
  // V36: HTML (trang) → ép lấy bản mới nhất, BỎ QUA cache trình duyệt; lỗi mạng mới dùng cache.
  const isDoc = req.mode === 'navigate' || (req.headers.get('accept') || '').indexOf('text/html') !== -1;
  const fetchReq = isDoc ? new Request(req.url, { cache: 'reload' }) : req;
  e.respondWith(
    fetch(fetchReq)
      .then((res) => {
        const copy = res.clone();
        caches.open(CACHE).then((c) => c.put(req, copy));
        return res;
      })
      .catch(() =>
        caches.match(req, { ignoreSearch: true })
          .then((r) => r || caches.match('./index.html'))
      )
  );
});
