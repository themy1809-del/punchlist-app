// Punch List Service Worker — V1 (fork từ FieldSnap V42)
// Mục tiêu: mở app KHÔNG CẦN MẠNG (cache app shell) + LUÔN lấy bản mới khi có mạng (network-first).
// HTML/sw luôn bỏ qua cache trình duyệt → cập nhật tức thì, không kẹt bản cũ.
// API backend là POST (không cache) — outbox trong app đã lo phần gửi lại.

const CACHE = 'punchlist-v21';   // v21 (16/10): pl-v55 nhiều dự án
const PREFIX = CACHE.indexOf('punchlist-beta-') === 0 ? 'punchlist-beta-' : 'punchlist-';
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
      // v13: chỉ xoá cache CŨ của CHÍNH app này (app chính và /beta/ chung 1 tên miền — trước đây xoá lẫn nhau)
      .then((keys) => Promise.all(keys.filter((k) => k !== CACHE && k.indexOf(PREFIX) === 0 &&
          (PREFIX === 'punchlist-beta-' || k.indexOf('punchlist-beta-') !== 0)).map((k) => caches.delete(k))))
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
  // v18: mạng chập chờn (có sóng nhưng treo) → quá 3 giây chưa có trả lời thì mở bản trong cache; mạng về sau vẫn lưu bản mới.
  const net = fetch(fetchReq).then((res) => {
    // v13: chỉ lưu trang TỐT (không lưu 404/5xx/trang lỗi wifi), lưu theo đường dẫn KHÔNG có ?nc= → không phình cache
    if (res.ok && res.type === 'basic') {
      const copy = res.clone();
      return caches.open(CACHE).then((c) => c.put(url.origin + url.pathname, copy)).catch(() => {}).then(() => res);
    }
    return res;
  });
  e.waitUntil(net.catch(() => {}));
  const fromCache = () => caches.match(req, { ignoreSearch: true })
    .then((r) => r || (isDoc ? caches.match('./index.html') : undefined));
  e.respondWith(
    Promise.race([net, new Promise((r) => setTimeout(() => r(null), 3000))])
      .then((r) => r || fromCache().then((c) => c || net))
      .catch(() => fromCache().then((c) => c || Response.error()))
  );
});
