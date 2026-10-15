# Giải pháp nâng cấp Punch List — nhiều người dùng, ổn định, xuất PDF trên Sheet

Ngày rà soát: 15/10/2026 · Bản app đang chạy: `pl-v37c-gonman` · SW cache `punchlist-v3`

## 0. Hiện trạng (đọc từ `index.html`)

| Thành phần | Hiện trạng | Đánh giá |
|---|---|---|
| Frontend | 1 file `index.html` (110 KB) + PWA (`sw.js`, `manifest.json`) | Tốt: chạy offline, tự cập nhật bản mới |
| Hàng chờ gửi | IndexedDB `punchlist_db` + `pump()` 20 giây/lần, thử lại 5 lần | Tốt: không mất ảnh khi mất mạng |
| Chống trùng | `reqId` gửi kèm mỗi mục, backend trả `duplicated` | Tốt |
| Ghi đồng thời | backend trả `busy` khi đang khoá → app tự gửi lại | Tốt (có LockService ở backend — **UNKNOWN, cần xem Code.gs**) |
| Backend | Apps Script `/exec` — mã nguồn **KHÔNG nằm trong repo** | Rủi ro: không có bản lưu, không soát được |
| Xác thực | `DEFAULT_KEY` viết cứng trong `index.html` + `deviceId` ngẫu nhiên | **Yếu**: repo công khai → ai đọc mã cũng có key |
| Xoá dữ liệu | PIN quản trị kiểm ở backend, lưu ở `localStorage` | Tạm được |
| Người dùng | Chỉ là ô nhập tên `iUser`, ai gõ tên gì cũng được | Không truy được ai làm gì |
| Xuất PDF | Có ở app (`punchPdf` → backend tạo PDF trên Drive) | Chưa có menu xuất ngay trong Sheet → **đã thêm** |

## 1. Xuất PDF ngay trên Google Sheet — ĐÃ LÀM

File mới: `backend/XuatPDF_Sheet.gs` (độc lập, mọi hàm có tiền tố `plPdf_`, không đụng `doPost`).

Menu **📄 XUẤT PDF** gồm:
- Tab đang mở — A4 ngang / A4 dọc
- Vùng đang chọn (bôi đen) — A4 ngang
- Tất cả tab đang hiện → 1 file PDF
- Mở thư mục PDF

Đặc điểm: vừa khổ ngang trang, lặp hàng tiêu đề (hàng cố định) mỗi trang, đánh số trang,
tên file có ngày giờ, lưu vào thư mục Drive `PUNCH LIST - PDF xuat tu Sheet`, tự thử lại khi Google báo bận (429).

Cài: Sheet → Tiện ích mở rộng → Apps Script → `+` Tập lệnh → dán file → chạy `plPdf_caiDat` 1 lần → tải lại Sheet.

## 2. Giải pháp nhiều người dùng — đề xuất (CHƯA làm, chờ duyệt)

Giữ nguyên kiến trúc **PWA + Apps Script + Google Sheet** (đủ cho đội 5–30 người, không tốn tiền).
Chỉ vá 4 điểm yếu:

### Bước 1 — Mã người dùng thay cho key chung (quan trọng nhất)
- Tab mới `NGUOI_DUNG` trong Sheet: `Mã | Tên | Vai trò (QC / Quản lý) | PIN | Hoạt động (Y/N)`.
- App: lần đầu mở → nhập Mã + PIN → backend trả **token** (lưu máy). Mỗi lần gửi kèm token thay `DEFAULT_KEY`.
- Backend: kiểm token → biết chắc ai gửi; cột "Người tạo/Người đóng" lấy từ token, không lấy ô gõ tay.
- Nghỉ việc / mất máy → đổi "Hoạt động" = N là khoá ngay, không phải đổi link /exec.
- Vai trò "Quản lý" mới được xoá / dọn trùng / xoá hết → bỏ PIN dùng chung.

### Bước 2 — Chịu tải khi nhiều người gửi cùng lúc
- Backend: `LockService.getScriptLock().tryLock(...)` chỉ bọc đoạn **ghi dòng + đánh số No.**;
  phần lưu ảnh lên Drive làm **trước**, ngoài khoá → khoá giữ ngắn, ít `busy`.
- App: thêm độ trễ ngẫu nhiên 0–5 giây trước khi `pump()` gửi lại (tránh 10 máy dồn cùng giây).
- Giới hạn Apps Script: ~30 lượt chạy đồng thời / 6 phút mỗi lượt → gói ảnh hiện 1600px/0.72 là hợp lý, giữ nguyên.

### Bước 3 — Đồng bộ danh sách giữa các máy
- `listOpenPunch` / `listSimple` trả thêm `version` (thời điểm sửa cuối). App chỉ tải lại khi version đổi
  → máy A đóng lỗi thì máy B thấy ngay lần làm mới sau, ít tốn quota.
- Đóng lỗi: backend kiểm "mục này đã đóng chưa" → nếu người khác đóng rồi thì báo, không ghi đè.

### Bước 4 — Bảo trì / an toàn
- Đưa `Code.gs` vào repo (thư mục `backend/`, **bỏ SECRET_KEY/ADMIN_PIN ra Script Properties**) để có bản lưu và soát lỗi.
- Tab `NHAT_KY` ghi: thời gian · người · thao tác · mã mục → truy vết khi có tranh chấp.
- Sheet dữ liệu để chia sẻ "Hạn chế", chỉ tài khoản chạy web app có quyền sửa (người dùng ghi qua app).
- Cập nhật web app luôn bằng Manage deployments → ✏️ → New version (giữ nguyên URL).

## 3. Khi nào mới cần đổi kiến trúc
Chỉ khi > 50 người dùng đồng thời hoặc > ~50.000 dòng/Sheet: chuyển dữ liệu sang Firebase/Supabase.
Hiện **chưa cần**.

## 4. Việc cần anh cung cấp
- Mã `Code.gs` của backend Punch List (để làm Bước 1–3) — **UNKNOWN — NEED USER CONFIRMATION**: backend đang gắn với Sheet (bound) hay đứng riêng?
- Số người dùng dự kiến và ai là "Quản lý".
