# HƯỚNG DẪN CÀI ĐẶT — 1 lần (16/10)

## 0. Bản mới nhất: Code.gs V62 (16/10) — V61 + 🧪 tự kiểm tra toàn hệ thống
File đầy đủ: `Code_V62.gs` (gửi riêng). Bản che khoá: `backend/Code_V62.CHE_KHOA.gs`.
Kiểm tra sau khi deploy: app → ⚙️ → 🧪 Tự kiểm tra TOÀN hệ thống (PIN) — hoặc trong Apps Script chọn hàm
**plTuKiemTra** → ▶ Run → Nhật ký thực thi. Chỉ đọc; thử thật link ảnh Drive, ảnh dán trong ô, ảnh nổi, tốc độ dashboard.

## 0b. Code.gs V61 (16/10) — nhiều dự án
File đầy đủ: `Code_V61.gs` (gửi riêng, có khoá thật). Bản che khoá: `backend/Code_V61.CHE_KHOA.gs`.
Dán đè Code.gs → Deploy → Manage deployments → ✏️ → New version. Menu Sheet: dán XuatPDF_Sheet.gs ('16/10 v61').
- Danh sách dự án do máy chủ trả (plProjects): thêm dự án = app → ⚙️ → ➕ Tạo dự án (cần PIN) → máy tạo tab
  "Punch list đơn giản-<DỰ ÁN>" chép đúng 3 dòng tiêu đề chuẩn + ghi vào Danh mục part.
- Ghi vào dự án CHƯA có tab → báo NO_PROJECT_TAB (không còn tự tạo tab lạ khi gõ sai tên).
- Dashboard mọi dự án: 1 lần gọi dashAll (có bộ nhớ đệm, tự xoá khi app ghi).
- App: không lẫn dữ liệu khi đổi dự án; ô nhập/bộ lọc nhớ riêng từng dự án; chọn dự án có tìm kiếm + gần đây.

## 0z. Code.gs V60 (16/10) — như V59 + ghi chú không đè cột ảnh
Tab không có cột Remarks (VIOLA): ghi chú / ghi chú đóng lỗi được nối vào ô Mô tả. App pl-v54 cần máy chủ V60.

## 0a. Code.gs V59 (16/10, nâng cấp lớn)
File đầy đủ: `Code_V60.gs` (gửi riêng trong chat, có khoá thật — KHÔNG lưu GitHub).
1. Mở Apps Script → tệp **Code.gs** → **Ctrl+A → dán Code_V60.gs → Ctrl+S**.
2. **Deploy → Manage deployments → ✏️ (bút chì) → Version: New version → Deploy**.
   ⚠️ KHÔNG bấm "New deployment" — phải giữ nguyên link /exec (app và mọi máy đang dùng link này).
3. Kiểm tra: mở app → chip góc phải phải hiện **"máy chủ V60"**.
   - Vẫn báo bản cũ → làm lại bước 2 (thường do chọn nhầm deployment hoặc quên chọn "New version").
   - App báo **"tệp cũ đang ghi đè simpleColMap"** → báo người quản lý (trong project có tệp thứ 2 chứa
     hàm simpleColMap cũ; sửa tệp thứ 2 cần user duyệt riêng).
4. Tự động lấy ảnh mỗi giờ: giữ như V58 (đã cài thì không cần cài lại). Kiểm tra trong app:
   ⚙️ → 🩺 Kiểm tra hệ thống → dòng "Tự lấy ảnh mỗi giờ: ĐANG BẬT · lần chạy cuối …".
V59 thêm: chống ghi trùng BỀN (dòng `REQ:<mã>` trong ghi chú ô SID — không thêm cột; đừng xoá các dòng REQ
trong ghi chú), action plPing (báo bản máy chủ) + plDiag (kiểm tra hệ thống, chỉ đọc), listSimple trả
Mức độ / Người khắc phục / Hạn xử lý (chỉ khi tab có đúng tiêu đề) + số ảnh dán/ảnh nổi chưa lên Drive,
lưu kết quả lần tự lấy ảnh, file ảnh lấy lỗi không còn bị bỏ lại mồ côi trên Drive (vào thùng rác, khôi phục được 30 ngày).

## 0b. (cũ) Code.gs V58 (16/10)
File đầy đủ: `Code_V58.gs` (gửi riêng trong chat, có khoá thật — KHÔNG lưu GitHub).
Bản lưu đã che khoá: `backend/Code_V60.CHE_KHOA.gs` (đã che cả SECRET_KEY, ADMIN_PIN, GEMINI_API_KEY, SHEET_ID).
1. Code.gs: **Ctrl+A → dán Code_V58.gs → Ctrl+S** → **Deploy → Manage deployments → ✏️ → New version → Deploy**.
2. Tự động lấy ảnh dán tay mỗi giờ (cài 1 lần): chọn hàm **plCaiTuDongLayAnh** → ▶ Run → cho phép.
   Tắt: chạy **plTatTuDongLayAnh**. Vẫn bấm tay được: app → ⚙️ → 🖼 Lấy ảnh DÁN TAY.
V53–V58: lấy ảnh dán trong ô + ảnh nổi (qua bản Excel tạm), không giữ khoá chung nhiều phút, đọc ảnh từ
tên file có gắn link / công thức IMAGE, không làm mất link ẩn khi thêm ảnh, chặn các lệnh nhạy cảm bằng PIN.

## 1. (cũ) Code.gs → V50 (BISON + tab lệch cột + ảnh TRƯỚC/SAU kể cả dán tay)
File đầy đủ: `Code_V50.gs` (gửi riêng trong chat, có khoá thật — KHÔNG lưu GitHub).
Bản lưu đã che khoá: `backend/Code_V50.CHE_KHOA.gs`.

1. Apps Script **Punch List Backend** → Code.gs → **Ctrl+End**: 3 dòng cuối phải là
   `return reply({ ok: false, error: String(err) });` · `}` · `}` (Code.gs kết thúc ở doPost).
2. Đúng → Code.gs: **Ctrl+A → Ctrl+V** file V49 → **Ctrl+S**.
   (Sai — còn `function` khác bên dưới → KHÔNG dán đè cả file, báo lại.)
3. **Deploy → Manage deployments → ✏️ → New version → Deploy**. KHÔNG "New deployment".
4. App → BISON → Báo cáo → 🔄 → Section 33 hiện 3 mục.
Quay lại: Manage deployments → ✏️ → chọn version cũ.

V49 sửa: tab theo tên dự án (giữ `Punch list đơn giản-BISON`) · tự nhận cột link/SID khi tiêu đề lệch ·
`closed`/`Close` = đóng · `listSimple` trả `afterFirst` · Đối chiếu / Sửa mục / Dọn trùng / đồng bộ đóng lỗi
đọc cột theo tên. Thử giả lập trên cấu trúc tab BISON thật: listSimple, listOpenPunch, punchAudit, punchDedupe đúng.

## 2. Menu Sheet "📄 XUẤT PDF" (tuỳ chọn)
File `backend/XuatPDF_Sheet.gs` → Apps Script: ＋ Tập lệnh `XuatPDF_Sheet` → dán → Lưu → chạy `plPdf_caiDat` → F5 Sheet.
Lệnh 🖼 "Sửa ảnh bị đè" đang KHOÁ.

## 3. App
Beta: https://themy1809-del.github.io/punchlist-app/beta/ (pl-v44). App chính chưa đổi.
Beta ổn → báo "thay app chính".

## 4. Bảo mật
Đổi khoá Gemini + ADMIN_PIN (đã lộ trong chat 15–16/10).
