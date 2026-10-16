# HƯỚNG DẪN CÀI ĐẶT — 1 lần (16/10)

## 1. Code.gs → V50 (BISON + tab lệch cột + ảnh TRƯỚC/SAU kể cả dán tay)
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
