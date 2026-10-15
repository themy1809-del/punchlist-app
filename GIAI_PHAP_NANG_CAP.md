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

## 1bis. Cải tiến trong app — ĐÃ LÀM (bản `pl-v38-nhieunguoi`, SW `punchlist-v4`)

Chỉ sửa `index.html` + `sw.js`, chạy được với backend HIỆN TẠI (không đổi API):
- **🖨 IN / PDF ĐANG LỌC** (màn Báo cáo): in đúng danh sách đang lọc (section / tình trạng / người lập / tìm kiếm)
  ra A4 ngang, có ảnh, tiêu đề lặp mỗi trang → chọn "Lưu dưới dạng PDF". Không cần máy chủ tạo file.
- **Nhiều cửa sổ trên 1 máy**: chỉ 1 cửa sổ gửi hàng chờ (Web Locks) → không gửi chồng.
- **Gửi lại lệch giờ ngẫu nhiên** khi Google lỗi tạm → nhiều máy không dồn cùng giây.
- **Mở lại app từ nền** → tự gửi hàng chờ + tải lại danh sách/số liệu (thấy ngay mục người khác vừa ghi/đóng), tối đa 1 lần/phút.
- **Bắt buộc tên người kiểm tra** trước khi lưu mục → truy được ai ghi.
- Vá 2 chỗ chèn dữ liệu chưa escape (mã ảnh Drive trong báo cáo, link PDF trong nút Copy).

## 1ter. Tối ưu giao diện + kỹ thuật — ĐÃ LÀM (bản `pl-v39-giaodien`, SW `punchlist-v5`)
- Bỏ màu tím: nút "PUNCH TED" + nhãn chế độ chụp đổi sang xanh dương.
- Màn Báo cáo: thanh nút dưới xếp lưới 3 cột, đủ 6 nút trên điện thoại (trước bị tràn, mất Excel/Đối chiếu);
  nút nguy hiểm (Dọn trùng, Xoá hết) nhỏ + viền, tách khỏi nút xuất.
- Nút "Xóa — chụp lại" ở từng mục: nhỏ, viền đỏ, nằm góc phải (trước là nút đỏ to, dễ bấm nhầm).
- Ô "Người kiểm tra" viền cam khi còn trống; viền xanh khi chọn bằng bàn phím.
- Danh sách Đóng lỗi vẽ 50 mục/lần + nút "Xem thêm" (vài trăm mục không giật).
- Ô tìm kiếm (Báo cáo, Đóng lỗi) chờ 0,18 giây sau phím cuối mới lọc → gõ mượt.
- Ảnh thu nhỏ Drive giải mã nền, có sẵn kích thước → trang không nhảy khi ảnh tải.

## 1quater. Mẫu PDF mới + sửa lỗi dự án — ĐÃ LÀM (bản `pl-v40-mau-pdf`, SW `punchlist-v6`)
Lỗi anh báo 15/10 và nguyên nhân (đối chiếu bản backend lưu trong lịch sử repo `PL_Code_1/2.txt`, ngày 30/07 —
bản đang deploy có thể khác, **UNKNOWN — NEED USER CONFIRMATION**):
| Lỗi | Nguyên nhân |
|---|---|
| Thêm 2 ảnh chỉ xuất 1 | `punchSimple` chỉ đặt `IMAGE()` của ảnh ĐẦU vào ô ảnh; `buildTedReport` đọc ô đó, bỏ cột "Photo link" (nơi có đủ link) |
| Xuất không nhận dự án | (1) `punchPdf` luôn đọc tab cố định + tên dự án cố định, bỏ qua `proj`; (2) app: `renderProjects()` tự nhảy về BISON và GHI ĐÈ lựa chọn khi danh mục chưa có dự án đó → **đã sửa trong app** |
| Mẫu xuất ra to | dòng cao 132px, ảnh 120×158px → ~5 mục/trang |

Đã làm trong app:
- Nút chính màn Báo cáo = **📄 XUẤT PDF — mẫu gọn, đủ ảnh**: dựng ngay trên máy từ danh sách đã tải
  (đúng dự án, ĐỦ ảnh mỗi mục), A4 ngang, ~9 mục/trang, gom theo Section, ô Tổng/Đang mở/Đã đóng, khối ký 3 bên.
- Nút cũ (server tạo PDF lên Drive) giữ lại: **☁️ PDF lên Drive (cũ)**.
- Giữ nguyên dự án đã chọn khi danh mục chưa tải xong.

Còn lại phía server (cần `Code.gs` đang deploy): `punchPdf` đọc đúng tab theo `proj` + in đủ ảnh từ cột "Photo link".

## 1quinquies. Ảnh TRƯỚC/SAU · xuất theo Section trên Sheet · màn hình chính gọn (15/10)
**App `pl-v42-gon` (SW `punchlist-v8`):**
- PDF trong app có 2 cột **Ảnh TRƯỚC (lỗi)** / **Ảnh SAU (khắc phục)**. Cách tách (`splitPhotos`):
  `afterIds`/`afterFirst` từ máy chủ → nếu chưa có thì dùng số ảnh đóng lỗi CHÍNH MÁY NÀY đã gửi →
  không biết thì in chung và ghi "chưa tách".
- Màn hình chính: hero thấp, ô tên 1 dòng (viền cam khi trống), 1 nút to "Chụp ảnh ghi lỗi" + lưới 2×2
  (Đóng lỗi · Báo cáo/PDF · Ghi không ảnh · Hàng chờ) — vừa 1 màn hình điện thoại, không cuộn.

**Sheet — `backend/XuatPDF_Sheet.gs` (`'15/10 pdf-section'`), menu 📄 XUẤT PDF:**
- ⭐ PUNCH LIST theo SECTION (gõ 1 hoặc nhiều section, trống = tất cả) · bản "chỉ mục ĐANG MỞ".
  Đọc tab "Punch list đơn giản - <dự án>" đang mở, dò cột theo TÊN (tab BISON lệch cột → dò cột có link Drive),
  tách ảnh trước/sau theo TÊN FILE (`PLS_` = lỗi, `PLSclose_`/`PLclose_` = khắc phục), tối đa 3 ảnh/bên,
  dựng sheet tạm → PDF A4 ngang vào thư mục `PUNCH LIST - PDF xuat tu Sheet` → xoá sheet tạm. Không sửa dữ liệu.
- 🧹 Làm gọn cột link ảnh: chỉ đổi định dạng (cắt chữ, chữ nhỏ xám, cột hẹp) — nội dung link giữ nguyên.

**Phát hiện trên Sheet thật (đọc 15/10, chưa sửa — cần anh duyệt):**
- Tab `Punch list đơn giản-BISON` lệch cột: cột I tiêu đề "SID" nhưng chứa link ảnh; SID thật ở cột K;
  "Người lập" (R) trống, tên nằm ở "Người khắc phục" (P).
- Tab `Gregory` dòng 2: ô link ghi tên file `PLS_20260922_162252_1.jpg` thay vì link → ảnh không hiện.
- Cột "Photo after" (J) trống ở các dòng đã đóng đọc được → máy chủ đang deploy có thể không ghi cột này.

**Sửa máy chủ đề xuất (để app tách ảnh trước/sau trên MỌI máy)** — trong nhánh `listSimple` của `Code.gs`,
trả thêm `afterIds` = các ID trong cột link có tên file chứa `close`
(hoặc nhẹ hơn: `afterFirst` = ID trong công thức cột "Photo after"). App đã sẵn sàng đọc 2 trường này.

## 1sexies. Dashboard · dự án mới tự nhận · màu Section/Đóng-Mở · ảnh đè khi lọc (15/10)
**Sheet (`XuatPDF_Sheet.gs` `'15/10 dashboard'`), menu 📄 XUẤT PDF thêm:**
- 📊 Cập nhật DASHBOARD + nhận dự án mới: tab `DASHBOARD PUNCH` (ghi đè tại chỗ; tên riêng để không đè tab của `Dasboard.gs` có sẵn) — mỗi dự án Tổng/Mở/Đóng/%/thanh tiến độ,
  mục mở > 7 ngày (ngày lấy từ SID), theo Section, theo người lập. Dự án = MỌI tab "Punch list đơn giản - <dự án>".
  Dự án chưa có trong "Danh mục part" → ghi thêm 1 dòng `<dự án> | | __DU_AN__` để app mọi máy thấy (app ẩn part giả).
- ⏱ Bật tự cập nhật Dashboard mỗi giờ (trigger thời gian, cài 1 lần).
- 🎨 Tô màu: mỗi Section 1 màu nhạt (5 màu xoay vòng) ở cột No.→Description; Tình trạng cam = mở / xanh = đóng.
  Chỉ là định dạng có điều kiện; quy tắc người dùng tự đặt được giữ.
- 🖼 Sửa ảnh bị đè khi lọc: ảnh NỔI (chèn đè lên ô) → đổi thành ảnh TRONG Ô `=IMAGE()` ở cột Photo theo link của dòng,
  rồi gỡ ảnh nổi. Dòng không có link: giữ nguyên. Có hỏi xác nhận trước.

**App `pl-v43-dashboard` (SW `punchlist-v9`):**
- Màn **📊 Dashboard**: mọi dự án trong danh sách — Mở/Đóng/Tổng, % đóng, mục mở > 7 ngày, chip theo Section,
  nút sang báo cáo dự án đó. Lưu bản gần nhất để xem khi mất mạng. Cảnh báo nếu máy chủ trả cùng dữ liệu cho mọi dự án.
- **Sửa lỗi đếm**: Sheet ghi "close"/"Close"/"closed" → trước đây "closed" bị tính là ĐANG MỞ. Nay mọi chữ bắt đầu "clos" = đã đóng.
- Màu theo Section (xếp theo thứ tự section → liền nhau khác màu) ở thẻ báo cáo + PDF; thẻ đã đóng mờ, nền xanh;
  nhãn "● ĐANG MỞ" / "✔ ĐÃ ĐÓNG".

## 1septies. Sự cố bản BACKUP mất ảnh + tính năng mới (15/10)
- **Sự cố:** anh báo bản `Punch List -Evapco - BACKUP 15-10` (bản sao lưu tạo 13:05) mất ảnh sau khi chạy các lệnh menu,
  trong đó có 🖼 Sửa ảnh bị đè. **File gốc không bị đụng** (sửa lần cuối 13:01, trước khi sao lưu).
  Ảnh Drive đều đang chia sẻ "ai có link xem được" → nguyên nhân chính xác: **UNKNOWN — NEED USER CONFIRMATION**.
  Đã **tạm khoá** lệnh 🖼 (gỡ khỏi menu, hàm chỉ báo "tạm khoá") — `'15/10 khoa-sua-anh'`.
  Khôi phục bản BACKUP: mở file → Tệp → Lịch sử phiên bản → chọn phiên bản lúc tạo (13:05) → Khôi phục.
- **App `pl-v44-tinhnang` (SW `punchlist-v10`), beta cập nhật theo:**
  đèn báo máy chủ (xanh/đỏ, bấm xem lỗi) · số ngày đang mở trên từng mục (đỏ > 7 ngày) · hàng nút Section có màu để lọc 1 chạm
  (ẩn ô chọn Section, bộ lọc gọn còn 2 hàng) · 📤 Gửi tóm tắt dự án (Zalo/email, dạng chữ) · Dashboard: bấm Section → báo cáo Section đó.

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
