# KẾT QUẢ RÀ SOÁT 10 TÁC TỬ (16/10)

77 phát hiện → 72 xác nhận (17 cao · 33 trung bình · 22 thấp), 5 bị bác bỏ khi kiểm chứng.
Đã sửa trong app pl-v47 / Code.gs V52 / XuatPDF 16/10 qa10: đánh dấu ✅. Còn lại: ⏳ (cần anh quyết hoặc làm tiếp).

| # | Mức | Mảng | Vấn đề | Trạng thái |
|---|---|---|---|---|
| 0 | Cao | app-logic | Typing an existing project name in different case leaves the select blank, so data is silently sent to the default BISON tab | ✅ đã sửa |
| 1 | Cao | offline-sync | Queued items are sent to whichever project and user is selected at SEND time, not the ones at capture time | ✅ đã sửa |
| 2 | Cao | offline-sync | IndexedDB writes are fire-and-forget: a failed write (quota, private mode, DB not opened) loses the item and its photos without any message | ✅ đã sửa |
| 3 | Cao | pdf-report | Before/after photos print about 15 mm wide and are cropped, so a client cannot read the defect | ✅ đã sửa |
| 4 | Cao | pdf-report | Report header (project, date, KPIs, Doc No) prints on page 1 only; pages have no page numbers | ✅ đã sửa |
| 5 | Cao | performance | Service worker stores a new 144 KB copy of index.html in Cache Storage on every app start (cache keeps growing) | ✅ đã sửa |
| 6 | Cao | performance | Dashboard loads projects one after another and downloads every item just to count them | ✅ đã sửa |
| 7 | Cao | security | API key for the /exec endpoint is public in the repo, so anyone online can call every server action | ⏳ |
| 8 | Cao | security | Destructive catalog actions skip the ADMIN_PIN check (catDelProj, descReplace, catalogBulk) | ✅ đã sửa |
| 9 | Cao | security | Stored XSS/JS injection in the audit screen through a part name from the Sheet | ✅ đã sửa |
| 10 | Cao | security | Sheet is anyone-with-link WRITER, which bypasses every server check and exposes user PINs and device IDs | ⏳ |
| 11 | Cao | server-scale | punchSimple / punchOpen: No. numbering + appendRow + getLastRow() with no LockService, so two QC sending at once get the same No. and photos/formatting land on the wrong row | ✅ đã sửa |
| 12 | Cao | server-scale | listSimple / listOpenPunch assign SIDs while reading, with no lock: two people reloading at the same time give the same row different IDs, and close/add-photo/delete then fail with 'Không tìm thấy mục' | ✅ đã sửa |
| 13 | Cao | sheet-data | Tab VIOLA: không có tiêu đề 'Remarks', E3:F3 gộp thành 'Photo', nên cột Photo trùng cột Remarks (E) và ảnh với ghi chú ghi đè lên nhau | ✅ đã sửa |
| 14 | Cao | sheet-data | Tab BISON: tiêu đề I3 ghi 'SID' nhưng chứa link ảnh, và lẫn mã SID cũ; cột K bị cấp SID MỚI hôm nay, nên mã của mục đã đổi | ⏳ |
| 15 | Cao | sheet-menu | Section PDF puts after-photos that are marked by note or column into the BEFORE column | ✅ đã sửa |
| 16 | Cao | ux-mobile | Choosing 'Re-inspect (kiểm lại)' still closes the item | ✅ đã sửa |
| 17 | Cao | ux-mobile | Close list does not show the defect photo or a clear Section/Part, so the field user cannot tell which item to close | ✅ đã sửa |
| 18 | TB | app-logic | Dashboard counts differ from the report and home stats for the same project | ✅ đã sửa |
| 19 | TB | app-logic | Open-age and 'quá 7 ngày' figures are wrong for rows not created by the app | ✅ đã sửa |
| 20 | TB | app-logic | Dashboard drill-down opens the wrong list (section '—', status, stale filters, wrong project) | ✅ đã sửa |
| 21 | TB | offline-sync | After 5 failed tries an item stops auto-sending for good, including after reconnecting or reopening the app | ✅ đã sửa |
| 22 | TB | offline-sync | Duplicate protection relies on CacheService (best-effort, 6h TTL); a send that times out and is retried late creates a duplicate row and photos | ⏳ |
| 23 | TB | offline-sync | Main and beta service workers delete each other's offline cache, and the SW caches error pages and unbounded ?nc= URLs | ✅ đã sửa |
| 24 | TB | offline-sync | Beta is not isolated from main: shared localStorage (API URL, project, open-list cache) and the same Web Lock name | ✅ đã sửa |
| 25 | TB | pdf-report | Missing or broken photos print silently after a 15 s timeout | ✅ đã sửa |
| 26 | TB | pdf-report | Print starts from an async callback with no user gesture; on an iOS PWA it can fail with no message | ⏳ |
| 27 | TB | performance | listSimple is requested repeatedly for the same data (startup, opening the report, returning to the app) | ⏳ |
| 28 | TB | performance | Every request opens the spreadsheet twice and scans 40 rows to find the photo column (worst on BISON) | ⏳ |
| 29 | TB | performance | listSimple and listOpenPunch write missing SIDs one cell at a time inside a read request | ✅ đã sửa |
| 30 | TB | performance | Section PDF report formats each row with ~8 separate Sheets calls (slow, risks the 6-minute limit) | ⏳ |
| 31 | TB | security | ADMIN_PIN stored in plain text in localStorage, no brute-force limit on verifyPin or PIN-gated actions | ⏳ |
| 32 | TB | server-correct | Rows whose SID starts with PL can only be closed through the VIOLA 'Punch List' tab: no fallback, and status checks are case/variant-sensitive | ⏳ |
| 33 | TB | server-correct | No LockService in sheet write paths: concurrent submits put the photo formula on the wrong row, and concurrent listSimple/listOpenPunch overwrite each other's new SIDs | ✅ đã sửa |
| 34 | TB | server-correct | punchFixImg after-only migration names files PLMIG_…, so the bound Sheet PDF (XuatPDF_Sheet.gs) classifies the migrated AFTER photos as BEFORE | ⏳ |
| 35 | TB | server-correct | Destructive catalog actions have no PIN, and sheetInfo returns the WRITER-shared Sheet URL, while the SECRET_KEY is public in index.html | ⏳ |
| 36 | TB | server-scale | punchFixImg has no time budget and no lock: on a tab with many pasted images it hits the 6-minute limit, leaves orphan Drive files, and double runs duplicate the images | ✅ đã sửa |
| 37 | TB | server-scale | punchAddPhoto and punchClose read a cell, append to it and write it back with no lock, so concurrent updates are lost | ⏳ |
| 38 | TB | server-scale | punchOpen swallows a failed simple-tab write and still returns ok, which is then cached as success for 6 hours | ⏳ |
| 39 | TB | sheet-data | listSimple tự cấp SID theo giờ ĐỌC (Date.now()), nên dashboard tính tuổi lỗi sai và một lệnh đọc lại ghi vào sheet mà không khoá | ⏳ |
| 40 | TB | sheet-data | Gregory mục 2: ô Photo link chứa TÊN FILE 'PLS_20260922_162252_1.jpg' thay vì link Drive | ⏳ |
| 41 | TB | sheet-data | BISON: cột 'Người lập' (R) trống, tên QC 'Mỹ' lại nằm ở 'Người khắc phục' (P); tên QC không thống nhất ('Mỹ' với 'QC Mỹ') | ⏳ |
| 42 | TB | sheet-data | Trạng thái ghi lẫn 'open'/'close'/'Close'/'closed' giữa các tab; còn 2 chỗ so sánh cứng, và cột trạng thái không có danh sách chọn (data validation) | ⏳ |
| 43 | TB | sheet-data | Sheet đang chia sẻ 'ai có link cũng SỬA được', trong khi app phụ thuộc vào tên tiêu đề dòng 3 và cột SID | ⏳ |
| 44 | TB | sheet-menu | Item sort has no section tie-break, so sections interleave and SECTION bands repeat | ✅ đã sửa |
| 45 | TB | sheet-menu | Long section exports can hit the 6-minute limit and leave a visible PDF_TAM_ sheet | ✅ đã sửa |
| 46 | TB | sheet-menu | DASHBOARD rewrite does not undo the previous run's merges | ✅ đã sửa |
| 47 | TB | ux-mobile | The bulk-close screen exists but nothing opens it | ⏳ |
| 48 | TB | ux-mobile | Report screen: 7 fixed bottom buttons, including admin and destructive ones, take a third of the phone screen | ✅ đã sửa |
| 49 | TB | ux-mobile | The per-item 🗑 delete button sits beside '+ Ảnh sau' and renumbers all items | ⏳ |
| 50 | TB | ux-mobile | Small grey text and labels are unreadable in sunlight | ✅ đã sửa |
| 51 | Thấp | app-logic | Project whose Sheet tab is missing is shown as '✔ Đã đóng hết 0 lỗi' (100% done look) | ⏳ |
| 52 | Thấp | app-logic | PDF prints after 15 s even if photos have not loaded, with no warning | ✅ đã sửa |
| 53 | Thấp | app-logic | fixPastedImg aborts at 3 minutes with a misleading 'sẽ tự gửi lại' message | ⏳ |
| 54 | Thấp | app-logic | splitPhotos mislabels a later 'before' photo as an 'after' photo on rows without an AFTER note | ⏳ |
| 55 | Thấp | offline-sync | Adding before/after photos is online-only and uses a new reqId on every attempt, so a retry after a timeout duplicates photos | ⏳ |
| 56 | Thấp | pdf-report | Section header row can print alone at the bottom of a page | ⏳ |
| 57 | Thấp | pdf-report | Item date column never prints because listSimple does not return a date | ⏳ |
| 58 | Thấp | pdf-report | Signature block has no date lines and fills 'Prepared by' with whoever is using the app | ⏳ |
| 59 | Thấp | performance | listSimple returns blank rows and makes 3 separate reads; the client then throws the blank rows away | ⏳ |
| 60 | Thấp | security | Attribute values escaped with esc() lack quote escaping (option values, data-u) and allow attribute injection from Sheet data | ✅ đã sửa |
| 61 | Thấp | security | Report PDFs/Excel folders are set to ANYONE_WITH_LINK view on every call | ⏳ |
| 62 | Thấp | server-correct | punchAddPhoto 'before' appends after existing close photos, so on rows without an AFTER note the app splits the new BEFORE photos into AFTER | ✅ đã sửa |
| 63 | Thấp | server-scale | punchDelete renumbers the whole tab one setValue per row, with column 1 hard-coded, and has no lock | ⏳ |
| 64 | Thấp | server-scale | plSimpleTabName_ runs openById + getSheets() on every request, even actions that never touch the simple tab | ⏳ |
| 65 | Thấp | sheet-data | Tab 'Dự án' đã cũ, không khớp dự án thật; 'Danh mục part' có section đáng ngờ | ⏳ |
| 66 | Thấp | sheet-menu | Temp sheet is not grown past its default 1000 rows | ⏳ |
| 67 | Thấp | sheet-menu | Writing __DU_AN__ to the catalog has no lock and does not clear the catalog cache | ✅ đã sửa |
| 68 | Thấp | sheet-menu | Section colour rule fails for decimal sections and gives one colour to text sections | ⏳ |
| 69 | Thấp | sheet-menu | Project name is worked out two different ways, which risks double counting BISON | ⏳ |
| 70 | Thấp | ux-mobile | Many tap targets are below 44px (section chips, dashboard rows, header buttons) | ⏳ |
| 71 | Thấp | ux-mobile | Confusing jargon on the camera and form: 'PUNCH TED', 'tab "Punch list đơn giản"', 'Project' vs 'Dự án' | ⏳ |
