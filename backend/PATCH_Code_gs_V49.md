# PATCH Code.gs V49 — tìm tab dự án theo tên (giữ tên "Punch list đơn giản-BISON")

Mục đích: BISON (dự án mặc định) trước đây đọc tab gốc `Punch list đơn giản` (không hậu tố)
→ tab đã đổi tên `Punch list đơn giản-BISON` nên app nhận 0 mục.
Sau patch: mọi dự án tìm tab `Punch list đơn giản - <dự án>`, so khớp KHÔNG phân biệt
dấu cách / dấu gạch / hoa thường → `…-BISON`, `… - BISON` đều nhận. Tab gốc cũ vẫn dùng được nếu còn.
Dự án mới chưa có tab → tạo tên chuẩn `Punch list đơn giản - <dự án>` (như cũ).
`PUNCH_TAB` (tab mẫu VIOLA) GIỮ NGUYÊN logic cũ.

## Sửa 1 — thay đúng 1 dòng trong doPost (dòng có chú thích V47)

TÌM dòng bắt đầu bằng:
```js
    (function(){ var _p = (data && data.proj != null) ? String(data.proj).trim() : ''; var _sfx = (_p && _p !== PUNCH_PROJECT_NAME) ? (' - ' + _p) : ''; SIMPLE_TAB = 'Punch list đơn giản' + _sfx; PUNCH_TAB = 'Punch List' + _sfx; })();
```
THAY bằng:
```js
    (function(){ var _p = (data && data.proj != null) ? String(data.proj).trim() : ''; var _sfx = (_p && _p !== PUNCH_PROJECT_NAME) ? (' - ' + _p) : ''; SIMPLE_TAB = plSimpleTabName_(_p); PUNCH_TAB = 'Punch List' + _sfx; })();   // V49: tab đơn giản tìm theo tên dự án
```

## Sửa 2 — dán hàm mới vào CUỐI file Code.gs

```js
// V49: tên tab "Punch list đơn giản" của 1 dự án — so khớp bỏ dấu cách/dấu gạch/hoa thường
// ("Punch list đơn giản-BISON" = "Punch list đơn giản - BISON"). Không thấy → tên chuẩn để tạo mới.
function plSimpleTabName_(p) {
  var proj = String(p || PUNCH_PROJECT_NAME || '').trim();
  var want = 'Punch list đơn giản - ' + proj;
  try {
    var shs = SpreadsheetApp.openById(SHEET_ID).getSheets();
    var norm = function (s) { return String(s).toLowerCase().replace(/[\s\-–_]+/g, ''); };
    var k = norm(want), base = null;
    for (var i = 0; i < shs.length; i++) {
      var nm = shs[i].getName();
      if (norm(nm) === k) return nm;
      if (nm === 'Punch list đơn giản') base = nm;
    }
    if (base && proj === PUNCH_PROJECT_NAME) return base;   // tab gốc cũ của dự án mặc định (nếu còn)
  } catch (e) {}
  return want;
}
```

## Sau khi sửa
1. Lưu (Ctrl+S).
2. Deploy → **Manage deployments** → ✏️ → Version: **New version** → Deploy.
   **KHÔNG** bấm "New deployment" (đổi URL → app hỏng).
3. Trên Sheet, ô **I3** của tab BISON: `SID` → `Photo link` (cột I đang ẩn; cột K mới là SID thật).

## Quay lại nếu cần
Deploy → Manage deployments → ✏️ → chọn lại version cũ → Deploy.
