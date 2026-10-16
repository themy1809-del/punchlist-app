// BẢN LƯU — ĐÃ CHE KHOÁ. Dùng file đầy đủ gửi riêng, hoặc Code_V50_TU_DANH_MUC.gs (thay từ dòng '// Danh mục' đến hết).
/**
 * Punch List Backend v1 — Google Apps Script  (FORK từ FieldSnap V44)
 * Nền tảng: JSON schema cứng, chống ảo giác, OCR đa hướng, đối chiếu chéo entities.
 * App RIÊNG — hạ tầng tách hoàn toàn khỏi FieldSnap (Sheet + Apps Script + repo riêng).
 * TODO Punch List: tab "Punch List" theo mẫu VIOLA, vòng đời open→close, ô nhập Vendor/Workshop.
 *   (chi tiết: README-PUNCH-LIST.md + PUNCH-LIST-DESIGN.md)
 *
 * CÀI ĐẶT app mới: xem README-PUNCH-LIST.md
 *
 * V49 (16/10): tab đơn giản tìm theo TÊN dự án (giữ "Punch list đơn giản-BISON");
 *   simpleColMap tự nhận cột link/SID khi tiêu đề lệch; "closed"/"Close" = đã đóng;
 *   listSimple trả thêm afterFirst (ảnh khắc phục đầu tiên) để app tách ảnh trước/sau;
 *   punchClose(đồng bộ VIOLA→đơn giản) / punchAudit / punchFix / punchDedupe đọc cột THEO TÊN (simpleColMap).
 * V50 (16/10): ẢNH SAU (khắc phục) được đánh dấu bằng GHI CHÚ "AFTER:id1,id2" trên ô Photo after
 *   (punchClose + punchFixImg ghi, listSimple trả afterIds) → PDF tách đúng ảnh TRƯỚC / SAU,
 *   kể cả ảnh DÁN TAY vào cột Photo after. Không thêm cột, không đổi dữ liệu.
 */

// ====== CẤU HÌNH — ⚠️ APP MỚI: PHẢI ĐIỀN LẠI CHO RIÊNG ======
var SECRET_KEY     = '***GIU_NGUYEN_CUA_ANH***';   // KHÁC FieldSnap — dùng đúng chuỗi này trong index.html app mới
var SHEET_ID       = '1B3GeWFxWwz_9_JViFDGBkavK_eD05WsKrkN1QyOxpGM';   // Sheet "Punch List - VIOLA" (themy1809)
var FOLDER_NAME    = 'Punch List Photos';        // thư mục Drive riêng cho ảnh app này
var ADMIN_PIN      = '***';  // V43: mã PIN quản trị cho thao tác XOÁ trên app (đổi tuỳ ý; để '' = tắt kiểm soát)
var GEMINI_API_KEY = '***';  // (tùy chọn) dán key Gemini nếu muốn AI đọc lỗi; để trống = tắt AI
// ====== V22 — CẢNH BÁO TỨC THÌ QUA TELEGRAM ======
// B1: Telegram → tìm @BotFather → /newbot → đặt tên → nhận TOKEN, dán vào dưới.
// B2: Mở bot vừa tạo, bấm Start + nhắn 1 câu bất kỳ.
// B3: Trong editor chạy hàm telegramGetChatId → xem log lấy chat id, dán vào dưới.
// Để trống 1 trong 2 = tắt tính năng. Nhóm Telegram cũng được (thêm bot vào nhóm).
var TELEGRAM_BOT_TOKEN = '';
var TELEGRAM_CHAT_ID   = '';
var TELEGRAM_NOTIFY_ALL = true;  // V24.1: true = báo Telegram MỌI báo cáo mới kèm ảnh (false = chỉ báo khi AI thấy lỗi)
// ====== V23 — MODEL AI (mới nhất + dự phòng khi lỗi/hết quota) ======
var GEMINI_MODEL          = 'gemini-3.5-flash';
var GEMINI_MODEL_FALLBACK = 'gemini-2.5-flash';
// V32: (tùy chọn) ảnh chữ ký Sếp cho khối nghiệm thu trong PDF — upload ảnh PNG chữ ký
// lên Drive, lấy File ID dán vào đây. Để trống = PDF chỉ in khung ký tay.
var SIGNATURE_FILE_ID = '';
// ========================================

// Danh mục — trọng tâm QC KÍCH THƯỚC kết cấu thép
var CATEGORIES = ['Kiểm tra kích thước', 'Chuẩn bị phôi / cắt', 'Gá lắp / định vị (fit-up)',
                  'Hàn', 'Kiểm tra mối hàn (VT/NDT)', 'Sơn / mạ / xử lý bề mặt',
                  'Lắp dựng tại công trường', 'Nghiệm thu / bàn giao', 'Khác'];

var HEADER = ['Thời gian', 'Dự án', 'Hạng mục', 'Mô tả AI', '⚠️ Cảnh báo', 'Ghi chú',
              'Số seri', 'Mối hàn', 'Kích thước', 'Số bản vẽ', 'Mác thép',
              'OCR thô', 'Độ tin cậy AI', 'Tọa độ GPS', 'Địa điểm', 'Link ảnh',
              'Người gửi', 'Report ID', 'Model', 'Cổng', 'Cấu kiện'];

var PROJECTS_TAB = 'Dự án';
var ALL_TAB = 'Tổng hợp';
var USERS_TAB = 'Người dùng';   // V18: danh sách tên chuẩn + vai trò + PIN
var UNITS_TAB = 'Cấu kiện';     // V29: kích thước thiết kế + dung sai để AI đối chiếu
var ISSUE_COL = 5; // cột ⚠️ Cảnh báo

// ====== PUNCH LIST (mẫu VIOLA) — tab riêng dạng danh sách lỗi open→close ======
var PUNCH_TAB = 'Punch List';
// Thông tin khối tiêu đề (sửa cho đúng dự án/khách của bạn; sau này có thể đưa vào tab Dự án)
var PUNCH_PROJECT_NAME = 'BISON';
var PUNCH_CUSTOMER     = 'EDC';
var PUNCH_DOC_NO       = '';   // để trống = không in dòng Doc No
var PUNCH_PREPARED_BY  = 'EDC VN TEAM';
// Cột: A COMMODITY | B DESCRIPTION(hạng mục) | C VENDOR | D STATUS | E DATE | F INSPECTED BY |
//      G ITEM# | H DESCRIPTION(lỗi) | I RESULT | J ACTION | K CLOSING DATE | L Open Photos |
//      M Closing photos | N REMARK WORKSHOP | O PUNCH_ID | P OPEN_IDS | Q CLOSE_IDS | R REPORT_ID (O–R ẩn)

/**
 * Khai báo danh sách dự án (chạy 1 lần từ editor để nạp vào tab "Dự án").
 * Thêm dự án mới: bổ sung vào mảng rows rồi Run lại. Tọa độ/bán kính để '' nếu
 * không cần tự gán GPS (vẫn chọn tay được trong app).
 */
function setupProjects() {
  var rows = [
    // [Tên dự án, Vĩ độ, Kinh độ, Bán kính (m)]
    ['Evapco Viola', '', '', ''],
    ['Cầu ống lồng di động', '', '', '']
  ];
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var sh = ss.getSheetByName(PROJECTS_TAB) || ss.insertSheet(PROJECTS_TAB);
  sh.getRange(1, 1, 1, 4)
    .setValues([['Tên dự án', 'Vĩ độ', 'Kinh độ', 'Bán kính (m)']]).setFontWeight('bold');
  sh.getRange(2, 1, rows.length, 4).setValues(rows);
  SpreadsheetApp.flush();
}

/**
 * V27 — SETUP CÔNG TY MỚI: chạy 1 LẦN duy nhất sau khi dán code + điền config.
 * Tự tạo: Sheet tabs (Dự án, Tổng hợp, Người dùng), thư mục Drive, trigger AI mỗi phút.
 * Sau đó: Deploy → New deployment → Web app (Execute as: Me, Access: Anyone) → lấy URL.
 */
function setupCompany() {
  var ss = SpreadsheetApp.openById(SHEET_ID);
  getProjects();        // tạo tab "Dự án" + header nếu chưa có
  getUsersSheet();      // tạo tab "Người dùng"
  if (!ss.getSheetByName(ALL_TAB)) {   // tạo tab Tổng hợp + header
    var shA = ss.insertSheet(ALL_TAB);
    shA.appendRow(HEADER);
    shA.getRange(1, 1, 1, HEADER.length).setFontWeight('bold')
       .setBackground('#4285f4').setFontColor('#ffffff');
    shA.setFrozenRows(1);
  }
  getOrCreateFolder(FOLDER_NAME);
  // trigger AI nền mỗi phút
  ScriptApp.getProjectTriggers().forEach(function (t) {
    if (t.getHandlerFunction() === 'processAiQueue') ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('processAiQueue').timeBased().everyMinutes(1).create();
  Logger.log('✅ Setup xong cho công ty mới!\nSheet: ' + ss.getUrl() +
    '\nTiếp theo: Deploy → New deployment → Web app (Execute as Me, Access Anyone) → copy URL /exec.' +
    '\nTạo MÃ KẾT NỐI cho nhân viên: chạy hàm makeJoinCode sau khi điền JOIN_URL bên trong nó.');
}

/**
 * V27 — TẠO MÃ KẾT NỐI + LINK CÀI ĐẶT cho nhân viên công ty này.
 * Điền URL /exec (sau khi deploy) vào JOIN_URL rồi chạy — log ra mã + link gửi nhân viên.
 */
function makeJoinCode() {
  var JOIN_URL = '';   // ← dán URL Web app /exec của công ty này vào đây rồi Run
  var JOIN_NAME = '';  // ← tên công ty hiển thị trong app (vd 'Công ty ABC')
  if (!JOIN_URL) { Logger.log('Chưa điền JOIN_URL — dán URL /exec rồi chạy lại.'); return; }
  var payload = JSON.stringify({ u: JOIN_URL, k: SECRET_KEY, n: JOIN_NAME });
  var code = Utilities.base64EncodeWebSafe(payload).replace(/=+$/, '');
  Logger.log('MÃ KẾT NỐI:\n' + code +
    '\n\nLINK GỬI NHÂN VIÊN (bấm là tự cấu hình):\nhttps://themy1809-del.github.io/fieldsnap-demo/?join=' + code);
}

// ============ CHỐNG GHI TRÙNG (V16) ============
// Các action ghi dữ liệu — cần bảo đảm gửi lại KHÔNG tạo thêm dòng/ảnh.
var IDEM_ACTS = { punchOpen: 1, punchSimple: 1, punchClose: 1 };
var _IDEM_KEY = '';   // khoá của lần gọi hiện tại (biến toàn cục reset mỗi lần chạy)
var _IDEM_TTL = 300;

/** Khoá nhận diện DÒNG TRÙNG trên tab đơn giản: Section + Part + Mô tả + ảnh (công thức/link). */
function pdupKey(section, part, desc, remarks, photoFormula, photoLink) {
  var img = String(photoFormula || '').match(/[-\w]{25,}/);
  var lnk = String(photoLink || '').match(/[-\w]{25,}/);
  function nz(x) { return String(x == null ? '' : x).trim().toLowerCase().replace(/\s+/g, ' '); }
  return [nz(section), nz(part), nz(desc), nz(remarks),
          img ? img[0] : (lnk ? lnk[0] : '')].join('|');
}

/** Khoá chống trùng: ưu tiên reqId app gửi; app CŨ chưa có reqId thì suy từ nội dung (TTL ngắn). */
function idemKey(data) {
  var a = String(data.action || '');
  if (data.reqId) return 'idem_' + a + '_' + String(data.reqId).replace(/\W/g, '').slice(0, 60);
  var sig = [a, String(data.punchId || ''), String(data.section || ''), String(data.part || ''),
             String(data.description || data.note || ''), String(data.remarks || ''),
             String(data.user || ''), String((data.images || []).length)].join('|');
  return 'idem_' + a + '_' + Utilities.base64EncodeWebSafe(
    Utilities.computeDigest(Utilities.DigestAlgorithm.MD5, sig)).replace(/\W/g, '');
}

// ============ ENDPOINTS ============

// V44: bản đồ cột tab đơn giản THEO TÊN TIÊU ĐỀ (dòng 3) — sheet chèn thêm cột vẫn đúng chỗ
// V49: tiêu đề lệch (vd tab BISON: cột link ảnh ghi nhầm "SID", có 2 cột "SID")
//      → dò cột link theo DỮ LIỆU (ô chứa drive.google.com), SID = cột "SID" còn lại.
function simpleColMap(src) {
  var lastCol = Math.max(9, src.getLastColumn());
  var hd = [];
  try { hd = src.getRange(3, 1, 1, lastCol).getValues()[0]; } catch (e) { hd = []; }
  function nrm(x) { return String(x == null ? '' : x).trim().toLowerCase(); }
  function find(names, dflt) {
    for (var i = 0; i < hd.length; i++) {
      var h = nrm(hd[i]);
      if (!h) continue;
      for (var j = 0; j < names.length; j++) if (h === names[j] || h.indexOf(names[j]) === 0) return i + 1;
    }
    return dflt;
  }
  var link = find(['photo link', 'link ảnh'], 0);
  var sids = [];
  for (var si = 0; si < hd.length; si++) if (nrm(hd[si]) === 'sid') sids.push(si + 1);
  if (!link) {
    try {
      var nR = Math.min(40, Math.max(0, src.getLastRow() - 3));
      if (nR) {
        var sv = src.getRange(4, 1, nR, lastCol).getDisplayValues(), best = 0, bestN = 0;
        for (var c = 0; c < lastCol; c++) {
          var k = 0;
          for (var r = 0; r < nR; r++) if (String(sv[r][c]).indexOf('drive.google.com') >= 0) k++;
          if (k > bestN) { bestN = k; best = c + 1; }
        }
        if (best) link = best;
      }
    } catch (eL) {}
  }
  if (!link) link = 7;
  var sidOk = sids.filter(function (c) { return c !== link; });
  var sid = sidOk.length ? sidOk[0] : 9;
  return {
    lastCol: lastCol,
    no:    find(['no.', 'no', 'stt'], 1),
    sec:   find(['section', 'sec'], 2),
    part:  find(['part', 'member no'], 3),
    desc:  find(['description', 'mô tả', 'mo ta'], 4),
    rem:   find(['remarks', 'remark', 'ghi chú'], 5),
    photo: find(['photo after'], 0) === 6 ? 6 : find(['photo', 'ảnh phát hiện', 'hình'], 6),
    link:  link,
    st:    find(['status', 'tình trạng'], 8),
    sid:   sid,
    after: find(['photo after', 'ảnh khắc phục'], 10),
    sev:   find(['mức độ', 'muc do', 'severity'], 11),
    asg:   find(['người khắc phục', 'nguoi khac phuc', 'assignee'], 12),
    due:   find(['hạn xử lý', 'han xu ly', 'due'], 13),
    insp:  find(['người lập', 'nguoi lap', 'người kiểm', 'nguoi kiem', 'inspector', 'inspected'], 0)
  };
}

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

// V49: "close" / "Close" / "closed" đều là ĐÃ ĐÓNG
function plIsClosed_(v) { return /^clos/i.test(String(v == null ? '' : v).trim()); }

// V50: thêm ID ảnh SAU vào ghi chú "AFTER:..." của ô Photo after (giữ ghi chú cũ, không trùng)
function plAddAfterNote_(cell, ids) {
  if (!cell || !ids || !ids.length) return;
  var note = String(cell.getNote() || '');
  var m = note.match(/AFTER:([-\w,]*)/);
  var cur = m && m[1] ? m[1].split(',').filter(String) : [];
  ids.forEach(function (id) { if (id && cur.indexOf(id) < 0) cur.push(id); });
  var line = 'AFTER:' + cur.join(',');
  cell.setNote(m ? note.replace(/AFTER:[-\w,]*/, line) : (note ? note + '\n' + line : line));
}

function doPost(e) {
  try {
    var data = JSON.parse(e.postData.contents);
    if (data.key !== SECRET_KEY) return reply({ ok: false, error: 'Sai khóa bảo mật' });

    // ===== V43 — KIỂM SOÁT XOÁ: action xoá dữ liệu phải kèm đúng mã PIN quản trị =====
    var PIN_ACTS = { punchDelete: 1, punchClearAll: 1, punchDedupe: 1 };
    if (ADMIN_PIN && PIN_ACTS[data.action] && !data.dryRun && String(data.pin || '') !== ADMIN_PIN)
      return reply({ ok: false, error: 'Sai hoặc thiếu mã PIN quản trị — chỉ người quản lý mới được xoá.', needPin: true });

    // ===== V16 — CHỐNG GHI TRÙNG =====
    // Mạng yếu / HTTP 404 / quá 90 giây → app tự gửi lại, trong khi máy chủ ĐÃ ghi xong dòng đó
    // → trước đây sinh dòng + ẢNH TRÙNG trên Sheet và trong PDF.
    // Nay: cùng reqId (app mới) hoặc cùng nội dung (app cũ) chỉ ghi 1 lần; lần gửi lại trả kết quả cũ.
    if (IDEM_ACTS[data.action]) {
      var idemK = idemKey(data);
      if (idemK) {
        var idemC = CacheService.getScriptCache();
        var idemV = idemC.get(idemK);
        if (idemV === 'RUNNING')
          return reply({ ok: false, busy: true,
            error: 'Mục này ĐANG được ghi (lần gửi trước chưa xong) — chờ ~15 giây rồi bấm 🔄 XEM BÁO CÁO để kiểm, đừng gửi lại.' });
        if (idemV) { var idemR = JSON.parse(idemV); idemR.duplicated = true; return reply(idemR); }
        idemC.put(idemK, 'RUNNING', 300);
        _IDEM_KEY = idemK;
        _IDEM_TTL = data.reqId ? 21600 : 300;   // có reqId → nhớ 6h; suy từ nội dung → chỉ 5 phút
      }
    }

    // ===== V18 — ĐA NGƯỜI DÙNG =====
    // Danh sách tên chuẩn cho app (chỉ trả tên, KHÔNG trả PIN/vai trò)
    // V47: dinh tuyen tab theo DU AN — BISON mac dinh giu tab cu; du an khac = '- <proj>'. Reset base moi request (chong warm-instance).
    (function(){ var _p = (data && data.proj != null) ? String(data.proj).trim() : ''; var _sfx = (_p && _p !== PUNCH_PROJECT_NAME) ? (' - ' + _p) : ''; SIMPLE_TAB = plSimpleTabName_(_p); PUNCH_TAB = 'Punch List' + _sfx; })();   // V49: tab đơn giản tìm theo tên dự án
    // V48: xoa dung KHOI 1 project trong 'Danh mục part' (Project ghi o dong dau, duoi ke thua). Co dryRun.
    if (data.action === 'catDelProj') {
      var cdName = String(data.project || '').trim().toUpperCase();
      if (!cdName) return reply({ ok: false, error: 'Thieu project' });
      var cdCat = SpreadsheetApp.openById(SHEET_ID).getSheetByName('Danh mục part');
      if (!cdCat) return reply({ ok: false, error: 'Khong co tab Danh muc part' });
      var cdLast = cdCat.getLastRow();
      if (cdLast < 2) return reply({ ok: true, deleted: 0 });
      var cdA = cdCat.getRange(1, 1, cdLast, 1).getValues();
      var cdCur = '', cdDel = [];
      for (var cdr = 2; cdr <= cdLast; cdr++) { var cda = String(cdA[cdr - 1][0] || '').trim(); if (cda) cdCur = cda; if (cdCur.toUpperCase() === cdName) cdDel.push(cdr); }
      if (data.dryRun) return reply({ ok: true, wouldDelete: cdDel.length, firstRow: cdDel[0] || 0, lastRow: cdDel[cdDel.length - 1] || 0 });
      for (var cdi = cdDel.length - 1; cdi >= 0; cdi--) cdCat.deleteRow(cdDel[cdi]);
      try { CacheService.getScriptCache().remove('catalog_v1'); } catch (eCd) {}
      return reply({ ok: true, deleted: cdDel.length });
    }

    if (data.action === 'listUsers') {
      return reply({ ok: true, users: listUserNames() });
    }
    // Thêm người mới từ app. Nếu đã thiết lập PIN Sếp thì bắt buộc kèm PIN đúng.
    if (data.action === 'addUser') {
      var nuName = String(data.name || '').trim().replace(/\s+/g, ' ');
      if (!nuName) return reply({ ok: false, error: 'Thiếu tên' });
      var nuUsers = getUsers();
      if (pinConfigured(nuUsers) && !validPin(nuUsers, data.pin))
        return reply({ ok: false, error: 'Cần mã PIN của Sếp để thêm người mới' });
      var nuCanon = canonicalUser(nuName);
      var nuExists = nuUsers.filter(function (u) { return u.name === nuCanon; }).length > 0;
      if (!nuExists) getUsersSheet().appendRow([nuName, 'QC', '', '', '', '']);
      return reply({ ok: true, users: listUserNames(), name: nuExists ? nuCanon : nuName });
    }
    // V24: xuất báo cáo PDF (chức năng Sếp — cần PIN nếu đã thiết lập)
    if (data.action === 'makePdf') {
      var mpUsers = getUsers();
      if (pinConfigured(mpUsers) && !validPin(mpUsers, data.pin))
        return reply({ ok: false, error: 'Chức năng này cần mã PIN của Sếp.' });
      try {
        var mp = makePdfReport(String(data.project || '__all__'), Number(data.days) || 0);
        return reply({ ok: true, url: mp.url, name: mp.name });
      } catch (em) {
        return reply({ ok: false, error: String(em.message || em) });
      }
    }
    // V26: Sếp duyệt báo cáo — verdict 'ok' (Đạt) | 'redo' (Chụp lại + lý do)
    if (data.action === 'reviewReport') {
      var rvUsers = getUsers();
      if (pinConfigured(rvUsers) && !validPin(rvUsers, data.pin))
        return reply({ ok: false, error: 'Chức năng này cần mã PIN của Sếp.' });
      var rvId = String(data.reportId || '').trim();
      if (!rvId) return reply({ ok: false, error: 'Thiếu reportId' });
      var rvText = (data.verdict === 'ok') ? 'Đạt' : ('Chụp lại: ' + String(data.reason || '').trim());
      var rvSs = SpreadsheetApp.openById(SHEET_ID);
      var rvRow = null;
      rvSs.getSheets().forEach(function (sh) {
        var nm = sh.getName();
        if (nm === PROJECTS_TAB || nm === 'Dashboard' || nm === AI_QUEUE_TAB || nm === USERS_TAB) return;
        if (String(sh.getRange(1, 1).getValue() || '').indexOf('Thời gian') === -1) return;
        var lastR = sh.getLastRow();
        if (lastR < 2) return;
        if (!String(sh.getRange(1, 22).getValue() || '')) sh.getRange(1, 22).setValue('Duyệt');
        var ids2 = sh.getRange(2, 18, lastR - 1, 1).getValues();
        for (var ri2 = 0; ri2 < ids2.length; ri2++) {
          if (String(ids2[ri2][0]) !== rvId) continue;
          sh.getRange(ri2 + 2, 22).setValue(rvText)
            .setBackground(data.verdict === 'ok' ? '#e6f4ea' : '#fce8e6');
          if (nm === ALL_TAB) rvRow = sh.getRange(ri2 + 2, 1, 1, 21).getValues()[0];
        }
      });
      if (!rvRow) return reply({ ok: false, error: 'Không tìm thấy báo cáo' });
      if (data.verdict !== 'ok') {
        try { sendRedoAlert(rvRow, String(data.reason || '')); } catch (er2) {}
      }
      return reply({ ok: true, review: rvText });
    }
    // V30: danh sách báo cáo CỦA TÔI đang bị yêu cầu chụp lại (cho banner nhắc trong app)
    if (data.action === 'myRedo') {
      var mrUsers = getUsers();
      var mrDev = userByDevice(mrUsers, String(data.deviceId || '').trim());
      var mrName = mrDev ? mrDev.name : canonicalUser(data.user);
      var mrOut = [];
      if (mrName) {
        var mrSs = SpreadsheetApp.openById(SHEET_ID);
        var mrSh = mrSs.getSheetByName(ALL_TAB);
        if (mrSh && mrSh.getLastRow() > 1) {
          var mrV = mrSh.getRange(2, 1, mrSh.getLastRow() - 1, 22).getValues();
          for (var mi = mrV.length - 1; mi >= 0 && mrOut.length < 10; mi--) {
            if (String(mrV[mi][21] || '').indexOf('Chụp lại') !== 0) continue;
            if (String(mrV[mi][16] || '') !== mrName) continue;
            mrOut.push({
              time: fmtCell(mrV[mi][0]), project: String(mrV[mi][1] || ''),
              unit: [mrV[mi][18], mrV[mi][19], mrV[mi][20]].map(function (x) { return String(x || '').trim(); })
                .filter(Boolean).join(' / '),
              reason: String(mrV[mi][21] || '')
            });
          }
        }
      }
      return reply({ ok: true, items: mrOut });
    }
    // Kiểm tra PIN Sếp để mở tab Quản lý. Chưa thiết lập PIN nào → mở tự do (open).
    if (data.action === 'verifyPin') {
      var vpUsers = getUsers();
      if (!pinConfigured(vpUsers)) return reply({ ok: true, open: true });
      if (validPin(vpUsers, data.pin)) return reply({ ok: true });
      return reply({ ok: false, error: 'Sai mã PIN' });
    }
    // V19: ĐĂNG KÝ TÊN THEO ID MÁY — nhập tên 1 lần, máy đó = người đó
    if (data.action === 'registerUser') {
      var rgName = String(data.name || '').trim().replace(/\s+/g, ' ');
      if (!rgName) return reply({ ok: false, error: 'Thiếu tên' });
      var rgDev = String(data.deviceId || '').trim();
      var rgUsers = getUsers();
      var rgSheet = getUsersSheet();
      if (!rgSheet.getRange(1, 7).getValue()) rgSheet.getRange(1, 7).setValue('Thiết bị (ID)');
      var rgOwner = userByDevice(rgUsers, rgDev);
      if (rgOwner && normName(rgOwner.name) !== normName(rgName)) {
        if (pinConfigured(rgUsers) && !validPin(rgUsers, data.pin))
          return reply({ ok: false, needPin: true, name: rgOwner.name,
            error: 'Máy này đã đăng ký cho "' + rgOwner.name + '" — đổi tên cần mã PIN của Sếp' });
        var rgKeep = rgOwner.devices.filter(function (x) { return x !== rgDev; });
        rgSheet.getRange(rgOwner.row, 7).setValue(rgKeep.join(', '));
        rgUsers = getUsers();
      }
      var rgU = null, rgN = normName(rgName);
      for (var rgi = 0; rgi < rgUsers.length && !rgU; rgi++) {
        if (normName(rgUsers[rgi].name) === rgN) { rgU = rgUsers[rgi]; break; }
        for (var rgj = 0; rgj < rgUsers[rgi].aliases.length; rgj++)
          if (normName(rgUsers[rgi].aliases[rgj]) === rgN) { rgU = rgUsers[rgi]; break; }
      }
      if (rgU) {
        if (rgDev && rgU.devices.indexOf(rgDev) === -1) {
          rgU.devices.push(rgDev);
          rgSheet.getRange(rgU.row, 7).setValue(rgU.devices.join(', '));
        }
        return reply({ ok: true, name: rgU.name });
      }
      rgSheet.appendRow([rgName, 'QC', '', '', '', '', rgDev]);
      return reply({ ok: true, name: rgName });
    }

    // Thêm dự án mới từ app (nút "➕ Thêm dự án")
    if (data.action === 'addProject') {
      var pname = String(data.name || '').trim();
      if (!pname) return reply({ ok: false, error: 'Thiếu tên dự án' });
      var existing = getProjects().map(function (p) { return p.name; });
      if (existing.indexOf(pname) === -1) {
        SpreadsheetApp.openById(SHEET_ID).getSheetByName(PROJECTS_TAB)
          .appendRow([pname, '', '', '']);
      }
      return reply({ ok: true, projects: getProjects().map(function (p) { return p.name; }) });
    }

    // Danh sách dự án kèm hạng mục riêng (cho ô chọn Hạng mục) + danh sách người dùng
    if (data.action === 'projectsFull') {
      return reply({ ok: true, projects: getProjects(), users: listUserNames() });
    }

    // Thêm hạng mục cho 1 dự án (nút ➕ cạnh ô Hạng mục)
    if (data.action === 'addCategory') {
      var acp = String(data.project || '').trim();
      var acc = String(data.category || '').trim();
      if (!acp || !acc) return reply({ ok: false, error: 'Thiếu dự án hoặc hạng mục' });
      var acsh = SpreadsheetApp.openById(SHEET_ID).getSheetByName(PROJECTS_TAB);
      var acrows = acsh.getDataRange().getValues();
      for (var aci = 1; aci < acrows.length; aci++) {
        if (String(acrows[aci][0]) === acp) {
          var cur = String(acrows[aci][5] || '').split(',').map(function (s) { return s.trim(); }).filter(Boolean);
          if (cur.indexOf(acc) === -1) cur.push(acc);
          acsh.getRange(aci + 1, 6).setValue(cur.join(', '));
          return reply({ ok: true, categories: cur });
        }
      }
      return reply({ ok: false, error: 'Không thấy dự án: ' + acp });
    }

    // Tìm thư mục Drive theo tên (để chọn trực tiếp trong app, không cần copy link)
    if (data.action === 'searchFolders') {
      var query = String(data.query || '').trim().replace(/'/g, "\\'");
      var it = query
        ? DriveApp.searchFolders("title contains '" + query + "' and trashed = false")
        : DriveApp.getFolders();
      var out = [];
      while (it.hasNext() && out.length < 30) {
        var f = it.next(), parent = '';
        try { var ps = f.getParents(); if (ps.hasNext()) parent = ps.next().getName(); } catch (e) {}
        out.push({ id: f.getId(), name: f.getName(), parent: parent });
      }
      out.sort(function (a, b) { return a.name.toLowerCase() < b.name.toLowerCase() ? -1 : 1; });
      return reply({ ok: true, folders: out });
    }

    // Đặt thư mục Drive lưu ảnh cho 1 dự án (nút "📁 Thư mục lưu")
    if (data.action === 'setProjectFolder') {
      var pn = String(data.project || '').trim();
      var link = String(data.folder || '').trim();
      var fidMatch = link.match(/folders\/([a-zA-Z0-9_-]+)/);
      var id = fidMatch ? fidMatch[1] : link.replace(/[^a-zA-Z0-9_-]/g, '');
      var fname = '';
      try { fname = DriveApp.getFolderById(id).getName(); }
      catch (e) { return reply({ ok: false, error: 'Không mở được thư mục (kiểm tra link và quyền truy cập).' }); }
      var sh = SpreadsheetApp.openById(SHEET_ID).getSheetByName(PROJECTS_TAB);
      var rows = sh.getDataRange().getValues();
      for (var i = 1; i < rows.length; i++) {
        if (String(rows[i][0]) === pn) { sh.getRange(i + 1, 5).setValue(id);
          return reply({ ok: true, folderName: fname }); }
      }
      return reply({ ok: false, error: 'Không tìm thấy dự án: ' + pn });
    }

    // Trợ lý AI hỏi-đáp dữ liệu báo cáo (cho sếp)
    if (data.action === 'ask') {
      if (!GEMINI_API_KEY) return reply({ ok: false, error: 'Chưa bật AI (thiếu GEMINI_API_KEY)' });
      // V18: nếu đã thiết lập PIN Sếp thì hỏi AI cũng cần PIN (chống tốn quota AI)
      var askUsers = getUsers();
      if (pinConfigured(askUsers) && !validPin(askUsers, data.pin))
        return reply({ ok: false, error: 'Chức năng này cần mã PIN của Sếp (mở tab Quản lý để nhập).' });
      var q = String(data.question || '').trim();
      if (!q) return reply({ ok: false, error: 'Thiếu câu hỏi' });
      var ssQ = SpreadsheetApp.openById(SHEET_ID);
      var shQ = ssQ.getSheetByName(ALL_TAB);
      if (!shQ || shQ.getLastRow() < 2) return reply({ ok: true, answer: 'Chưa có dữ liệu báo cáo.' });
      var nQ = Math.min(shQ.getLastRow() - 1, 500);
      var valsQ = shQ.getRange(2, 1, nQ, HEADER.length).getValues();
      var lines = valsQ.map(function (r) {
        return [String(r[0]), 'DỰÁN:' + r[1], 'HẠNGMỤC:' + r[2], 'NGƯỜI:' + r[16],
                r[20] ? 'CẤUKIỆN:' + r[20] : '', r[4] ? 'CẢNHBÁO:' + r[4] : '']
               .filter(Boolean).join(' | ');
      });
      var prompt = 'Bạn là trợ lý quản lý QC hiện trường. Dưới đây là dữ liệu báo cáo (mỗi dòng 1 báo cáo). ' +
        'Trả lời câu hỏi của sếp NGẮN GỌN, chính xác, bằng tiếng Việt, CHỈ dựa trên dữ liệu. ' +
        'Khi cần đếm/tổng hợp hãy tính chính xác. Hôm nay là ' +
        Utilities.formatDate(new Date(), 'GMT+7', 'yyyy-MM-dd') + '.\n\nDỮ LIỆU:\n' +
        lines.join('\n') + '\n\nCÂU HỎI: ' + q;
      var askOpts = { method: 'post', contentType: 'application/json', muteHttpExceptions: true,
          payload: JSON.stringify({ contents: [{ parts: [{ text: prompt }] }],
                                    generationConfig: { temperature: 0.2 } }) };
      var ans = '';
      // V23: lần 1 model mới, lần 2 model dự phòng (kiêm thử lại khi dính giới hạn tốc độ)
      for (var at = 0; at < 2 && !ans; at++) {
        if (at > 0) Utilities.sleep(3000);
        var askModel = (at === 0) ? GEMINI_MODEL : GEMINI_MODEL_FALLBACK;
        var askUrl = 'https://generativelanguage.googleapis.com/v1beta/models/' + askModel + ':generateContent?key=' + GEMINI_API_KEY;
        try {
          var bd = JSON.parse(UrlFetchApp.fetch(askUrl, askOpts).getContentText());
          if (bd.candidates && bd.candidates[0] && bd.candidates[0].content)
            ans = bd.candidates[0].content.parts[0].text;
        } catch (e) {}
      }
      if (!ans) return reply({ ok: false, error: 'AI đang bận (giới hạn tốc độ). Thử lại sau vài giây.' });
      return reply({ ok: true, answer: ans });
    }

    // Trả danh sách báo cáo cho tab Tra cứu (mới nhất trước, tối đa 300)
    if (data.action === 'listReports') {
      var proj = String(data.project || '').trim();
      var ssR = SpreadsheetApp.openById(SHEET_ID);
      var shR = (proj && proj !== '__all__') ? ssR.getSheetByName(proj) : ssR.getSheetByName(ALL_TAB);
      if (!shR || shR.getLastRow() < 2) return reply({ ok: true, reports: [] });
      var vals = shR.getRange(2, 1, shR.getLastRow() - 1, HEADER.length + 1).getValues(); // +1: cột 22 Duyệt (V26)
      // V27: phân trang — limit/offset (mặc định 300 như cũ, app cũ vẫn chạy)
      var rpLimit = Math.min(Number(data.limit) || 300, 300);
      var rpOffset = Math.max(Number(data.offset) || 0, 0);
      var reports = [];
      for (var ri = vals.length - 1 - rpOffset; ri >= 0 && reports.length < rpLimit; ri--) {
        var rr = vals[ri];
        reports.push({
          time: fmtCell(rr[0]), project: String(rr[1]), category: String(rr[2]),
          description: String(rr[3]), issues: String(rr[4]), note: String(rr[5]),
          images: String(rr[15] || '').split('\n').filter(function (x) { return x; }),
          user: String(rr[16]), model: String(rr[18] || ''),
          gate: String(rr[19] || ''), unit: String(rr[20] || ''),
          reportId: String(rr[17] || ''),   // V26: để Sếp duyệt từng báo cáo
          review: String(rr[21] || '')      // V26: '' | 'Đạt' | 'Chụp lại: <lý do>'
        });
      }
      return reply({ ok: true, reports: reports, total: vals.length,
        hasMore: (rpOffset + reports.length) < vals.length });  // V27
    }

    // ============ PUNCH LIST — MỞ LỖI (chụp hiện trường → 1 dòng 'open' theo mẫu VIOLA) ============
    if (data.action === 'punchOpen') {
      var poProject = (data.project && data.project !== 'auto') ? data.project
                    : (matchProjectByGps(data.lat, data.lng) || 'Chưa phân loại');
      var poImgs = data.images || (data.imageBase64 ? [data.imageBase64] : []);
      var poOpenIds = [], poOpenUrls = [];
      if (poImgs.length) {
        var poMonth = Utilities.formatDate(new Date(), 'GMT+7', 'yyyy-MM');
        var poDay   = Utilities.formatDate(new Date(), 'GMT+7', 'yyyy-MM-dd');
        var poFolder = getReportFolder(poProject, poMonth, poDay);
        var poStamp = Utilities.formatDate(new Date(), 'GMT+7', 'yyyyMMdd_HHmmss');
        poImgs.forEach(function (b64, i) {
          var f = poFolder.createFile(Utilities.newBlob(Utilities.base64Decode(b64), 'image/jpeg', 'PL_' + poStamp + '_' + (i + 1) + '.jpg'));
          poOpenIds.push(f.getId()); poOpenUrls.push(f.getUrl());
        });
      }
      var poSs = SpreadsheetApp.openById(SHEET_ID);
      var poQ = ensurePunchList(poSs);
      var poNo = nextPunchNo(poQ);
      var poId = 'PL' + Utilities.formatDate(new Date(), 'GMT+7', 'yyMMddHHmmss') + Math.floor(Math.random() * 90 + 10);
      poQ.appendRow([
        data.commodityCode || '',                                   // A COMMODITY CODE
        data.category || '',                                        // B DESCRIPTION (hạng mục)
        data.vendor || '',                                          // C VENDOR
        'open',                                                     // D STATUS
        Utilities.formatDate(new Date(), 'GMT+7', 'yyyy-MM-dd'),    // E DATE
        senderName(data),                                           // F INSPECTED BY
        '#' + poNo,                                                 // G ITEM #
        data.description || data.note || '',                        // H DESCRIPTION (lỗi)
        '', '', '',                                                 // I RESULT | J ACTION | K CLOSING DATE
        '',                                                         // L Open Photos (điền công thức bên dưới)
        '',                                                         // M Closing photos
        data.workshop || data.remark || '',                        // N REMARK ACTION WORKSHOP
        poId,                                                       // O PUNCH_ID
        poOpenIds.join(','),                                        // P OPEN_IDS
        '',                                                         // Q CLOSE_IDS
        data.reportId || ''                                         // R REPORT_ID
      ]);
      var poRow = poQ.getLastRow();
      if (poOpenIds.length) { poQ.getRange(poRow, 12).setFormula(punchImage(poOpenIds[0])); poQ.setRowHeight(poRow, 130); }
      // Ghi ĐỒNG THỜI vào tab ĐƠN GIẢN (cùng ID → đóng 1 lần là cả 2 sheet cùng close)
      try {
        var poSp = ensurePunchSimple(poSs);
        upgradeSimpleCols(poSp);
        var poSpNo = nextSimpleNo(poSp);
        var poPart = canonPart(data.catProject || data.project, data.section, data.part);
        // V46: ghi tab đơn giản THEO TÊN cột (khớp cả khi user chèn thêm cột G/H) + cột Người lập (đa QC)
        var poM = simpleColMap(poSp);
        var poArr = [];
        for (var poc = 0; poc < poM.lastCol; poc++) poArr.push('');
        poArr[poM.no - 1]   = poSpNo;
        poArr[poM.sec - 1]  = data.section || '';
        poArr[poM.part - 1] = poPart;
        poArr[poM.desc - 1] = data.description || data.note || '';
        poArr[poM.rem - 1]  = data.remarks || data.workshop || '';
        poArr[poM.link - 1] = poOpenUrls.join('\n');
        poArr[poM.st - 1]   = 'open';
        poArr[poM.sid - 1]  = poId;
        if (poM.sev)  poArr[poM.sev - 1]  = data.severity || '';
        if (poM.asg)  poArr[poM.asg - 1]  = data.assignee || '';
        if (poM.due)  poArr[poM.due - 1]  = data.due || '';
        if (poM.insp) poArr[poM.insp - 1] = senderName(data);
        poSp.appendRow(poArr);
        var poSpRow = poSp.getLastRow();
        poSp.getRange(poSpRow, 1, 1, poM.lastCol).setBorder(true, true, true, true, true, true).setVerticalAlignment('middle');
        poSp.getRange(poSpRow, poM.desc).setWrap(true);
        poSp.getRange(poSpRow, 1, 1, 3).setHorizontalAlignment('center');
        poSp.getRange(poSpRow, poM.st).setHorizontalAlignment('center');
        if (poOpenIds.length) { poSp.getRange(poSpRow, poM.photo).setFormula(punchImage(poOpenIds[0])); poSp.setRowHeight(poSpRow, 110); }
      } catch (ePoSp) {}
      return reply({ ok: true, punchId: poId, itemNo: '#' + poNo, project: poProject,
                     openPhotos: poOpenUrls, photoCount: poOpenUrls.length });
    }

    // ============ PUNCH LIST — DANH SÁCH MỤC ĐANG MỞ (cho app chọn để đóng) ============
    if (data.action === 'listOpenPunch') {
      var lpSs = SpreadsheetApp.openById(SHEET_ID);
      var lpQ = lpSs.getSheetByName(PUNCH_TAB);
      var lpOut = [];
      // V42: tab VIOLA trống thì KHÔNG return sớm nữa — vẫn phải liệt kê tab đơn giản bên dưới
      if (lpQ && lpQ.getLastRow() >= 8) {
        var lpVals = lpQ.getRange(8, 1, lpQ.getLastRow() - 7, 18).getValues();
        for (var lpi = lpVals.length - 1; lpi >= 0 && lpOut.length < 100; lpi--) {
          if (String(lpVals[lpi][3]).toLowerCase() !== 'open') continue;   // D STATUS
          lpOut.push({
            punchId: String(lpVals[lpi][14]),  itemNo: String(lpVals[lpi][6]),
            project: String(lpVals[lpi][1]),   vendor: String(lpVals[lpi][2]),
            description: String(lpVals[lpi][7]), date: fmtCell(lpVals[lpi][4]),
            workshop: String(lpVals[lpi][13]), inspector: String(lpVals[lpi][5])
          });
        }
      }
      // Cộng thêm mục OPEN của tab ĐƠN GIẢN (đóng chung 1 màn trong app)
      try {
        var lsQ = lpSs.getSheetByName(SIMPLE_TAB);
        if (lsQ && lsQ.getLastRow() >= 4) {
          var lsM = simpleColMap(lsQ);   // V44: cột theo tên tiêu đề
          var lsV = lsQ.getRange(4, 1, lsQ.getLastRow() - 3, lsM.lastCol).getValues();
          // V43: dòng chưa có ID bền → cấp ngay; đóng/xoá bám ID, không bám số dòng (chống lệch khi có người xoá)
          var lsStamp = 'SP' + Date.now();
          for (var lb = 0; lb < lsV.length; lb++) {
            if (!String(lsV[lb][lsM.sid - 1] || '').trim() && String(lsV[lb].join('')).trim()) {
              var lsNid = lsStamp + '_' + (4 + lb);
              lsQ.getRange(4 + lb, lsM.sid).setValue(lsNid);
              lsV[lb][lsM.sid - 1] = lsNid;
            }
          }
          for (var lsi = lsV.length - 1; lsi >= 0 && lpOut.length < 150; lsi--) {
            if (plIsClosed_(lsV[lsi][lsM.st - 1])) continue;   // V49: close/Close/closed
            if (String(lsV[lsi][lsM.sid - 1] || '').indexOf('PL') === 0) continue; // dòng liên kết VIOLA — đã hiện ở trên, tránh trùng
            lpOut.push({ punchId: String(lsV[lsi][lsM.sid - 1] || ('SPROW' + (4 + lsi))), itemNo: 'No.' + lsV[lsi][lsM.no - 1],
              project: 'Đơn giản', vendor: String(lsV[lsi][lsM.part - 1]), description: String(lsV[lsi][lsM.desc - 1]),
              date: '', workshop: String(lsV[lsi][lsM.sec - 1]), inspector: '', simple: true });
          }
        }
      } catch (eLs) {}
      return reply({ ok: true, items: lpOut });
    }

    // ============ PUNCH LIST — ĐÓNG LỖI (chụp ảnh nghiệm thu → cập nhật đúng dòng) ============
    if (data.action === 'punchClose') {
      var pcId = String(data.punchId || '').trim();
      if (!pcId) return reply({ ok: false, error: 'Thiếu punchId' });
      // Mục của tab ĐƠN GIẢN (id SP.../SPROW...) → đóng tại tab đơn giản
      if (/^SP/.test(pcId)) {
        var scSs = SpreadsheetApp.openById(SHEET_ID);
        var scQ = scSs.getSheetByName(SIMPLE_TAB);
        if (!scQ) return reply({ ok: false, error: 'Chưa có tab đơn giản' });
        var scM = simpleColMap(scQ);   // V44: cột theo tên tiêu đề
        var scRow = -1;
        if (pcId.indexOf('SPROW') === 0) scRow = parseInt(pcId.slice(5), 10);
        else {
          var scIds = scQ.getRange(4, scM.sid, Math.max(1, scQ.getLastRow() - 3), 1).getValues();
          for (var sci = 0; sci < scIds.length; sci++) if (String(scIds[sci][0]) === pcId) { scRow = 4 + sci; break; }
        }
        if (scRow < 4 || scRow > scQ.getLastRow()) return reply({ ok: false, error: 'Không tìm thấy mục ' + pcId });
        if (plIsClosed_(scQ.getRange(scRow, scM.st).getValue()))   // V49: close/Close/closed
          return reply({ ok: false, error: 'Mục này đã đóng trước đó' });
        var scImgs = data.images || (data.imageBase64 ? [data.imageBase64] : []);
        var scUrls = [], scCloseIds = [];   // V50: giữ ID ảnh khắc phục để ghi chú AFTER
        if (scImgs.length) {
          var scMonth = Utilities.formatDate(new Date(), 'GMT+7', 'yyyy-MM');
          var scDay = Utilities.formatDate(new Date(), 'GMT+7', 'yyyy-MM-dd');
          var scFolder = getReportFolder('Don gian', scMonth, scDay);
          var scStamp = Utilities.formatDate(new Date(), 'GMT+7', 'yyyyMMdd_HHmmss');
          scImgs.forEach(function (b64, i) {
            var f = scFolder.createFile(Utilities.newBlob(Utilities.base64Decode(b64), 'image/jpeg', 'PLSclose_' + scStamp + '_' + (i + 1) + '.jpg'));
            scUrls.push(f.getUrl()); scCloseIds.push(f.getId());
          });
        }
        scQ.getRange(scRow, scM.st).setValue('close');
        var scRem = String(scQ.getRange(scRow, scM.rem).getValue() || '');
        var scNote = 'CLOSED ' + Utilities.formatDate(new Date(), 'GMT+7', 'yyyy-MM-dd') + (data.actionText ? (': ' + data.actionText) : '');
        scQ.getRange(scRow, scM.rem).setValue(scRem ? (scRem + '\n' + scNote) : scNote).setWrap(true);
        if (scUrls.length) {
          var scLinks = String(scQ.getRange(scRow, scM.link).getValue() || '');
          scQ.getRange(scRow, scM.link).setValue((scLinks ? scLinks + '\n' : '') + scUrls.join('\n'));
          upgradeSimpleCols(scQ);
          var scFid = (String(scUrls[0]).match(/[-\w]{25,}/) || [''])[0];
          if (scFid && scM.after) scQ.getRange(scRow, scM.after).setFormula(punchImage(scFid));   // ẢNH KHẮC PHỤC (before/after)
          if (scM.after) { try { plAddAfterNote_(scQ.getRange(scRow, scM.after), scCloseIds); } catch (eAn) {} }   // V50
        }
        return reply({ ok: true, punchId: pcId, closePhotos: scUrls, photoCount: scUrls.length });
      }
      var pcSs = SpreadsheetApp.openById(SHEET_ID);
      var pcQ = pcSs.getSheetByName(PUNCH_TAB);
      if (!pcQ || pcQ.getLastRow() < 8) return reply({ ok: false, error: 'Chưa có mục Punch List nào' });
      var pcIds = pcQ.getRange(8, 15, pcQ.getLastRow() - 7, 1).getValues();   // O PUNCH_ID
      var pcTarget = -1;
      for (var pci = 0; pci < pcIds.length; pci++) { if (String(pcIds[pci][0]) === pcId) { pcTarget = 8 + pci; break; } }
      if (pcTarget < 0) return reply({ ok: false, error: 'Không tìm thấy mục ' + pcId });
      if (String(pcQ.getRange(pcTarget, 4).getValue()).toLowerCase() === 'close')
        return reply({ ok: false, error: 'Mục này đã đóng trước đó' });
      var pcProject = String(pcQ.getRange(pcTarget, 2).getValue() || 'Chưa phân loại');
      var pcImgs = data.images || (data.imageBase64 ? [data.imageBase64] : []);
      var pcCloseIds = [], pcCloseUrls = [];
      if (pcImgs.length) {
        var pcMonth = Utilities.formatDate(new Date(), 'GMT+7', 'yyyy-MM');
        var pcDay   = Utilities.formatDate(new Date(), 'GMT+7', 'yyyy-MM-dd');
        var pcFolder = getReportFolder(pcProject, pcMonth, pcDay);
        var pcStamp = Utilities.formatDate(new Date(), 'GMT+7', 'yyyyMMdd_HHmmss');
        pcImgs.forEach(function (b64, i) {
          var f = pcFolder.createFile(Utilities.newBlob(Utilities.base64Decode(b64), 'image/jpeg', 'PLclose_' + pcStamp + '_' + (i + 1) + '.jpg'));
          pcCloseIds.push(f.getId()); pcCloseUrls.push(f.getUrl());
        });
      }
      pcQ.getRange(pcTarget, 4).setValue('close');                                              // D STATUS
      pcQ.getRange(pcTarget, 9).setValue(data.result || 'Accepted');                            // I RESULT
      if (data.actionText) pcQ.getRange(pcTarget, 10).setValue(String(data.actionText));        // J ACTION
      pcQ.getRange(pcTarget, 11).setValue(Utilities.formatDate(new Date(), 'GMT+7', 'yyyy-MM-dd')); // K CLOSING DATE
      if (pcCloseIds.length) {
        pcQ.getRange(pcTarget, 13).setFormula(punchImage(pcCloseIds[0]));                       // M Closing photos
        pcQ.getRange(pcTarget, 17).setValue(pcCloseIds.join(','));                              // Q CLOSE_IDS
        pcQ.setRowHeight(pcTarget, 130);
      }
      // Đồng bộ CLOSE sang tab ĐƠN GIẢN (dòng cùng ID, nếu có)
      try {
        var pcSp = pcSs.getSheetByName(SIMPLE_TAB);
        if (pcSp && pcSp.getLastRow() >= 4) {
          var pcSm = simpleColMap(pcSp);   // V49: cột theo tên tiêu đề
          var pcSids = pcSp.getRange(4, pcSm.sid, pcSp.getLastRow() - 3, 1).getValues();
          for (var pcsi = 0; pcsi < pcSids.length; pcsi++) {
            if (String(pcSids[pcsi][0]) === pcId) {
              var pcSpRow = 4 + pcsi;
              pcSp.getRange(pcSpRow, pcSm.st).setValue('close');
              var pcRem = String(pcSp.getRange(pcSpRow, pcSm.rem).getValue() || '');
              var pcNote = 'CLOSED ' + Utilities.formatDate(new Date(), 'GMT+7', 'yyyy-MM-dd') + (data.actionText ? (': ' + String(data.actionText)) : '');
              pcSp.getRange(pcSpRow, pcSm.rem).setValue(pcRem ? (pcRem + '\n' + pcNote) : pcNote).setWrap(true);
              if (pcCloseUrls.length) {
                var pcLk = String(pcSp.getRange(pcSpRow, pcSm.link).getValue() || '');
                pcSp.getRange(pcSpRow, pcSm.link).setValue((pcLk ? pcLk + '\n' : '') + pcCloseUrls.join('\n'));
                upgradeSimpleCols(pcSp);
                if (pcCloseIds.length && pcSm.after) pcSp.getRange(pcSpRow, pcSm.after).setFormula(punchImage(pcCloseIds[0]));
                if (pcSm.after) { try { plAddAfterNote_(pcSp.getRange(pcSpRow, pcSm.after), pcCloseIds); } catch (eAn2) {} }   // V50
              }
              break;
            }
          }
        }
      } catch (ePcSp) {}
      return reply({ ok: true, punchId: pcId, closePhotos: pcCloseUrls, photoCount: pcCloseUrls.length });
    }

    // ============ PUNCH LIST — FORM ĐƠN GIẢN (No./Section/Part/Description/Remarks/Photo) ============
    if (data.action === 'punchSimple') {
      if (!String(data.description || '').trim()) return reply({ ok: false, error: 'Thiếu mô tả (Description)' });
      var spImgs = data.images || (data.imageBase64 ? [data.imageBase64] : []);
      var spIds = [], spUrls = [];
      if (spImgs.length) {
        var spMonth = Utilities.formatDate(new Date(), 'GMT+7', 'yyyy-MM');
        var spDay   = Utilities.formatDate(new Date(), 'GMT+7', 'yyyy-MM-dd');
        var spFolder = getReportFolder('Don gian', spMonth, spDay);
        var spStamp = Utilities.formatDate(new Date(), 'GMT+7', 'yyyyMMdd_HHmmss');
        spImgs.forEach(function (b64, i) {
          var f = spFolder.createFile(Utilities.newBlob(Utilities.base64Decode(b64), 'image/jpeg', 'PLS_' + spStamp + '_' + (i + 1) + '.jpg'));
          spIds.push(f.getId()); spUrls.push(f.getUrl());
        });
      }
      var spQ = ensurePunchSimple(SpreadsheetApp.openById(SHEET_ID));
      var spNo = nextSimpleNo(spQ);
      var spId = 'SP' + Utilities.formatDate(new Date(), 'GMT+7', 'yyMMddHHmmss') + Math.floor(Math.random() * 90 + 10);
      upgradeSimpleCols(spQ);
      var spPart = canonPart(data.catProject || data.project, data.section, data.part);
      // V44: ghi theo TÊN cột tiêu đề — user chèn thêm cột trên sheet vẫn ghi đúng chỗ
      var spM = simpleColMap(spQ);
      var spArr = [];
      for (var spc = 0; spc < spM.lastCol; spc++) spArr.push('');
      spArr[spM.no - 1] = spNo;
      spArr[spM.sec - 1] = data.section || '';
      spArr[spM.part - 1] = spPart;
      spArr[spM.desc - 1] = data.description || '';
      spArr[spM.rem - 1] = data.remarks || '';
      spArr[spM.link - 1] = spUrls.join('\n');
      spArr[spM.st - 1] = 'open';
      spArr[spM.sid - 1] = spId;
      if (spM.sev) spArr[spM.sev - 1] = data.severity || '';
      if (spM.asg) spArr[spM.asg - 1] = data.assignee || '';
      if (spM.due) spArr[spM.due - 1] = data.due || '';
      if (spM.insp) spArr[spM.insp - 1] = senderName(data);   // V46: ghi NGƯỜI LẬP (đa QC)
      spQ.appendRow(spArr);
      var spRow = spQ.getLastRow();
      spQ.getRange(spRow, 1, 1, spM.lastCol).setBorder(true, true, true, true, true, true).setVerticalAlignment('middle');
      spQ.getRange(spRow, spM.desc).setWrap(true);
      spQ.getRange(spRow, 1, 1, 3).setHorizontalAlignment('center');
      spQ.getRange(spRow, spM.st).setHorizontalAlignment('center');
      if (spIds.length) { spQ.getRange(spRow, spM.photo).setFormula(punchImage(spIds[0])); spQ.setRowHeight(spRow, 110); }
      return reply({ ok: true, no: spNo, id: spId, photoCount: spUrls.length, photos: spUrls });
    }

    // ============ PUNCH LIST — XEM BÁO CÁO NHÁP (đọc tab đơn giản về app) ============
    if (data.action === 'listSimple') {
      var lsSs2 = SpreadsheetApp.openById(SHEET_ID);
      var lsQ2 = lsSs2.getSheetByName(SIMPLE_TAB);
      if (!lsQ2 || lsQ2.getLastRow() < 4) return reply({ ok: true, items: [] });
      var lsM2 = simpleColMap(lsQ2);   // V44: cột theo tên tiêu đề
      var lsV2 = lsQ2.getRange(4, 1, lsQ2.getLastRow() - 3, lsM2.lastCol).getValues();
      // V49: công thức cột "Photo after" → ID ảnh khắc phục đầu tiên (app tách ảnh TRƯỚC / SAU)
      var lsAf2 = [], lsAn2 = [];
      try { if (lsM2.after) lsAf2 = lsQ2.getRange(4, lsM2.after, lsV2.length, 1).getFormulas(); } catch (eAf) { lsAf2 = []; }
      try { if (lsM2.after) lsAn2 = lsQ2.getRange(4, lsM2.after, lsV2.length, 1).getNotes(); } catch (eAn) { lsAn2 = []; }   // V50
      // V43: cấp ID bền cho dòng chưa có (app luôn thao tác theo ID → khớp tuyệt đối với sheet)
      var lsStamp2 = 'SP' + Date.now();
      for (var lb2 = 0; lb2 < lsV2.length; lb2++) {
        if (!String(lsV2[lb2][lsM2.sid - 1] || '').trim() && String(lsV2[lb2].join('')).trim()) {
          var lsNid2 = lsStamp2 + '_' + (4 + lb2);
          lsQ2.getRange(4 + lb2, lsM2.sid).setValue(lsNid2);
          lsV2[lb2][lsM2.sid - 1] = lsNid2;
        }
      }
      var lsOut2 = [];
      for (var li2 = 0; li2 < lsV2.length; li2++) {
        var lr = lsV2[li2];
        var lIds = String(lr[lsM2.link - 1] || '').split('\n').map(function (u) { var m = String(u).match(/[-\w]{25,}/); return m ? m[0] : ''; }).filter(String);
        var lAf = (lsAf2[li2] && String(lsAf2[li2][0] || '').match(/[-\w]{25,}/)) || [''];
        var lAnM = lsAn2[li2] ? String(lsAn2[li2][0] || '').match(/AFTER:([-\w,]+)/) : null;
        var lAfterIds = lAnM ? lAnM[1].split(',').filter(String) : [];   // V50
        lsOut2.push({ no: lr[lsM2.no - 1], section: String(lr[lsM2.sec - 1]), part: String(lr[lsM2.part - 1]), description: String(lr[lsM2.desc - 1]),
          remarks: String(lr[lsM2.rem - 1]), status: String(lr[lsM2.st - 1] || 'open'), sid: String(lr[lsM2.sid - 1] || ''),
          inspector: (lsM2.insp ? String(lr[lsM2.insp - 1] || '') : ''), photoIds: lIds, afterFirst: lAf[0], afterIds: lAfterIds });
      }
      return reply({ ok: true, items: lsOut2 });
    }

    // ============ V46 — TÔ MÀU Ô TÌNH TRẠNG theo form Google Sheet (close=xanh, open=cam) ============
    if (data.action === 'punchColorStatus') {
      var csSs = SpreadsheetApp.openById(SHEET_ID);
      var csQ = csSs.getSheetByName(SIMPLE_TAB);
      if (!csQ || csQ.getLastRow() < 4) return reply({ ok: true, painted: 0 });
      var csM = simpleColMap(csQ);
      ensureStatusFormat(csQ, csM.st);
      return reply({ ok: true, painted: csQ.getLastRow() - 3, col: csM.st });
    }

    // ============ PUNCH LIST — XÓA MỤC ĐỂ CHỤP LẠI (xóa cả 2 tab nếu chung ID) + ĐÁNH LẠI SỐ THỨ TỰ ============
    if (data.action === 'punchDelete') {
      var pdId = String(data.id || '').trim();
      if (!pdId) return reply({ ok: false, error: 'Thiếu id' });
      var pdSs = SpreadsheetApp.openById(SHEET_ID);
      var pdQ = pdSs.getSheetByName(SIMPLE_TAB);
      var pdDel = 0;
      if (pdQ && pdQ.getLastRow() >= 4) {
        var pdM = simpleColMap(pdQ);   // V44: cột theo tên tiêu đề
        var pdSids = pdQ.getRange(4, pdM.sid, pdQ.getLastRow() - 3, 1).getValues();
        for (var pdi = pdSids.length - 1; pdi >= 0; pdi--) {
          if (String(pdSids[pdi][0]) === pdId) { pdQ.deleteRow(4 + pdi); pdDel++; break; }
        }
        var pdLast = pdQ.getLastRow();
        for (var pdr = 4; pdr <= pdLast; pdr++) pdQ.getRange(pdr, 1).setValue(pdr - 3);   // No. = 1..n liên tục
      }
      if (/^PL/.test(pdId)) {
        var pvQ = pdSs.getSheetByName(PUNCH_TAB);
        if (pvQ && pvQ.getLastRow() >= 8) {
          var pvIds = pvQ.getRange(8, 15, pvQ.getLastRow() - 7, 1).getValues();
          for (var pvi = pvIds.length - 1; pvi >= 0; pvi--) {
            if (String(pvIds[pvi][0]) === pdId) { pvQ.deleteRow(8 + pvi); pdDel++; break; }
          }
          var pvLast = pvQ.getLastRow();
          for (var pvr = 8; pvr <= pvLast; pvr++) pvQ.getRange(pvr, 7).setValue('#' + (pvr - 7));  // ITEM # = #1..#n
        }
      }
      if (!pdDel) return reply({ ok: false, error: 'Không tìm thấy mục ' + pdId });
      return reply({ ok: true, deleted: pdDel });
    }

    // ============ V17 — DANH MỤC PART & MẪU MÔ TẢ (app tải về để QC bấm chọn) ============
    if (data.action === 'dbgImg') {
      var dss = SpreadsheetApp.openById(SHEET_ID);
      var dsh = dss.getSheetByName(String(data.sheet || 'Punch list đơn giản'));
      if (!dsh) return reply({ ok: false, error: 'no sheet' });
      var dog = [];
      try { var gg = dsh.getImages(); for (var g1 = 0; g1 < gg.length; g1++) { var ac = gg[g1].getAnchorCell(); dog.push(ac.getRow() + '/' + ac.getColumn()); } } catch (eg) { dog.push('ERR ' + eg); }
      var dlast = dsh.getLastRow();
      var dn = Math.max(0, dlast - 3);
      var dv = dn ? dsh.getRange(4, 1, dn, dsh.getLastColumn()).getValues() : [];
      var df = dn ? dsh.getRange(4, 1, dn, dsh.getLastColumn()).getFormulas() : [];
      var dcell = [];
      for (var r1 = 0; r1 < dv.length; r1++) {
        for (var c1 = 0; c1 < dv[r1].length; c1++) {
          var val = dv[r1][c1];
          if (val && typeof val === 'object' && !(val instanceof Date)) {
            var s = (r1 + 4) + '/' + (c1 + 1) + ' ' + (val.toString ? String(val).slice(0, 20) : 'obj');
            try { if (typeof val.getContentUrl === 'function') s += ' CURL=' + String(val.getContentUrl()).slice(0, 70); } catch (e1) { s += ' CURLERR'; }
            try { if (typeof val.getUrl === 'function') s += ' URL=' + String(val.getUrl()).slice(0, 70); } catch (e2) { s += ' URLERR'; }
            dcell.push(s);
          } else if (df[r1][c1] && String(df[r1][c1]).toUpperCase().indexOf('IMAGE') >= 0) {
            dcell.push((r1 + 4) + '/' + (c1 + 1) + ' F=' + String(df[r1][c1]).slice(0, 60));
          }
        }
      }
      // V39: thử tải ảnh trong ô đầu tiên — báo mã HTTP (thường / kèm OAuth) để chẩn đoán
      var dft = '';
      try {
        var dfu = '';
        for (var r2 = 0; r2 < dv.length && !dfu; r2++)
          for (var c2 = 0; c2 < dv[r2].length && !dfu; c2++) {
            var v2 = dv[r2][c2];
            if (v2 && typeof v2 === 'object' && !(v2 instanceof Date) && typeof v2.getContentUrl === 'function')
              try { dfu = v2.getContentUrl() || ''; } catch (e2) {}
          }
        if (dfu) {
          var dr1 = UrlFetchApp.fetch(dfu, { muteHttpExceptions: true });
          var dr2 = UrlFetchApp.fetch(dfu, { muteHttpExceptions: true,
            headers: { Authorization: 'Bearer ' + ScriptApp.getOAuthToken() } });
          dft = 'plain=' + dr1.getResponseCode() + '/' + dr1.getBlob().getContentType()
              + ' bearer=' + dr2.getResponseCode() + '/' + dr2.getBlob().getContentType();
        } else dft = 'no cell image';
      } catch (e3) { dft = 'ERR ' + e3; }
      return reply({ ok: true, lastRow: dlast, overgrid: dog, cellImages: dcell, fetchTest: dft,
                     tabs: dss.getSheets().map(function (x) { return x.getName(); }) });
    }

    // ============ V45 — LINK THƯ MỤC BÁO CÁO trên Drive (mở xem mọi PDF/Excel đã xuất) ============
    if (data.action === 'reportFolder') {
      var rfRoot = getOrCreateFolder(FOLDER_NAME);
      var rfP = getOrCreateSubfolder(rfRoot, 'Bao cao PDF');
      var rfX = getOrCreateSubfolder(rfRoot, 'Bao cao Excel');
      try { rfP.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW); } catch (eR1) {}
      try { rfX.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW); } catch (eR2) {}
      return reply({ ok: true, pdf: rfP.getUrl(), xlsx: rfX.getUrl() });
    }

    // ============ V44 — SỬA DỮ LIỆU GHI NHẦM CỘT (khi sheet bị chèn thêm cột) ============
    if (data.action === 'punchRepair') {
      var rpSs2 = SpreadsheetApp.openById(SHEET_ID);
      var rpQ2 = rpSs2.getSheetByName(SIMPLE_TAB);
      if (!rpQ2 || rpQ2.getLastRow() < 4) return reply({ ok: true, fixed: [] });
      var rpM = simpleColMap(rpQ2);
      var rpN = rpQ2.getLastRow() - 3;
      var rpV = rpQ2.getRange(4, 1, rpN, rpM.lastCol).getValues();
      var rpF = rpQ2.getRange(4, 1, rpN, rpM.lastCol).getFormulas();
      var rpFixed = [];
      for (var rr = 0; rr < rpN; rr++) {
        var rFix = 0;
        // 1) công thức IMAGE nằm nhầm trong ô Status → chuyển sang Photo after
        var rf = String(rpF[rr][rpM.st - 1] || '');
        if (rf.toUpperCase().indexOf('=IMAGE') === 0) {
          if (rpM.after && !String(rpF[rr][rpM.after - 1] || '')) rpQ2.getRange(4 + rr, rpM.after).setFormula(rf);
          rpQ2.getRange(4 + rr, rpM.st).setValue('');
          rpV[rr][rpM.st - 1] = ''; rFix++;
        }
        // 2) chữ open/close nằm nhầm cột khác → chuyển về cột Status
        for (var rc = 1; rc <= rpM.lastCol; rc++) {
          if (rc === rpM.st) continue;
          if (String(rpF[rr][rc - 1] || '')) continue;
          var rv = String(rpV[rr][rc - 1] || '').trim();
          if (/^(open|close)$/i.test(rv)) {
            if (!String(rpV[rr][rpM.st - 1] || '').trim()) { rpQ2.getRange(4 + rr, rpM.st).setValue(rv.toLowerCase()); rpV[rr][rpM.st - 1] = rv; }
            rpQ2.getRange(4 + rr, rc).setValue('');
            rFix++;
          }
        }
        if (rFix) rpFixed.push('dòng ' + (4 + rr) + ' (' + rFix + ' ô)');
      }
      return reply({ ok: true, fixed: rpFixed });
    }

    // ============ V41 — AI KHỚP ẢNH ĐÓNG HÀNG LOẠT: đọc mã part trên ảnh → gợi ý mục đang mở ============
    if (data.action === 'punchMatch') {
      if (!GEMINI_API_KEY) return reply({ ok: true, ai: false, matches: [] });
      var pmImgs = data.images || [];
      if (!pmImgs.length) return reply({ ok: false, error: 'Không có ảnh' });
      if (pmImgs.length > 10) return reply({ ok: false, error: 'Tối đa 10 ảnh mỗi lần' });
      var pmSs = SpreadsheetApp.openById(SHEET_ID);
      var pmQ = pmSs.getSheetByName(SIMPLE_TAB);
      var pmOpen = [];
      if (pmQ && pmQ.getLastRow() >= 4) {
        var pmM = simpleColMap(pmQ);   // V44: cột theo tên tiêu đề
        var pmV = pmQ.getRange(4, 1, pmQ.getLastRow() - 3, pmM.lastCol).getValues();
        for (var pmi = 0; pmi < pmV.length; pmi++) {
          if (plIsClosed_(pmV[pmi][pmM.st - 1])) continue;   // V49: close/Close/closed
          if (!String(pmV[pmi][pmM.part - 1] || '').trim() && !String(pmV[pmi][pmM.desc - 1] || '').trim()) continue;
          pmOpen.push({ punchId: String(pmV[pmi][pmM.sid - 1] || ('SPROW' + (4 + pmi))), itemNo: 'No.' + pmV[pmi][pmM.no - 1],
                        part: String(pmV[pmi][pmM.part - 1] || ''), desc: String(pmV[pmi][pmM.desc - 1] || '').slice(0, 80) });
        }
      }
      if (!pmOpen.length) return reply({ ok: true, ai: true, matches: [], note: 'Không còn mục đang mở' });
      var pmParts = [{ text:
        'Mỗi ảnh chụp một chi tiết thép có mã part VIẾT TAY bằng phấn/sơn hoặc in sẵn (ví dụ TED1043, RSD1006, B75). ' +
        'Với TỪNG ảnh theo đúng thứ tự gửi lên, đọc mã part rõ nhất trong ảnh. ' +
        'Chỉ trả JSON dạng {"parts":["TED1043",null,...]} — phần tử null nếu không đọc được mã.' }];
      pmImgs.forEach(function (b64) { pmParts.push({ inline_data: { mime_type: 'image/jpeg', data: b64 } }); });
      var pmPayload = JSON.stringify({ contents: [{ parts: pmParts }],
        generationConfig: { temperature: 0, response_mime_type: 'application/json' } });
      var pmModels = [GEMINI_MODEL, GEMINI_MODEL_FALLBACK], pmRes = null, pmErr = '';
      for (var pmm = 0; pmm < pmModels.length; pmm++) {
        pmRes = UrlFetchApp.fetch('https://generativelanguage.googleapis.com/v1beta/models/' + pmModels[pmm] + ':generateContent?key=' + GEMINI_API_KEY,
          { method: 'post', contentType: 'application/json', muteHttpExceptions: true, payload: pmPayload });
        if (pmRes.getResponseCode() === 200) break;
        pmErr = pmModels[pmm] + ' HTTP ' + pmRes.getResponseCode();
      }
      if (pmRes.getResponseCode() !== 200) return reply({ ok: false, error: 'Gemini lỗi: ' + pmErr });
      var pmTxt = '';
      try {
        var pmJ = JSON.parse(pmRes.getContentText());
        pmTxt = pmJ.candidates[0].content.parts[0].text || '';
      } catch (ePm) {}
      var pmOut = safeParseJson(pmTxt) || {};
      var pmRead = pmOut.parts || [];
      var pmMatches = [];
      for (var pk = 0; pk < pmImgs.length; pk++) {
        var pmP = normPart(pmRead[pk] || '');
        if (!pmP) { pmMatches.push(null); continue; }
        var pmHit = null, pmN = 0;
        for (var ph = 0; ph < pmOpen.length; ph++)
          if (normPart(pmOpen[ph].part) === pmP) { pmN++; if (!pmHit) pmHit = pmOpen[ph]; }
        pmMatches.push(pmHit ? { part: String(pmRead[pk] || ''), punchId: pmHit.punchId, itemNo: pmHit.itemNo,
                                 desc: pmHit.desc, many: pmN > 1 }
                             : { part: String(pmRead[pk] || ''), punchId: '', itemNo: '', many: false });
      }
      return reply({ ok: true, ai: true, matches: pmMatches });
    }

    // ============ V40 — DI TRÚ ẢNH DÁN TAY → Drive + cột "Photo link" (ảnh bền, tự co trong báo cáo) ============
    if (data.action === 'punchFixImg') {
      var fxSs = SpreadsheetApp.openById(SHEET_ID);
      var fxSh = fxSs.getSheetByName(String(data.sheet || 'Punch list đơn giản'));
      if (!fxSh) return reply({ ok: false, error: 'no sheet' });
      var fxLastC = fxSh.getLastColumn(), fxLastR = fxSh.getLastRow();
      var fxMap = simpleColMap(fxSh);         // V44/V50: theo tên tiêu đề
      var fxLink = fxMap.link, fxAfterCol = fxMap.after || 0;
      var fxNotes = [];
      try { if (fxAfterCol && fxLastR >= 4) fxNotes = fxSh.getRange(4, fxAfterCol, fxLastR - 3, 1).getNotes(); } catch (eNt) { fxNotes = []; }
      var fxMonth = Utilities.formatDate(new Date(), 'GMT+7', 'yyyy-MM');
      var fxDay = Utilities.formatDate(new Date(), 'GMT+7', 'yyyy-MM-dd');
      var fxFol = getReportFolder('Don gian', fxMonth, fxDay);
      var fxN = Math.max(0, fxLastR - 3);
      var fxVals = fxN ? fxSh.getRange(4, 1, fxN, fxLastC).getValues() : [];
      var fxOg = {};
      try {
        var fxIms = fxSh.getImages();
        for (var fo = 0; fo < fxIms.length; fo++) {
          var fa = fxIms[fo].getAnchorCell().getRow();
          (fxOg[fa] = fxOg[fa] || []).push(fxIms[fo]);
        }
      } catch (eFo) {}
      var fxSaved = 0, fxFail = [], fxRows = 0, fxSkip = 0, fxStamp = Utilities.formatDate(new Date(), 'GMT+7', 'yyyyMMdd_HHmmss');
      var fxIsImg = function (v) { return v && typeof v === 'object' && !(v instanceof Date) && typeof v.getContentUrl === 'function'; };
      for (var fr = 0; fr < fxN; fr++) {
        var fxCur = String(fxVals[fr][fxLink - 1] || '').trim();
        var fxHasImg = false, fxHasAfter = false, fxIds = [], fxAfterIds = [];
        for (var fcc = 0; fcc < fxLastC; fcc++) {
          if (fxIsImg(fxVals[fr][fcc])) { fxHasImg = true; if (fcc + 1 === fxAfterCol) fxHasAfter = true; }
        }
        var fxG = fxOg[fr + 4] || [];
        fxG.forEach(function (g) { try { if (g.getAnchorCell().getColumn() === fxAfterCol) fxHasAfter = true; } catch (eAc) {} });
        if (!fxHasImg && !fxG.length) continue;
        // V50: dòng đã có link → chỉ lấy thêm ẢNH SAU (cột Photo after) nếu chưa ghi chú AFTER; còn lại bỏ qua
        var fxAfterOnly = false;
        if (fxCur && !data.force) {
          var fxNoted = fxNotes[fr] && /AFTER:/.test(String(fxNotes[fr][0] || ''));
          if (fxHasAfter && !fxNoted) fxAfterOnly = true; else { fxSkip++; continue; }
        }
        for (fcc = 0; fcc < fxLastC; fcc++) {
          var fv2 = fxVals[fr][fcc];
          if (!fxIsImg(fv2)) continue;
          var fxIsAfter = (fcc + 1 === fxAfterCol);
          if (fxAfterOnly && !fxIsAfter) continue;
          var fu = ''; try { fu = fv2.getContentUrl() || ''; } catch (eFu) {}
          var fb = fu ? plFetchCellImg(fu) : null;
          if (fb) {
            var ff = fxFol.createFile(fb.setName('PLMIG_' + fxStamp + '_r' + (fr + 4) + 'c' + (fcc + 1) + '.png'));
            fxIds.push(ff.getId()); if (fxIsAfter) fxAfterIds.push(ff.getId()); fxSaved++;
          } else fxFail.push('dòng ' + (fr + 4) + ' cột ' + (fcc + 1) + ' (ảnh trong ô)');
        }
        for (var fg = 0; fg < fxG.length; fg++) {
          var fgAfter = false;
          try { fgAfter = (fxG[fg].getAnchorCell().getColumn() === fxAfterCol); } catch (eAc2) {}
          if (fxAfterOnly && !fgAfter) continue;
          try {
            var fb2 = fxG[fg].getBlob();
            var ff2 = fxFol.createFile(fb2.setName('PLMIG_' + fxStamp + '_r' + (fr + 4) + 'og' + (fg + 1) + '.png'));
            fxIds.push(ff2.getId()); if (fgAfter) fxAfterIds.push(ff2.getId()); fxSaved++;
          } catch (eFg) { fxFail.push('dòng ' + (fr + 4) + ' ảnh nổi ' + (fg + 1)); }
        }
        if (fxIds.length) {
          var fxUrls = fxIds.map(function (id) { return 'https://drive.google.com/file/d/' + id + '/view'; });
          fxSh.getRange(fr + 4, fxLink).setValue(fxCur ? (fxCur + '\n' + fxUrls.join('\n')) : fxUrls.join('\n'));
          if (fxAfterIds.length && fxAfterCol) { try { plAddAfterNote_(fxSh.getRange(fr + 4, fxAfterCol), fxAfterIds); } catch (eAn3) {} }   // V50
          fxRows++;
        }
      }
      return reply({ ok: true, rows: fxRows, saved: fxSaved, skipped: fxSkip, failed: fxFail });
    }

    if (data.action === 'listCatalog') {
      var lcC = readCatalog(!!data.force);
      return reply({ ok: true, projects: lcC.projects, tree: lcC.tree, descs: lcC.descs, descg: lcC.descg || [] });
    }

    // NẠP HÀNG LOẠT part cho 1 project: data.rows = [[section, part, ghi chú], ...]
    // Mặc định XOÁ hết dòng cũ của project đó rồi ghi mới (replace:false = ghi thêm).
    if (data.action === 'catalogBulk') {
      var cbPj = String(data.project || '').trim();
      var cbIn = data.rows || [];
      if (!cbPj) return reply({ ok: false, error: 'Thiếu project' });
      if (!cbIn.length) return reply({ ok: false, error: 'Không có dòng nào' });
      var cbSh = ensureCatalog(SpreadsheetApp.openById(SHEET_ID)).cat;
      var cbDel = 0;
      if (data.replace !== false && cbSh.getLastRow() >= 2) {
        var cbOld = cbSh.getRange(2, 1, cbSh.getLastRow() - 1, 1).getValues();
        for (var cbi = cbOld.length - 1; cbi >= 0; cbi--) {
          if (String(cbOld[cbi][0]).trim().toUpperCase() === cbPj.toUpperCase()) { cbSh.deleteRow(2 + cbi); cbDel++; }
        }
      }
      var cbSeen = {}, cbOut = [], cbSecs = {};
      for (var cbj = 0; cbj < cbIn.length; cbj++) {
        var cbR = cbIn[cbj] || [];
        var cbS = normSection(cbR[0]);
        var cbP = String(cbR[1] || '').trim();
        if (!cbP) continue;
        var cbK = cbS + '|' + cbP.toUpperCase();
        if (cbSeen[cbK]) continue;
        cbSeen[cbK] = 1; cbSecs[cbS] = 1;
        cbOut.push([cbPj, cbS, cbP, String(cbR[2] || '').trim()]);
      }
      if (!cbOut.length) return reply({ ok: false, error: 'Không có dòng hợp lệ (thiếu cột Part)' });
      var cbStart = Math.max(2, cbSh.getLastRow() + 1);
      cbSh.getRange(cbStart, 1, cbOut.length, 4).setValues(cbOut);
      cbSh.getRange(cbStart, 1, cbOut.length, 4).setBorder(true, true, true, true, true, true);
      cbSh.getRange(cbStart, 2, cbOut.length, 2).setHorizontalAlignment('center');
      catalogFormat(cbSh);   // kẻ lại tiêu đề + hướng dẫn SAU khi xoá/ghi, tránh mất chữ
      try { CacheService.getScriptCache().remove('catalog_v1'); } catch (eCb) {}
      return reply({ ok: true, project: cbPj, written: cbOut.length,
                     sections: Object.keys(cbSecs).length, removedOld: cbDel });
    }

    // Ghi ĐÈ toàn bộ tab "Mẫu mô tả" bằng danh sách đã gom gọn: data.items = [[câu, nhóm], ...]
    if (data.action === 'descReplace') {
      var drIt = data.items || [];
      if (!drIt.length) return reply({ ok: false, error: 'Thiếu items' });
      var drSh = ensureCatalog(SpreadsheetApp.openById(SHEET_ID)).desc;
      if (drSh.getLastRow() >= 2) drSh.deleteRows(2, drSh.getLastRow() - 1);
      var drRows = drIt.map(function (x) {
        return [String((x[0] !== undefined ? x[0] : x.text) || '').trim(),
                String((x[1] !== undefined ? x[1] : x.group) || '').trim()];
      }).filter(function (r) { return r[0]; });
      if (!drRows.length) return reply({ ok: false, error: 'items rỗng' });
      drSh.getRange(2, 1, drRows.length, 2).setValues(drRows);
      drSh.getRange(2, 1, drRows.length, 2).setWrap(true).setVerticalAlignment('middle');
      try { CacheService.getScriptCache().remove('catalog_v1'); } catch (eDr) {}
      return reply({ ok: true, written: drRows.length });
    }

    // Thêm 1 part mới vào danh mục (khi QC gõ part chưa có trong danh sách) — không trùng.
    if (data.action === 'catalogAdd') {
      var caPj = String(data.project || '').trim();
      var caSc = String(data.section || '').trim();
      var caPt = String(data.part || '').trim();
      if (!caPj || !caPt) return reply({ ok: false, error: 'Thiếu project hoặc part' });
      var caT = ensureCatalog(SpreadsheetApp.openById(SHEET_ID)).cat;
      if (caT.getLastRow() >= 2) {
        var caV = caT.getRange(2, 1, caT.getLastRow() - 1, 3).getValues();
        for (var cai = 0; cai < caV.length; cai++) {
          if (String(caV[cai][0]).trim() === caPj && String(caV[cai][1]).trim() === caSc
              && String(caV[cai][2]).trim() === caPt) return reply({ ok: true, added: 0, existed: true });
        }
      }
      caT.appendRow([caPj, caSc, caPt, 'Thêm từ app ' + Utilities.formatDate(new Date(), 'GMT+7', 'yyyy-MM-dd')]);
      try { CacheService.getScriptCache().remove('catalog_v1'); } catch (eCa) {}
      return reply({ ok: true, added: 1 });
    }

    // ============ V23 — ĐỐI CHIẾU PUNCH LIST VỚI DANH MỤC PART ============
    // Soát từng mục: part có trong danh mục không, có đúng section không, có sai cách viết không.
    if (data.action === 'punchAudit') {
      var auPj = String(data.catProject || data.project || '').trim() || PUNCH_PROJECT_NAME;
      var auTree = (readCatalog(!!data.force).tree || {})[auPj] || {};
      var auMap = {}, auCanon = {};
      Object.keys(auTree).forEach(function (sc) {
        (auTree[sc] || []).forEach(function (pt) {
          var k = normPart(pt);
          if (!auMap[k]) auMap[k] = [];
          if (auMap[k].indexOf(sc) < 0) auMap[k].push(sc);
          auCanon[sc + '|' + k] = pt;
        });
      });
      var auSs = SpreadsheetApp.openById(SHEET_ID);
      var auQ = auSs.getSheetByName(SIMPLE_TAB);
      if (!auQ || auQ.getLastRow() < 4) return reply({ ok: true, total: 0, items: [], project: auPj });
      var auN = auQ.getLastRow() - 3;
      var auM = simpleColMap(auQ);   // V49: cột theo tên tiêu đề
      var auV = auQ.getRange(4, 1, auN, auM.lastCol).getValues();
      var auOut = [], auOk = 0;
      for (var aui = 0; aui < auN; aui++) {
        var auSec = normSection(auV[aui][auM.sec - 1]);
        var auPt = String(auV[aui][auM.part - 1] || '').trim();
        var auBase = { row: 4 + aui, no: auV[aui][auM.no - 1], section: auSec, part: auPt,
                       description: String(auV[aui][auM.desc - 1] || '').slice(0, 60), sid: String(auV[aui][auM.sid - 1] || '') };
        if (!auPt) { auBase.issue = 'thieu_part'; auBase.suggest = ''; auOut.push(auBase); continue; }
        var auK = normPart(auPt);
        var auIn = auMap[auK];
        if (!auIn) { auBase.issue = 'khong_co_trong_danh_muc'; auBase.suggest = ''; auOut.push(auBase); continue; }
        if (auIn.indexOf(auSec) < 0) {
          auBase.issue = 'sai_section'; auBase.suggest = auIn.join(', '); auOut.push(auBase); continue;
        }
        var auC = auCanon[auSec + '|' + auK];
        if (auC && auC !== auPt) { auBase.issue = 'sai_cach_viet'; auBase.suggest = auC; auOut.push(auBase); continue; }
        auOk++;
      }
      return reply({ ok: true, project: auPj, total: auN, okCount: auOk, items: auOut });
    }

    // Sửa 1 mục theo kết quả đối chiếu: đổi section và/hoặc part (đồng bộ luôn dòng VIOLA cùng ID).
    if (data.action === 'punchFix') {
      var pfSid = String(data.sid || '').trim();
      var pfRow = parseInt(data.row, 10) || 0;
      var pfSs = SpreadsheetApp.openById(SHEET_ID);
      var pfQ = pfSs.getSheetByName(SIMPLE_TAB);
      if (!pfQ || pfQ.getLastRow() < 4) return reply({ ok: false, error: 'Chưa có dữ liệu' });
      var pfM = simpleColMap(pfQ);   // V49: cột theo tên tiêu đề
      if (pfSid) {
        var pfIds = pfQ.getRange(4, pfM.sid, pfQ.getLastRow() - 3, 1).getValues();
        pfRow = 0;
        for (var pfi = 0; pfi < pfIds.length; pfi++) if (String(pfIds[pfi][0]) === pfSid) { pfRow = 4 + pfi; break; }
      }
      if (pfRow < 4 || pfRow > pfQ.getLastRow()) return reply({ ok: false, error: 'Không tìm thấy dòng cần sửa' });
      var pfNewSec = (data.section !== undefined && data.section !== null) ? normSection(data.section) : null;
      var pfNewPt = (data.part !== undefined && data.part !== null) ? String(data.part).trim() : null;
      if (pfNewSec !== null && pfNewSec !== '') pfQ.getRange(pfRow, pfM.sec).setValue(pfNewSec);
      if (pfNewPt !== null && pfNewPt !== '') pfQ.getRange(pfRow, pfM.part).setValue(pfNewPt);
      try {
        if (/^PL/.test(pfSid)) {
          var pfV = pfSs.getSheetByName(PUNCH_TAB);
          if (pfV && pfV.getLastRow() >= 8) {
            var pfVIds = pfV.getRange(8, 15, pfV.getLastRow() - 7, 1).getValues();
            for (var pfj = 0; pfj < pfVIds.length; pfj++) {
              if (String(pfVIds[pfj][0]) === pfSid) {
                if (pfNewPt) pfV.getRange(8 + pfj, 1).setValue(pfNewPt);
                break;
              }
            }
          }
        }
      } catch (ePf) {}
      return reply({ ok: true, row: pfRow, section: pfNewSec, part: pfNewPt });
    }

    // ============ PUNCH LIST — DỌN DÒNG TRÙNG (V16) ============
    // Xoá các dòng bị lặp do gửi lại khi mạng lỗi: giữ dòng ĐẦU, xoá dòng sau có cùng
    // Section + Part + Mô tả + cùng ảnh. Xoá kèm dòng VIOLA cùng ID rồi đánh lại số thứ tự.
    // data.dryRun = true → chỉ ĐẾM, không xoá (để app hỏi xác nhận trước).
    if (data.action === 'punchDedupe') {
      // V24 — 2 kiểu trùng:
      //   'photo'   = cùng Section+Part+Mô tả+Remarks VÀ cùng file ảnh  (chắc chắn là gửi lặp)
      //   'content' = cùng Section+Part+Mô tả+Remarks, ẢNH KHÁC NHAU     (chụp lại 2 lần cùng 1 lỗi)
      var ddMode = String(data.mode || 'photo').toLowerCase();
      var ddSs = SpreadsheetApp.openById(SHEET_ID);
      var ddQ = ddSs.getSheetByName(SIMPLE_TAB);
      if (!ddQ || ddQ.getLastRow() < 4)
        return reply({ ok: true, found: 0, foundPhoto: 0, foundContent: 0, deleted: 0, items: [] });
      var ddN = ddQ.getLastRow() - 3;
      var ddM = simpleColMap(ddQ);   // V49: cột theo tên tiêu đề
      var ddAll = ddQ.getRange(4, 1, ddN, ddM.lastCol).getValues();
      var ddAllF = ddQ.getRange(4, 1, ddN, ddM.lastCol).getFormulas();
      // dựng lại mảng theo THỨ TỰ CŨ [no, sec, part, desc, rem, photo, link, st, sid] → giữ nguyên logic bên dưới
      var ddVals = ddAll.map(function (r) {
        return [r[ddM.no - 1], r[ddM.sec - 1], r[ddM.part - 1], r[ddM.desc - 1], r[ddM.rem - 1],
                r[ddM.photo - 1], r[ddM.link - 1], r[ddM.st - 1], r[ddM.sid - 1]];
      });
      var ddFmls = ddAllF.map(function (r) { return [r[ddM.photo - 1], r[ddM.link - 1]]; });
      var ddSeenP = {}, ddSeenC = {}, ddDupP = [], ddDupC = [];
      var ddRef = {};   // key -> {no, photoIds} của DÒNG ĐẦU (dòng được giữ)
      for (var ddi = 0; ddi < ddN; ddi++) {
        var ddPh = String(ddVals[ddi][6] || '').split('\n').map(function (u) {
          var m = String(u).match(/[-\w]{25,}/); return m ? m[0] : '';
        }).filter(String);
        var ddIt = { row: 4 + ddi, no: ddVals[ddi][0], section: String(ddVals[ddi][1]),
                     part: String(ddVals[ddi][2]), description: String(ddVals[ddi][3]),
                     sid: String(ddVals[ddi][8] || ''), photoIds: ddPh };
        var ddKP = pdupKey(ddVals[ddi][1], ddVals[ddi][2], ddVals[ddi][3], ddVals[ddi][4], ddFmls[ddi][0], ddVals[ddi][6]);
        var ddKC = pdupKey(ddVals[ddi][1], ddVals[ddi][2], ddVals[ddi][3], ddVals[ddi][4], '', '');
        if (ddSeenP[ddKP]) { ddIt.keepNo = ddSeenP[ddKP]; ddIt.keepPhotoIds = (ddRef[ddKP] || {}).photoIds || []; ddDupP.push(ddIt); }
        else { ddSeenP[ddKP] = ddVals[ddi][0]; ddRef[ddKP] = { sid: ddIt.sid, photoIds: ddPh }; }
        if (ddSeenC[ddKC]) {
          var ddR = ddRef[ddKC] || {};
          ddDupC.push({ row: ddIt.row, no: ddIt.no, section: ddIt.section, part: ddIt.part,
                        description: ddIt.description, sid: ddIt.sid, photoIds: ddPh,
                        keepNo: ddSeenC[ddKC], keepSid: ddR.sid || '', keepPhotoIds: ddR.photoIds || [] });
        } else { ddSeenC[ddKC] = ddVals[ddi][0]; ddRef[ddKC] = { sid: ddIt.sid, photoIds: ddPh }; }
      }
      var ddDup = (ddMode === 'content') ? ddDupC : ddDupP;
      if (data.dryRun)
        return reply({ ok: true, mode: ddMode, found: ddDup.length,
                       foundPhoto: ddDupP.length, foundContent: ddDupC.length, deleted: 0, items: ddDup });
      if (!ddDup.length)
        return reply({ ok: true, mode: ddMode, found: 0, foundPhoto: ddDupP.length, foundContent: ddDupC.length,
                       deleted: 0, items: [] });
      var ddSids = [];
      for (var ddj = ddDup.length - 1; ddj >= 0; ddj--) {
        if (ddDup[ddj].sid) ddSids.push(ddDup[ddj].sid);
        ddQ.deleteRow(ddDup[ddj].row);
      }
      var ddLast = ddQ.getLastRow();
      for (var ddr = 4; ddr <= ddLast; ddr++) ddQ.getRange(ddr, ddM.no).setValue(ddr - 3);
      var ddV = ddSs.getSheetByName(PUNCH_TAB);
      if (ddV && ddV.getLastRow() >= 8 && ddSids.length) {
        var ddVIds = ddV.getRange(8, 15, ddV.getLastRow() - 7, 1).getValues();
        for (var ddvi = ddVIds.length - 1; ddvi >= 0; ddvi--) {
          if (ddSids.indexOf(String(ddVIds[ddvi][0])) >= 0) ddV.deleteRow(8 + ddvi);
        }
        var ddVLast = ddV.getLastRow();
        for (var ddvr = 8; ddvr <= ddVLast; ddvr++) ddV.getRange(ddvr, 7).setValue('#' + (ddvr - 7));
      }
      return reply({ ok: true, mode: ddMode, found: ddDup.length, deleted: ddDup.length, items: ddDup });
    }

    // ============ PUNCH LIST — DỌN SẠCH DỮ LIỆU NHÁP (xóa hết dòng cả 2 tab, giữ header) ============
    if (data.action === 'punchClearAll') {
      // V21 — CHỐNG XOÁ NHẦM: bắt buộc gõ đúng mã, và LUÔN sao lưu ra tab ẩn trước khi xoá.
      if (String(data.code || '').trim().toUpperCase() !== 'XOA')
        return reply({ ok: false, error: 'Cần gõ đúng chữ XOA để xác nhận. Đây là xoá dữ liệu TRÊN GOOGLE SHEET, không phải chỉ trên máy.' });
      var caSs = SpreadsheetApp.openById(SHEET_ID);
      var caStamp = Utilities.formatDate(new Date(), 'GMT+7', 'yyMMdd_HHmm');
      var caN = 0, caBk = [];
      var caQ = caSs.getSheetByName(SIMPLE_TAB);
      if (caQ && caQ.getLastRow() >= 4) {
        var caB1 = caQ.copyTo(caSs).setName('SAOLUU_' + caStamp + '_dongian');
        try { caB1.hideSheet(); } catch (eH1) {}
        caBk.push(caB1.getName());
        caN += caQ.getLastRow() - 3; caQ.deleteRows(4, caQ.getLastRow() - 3);
      }
      var caV = caSs.getSheetByName(PUNCH_TAB);
      if (caV && caV.getLastRow() >= 8) {
        var caB2 = caV.copyTo(caSs).setName('SAOLUU_' + caStamp + '_viola');
        try { caB2.hideSheet(); } catch (eH2) {}
        caBk.push(caB2.getName());
        caN += caV.getLastRow() - 7; caV.deleteRows(8, caV.getLastRow() - 7);
      }
      return reply({ ok: true, deleted: caN, backup: caBk });
    }

    // ============ V22 — XUẤT BÁO CÁO TED: PDF hoặc EXCEL (dựng sheet tạm trình bày đẹp) ============
    if (data.action === 'punchPdf' || data.action === 'punchXlsx') {
      var ppSs = SpreadsheetApp.openById(SHEET_ID);
      var ppSec = String(data.section || '').trim();
      var ppDate = Utilities.formatDate(new Date(), 'GMT+7', 'yyyy-MM-dd');
      var ppMonth = Utilities.formatDate(new Date(), 'GMT+7', 'yyyy-MM');

      // Mẫu EVAPCO (tab VIOLA) giữ nguyên cách cũ: xuất thẳng tab đó
      if (data.tab === 'viola') {
        var pvSh = ppSs.getSheetByName(PUNCH_TAB);
        if (!pvSh) return reply({ ok: false, error: 'Chưa có tab "' + PUNCH_TAB + '".' });
        var pvUrl = 'https://docs.google.com/spreadsheets/d/' + SHEET_ID + '/export?format=pdf&gid=' + pvSh.getSheetId() +
          '&size=A4&portrait=false&fitw=true&sheetnames=false&printtitle=false&pagenumbers=true&gridlines=false&fzr=true' +
          '&top_margin=0.25&bottom_margin=0.25&left_margin=0.25&right_margin=0.25';
        var pvRes = UrlFetchApp.fetch(pvUrl, { headers: { Authorization: 'Bearer ' + ScriptApp.getOAuthToken() }, muteHttpExceptions: true });
        if (pvRes.getResponseCode() !== 200) return reply({ ok: false, error: 'Không xuất được PDF (HTTP ' + pvRes.getResponseCode() + ')' });
        var pvName = 'Punchlist EVAPCO - ' + ppDate + '.pdf';
        var pvFile = getReportFolder('Bao cao PDF', ppMonth, ppDate).createFile(pvRes.getBlob().setName(pvName));
        try { pvFile.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW); } catch (ePv) {}
        return reply({ ok: true, url: pvFile.getUrl(), name: pvName });
      }

      // V46: bấm lại trong 90 giây → trả ngay kết quả vừa tạo, không chạy lại từ đầu
      var ppCKey = 'ppres_' + data.action + '_' + ppSec;
      var ppCached = CacheService.getScriptCache().get(ppCKey);
      if (ppCached) { try { return reply(JSON.parse(ppCached)); } catch (eCc) {} }
      // V43/V46: khoá chống 2 người xuất chồng nhau — chờ tối đa 90 giây (xếp hàng thay vì báo lỗi)
      var ppLock = LockService.getScriptLock();
      if (!ppLock.tryLock(90000)) return reply({ ok: false, error: 'Máy chủ đang bận — chờ 1 phút rồi bấm lại giúp anh.' });
      ppCached = CacheService.getScriptCache().get(ppCKey);   // người trước vừa xuất xong đúng bản này → dùng luôn
      if (ppCached) { try { ppLock.releaseLock(); return reply(JSON.parse(ppCached)); } catch (eC2) {} }
      try {
        var ppShs = ppSs.getSheets();
        for (var ppsi = 0; ppsi < ppShs.length; ppsi++)
          if (ppShs[ppsi].getName().indexOf('BAOCAO_TAM_') === 0) { try { ppSs.deleteSheet(ppShs[ppsi]); } catch (ePs) {} }
      } catch (ePc) {}
      var ppTmp = null;
      try {
        var ppBuilt = buildTedReport(ppSs, {
          section: ppSec,
          project: String(data.project || '').trim() || PUNCH_PROJECT_NAME,
          preparedBy: String(data.preparedBy || data.user || '').trim() || PUNCH_PREPARED_BY,
          docNo: String(data.docNo || '').trim() || PUNCH_DOC_NO,
          rev: String(data.rev || '').trim() || TED_REV
        });
        if (!ppBuilt || !ppBuilt.sh)
          return reply({ ok: false, error: ppSec ? ('Section "' + ppSec + '" chưa có mục nào.') : 'Chưa có mục nào trên Sheet.' });
        ppTmp = ppBuilt.sh;
        var ppBase = 'Punchlist TED' + (ppSec ? (' - Section ' + ppSec) : '') + ' - ' + ppDate;

        if (data.action === 'punchXlsx') {
          // copy sheet báo cáo sang FILE MỚI rồi xuất .xlsx (export xlsx theo gid không tách được 1 sheet)
          var xTmpSs = SpreadsheetApp.create('TMP_' + ppBase);
          var xCopied = ppTmp.copyTo(xTmpSs);
          xCopied.setName('Punch list');
          try { xTmpSs.deleteSheet(xTmpSs.getSheetByName('Sheet1')); } catch (eX1) {}
          SpreadsheetApp.flush();
          var xUrl = 'https://docs.google.com/spreadsheets/d/' + xTmpSs.getId() + '/export?format=xlsx';
          var xRes = null;
          for (var xTry = 0; xTry < 3; xTry++) {   // V43: tự thử lại khi export lỗi thoáng qua
            xRes = UrlFetchApp.fetch(xUrl, { headers: { Authorization: 'Bearer ' + ScriptApp.getOAuthToken() }, muteHttpExceptions: true });
            if (xRes.getResponseCode() === 200) break;
            Utilities.sleep(2500);
          }
          var xName = ppBase + '.xlsx';
          if (xRes.getResponseCode() !== 200) {
            try { DriveApp.getFileById(xTmpSs.getId()).setTrashed(true); } catch (eX2) {}
            return reply({ ok: false, error: 'Không xuất được Excel (HTTP ' + xRes.getResponseCode() + ')' });
          }
          var xFile = getReportFolder('Bao cao Excel', ppMonth, ppDate).createFile(xRes.getBlob().setName(xName));
          try { xFile.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW); } catch (eX3) {}
          try { DriveApp.getFileById(xTmpSs.getId()).setTrashed(true); } catch (eX4) {}
          var xFolderUrl = ''; try { xFolderUrl = getReportFolder('Bao cao Excel', ppMonth, ppDate).getUrl(); } catch (eXo) {}
          var xOut = { ok: true, url: xFile.getUrl(), name: xName, rows: ppBuilt.rows, skipped: ppBuilt.skipped, kind: 'xlsx', folder: xFolderUrl };
          try { CacheService.getScriptCache().put(ppCKey, JSON.stringify(xOut), 90); } catch (eCx) {}
          return reply(xOut);
        }

        var ppUrl = 'https://docs.google.com/spreadsheets/d/' + SHEET_ID + '/export?format=pdf&gid=' + ppTmp.getSheetId() +
          '&size=A4&portrait=false&fitw=true&sheetnames=false&printtitle=false&pagenumbers=true&gridlines=false&fzr=true' +
          '&top_margin=0.3&bottom_margin=0.55&left_margin=0.25&right_margin=0.25';
        SpreadsheetApp.flush(); Utilities.sleep(1200);   // V43: chờ ảnh IMAGE() dựng xong mới xuất
        var ppRes = null;
        for (var ppTry = 0; ppTry < 3; ppTry++) {   // V43: export hay dính 429/500 thoáng qua — tự thử lại
          ppRes = UrlFetchApp.fetch(ppUrl, { headers: { Authorization: 'Bearer ' + ScriptApp.getOAuthToken() }, muteHttpExceptions: true });
          if (ppRes.getResponseCode() === 200) break;
          Utilities.sleep(2500);
        }
        if (ppRes.getResponseCode() !== 200) return reply({ ok: false, error: 'Không xuất được PDF (HTTP ' + ppRes.getResponseCode() + ') — bấm xuất lại giúp em.' });
        var ppName = ppBase + '.pdf';
        var ppFile = getReportFolder('Bao cao PDF', ppMonth, ppDate).createFile(ppRes.getBlob().setName(ppName));
        try { ppFile.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW); } catch (ePp) {}
        var ppFolderUrl = ''; try { ppFolderUrl = getReportFolder('Bao cao PDF', ppMonth, ppDate).getUrl(); } catch (eFo) {}
        var ppOut = { ok: true, url: ppFile.getUrl(), name: ppName, section: ppSec, folder: ppFolderUrl,
                      rows: ppBuilt.rows, skipped: ppBuilt.skipped, open: ppBuilt.open, close: ppBuilt.close, kind: 'pdf' };
        try { CacheService.getScriptCache().put(ppCKey, JSON.stringify(ppOut), 90); } catch (eCp) {}
        return reply(ppOut);
      } finally {
        try { if (ppTmp) ppSs.deleteSheet(ppTmp); } catch (eDel) {}
        try { ppLock.releaseLock(); } catch (eLk) {}
      }
    }

    // V18: action lạ (vd app mới gọi backend cũ/chưa deploy) → BÁO LỖI RÕ,
    // tuyệt đối không rơi xuống nhánh ghi báo cáo tạo dòng rác.
    if (data.action) {
      return reply({ ok: false, error: 'Backend chưa hỗ trợ action "' + data.action + '" — cần deploy New version.' });
    }

    // Idempotency: cùng reportId trong 6h chỉ ghi 1 lần (chống queue gửi trùng)
    var cache = CacheService.getScriptCache();
    if (data.reportId) {
      if (cache.get('rid_' + data.reportId)) {
        return reply({ ok: true, duplicated: true, project: '', category: '',
                       description: 'Báo cáo đã được ghi trước đó', issues: '', imageUrl: '' });
      }
      cache.put('rid_' + data.reportId, '1', 21600);
    }

    // 1. Xác định dự án TRƯỚC để xếp ảnh đúng thư mục:
    //    người dùng chọn > GPS > Chưa phân loại
    var project = (data.project && data.project !== 'auto') ? data.project
                : matchProjectByGps(data.lat, data.lng) || 'Chưa phân loại';

    // 2. Lưu ảnh Drive (hỗ trợ NHIỀU ảnh / 1 báo cáo)
    //    Nếu dự án khai báo Folder ID riêng → lưu vào thư mục đó / <yyyy-MM>.
    //    Nếu không → mặc định: FieldSnap Photos / <dự án> / <yyyy-MM>.
    var images = data.images || (data.imageBase64 ? [data.imageBase64] : []);
    var imageUrls = [];
    var firstFileId = '';
    if (images.length) {
      var month = Utilities.formatDate(new Date(), 'GMT+7', 'yyyy-MM');
      var day = Utilities.formatDate(new Date(), 'GMT+7', 'yyyy-MM-dd');
      // V42: tái dùng thư mục ngày (cache 6h) + chia sẻ 1 LẦN ở cấp thư mục → gửi nhanh hơn nhiều
      var folder = getReportFolder(project, month, day);
      var stamp = Utilities.formatDate(new Date(), 'GMT+7', 'yyyyMMdd_HHmmss');
      images.forEach(function (b64, i) {
        var fileName = 'IMG_' + stamp + '_' + (i + 1) + '.jpg';
        var blob = Utilities.newBlob(Utilities.base64Decode(b64), 'image/jpeg', fileName);
        var file = folder.createFile(blob);   // KHÔNG setSharing từng ảnh (thư mục đã chia sẻ)
        imageUrls.push(file.getUrl());
        if (i === 0) firstFileId = file.getId();
      });
    }

    // 3. AI phân tích CHẠY NỀN (async): không phân tích ngay để upload nhanh.
    //    Chỉ lưu ảnh + ghi Sheet, rồi xếp hàng cho trigger phân tích sau (~1 phút).
    var clientEntities = data.entities || [];
    var ai = emptyAnalysis();
    var merged = mergeEntities(clientEntities, []);

    // 5. Ghi Sheet
    var ss = SpreadsheetApp.openById(SHEET_ID);
    var mapLink = (data.lat && data.lng)
      ? '=HYPERLINK("https://www.google.com/maps?q=' + data.lat + ',' + data.lng + '","' +
        String(data.address || 'Xem bản đồ').replace(/"/g, '""') + '")'
      : '';
    var issuesText = formatIssues(ai.issues);

    var row = [
      Utilities.formatDate(new Date(), 'GMT+7', 'yyyy-MM-dd HH:mm:ss'),
      project,
      data.category || ai.category,   // hạng mục người dùng chọn (nếu có) > AI
      ai.description,
      issuesText,
      data.note || '',
      pickEntities(merged, 'serialNumber'),
      pickEntities(merged, 'weldId'),
      pickEntities(merged, 'dimension'),
      pickEntities(merged, 'drawingNumber'),
      pickEntities(merged, 'materialGrade'),
      data.ocrText || ai.rawText || '',
      ai.confidence,
      (data.lat && data.lng) ? (data.lat + ', ' + data.lng) : '',
      mapLink,
      imageUrls.join('\n'),
      senderName(data),           // V19: nhận diện theo ID máy trước, rồi mới tới tên theo tab Người dùng
      data.reportId || '',
      data.model || '',
      data.gate || '',
      data.unit || ''
    ];

    appendRow(ss, project, row);
    appendRow(ss, ALL_TAB, row);

    // Xếp hàng cho AI phân tích nền (cần reportId để cập nhật đúng dòng sau này)
    if (GEMINI_API_KEY && firstFileId && data.reportId) {
      enqueueAi(ss, data.reportId, firstFileId);
    }

    // V24.1: báo Telegram NGAY khi nhận báo cáo mới (kèm ảnh) — lỗi gửi tin không làm hỏng việc lưu
    if (TELEGRAM_NOTIFY_ALL) {
      try { sendNewReportAlert(row); } catch (etg) {}
    }

    return reply({ ok: true, imageUrl: imageUrls[0] || '', imageUrls: imageUrls,
                   photoCount: imageUrls.length, project: project,
                   aiPending: !!(GEMINI_API_KEY && firstFileId && data.reportId) });
  } catch (err) {
    return reply({ ok: false, error: String(err) });
  }
}
