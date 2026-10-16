// BẢN LƯU — ĐÃ CHE KHOÁ.
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
 * V51 (16/10): punchAddPhoto — thêm ảnh TRƯỚC/SAU vào 1 mục có sẵn từ app web (theo SID);
 *   sheetInfo — link mở Google Sheet đúng tab dự án (không ghi cứng link trong app công khai).
 * V52 (16/10, rà soát 10 tác tử): khoá ghi LockService cho mọi lệnh ghi (chống trùng No./ghi nhầm dòng);
 *   PIN cho catDelProj/descReplace/catalogBulk; simpleColMap không để Photo trùng Remarks (tab VIOLA);
 *   cấp SID 1 lần cho cả cột; punchFixImg dừng an toàn ~4,5 phút; listOpenPunch trả section/part/ảnh;
 *   ghi Result vào ghi chú đóng lỗi; punchAddPhoto ảnh trước giữ đúng ảnh sau cũ.
 * V53 (16/10): thêm plFetchCellImg (nút "Lấy ảnh DÁN TAY" lỗi vì thiếu hàm này);
 *   ảnh dán vào chính cột Photo link không còn bị coi là "đã có link" rồi bỏ qua.
 * V54 (16/10): "Lấy ảnh DÁN TAY" lấy được cả ẢNH NỔI (dán đè lên ô) qua bản xuất Excel tạm.
 * V55 (16/10): "Lấy ảnh DÁN TAY" không giữ khoá ghi chung nhiều phút (QC khác không bị báo bận);
 *   mỗi đợt ~2,5 phút, app tự chạy tiếp; chặn bấm chồng 2 lần; kiểm tra dòng trước khi ghi link.
 * V56 (16/10): listSimple lấy thêm ảnh từ ô chữ có gắn link (tên file) và công thức IMAGE()/HYPERLINK()
 *   ở cột Photo / Photo after (ảnh ở Photo after = ảnh SAU).
 * V57 (16/10): lấy ảnh nổi nhanh hơn — bản Excel tạm dựng 1 lần, các đợt sau dùng lại (30 phút).
 * V58 (16/10): tự động lấy ảnh dán tay mỗi giờ (cài 1 lần: chạy plCaiTuDongLayAnh).
 * V58 (16/10, rà soát 10 tác tử lần 2): thêm link ảnh KHÔNG làm mất link ẩn/công thức (plAppendLinks_);
 *   tên file không bị coi là ID ảnh; dòng lấy ảnh lỗi 1 phần không ghi link (lần sau lấy lại);
 *   listSimple/listOpenPunch không giữ khoá chung, cấp SID dưới khoá ngắn + đọc lại; danh sách đóng lỗi
 *   tới 2000 mục + cờ truncated, bỏ dòng trống/chữ "CellImage"; SID mới khó trùng (uuid);
 *   punchClose báo final/code khi đã đóng/không thấy; punchOpen báo simpleError; PIN không bỏ qua bởi dryRun
 *   (trừ catDelProj/punchDedupe); listReports chỉ mở tab dự án; dbgImg/searchFolders/setProjectFolder cần PIN;
 *   punchFixImg chỉ nhận tab "Punch list đơn giản…", so thêm SID, không dựng Excel tạm sau 60 giây,
 *   nhớ lần xuất Excel lỗi; bản Excel tạm kiểm theo vị trí ảnh nổi; punchDelete đánh số đúng cột No. 1 lần.
 * V59 (16/10, nâng cấp lớn): chống trùng BỀN (REQ:<reqId> trong ghi chú ô SID); plPing (báo bản máy chủ) + plDiag (kiểm tra hệ thống); listSimple thêm Mức độ/Người khắc phục/Hạn (chỉ khi có tiêu đề) + đếm ảnh dán/ảnh nổi chưa lên Drive; lưu trạng thái tự lấy ảnh; bỏ file PLMIG_ mồ côi khi dòng lấy ảnh lỗi.
 */

// ====== CẤU HÌNH — ⚠️ APP MỚI: PHẢI ĐIỀN LẠI CHO RIÊNG ======
var SECRET_KEY     = '***GIU_NGUYEN_CUA_ANH***';   // KHÁC FieldSnap — dùng đúng chuỗi này trong index.html app mới
var SHEET_ID       = '***ID_GOOGLE_SHEET_CUA_ANH***';   // Sheet "Punch List - VIOLA" (themy1809)
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
var IDEM_ACTS = { punchOpen: 1, punchSimple: 1, punchClose: 1, punchAddPhoto: 1 };   // V51: + punchAddPhoto
var _IDEM_KEY = '';   // khoá của lần gọi hiện tại (biến toàn cục reset mỗi lần chạy)
var _IDEM_TTL = 300;
// V59: bản máy chủ — app hỏi bằng action plPing để biết bản mới đã Deploy chưa (tăng cả 2 mỗi lần phát hành)
var PL_SERVER_VER = 'V59';
var PL_SERVER_NOTE = 'V59 16/10 nâng cấp lớn';

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
  // PL_COLMAP_V59 — GIỮ NGUYÊN dòng này: plPing dùng để phát hiện tệp khác có simpleColMap cũ ghi đè.
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
  // V52: tab thiếu tiêu đề "Remarks" mà ô "Photo" gộp 2 cột (vd VIOLA: E3:F3 = Photo) → Photo trùng Remarks.
  //      Khi đó Remarks = cột đầu, Photo = cột kế bên (ô tiêu đề trống do gộp) — không ghi ảnh đè lên ghi chú.
  var rem = find(['remarks', 'remark', 'ghi chú'], 5);
  var photo = find(['photo after'], 0) === 6 ? 6 : find(['photo', 'ảnh phát hiện', 'hình'], 6);
  if (photo === rem && !nrm(hd[photo]) && photo + 1 <= lastCol) photo = photo + 1;
  return {
    lastCol: lastCol,
    no:    find(['no.', 'no', 'stt'], 1),
    sec:   find(['section', 'sec'], 2),
    part:  find(['part', 'member no'], 3),
    desc:  find(['description', 'mô tả', 'mo ta'], 4),
    rem:   rem,
    photo: photo,
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
// V52: ô chứa ẢNH DÁN (CellImage) → trả chuỗi rỗng, không gửi chữ "CellImage" về app
function plTxt_(v) { return (v && typeof v === 'object' && !(v instanceof Date)) ? '' : String(v == null ? '' : v); }
// V59: ô tiêu đề (dòng 3) của cột col có đúng tên mong đợi không
function plHdrOk_(sh, col, re) {
  if (!col) return false;
  try { return re.test(String(sh.getRange(3, col).getValue() || '').trim().toLowerCase()); } catch (e) { return false; }
}
// V56: ID ảnh Drive trong 1 chuỗi (link /d/ID, ?id=ID, thumbnail?id=ID) — dùng cho công thức IMAGE/HYPERLINK
function plDriveIds_(s) {
  var out = [], re = /(?:\/d\/|[?&]id=)([-\w]{25,})/g, m;
  s = String(s || '');
  while ((m = re.exec(s))) if (out.indexOf(m[1]) < 0) out.push(m[1]);
  return out;
}
// V56: ID ảnh từ đường link ẩn của ô chữ có gắn link (RichTextValue)
function plLinkIds_(rt) {
  var out = [];
  if (!rt) return out;
  try {
    var add = function (u) { plDriveIds_(u).forEach(function (x) { if (out.indexOf(x) < 0) out.push(x); }); };
    add(rt.getLinkUrl());
    rt.getRuns().forEach(function (r) { add(r.getLinkUrl()); });
  } catch (e) {}
  return out;
}
// V58: ID ảnh trong CHỮ của ô link — chỉ lấy theo ngữ cảnh URL (/d/ID, ?id=ID) hoặc dòng là ID trần.
//      Tên file "PLSclose_…_1.jpg" (≥25 ký tự) KHÔNG còn bị coi là ID ảnh.
function plTextIds_(s) {
  var out = [];
  String(s || '').split('\n').forEach(function (ln) {
    var ids = plDriveIds_(ln), t = String(ln).trim();
    if (!ids.length && /^[-\w]{25,}$/.test(t)) ids = [t];
    ids.forEach(function (x) { if (out.indexOf(x) < 0) out.push(x); });
  });
  return out;
}
// V58: mọi ID ảnh của 1 ô link: chữ + link ẩn (RichText) + công thức HYPERLINK/IMAGE
function plCellIds_(range) {
  var ids = plTextIds_(plTxt_(range.getValue()));
  var add = function (x) { if (ids.indexOf(x) < 0) ids.push(x); };
  try { plLinkIds_(range.getRichTextValue()).forEach(add); } catch (eR) {}
  try { plDriveIds_(range.getFormula()).forEach(add); } catch (eF) {}
  return ids;
}
// V58: THÊM link ảnh vào ô — giữ dòng chữ cũ, link ẩn/công thức đổi thành dòng link rõ (setValue xoá link ẩn),
//      rồi nối link mới, không trùng ID.
function plAppendLinks_(range, urls) {
  var lines = plTxt_(range.getValue()).split('\n').filter(function (x) { return String(x).trim(); });
  var have = plTextIds_(lines.join('\n'));
  plCellIds_(range).forEach(function (id) {
    if (have.indexOf(id) < 0) { have.push(id); lines.push('https://drive.google.com/file/d/' + id + '/view'); }
  });
  (urls || []).forEach(function (u) {
    var id = plDriveIds_(u)[0] || '';
    if (id && have.indexOf(id) >= 0) return;
    if (id) have.push(id);
    lines.push(u);
  });
  range.setValue(lines.join('\n'));
}
// V58: cấp SID cho dòng chưa có — chỉ dưới khoá ngắn, đọc lại cột SID ngay trước khi ghi,
//      chỉ điền ô CÒN TRỐNG và dòng chưa bị dời (số dòng + mô tả như lúc đọc). Không có khoá → bỏ qua, không ghi.
//      vals = ảnh chụp getValues từ dòng 4 (được cập nhật SID đã ghi / SID người khác vừa cấp).
function plStampSids_(sh, m, vals) {
  var need = [];
  for (var i = 0; i < vals.length; i++)
    if (!String(vals[i][m.sid - 1] || '').trim() && String(vals[i].join('')).trim()) need.push(i);
  if (!need.length) return;
  var lk = LockService.getScriptLock();
  if (!lk.tryLock(10000)) return;
  try {
    if (sh.getLastRow() - 3 !== vals.length) return;   // có người thêm/xoá dòng → lần sau cấp
    var cur = sh.getRange(4, m.sid, vals.length, 1).getValues();
    var dsc = sh.getRange(4, m.desc, vals.length, 1).getValues();
    var stamp = 'SP' + Date.now(), chg = false;
    need.forEach(function (i) {
      var now = String(cur[i][0] || '').trim();
      if (now) { vals[i][m.sid - 1] = now; return; }   // người khác vừa cấp → dùng SID đó
      if (plTxt_(dsc[i][0]).trim() !== plTxt_(vals[i][m.desc - 1]).trim()) return;
      cur[i][0] = stamp + '_' + (4 + i); vals[i][m.sid - 1] = cur[i][0]; chg = true;
    });
    if (chg) { sh.getRange(4, m.sid, vals.length, 1).setValues(cur); SpreadsheetApp.flush(); }
  } finally { try { lk.releaseLock(); } catch (eL) {} }
}

// V54: ẢNH NỔI (dán đè lên ô) — Apps Script không đọc được nội dung ảnh nổi. Cách làm:
//      chép tab sang 1 file tạm → xuất .xlsx → giải nén → lấy ảnh trong xl/media theo vị trí
//      hàng/cột ghi trong xl/drawings. File tạm bỏ vào thùng rác ngay. KHÔNG sửa tab gốc.
//      Trả về { số hàng: [{ c: cột, blob, u: đã dùng }] } hoặc null nếu không lấy được.
function plOgFromXlsx_(sh) {
  var tmp = null, cache = CacheService.getScriptCache(), ck = 'pl_xlsx_' + sh.getSheetId();
  try {
    // V57: bản Excel tạm giữ lại 30 phút cho các đợt sau (trước đây đợt nào cũng dựng lại → chậm, không kịp xong)
    //      Số dòng của tab đổi (có người thêm/xoá dòng) → dựng lại để ảnh không lệch hàng.
    // V58: khoá so = số dòng + vị trí (hàng,cột) từng ảnh nổi + số ảnh → ảnh bị thay/dời thì dựng lại, không lấy ảnh cũ
    var sig = String(sh.getLastRow());
    try {
      var sgI = sh.getImages(), sgA = [];
      for (var sgk = 0; sgk < sgI.length; sgk++) { var sgC = sgI[sgk].getAnchorCell(); sgA.push(sgC.getRow() + ',' + sgC.getColumn()); }
      sgA.sort();
      sig += ':' + sgA.join(';') + ':' + sgI.length;
    } catch (eSg) {}
    var zb = null, ckv = String(cache.get(ck) || '').split('|');
    if (ckv[0] === 'FAIL' && ckv[1] === sig) return null;   // V58: vừa xuất lỗi (30') → không dựng lại mỗi đợt
    if (ckv[0] && ckv[0] !== 'FAIL' && ckv[1] === sig) { try { zb = DriveApp.getFileById(ckv[0]).getBlob(); } catch (eC) { zb = null; } }
    if (!zb) {
      tmp = SpreadsheetApp.create('PL_TAM_ANH_NOI_' + Date.now());
      var cp = sh.copyTo(tmp);
      tmp.getSheets().forEach(function (x) { if (x.getSheetId() !== cp.getSheetId()) tmp.deleteSheet(x); });
      SpreadsheetApp.flush();
      var rs = UrlFetchApp.fetch('https://docs.google.com/spreadsheets/d/' + tmp.getId() + '/export?format=xlsx',
        { headers: { Authorization: 'Bearer ' + ScriptApp.getOAuthToken() }, muteHttpExceptions: true });
      if (rs.getResponseCode() !== 200) { try { cache.put(ck, 'FAIL|' + sig, 1800); } catch (eFp) {} return null; }   // V58
      zb = rs.getBlob();
      try {
        plOgXlsxDone_(sh);
        var keep = DriveApp.createFile(zb.copyBlob().setName('PL_TAM_XLSX_' + sh.getSheetId() + '.xlsx'));
        cache.put(ck, keep.getId() + '|' + sig, 1800);   // V58: id|khoá so
      } catch (eK) {}
    }
    zb.setContentType('application/zip');
    var by = {};
    Utilities.unzip(zb).forEach(function (f) { by[f.getName()] = f; });
    var out = {}, n = 0;
    Object.keys(by).forEach(function (nm) {
      var m = nm.match(/^xl\/drawings\/(drawing\d+)\.xml$/);
      if (!m) return;
      var rels = {}, rf = by['xl/drawings/_rels/' + m[1] + '.xml.rels'];
      if (rf) rf.getDataAsString().replace(/<Relationship\b[^>]*>/g, function (t) {
        var id = (t.match(/Id="([^"]+)"/) || [])[1], tg = (t.match(/Target="([^"]+)"/) || [])[1];
        if (id && tg) rels[id] = tg;
        return t;
      });
      var xml = by[nm].getDataAsString();
      var re = /<(?:\w+:)?(twoCellAnchor|oneCellAnchor)\b[\s\S]*?<\/(?:\w+:)?\1>/g, a;
      while ((a = re.exec(xml))) {
        var blk = a[0];
        var from = (blk.match(/<(?:\w+:)?from>([\s\S]*?)<\/(?:\w+:)?from>/) || [])[1] || '';
        var col = (from.match(/<(?:\w+:)?col>(\d+)</) || [])[1];
        var row = (from.match(/<(?:\w+:)?row>(\d+)</) || [])[1];
        var rid = (blk.match(/r:embed="([^"]+)"/) || [])[1];
        if (col == null || row == null || !rid || !rels[rid]) continue;
        var mb = by['xl/' + rels[rid].replace(/^\.\.\//, '').replace(/^\/?xl\//, '')];
        if (!mb) continue;
        var r1 = Number(row) + 1;
        (out[r1] = out[r1] || []).push({ c: Number(col) + 1, blob: mb.copyBlob(), u: false });
        n++;
      }
    });
    return n ? out : null;
  } catch (e) {
    try { if (sig != null) cache.put(ck, 'FAIL|' + sig, 1800); } catch (eFc) {}   // V58: nhớ lần lỗi, đợt sau không dựng lại
    return null;
  } finally {
    if (tmp) try { DriveApp.getFileById(tmp.getId()).setTrashed(true); } catch (e2) {}
  }
}
// ============ V58 — TỰ ĐỘNG LẤY ẢNH DÁN TAY mỗi giờ (user duyệt 16/10: tự động + vẫn bấm tay được) ============
// Cài 1 lần: trong Apps Script chọn hàm plCaiTuDongLayAnh → ▶ Run → cho phép. Tắt: chạy plTatTuDongLayAnh.
// Mỗi giờ quét mọi tab "Punch list đơn giản…": dòng có ảnh dán (trong ô / ảnh nổi) mà chưa có link
// → chép lên Drive + ghi link (dùng lại đúng lệnh punchFixImg của nút trong app, không viết lại logic).
function plCaiTuDongLayAnh() {
  plTatTuDongLayAnh();
  ScriptApp.newTrigger('plTuDongLayAnh').timeBased().everyHours(1).create();
  return 'Đã bật tự động lấy ảnh dán tay mỗi giờ';
}
function plTatTuDongLayAnh() {
  ScriptApp.getProjectTriggers().forEach(function (t) {
    if (t.getHandlerFunction() === 'plTuDongLayAnh') ScriptApp.deleteTrigger(t);
  });
}
function plTuDongLayAnh() {
  var t0 = new Date().getTime(), norm = function (s) { return String(s).toLowerCase().replace(/[\s\-–_]+/g, ''); };
  var log = [], err = '';
  try {
    var tabs = SpreadsheetApp.openById(SHEET_ID).getSheets().map(function (s) { return s.getName(); })
      .filter(function (n) { return norm(n).indexOf(norm('Punch list đơn giản')) === 0; });
    if (!tabs.length) { log.push('Không có tab "Punch list đơn giản…"'); return; }
    var props = PropertiesService.getScriptProperties();
    var start = (Number(props.getProperty('PL_AUTO_IMG_I')) || 0) % tabs.length;   // lần sau bắt đầu từ tab kế tiếp
    for (var k = 0; k < tabs.length; k++) {
      if (new Date().getTime() - t0 > 150000) break;   // V59: 180s → 150s — chừa thời gian, Google cắt ở 6 phút
      var nm = tabs[(start + k) % tabs.length];
      var out = plDoPost_({ postData: { contents: JSON.stringify({ key: SECRET_KEY, action: 'punchFixImg', sheet: nm, _auto: true }) } });
      var d = out;
      try { if (out && typeof out.getContent === 'function') d = JSON.parse(out.getContent()); } catch (eJ) {}
      log.push(nm + ': ' + ((d && d.saved) || 0) + ' ảnh' + (d && d.partial ? ' (còn tiếp)' : '') + (d && d.error ? ' ' + d.error : ''));
      props.setProperty('PL_AUTO_IMG_I', String((start + k + (d && d.partial ? 0 : 1)) % tabs.length));
      if (d && d.partial) break;   // tab này chưa xong → giờ sau làm tiếp đúng tab này
    }
  } catch (eA) {
    err = String((eA && eA.message) || eA);
  } finally {
    Logger.log(log.join('\n') + (err ? '\nLỖI: ' + err : ''));
    // V59: lưu trạng thái lần tự lấy ảnh gần nhất (plDiag hiển thị "lần chạy cuối")
    try {
      PropertiesService.getScriptProperties().setProperty('PL_AUTO_IMG_LAST', JSON.stringify({
        t: Utilities.formatDate(new Date(), 'GMT+7', 'yyyy-MM-dd HH:mm'), ms: new Date().getTime() - t0,
        tabs: log.slice(0, 20).map(function (x) { return String(x).slice(0, 120); }), err: err.slice(0, 300) }));
    } catch (eS) {}
  }
}

// V57: xong hết → bỏ bản Excel tạm vào thùng rác
function plOgXlsxDone_(sh) {
  var cache = CacheService.getScriptCache(), ck = 'pl_xlsx_' + sh.getSheetId();
  var v = String(cache.get(ck) || '').split('|')[0];
  if (v) { try { DriveApp.getFileById(v).setTrashed(true); } catch (e) {} }
  try { cache.remove(ck); } catch (e2) {}
}
// Lấy 1 ảnh nổi của hàng r: ưu tiên đúng cột, chưa dùng
function plOgPick_(map, r, c) {
  var L = (map && map[r]) || [], i;
  for (i = 0; i < L.length; i++) if (!L[i].u && L[i].c === c) { L[i].u = true; return L[i].blob; }
  for (i = 0; i < L.length; i++) if (!L[i].u) { L[i].u = true; return L[i].blob; }
  return null;
}

// V53: tải ảnh DÁN TAY trong ô (CellImage) → blob để lưu lên Drive. Thiếu hàm này nên nút
//      "Lấy ảnh DÁN TAY" trước đây lỗi ngay khi gặp ảnh trong ô. Thử kèm OAuth trước, rồi không kèm.
function plFetchCellImg(u) {
  var tries = [{ muteHttpExceptions: true, headers: { Authorization: 'Bearer ' + ScriptApp.getOAuthToken() } },
               { muteHttpExceptions: true }];
  for (var i = 0; i < tries.length; i++) {
    try {
      var rs = UrlFetchApp.fetch(u, tries[i]);
      if (rs.getResponseCode() !== 200) continue;
      var bl = rs.getBlob();
      if (/^image\//i.test(String(bl.getContentType() || ''))) return bl;
    } catch (e) {}
  }
  return null;
}
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

// ============ V59 — CHỐNG TRÙNG BỀN (không thêm cột) ============
// Mỗi lần ghi thành công: thêm dòng "REQ:<reqId>" vào GHI CHÚ ô SID của dòng vừa tạo/sửa (giữ chữ ghi chú khác,
// tối đa 5 dòng REQ mới nhất). Cache 6 giờ bị xoá/hết hạn → vẫn tra được trong ghi chú → không ghi lần 2.
function plReqId_(data) { return (data && data.reqId) ? String(data.reqId).replace(/\W/g, '').slice(0, 60) : ''; }
// keep non-REQ note text; keep at most 5 REQ lines (newest last)
function plReqMark_(cell, rid) {
  if (!cell || !rid) return;
  try {
    var lines = String(cell.getNote() || '').split('\n');
    var other = lines.filter(function (l) { return l !== '' && !/^REQ:/.test(l); });
    var reqs = lines.filter(function (l) { return /^REQ:/.test(l) && l !== 'REQ:' + rid; });
    reqs.push('REQ:' + rid);
    cell.setNote(other.concat(reqs.slice(-5)).join('\n'));
  } catch (e) {}   // a failed note must never fail the write
}
function plReqFind_(sh, col, firstRow, rid) {
  if (!sh || !col || !rid || sh.getLastRow() < firstRow) return -1;
  var notes = sh.getRange(firstRow, col, sh.getLastRow() - firstRow + 1, 1).getNotes();   // ONE call
  var re = new RegExp('(^|\\n)REQ:' + rid + '(\\n|$)');
  for (var i = 0; i < notes.length; i++) if (notes[i][0] && re.test(String(notes[i][0]))) return firstRow + i;
  return -1;
}
// Trả kết quả kiểu cũ (như lần ghi đầu) khi reqId đã có trong ghi chú; không thấy → null (ghi bình thường).
function plReqLookup_(data, rid) {
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var sp = ss.getSheetByName(SIMPLE_TAB);
  var m = sp ? simpleColMap(sp) : null;
  var q = null;
  var inSimple = function () { return m ? plReqFind_(sp, m.sid, 4, rid) : -1; };
  var inViola = function () { q = ss.getSheetByName(PUNCH_TAB); return q ? plReqFind_(q, 15, 8, rid) : -1; };
  var act = String(data.action || ''), r, v;
  if (act === 'punchSimple') {
    r = inSimple();
    if (r > 0) return { ok: true, no: plTxt_(sp.getRange(r, m.no).getValue()), id: plTxt_(sp.getRange(r, m.sid).getValue()), photoCount: 0, photos: [] };
    return null;
  }
  if (act === 'punchOpen') {
    v = inViola();
    if (v > 0) return { ok: true, punchId: plTxt_(q.getRange(v, 15).getValue()), itemNo: plTxt_(q.getRange(v, 7).getValue()),
      project: '', openPhotos: [], photoCount: 0, simpleError: '' };
    r = inSimple();
    if (r > 0) return { ok: true, punchId: plTxt_(sp.getRange(r, m.sid).getValue()), itemNo: 'No.' + plTxt_(sp.getRange(r, m.no).getValue()),
      project: '', openPhotos: [], photoCount: 0, simpleError: '' };
    return null;
  }
  if (act === 'punchClose') {
    var id = String(data.punchId || '').trim();
    if (/^SP/.test(id)) r = inSimple();
    else { r = inViola(); if (!(r > 0)) r = inSimple(); }
    if (r > 0) return { ok: true, punchId: id, closePhotos: [], photoCount: 0 };
    return null;
  }
  if (act === 'punchAddPhoto') {
    r = inSimple();
    if (r > 0) return { ok: true, sid: String(data.sid || '').trim(), kind: (String(data.kind) === 'after' ? 'after' : 'before'), added: 0, row: r };
    return null;
  }
  return null;
}
// V59: bỏ vào thùng rác (khôi phục được 30 ngày) các file vừa tạo trong LẦN CHẠY NÀY cho 1 dòng mà link KHÔNG được ghi
function plTrashIds_(ids) { (ids || []).forEach(function (id) { try { DriveApp.getFileById(id).setTrashed(true); } catch (e) {} }); }
// V59: số cột → chữ cột (1 → A, 27 → AA)
function plColL_(c) {
  var s = ''; c = Number(c) || 0;
  while (c > 0) { var k = (c - 1) % 26; s = String.fromCharCode(65 + k) + s; c = Math.floor((c - 1) / 26); }
  return s;
}

// ============ V59 — KIỂM TRA HỆ THỐNG (plDiag): CHỈ ĐỌC, không ghi, không gọi Drive ============
function plDiag_() {
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var props = null;
  try { props = PropertiesService.getScriptProperties(); } catch (eP) { props = null; }
  var pj = function (k) { try { return props ? JSON.parse(props.getProperty(k) || 'null') : null; } catch (e) { return null; } };
  var auto = { triggerOn: false, last: pj('PL_AUTO_IMG_LAST') };
  try { auto.triggerOn = ScriptApp.getProjectTriggers().some(function (t) { return t.getHandlerFunction() === 'plTuDongLayAnh'; }); } catch (eT) { auto.triggerOn = false; }
  var counts = { noSid: 0, dupSid: 0, pastedNoLink: 0, floatNoLink: 0, filenameOnly: 0 };
  var samples = { pastedNoLink: [], floatNoLink: [], filenameOnly: [], noSid: [], dupSid: [] };
  var out = { ok: true, server: PL_SERVER_VER, tab: SIMPLE_TAB, tabFound: false, rows: 0, cols: [], warnings: [],
              counts: counts, samples: samples, auto: auto, lastFix: null };
  var sh = ss.getSheetByName(SIMPLE_TAB);
  if (!sh) {
    var norm = function (x) { return String(x).toLowerCase().replace(/[\s\-–_]+/g, ''); };
    var base = norm('Punch list đơn giản');
    out.tabsSeen = ss.getSheets().map(function (x) { return x.getName(); }).filter(function (n) { return norm(n).indexOf(base) === 0; });
    out.warnings.push('Không thấy tab "' + SIMPLE_TAB + '" — kiểm tra tên tab dự án.');
    return out;
  }
  out.tabFound = true;
  out.lastFix = pj('PL_FIX_LAST_' + sh.getSheetId());
  var m = simpleColMap(sh);
  var hd = [];
  try { hd = sh.getRange(3, 1, 1, m.lastCol).getDisplayValues()[0]; } catch (eH) { hd = []; }
  var hn = function (c) { return String(hd[c - 1] || '').trim().toLowerCase(); };
  var RX = {
    no: /^(no\.?|stt)$/, sec: /^(section|sec)/, part: /^(part|member no)/, desc: /^(description|mô tả|mo ta)/,
    rem: /^(remarks?|ghi chú)/, photo: /^(photo|ảnh phát hiện|hình)/, link: /^(photo link|link ảnh)/,
    st: /^(status|tình trạng)/, sid: /^sid$/, after: /^(photo after|ảnh khắc phục)/, sev: /^(mức độ|muc do|severity)/,
    asg: /^(người khắc phục|nguoi khac phuc|assignee)/, due: /^(hạn xử lý|han xu ly|due)/,
    insp: /^(người lập|nguoi lap|người kiểm|nguoi kiem|inspect)/
  };
  var isPhotoHd = function (h) { return RX.photo.test(h) && !/^photo (after|link)/.test(h); };
  var keys = ['no', 'sec', 'part', 'desc', 'rem', 'photo', 'link', 'st', 'sid', 'after', 'sev', 'asg', 'due', 'insp'];
  var byH = {};
  keys.forEach(function (k) {
    var c = m[k] || 0;
    if (!c) { if (k === 'insp') out.cols.push({ key: 'insp', col: 0, letter: '', header: '', byHeader: false }); return; }
    var h = hn(c), ok;
    if (k === 'photo') ok = isPhotoHd(h) || (h === '' && c > 1 && isPhotoHd(hn(c - 1)));   // V52: tiêu đề Photo gộp 2 ô
    else ok = RX[k].test(h);
    byH[k] = ok;
    out.cols.push({ key: k, col: c, letter: plColL_(c), header: String(hd[c - 1] || ''), byHeader: ok });
  });
  // --- cảnh báo cấu trúc ---
  var nSid = 0;
  for (var hi = 0; hi < hd.length; hi++) if (String(hd[hi] || '').trim().toLowerCase() === 'sid') nSid++;
  if (nSid > 1) out.warnings.push('Có ' + nSid + ' cột tiêu đề "SID" — máy chủ dùng cột ' + plColL_(m.sid) + ' làm SID, cột ' + plColL_(m.link) + ' làm link ảnh.');
  if (!byH.link) out.warnings.push('Cột link ảnh (' + plColL_(m.link) + ') được đoán theo dữ liệu, tiêu đề đang là "' + String(hd[m.link - 1] || '') + '" — nên đổi tiêu đề thành "Photo link".');
  var hasAfterHd = false;
  for (var ha = 0; ha < hd.length; ha++) if (RX.after.test(String(hd[ha] || '').trim().toLowerCase())) hasAfterHd = true;
  if (!hasAfterHd) out.warnings.push('Thiếu tiêu đề "Photo after" — ảnh SAU có thể bị ghi vào cột ' + plColL_(m.after) + '.');
  if (!m.insp) out.warnings.push('Thiếu cột "Người lập" — không lọc được theo QC.');
  var ck = ['no', 'sec', 'part', 'desc', 'rem', 'photo', 'link', 'st', 'sid', 'after', 'sev', 'asg', 'due'], seenC = {};
  ck.forEach(function (k) {
    var c = m[k] || 0;
    if (!c) return;
    if (seenC[c]) out.warnings.push('Cột ' + plColL_(c) + ' đang được dùng cho 2 mục: ' + seenC[c] + ' và ' + k + ' (thiếu tiêu đề) — có thể ghi đè dữ liệu.');
    else seenC[c] = k;
  });
  // --- quét dữ liệu: 1 lần getValues + 1 lần chữ hiển thị cột link + 1 lần link ẩn cột link + 1 lần ảnh nổi ---
  var n = Math.max(0, sh.getLastRow() - 3);
  var vals = [], ldv = [], lrt = [], flt = {};
  if (n) {
    vals = sh.getRange(4, 1, n, m.lastCol).getValues();
    try { ldv = sh.getRange(4, m.link, n, 1).getDisplayValues(); } catch (eD) { ldv = []; }
    try { lrt = sh.getRange(4, m.link, n, 1).getRichTextValues(); } catch (eR) { lrt = []; }
  }
  try { sh.getImages().forEach(function (g) { var r = g.getAnchorCell().getRow(); if (r >= 4) flt[r] = (flt[r] || 0) + 1; }); } catch (eG) {}
  var isImg = function (v) { return v && typeof v === 'object' && !(v instanceof Date) && typeof v.getContentUrl === 'function'; };
  var sidSeen = {}, rowsN = 0;
  var push = function (k, row) { counts[k]++; if (samples[k].length < 5) samples[k].push(row); };
  for (var i = 0; i < n; i++) {
    var row = i + 4, lr = vals[i];
    var hasCellImg = false;
    for (var c2 = 0; c2 < lr.length; c2++) if (isImg(lr[c2])) { hasCellImg = true; break; }
    var lt = String((ldv[i] && ldv[i][0]) || plTxt_(lr[m.link - 1]) || '');
    if (lt === 'CellImage') lt = '';
    var hasContent = plTxt_(lr[m.desc - 1]).trim() || plTxt_(lr[m.part - 1]).trim() || lt.trim() || hasCellImg || flt[row];
    if (!hasContent) continue;
    rowsN++;
    var rich = plLinkIds_(lrt[i] && lrt[i][0]);
    var hasId = plTextIds_(lt).length > 0 || rich.length > 0;
    var sidV = plTxt_(lr[m.sid - 1]).trim();
    if (!sidV) push('noSid', row);
    else if (sidSeen[sidV]) push('dupSid', row);
    else sidSeen[sidV] = 1;
    if (/\.(jpe?g|png|heic)\b/i.test(lt) && !plTextIds_(lt).length && !rich.length) push('filenameOnly', row);
    if (hasCellImg && !hasId) push('pastedNoLink', row);
    if (flt[row] && !hasId) push('floatNoLink', row);
  }
  out.rows = rowsN;
  var rowsTxt = function (k) { return 'dòng ' + samples[k].join(', ') + (counts[k] > samples[k].length ? ', …' : ''); };
  if (counts.noSid) out.warnings.push(counts.noSid + ' dòng chưa có SID (app tự cấp khi mở Báo cáo): ' + rowsTxt('noSid') + '.');
  if (counts.dupSid) out.warnings.push(counts.dupSid + ' dòng có SID trùng dòng khác — đóng lỗi/thêm ảnh có thể vào nhầm dòng: ' + rowsTxt('dupSid') + '.');
  if (counts.pastedNoLink) out.warnings.push(counts.pastedNoLink + ' dòng có ảnh dán trong ô chưa lên Drive (tự lấy mỗi giờ hoặc bấm "Lấy ảnh dán tay"): ' + rowsTxt('pastedNoLink') + '.');
  if (counts.floatNoLink) out.warnings.push(counts.floatNoLink + ' dòng có ảnh nổi (dán đè lên ô) chưa lên Drive: ' + rowsTxt('floatNoLink') + '.');
  if (counts.filenameOnly) out.warnings.push(counts.filenameOnly + ' dòng ô link ảnh chỉ có tên file, không có link Drive — app không mở được ảnh: ' + rowsTxt('filenameOnly') + '.');
  return out;
}

// V52: KHOÁ GHI — mọi lệnh ghi Sheet chạy lần lượt (2 QC gửi cùng lúc không trùng No. / không ghi nhầm dòng).
// (V58: lệnh đọc listSimple/listOpenPunch không còn xin khoá ở đây — xem plStampSids_.)
var PL_WRITE_ACTS = { punchSimple: 1, punchOpen: 1, punchClose: 1, punchAddPhoto: 1, punchDelete: 1, punchDedupe: 1,
  punchClearAll: 1, punchFix: 1, punchRepair: 1, catalogAdd: 1, catalogBulk: 1, descReplace: 1, catDelProj: 1 };
// V58: listSimple/listOpenPunch KHÔNG giữ khoá chung khi đọc nữa (QC khác không bị báo bận);
//      chỉ xin khoá ngắn lúc cấp SID (plStampSids_).
var PL_READ_LOCK_ACTS = {};
function doPost(e) {
  var act = '';
  try { act = String(JSON.parse(e.postData.contents).action || ''); } catch (eP) {}
  var lock = null;
  if (PL_WRITE_ACTS[act] || PL_READ_LOCK_ACTS[act]) {
    lock = LockService.getScriptLock();
    if (!lock.tryLock(PL_WRITE_ACTS[act] ? 25000 : 8000)) {
      lock = null;
      if (PL_WRITE_ACTS[act]) return reply({ ok: false, busy: true, error: 'Máy chủ đang ghi cho người khác — app sẽ tự gửi lại sau vài giây.' });
    }
  }
  try { return plDoPost_(e); }
  finally { if (lock) { try { lock.releaseLock(); } catch (eL) {} } }
}

function plDoPost_(e) {
  try {
    _IDEM_KEY = '';   // V59: phiên chạy ấm (plTuDongLayAnh gọi plDoPost_ nhiều lần) — không dùng lại khoá của lần trước
    var data = JSON.parse(e.postData.contents);
    if (data.key !== SECRET_KEY) return reply({ ok: false, error: 'Sai khóa bảo mật' });

    // ===== V43 — KIỂM SOÁT XOÁ: action xoá dữ liệu phải kèm đúng mã PIN quản trị =====
    var PIN_ACTS = { punchDelete: 1, punchClearAll: 1, punchDedupe: 1, catDelProj: 1, descReplace: 1, catalogBulk: 1 };   // V52: + lệnh xoá/ghi đè danh mục
    var DRY_OK = { catDelProj: 1, punchDedupe: 1 };   // V58: chỉ 2 lệnh này có chạy thử thật — lệnh khác dryRun vẫn phải có PIN
    if (ADMIN_PIN && PIN_ACTS[data.action] && !(data.dryRun && DRY_OK[data.action]) && String(data.pin || '') !== ADMIN_PIN)
      return reply({ ok: false, error: 'Sai hoặc thiếu mã PIN quản trị — chỉ người quản lý mới được xoá.', needPin: true });
    // V58: lệnh quản trị/chẩn đoán (app không dùng) — cần PIN quản trị
    var PIN_ADMIN_ACTS = { dbgImg: 1, searchFolders: 1, setProjectFolder: 1 };
    if (ADMIN_PIN && PIN_ADMIN_ACTS[data.action] && String(data.pin || '') !== ADMIN_PIN)
      return reply({ ok: false, error: 'Chức năng này cần mã PIN quản trị.', needPin: true });

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
    // V59: CHỐNG TRÙNG BỀN — cache không còn nhớ (hết 6h / bị xoá) mà reqId đã có trong ghi chú ô SID
    //      → trả kết quả cũ, KHÔNG ghi lần 2 (đóng lỗi gửi lại = ok/duplicated, không phải ALREADY_CLOSED).
    if (_IDEM_KEY && IDEM_ACTS[data.action]) {
      var rqId = plReqId_(data);
      if (rqId) {
        var rqRes = null;
        try { rqRes = plReqLookup_(data, rqId); } catch (eRq) { rqRes = null; }
        if (rqRes) {
          try { CacheService.getScriptCache().put(_IDEM_KEY, JSON.stringify(rqRes), 21600); } catch (eRc) {}   // replaces 'RUNNING'
          rqRes.duplicated = true;
          return reply(rqRes);
        }
      }
    }
    // V59: plPing — báo bản máy chủ đang chạy (chỉ đọc, không PIN, không khoá; không trả id/khoá/PIN)
    if (data.action === 'plPing') return reply({ ok: true, server: PL_SERVER_VER, serverNote: PL_SERVER_NOTE,
      colMapOk: String(simpleColMap).indexOf('PL_COLMAP_V59') >= 0, idemNote: typeof plReqFind_ === 'function',
      acts: ['plPing', 'plDiag', 'punchFixImg', 'punchAddPhoto'], tz: Session.getScriptTimeZone(),
      now: Utilities.formatDate(new Date(), 'GMT+7', 'yyyy-MM-dd HH:mm:ss') });
    // V59: plDiag — kiểm tra hệ thống cho tab dự án đang chọn (chỉ đọc)
    if (data.action === 'plDiag') return reply(plDiag_());
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
      // V58: chỉ mở TAB DỰ ÁN (không cho đọc tab Người dùng / Dự án / sao lưu… theo tên tuỳ ý)
      if (proj && proj !== '__all__') {
        var rpBad = [USERS_TAB, PROJECTS_TAB, AI_QUEUE_TAB, 'Dashboard', 'Danh mục part', 'Mẫu mô tả'];
        if (rpBad.indexOf(proj) >= 0 || proj.indexOf('SAOLUU_') === 0 ||
            getProjects().map(function (p) { return p.name; }).indexOf(proj) === -1)
          return reply({ ok: false, error: 'Dự án không hợp lệ' });
      }
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
      var poId = 'PL' + Utilities.formatDate(new Date(), 'GMT+7', 'yyMMddHHmmss') + '_' + Utilities.getUuid().replace(/-/g, '').slice(0, 8);   // V58: khó trùng (giữ PL + 12 số ngày giờ)
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
      plReqMark_(poQ.getRange(poRow, 15), plReqId_(data));   // V59: chống trùng bền (ô O PUNCH_ID)
      if (poOpenIds.length) { poQ.getRange(poRow, 12).setFormula(punchImage(poOpenIds[0])); poQ.setRowHeight(poRow, 130); }
      // Ghi ĐỒNG THỜI vào tab ĐƠN GIẢN (cùng ID → đóng 1 lần là cả 2 sheet cùng close)
      var poSpErr = '';
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
        // V59: chỉ ghi Mức độ / Người khắc phục / Hạn khi tab CÓ tiêu đề đó (không thì cột mặc định 11-13 có thể là Photo after)
        if (plHdrOk_(poSp, poM.sev, /^(mức độ|muc do|severity)/))  poArr[poM.sev - 1]  = data.severity || '';
        if (plHdrOk_(poSp, poM.asg, /^(người khắc phục|nguoi khac phuc|assignee)/))  poArr[poM.asg - 1]  = data.assignee || '';
        if (plHdrOk_(poSp, poM.due, /^(hạn xử lý|han xu ly|due)/))  poArr[poM.due - 1]  = data.due || '';
        if (poM.insp) poArr[poM.insp - 1] = senderName(data);
        poSp.appendRow(poArr);
        var poSpRow = poSp.getLastRow();
        plReqMark_(poSp.getRange(poSpRow, poM.sid), plReqId_(data));   // V59
        poSp.getRange(poSpRow, 1, 1, poM.lastCol).setBorder(true, true, true, true, true, true).setVerticalAlignment('middle');
        poSp.getRange(poSpRow, poM.desc).setWrap(true);
        poSp.getRange(poSpRow, 1, 1, 3).setHorizontalAlignment('center');
        poSp.getRange(poSpRow, poM.st).setHorizontalAlignment('center');
        if (poOpenIds.length) { poSp.getRange(poSpRow, poM.photo).setFormula(punchImage(poOpenIds[0])); poSp.setRowHeight(poSpRow, 110); }
      } catch (ePoSp) { poSpErr = String(ePoSp); }   // V58: báo lỗi ghi tab đơn giản (không nuốt im lặng)
      return reply({ ok: true, punchId: poId, itemNo: '#' + poNo, project: poProject,
                     openPhotos: poOpenUrls, photoCount: poOpenUrls.length, simpleError: poSpErr });
    }

    // ============ PUNCH LIST — DANH SÁCH MỤC ĐANG MỞ (cho app chọn để đóng) ============
    if (data.action === 'listOpenPunch') {
      var lpSs = SpreadsheetApp.openById(SHEET_ID);
      var lpQ = lpSs.getSheetByName(PUNCH_TAB);
      var lpOut = [], lpTrunc = false;   // V58: lpTrunc = danh sách bị cắt (còn mục cũ hơn)
      // V42: tab VIOLA trống thì KHÔNG return sớm nữa — vẫn phải liệt kê tab đơn giản bên dưới
      if (lpQ && lpQ.getLastRow() >= 8) {
        var lpVals = lpQ.getRange(8, 1, lpQ.getLastRow() - 7, 18).getValues();
        for (var lpi = lpVals.length - 1; lpi >= 0; lpi--) {
          if (String(lpVals[lpi][3]).toLowerCase() !== 'open') continue;   // D STATUS
          if (lpOut.length >= 100) { lpTrunc = true; break; }   // V58: giữ trần 100 của tab VIOLA, báo bị cắt
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
          // V59: Mức độ / Người khắc phục / Hạn — chỉ khi có tiêu đề thật
          var lpHd = []; try { lpHd = lsQ.getRange(3, 1, 1, lsM.lastCol).getDisplayValues()[0]; } catch (eLh) {}
          var lpHok = function (col, re) { return !!col && re.test(String(lpHd[col - 1] || '').trim().toLowerCase()); };
          var lpSevOk = lpHok(lsM.sev, /^(mức độ|muc do|severity)/), lpAsgOk = lpHok(lsM.asg, /^(người khắc phục|nguoi khac phuc|assignee)/),
              lpDueOk = lpHok(lsM.due, /^(hạn xử lý|han xu ly|due)/);
          var lsV = lsQ.getRange(4, 1, lsQ.getLastRow() - 3, lsM.lastCol).getValues();
          // V43: dòng chưa có ID bền → cấp ngay; đóng/xoá bám ID, không bám số dòng (chống lệch khi có người xoá)
          // V58: cấp dưới khoá ngắn + đọc lại, chỉ điền ô còn trống (plStampSids_)
          plStampSids_(lsQ, lsM, lsV);
          for (var lsi = lsV.length - 1; lsi >= 0; lsi--) {
            if (plIsClosed_(lsV[lsi][lsM.st - 1])) continue;   // V49: close/Close/closed
            if (String(lsV[lsi][lsM.sid - 1] || '').indexOf('PL') === 0) continue; // dòng liên kết VIOLA — đã hiện ở trên, tránh trùng
            // V58: bỏ dòng trống (chỉ có No./status) — khớp số với Báo cáo/Dashboard; ô link có ảnh dán vẫn giữ
            if (!plTxt_(lsV[lsi][lsM.desc - 1]).trim() && !plTxt_(lsV[lsi][lsM.part - 1]).trim() && !String(lsV[lsi][lsM.link - 1] || '').trim()) continue;
            if (lpOut.length >= 2000) { lpTrunc = true; break; }   // V58: trần 2000 (trước 150 → mục CŨ bị mất)
            var lpIds = plTextIds_(plTxt_(lsV[lsi][lsM.link - 1]));   // V58: tên file không bị coi là ID ảnh
            lpOut.push({ punchId: String(lsV[lsi][lsM.sid - 1] || ('SPROW' + (4 + lsi))), itemNo: 'No.' + lsV[lsi][lsM.no - 1],
              project: 'Đơn giản', vendor: plTxt_(lsV[lsi][lsM.part - 1]), description: plTxt_(lsV[lsi][lsM.desc - 1]),
              date: '', workshop: plTxt_(lsV[lsi][lsM.sec - 1]), inspector: '', simple: true,
              section: plTxt_(lsV[lsi][lsM.sec - 1]), part: plTxt_(lsV[lsi][lsM.part - 1]), photoIds: lpIds.slice(0, 3) });   // V52
            var lpDue = lpDueOk ? lsV[lsi][lsM.due - 1] : '';   // V59
            var lpIt = lpOut[lpOut.length - 1];
            lpIt.severity = lpSevOk ? plTxt_(lsV[lsi][lsM.sev - 1]).trim() : '';
            lpIt.assignee = lpAsgOk ? plTxt_(lsV[lsi][lsM.asg - 1]).trim() : '';
            lpIt.due = (lpDue instanceof Date) ? Utilities.formatDate(lpDue, 'GMT+7', 'yyyy-MM-dd') : plTxt_(lpDue).trim();
          }
        }
      } catch (eLs) {}
      return reply({ ok: true, items: lpOut, truncated: lpTrunc });
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
        // V58: final + code → app không gửi lại mãi (gửi lại cũng không bao giờ thành công)
        if (scRow < 4 || scRow > scQ.getLastRow()) return reply({ ok: false, final: true, code: 'NOT_FOUND', error: 'Không tìm thấy mục ' + pcId });
        if (plIsClosed_(scQ.getRange(scRow, scM.st).getValue()))   // V49: close/Close/closed
          return reply({ ok: false, final: true, code: 'ALREADY_CLOSED', error: 'Mục này đã đóng trước đó' });
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
        plReqMark_(scQ.getRange(scRow, scM.sid), plReqId_(data));   // V59: chống trùng bền
        var scRem = String(scQ.getRange(scRow, scM.rem).getValue() || '');
        var scNote = 'CLOSED ' + Utilities.formatDate(new Date(), 'GMT+7', 'yyyy-MM-dd') + (data.result ? (' [' + data.result + ']') : '') + (data.actionText ? (': ' + data.actionText) : '');   // V52: + Result
        scQ.getRange(scRow, scM.rem).setValue(scRem ? (scRem + '\n' + scNote) : scNote).setWrap(true);
        if (scUrls.length) {
          plAppendLinks_(scQ.getRange(scRow, scM.link), scUrls);   // V58: giữ link ẩn/công thức cũ
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
      if (pcTarget < 0) return reply({ ok: false, final: true, code: 'NOT_FOUND', error: 'Không tìm thấy mục ' + pcId });   // V58
      if (String(pcQ.getRange(pcTarget, 4).getValue()).toLowerCase() === 'close')
        return reply({ ok: false, final: true, code: 'ALREADY_CLOSED', error: 'Mục này đã đóng trước đó' });   // V58
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
      plReqMark_(pcQ.getRange(pcTarget, 15), plReqId_(data));   // V59: chống trùng bền
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
              plReqMark_(pcSp.getRange(pcSpRow, pcSm.sid), plReqId_(data));   // V59
              var pcRem = String(pcSp.getRange(pcSpRow, pcSm.rem).getValue() || '');
              var pcNote = 'CLOSED ' + Utilities.formatDate(new Date(), 'GMT+7', 'yyyy-MM-dd') + (data.actionText ? (': ' + String(data.actionText)) : '');
              pcSp.getRange(pcSpRow, pcSm.rem).setValue(pcRem ? (pcRem + '\n' + pcNote) : pcNote).setWrap(true);
              if (pcCloseUrls.length) {
                plAppendLinks_(pcSp.getRange(pcSpRow, pcSm.link), pcCloseUrls);   // V58: giữ link ẩn/công thức cũ
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
      var spId = 'SP' + Utilities.formatDate(new Date(), 'GMT+7', 'yyMMddHHmmss') + '_' + Utilities.getUuid().replace(/-/g, '').slice(0, 8);   // V58: khó trùng (giữ SP + 12 số ngày giờ)
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
      // V59: chỉ ghi khi tab CÓ tiêu đề (tránh ghi chữ vào cột Photo after / cột khác)
      if (plHdrOk_(spQ, spM.sev, /^(mức độ|muc do|severity)/)) spArr[spM.sev - 1] = data.severity || '';
      if (plHdrOk_(spQ, spM.asg, /^(người khắc phục|nguoi khac phuc|assignee)/)) spArr[spM.asg - 1] = data.assignee || '';
      if (plHdrOk_(spQ, spM.due, /^(hạn xử lý|han xu ly|due)/)) spArr[spM.due - 1] = data.due || '';
      if (spM.insp) spArr[spM.insp - 1] = senderName(data);   // V46: ghi NGƯỜI LẬP (đa QC)
      spQ.appendRow(spArr);
      var spRow = spQ.getLastRow();
      plReqMark_(spQ.getRange(spRow, spM.sid), plReqId_(data));   // V59: chống trùng bền
      spQ.getRange(spRow, 1, 1, spM.lastCol).setBorder(true, true, true, true, true, true).setVerticalAlignment('middle');
      spQ.getRange(spRow, spM.desc).setWrap(true);
      spQ.getRange(spRow, 1, 1, 3).setHorizontalAlignment('center');
      spQ.getRange(spRow, spM.st).setHorizontalAlignment('center');
      if (spIds.length) { spQ.getRange(spRow, spM.photo).setFormula(punchImage(spIds[0])); spQ.setRowHeight(spRow, 110); }
      return reply({ ok: true, no: spNo, id: spId, photoCount: spUrls.length, photos: spUrls });
    }

    // ============ V51 — LINK MỞ GOOGLE SHEET đúng tab dự án ============
    if (data.action === 'sheetInfo') {
      var siSs = SpreadsheetApp.openById(SHEET_ID);
      var siQ = siSs.getSheetByName(SIMPLE_TAB);
      return reply({ ok: true, url: siSs.getUrl() + (siQ ? ('#gid=' + siQ.getSheetId()) : ''), tab: siQ ? SIMPLE_TAB : '' });
    }

    // ============ V51 — THÊM ẢNH TRƯỚC / SAU vào 1 mục có sẵn (app web) ============
    if (data.action === 'punchAddPhoto') {
      var apSid = String(data.sid || '').trim();
      var apKind = (String(data.kind || 'before') === 'after') ? 'after' : 'before';
      var apImgs = data.images || [];
      if (!apSid) return reply({ ok: false, error: 'Thiếu SID của mục' });
      if (!apImgs.length) return reply({ ok: false, error: 'Chưa có ảnh' });
      var apSs = SpreadsheetApp.openById(SHEET_ID);
      var apQ = apSs.getSheetByName(SIMPLE_TAB);
      if (!apQ || apQ.getLastRow() < 4) return reply({ ok: false, error: 'Không thấy tab dự án' });
      var apM = simpleColMap(apQ);
      var apIdsCol = apQ.getRange(4, apM.sid, apQ.getLastRow() - 3, 1).getValues(), apRow = -1;
      for (var api = 0; api < apIdsCol.length; api++) if (String(apIdsCol[api][0]) === apSid) { apRow = 4 + api; break; }
      if (apRow < 0) return reply({ ok: false, error: 'Không tìm thấy mục ' + apSid + ' (bấm 🔄 tải lại báo cáo)' });
      var apMonth = Utilities.formatDate(new Date(), 'GMT+7', 'yyyy-MM');
      var apDay = Utilities.formatDate(new Date(), 'GMT+7', 'yyyy-MM-dd');
      var apFolder = getReportFolder('Don gian', apMonth, apDay);
      var apStamp = Utilities.formatDate(new Date(), 'GMT+7', 'yyyyMMdd_HHmmss');
      var apNew = [], apUrls = [];
      apImgs.forEach(function (b64, i) {
        var f = apFolder.createFile(Utilities.newBlob(Utilities.base64Decode(b64), 'image/jpeg',
          (apKind === 'after' ? 'PLSclose_' : 'PLSadd_') + apStamp + '_' + (i + 1) + '.jpg'));
        apNew.push(f.getId()); apUrls.push(f.getUrl());
      });
      var apLkR = apQ.getRange(apRow, apM.link);
      if (apKind === 'before') {
        // V52: dòng có ảnh SAU cũ (công thức Photo after) mà chưa có ghi chú AFTER → ghi chú trước, để ảnh TRƯỚC mới không bị xếp nhầm sang SAU
        // V58: ghi chú TRƯỚC khi nối link; dòng đóng thời V49 (nhiều ảnh đóng) → mọi ảnh từ ảnh SAU đầu tiên trở đi là ảnh SAU
        if (apM.after) { try { var apOld = apQ.getRange(apRow, apM.after); var apOid = (String(apOld.getFormula() || '').match(/[-\w]{25,}/) || [''])[0];
          if (apOid && !/AFTER:/.test(String(apOld.getNote() || ''))) {
            var apOldIds = plCellIds_(apLkR), apK = apOldIds.indexOf(apOid);
            plAddAfterNote_(apOld, apK >= 0 ? apOldIds.slice(apK) : [apOid]);
          } } catch (eAo) {} }
      }
      plAppendLinks_(apLkR, apUrls);   // V58: giữ link ẩn/công thức cũ
      plReqMark_(apQ.getRange(apRow, apM.sid), plReqId_(data));   // V59: chống trùng bền
      if (apKind === 'before') {
        var apPh = apQ.getRange(apRow, apM.photo);
        if (!apPh.getFormula() && !String(apPh.getDisplayValue() || '').trim()) { apPh.setFormula(punchImage(apNew[0])); if (apQ.getRowHeight(apRow) < 90) apQ.setRowHeight(apRow, 110); }
      } else if (apM.after) {
        var apAf = apQ.getRange(apRow, apM.after);
        if (!apAf.getFormula() && !String(apAf.getDisplayValue() || '').trim()) apAf.setFormula(punchImage(apNew[0]));
        try { plAddAfterNote_(apAf, apNew); } catch (eAp) {}
      }
      return reply({ ok: true, sid: apSid, kind: apKind, added: apNew.length, row: apRow });
    }

    // ============ PUNCH LIST — XEM BÁO CÁO NHÁP (đọc tab đơn giản về app) ============
    if (data.action === 'listSimple') {
      var lsSs2 = SpreadsheetApp.openById(SHEET_ID);
      var lsQ2 = lsSs2.getSheetByName(SIMPLE_TAB);
      if (!lsQ2 || lsQ2.getLastRow() < 4) return reply({ ok: true, items: [] });
      var lsM2 = simpleColMap(lsQ2);   // V44: cột theo tên tiêu đề
      // V59: Mức độ / Người khắc phục / Hạn — CHỈ đọc khi có tiêu đề thật (simpleColMap trả cột mặc định 11/12/13 khi thiếu)
      var lsHd = []; try { lsHd = lsQ2.getRange(3, 1, 1, lsM2.lastCol).getDisplayValues()[0]; } catch (eH) {}
      function lsHok(col, re) { return !!col && re.test(String(lsHd[col - 1] || '').trim().toLowerCase()); }
      var lsSevOk = lsHok(lsM2.sev, /^(mức độ|muc do|severity)/), lsAsgOk = lsHok(lsM2.asg, /^(người khắc phục|nguoi khac phuc|assignee)/),
          lsDueOk = lsHok(lsM2.due, /^(hạn xử lý|han xu ly|due)/);
      var lsV2 = lsQ2.getRange(4, 1, lsQ2.getLastRow() - 3, lsM2.lastCol).getValues();
      // V49: công thức cột "Photo after" → ID ảnh khắc phục đầu tiên (app tách ảnh TRƯỚC / SAU)
      // V58: không đọc riêng công thức cột Photo after nữa — lấy từ lsF2 (đã gồm mọi cột)
      var lsAn2 = [];
      try { if (lsM2.after) lsAn2 = lsQ2.getRange(4, lsM2.after, lsV2.length, 1).getNotes(); } catch (eAn) { lsAn2 = []; }   // V50
      // V43: cấp ID bền cho dòng chưa có (app luôn thao tác theo ID → khớp tuyệt đối với sheet)
      // V58: cấp dưới khoá ngắn + đọc lại, chỉ điền ô còn trống (plStampSids_)
      plStampSids_(lsQ2, lsM2, lsV2);
      // V56: thêm ảnh bị sót — (1) ô link là CHỮ có gắn link (tên file "PLS_….jpg"), (2) công thức
      //      =IMAGE(...)/=HYPERLINK(...) ở cột Photo / Photo after. Ảnh ở cột Photo after = ảnh SAU.
      var lsF2 = [], lsRt2 = [];
      try { lsF2 = lsQ2.getRange(4, 1, lsV2.length, lsM2.lastCol).getFormulas(); } catch (eF2) { lsF2 = []; }
      try { lsRt2 = lsQ2.getRange(4, lsM2.link, lsV2.length, 1).getRichTextValues(); } catch (eRt) { lsRt2 = []; }
      var lsOut2 = [], lsPendN = 0;
      // V59: ảnh NỔI theo hàng (đếm, để app biết dòng có ảnh dán chưa lên Drive)
      var lsFlt = {}; try { var lsImgs = lsQ2.getImages(); if (lsImgs.length <= 300) lsImgs.forEach(function (g) { var ac = g.getAnchorCell(), r = ac.getRow(); if (r < 4) return; var o = lsFlt[r] = lsFlt[r] || { b: 0, a: 0 }; (ac.getColumn() === lsM2.after ? o.a++ : o.b++); }); } catch (eG) {}
      for (var li2 = 0; li2 < lsV2.length; li2++) {
        var lr = lsV2[li2];
        var lIds = plTextIds_(plTxt_(lr[lsM2.link - 1]));   // V58: chỉ ID theo URL / dòng ID trần (tên file .jpg không phải ID)
        var lAf = (lsM2.after && lsF2[li2] && String(lsF2[li2][lsM2.after - 1] || '').match(/[-\w]{25,}/)) || [''];
        var lAnM = lsAn2[li2] ? String(lsAn2[li2][0] || '').match(/AFTER:([-\w,]+)/) : null;
        var lAfterIds = lAnM ? lAnM[1].split(',').filter(String) : [];   // V50
        plLinkIds_(lsRt2[li2] && lsRt2[li2][0]).forEach(function (x) { if (lIds.indexOf(x) < 0) lIds.push(x); });
        // V58: dòng đóng thời V49 (không có ghi chú AFTER): mọi ảnh trong ô link từ ảnh SAU đầu tiên trở đi là ảnh SAU
        var lK = (!lAnM && lAf[0]) ? lIds.indexOf(lAf[0]) : -1;
        if (lK >= 0) lIds.slice(lK).forEach(function (x) { if (lAfterIds.indexOf(x) < 0) lAfterIds.push(x); });
        var lFr = lsF2[li2] || [];
        for (var lfc = 0; lfc < lFr.length; lfc++) {
          plDriveIds_(lFr[lfc]).forEach(function (x) {
            if (lIds.indexOf(x) < 0) lIds.push(x);
            if (lfc + 1 === lsM2.after && lAfterIds.indexOf(x) < 0) lAfterIds.push(x);
          });
        }
        lsOut2.push({ no: plTxt_(lr[lsM2.no - 1]), section: plTxt_(lr[lsM2.sec - 1]), part: plTxt_(lr[lsM2.part - 1]), description: plTxt_(lr[lsM2.desc - 1]),
          remarks: plTxt_(lr[lsM2.rem - 1]), status: plTxt_(lr[lsM2.st - 1]) || 'open', sid: plTxt_(lr[lsM2.sid - 1]),
          inspector: (lsM2.insp ? String(lr[lsM2.insp - 1] || '') : ''), photoIds: lIds, afterFirst: lAf[0], afterIds: lAfterIds });
        // V59: thêm Mức độ / Người khắc phục / Hạn + số ảnh dán (trong ô) / ảnh nổi của dòng
        var lIt = lsOut2[lsOut2.length - 1];
        var lDue = lsDueOk ? lr[lsM2.due - 1] : '';
        lIt.severity = lsSevOk ? plTxt_(lr[lsM2.sev - 1]).trim() : '';
        lIt.assignee = lsAsgOk ? plTxt_(lr[lsM2.asg - 1]).trim() : '';
        lIt.due = (lDue instanceof Date) ? Utilities.formatDate(lDue, 'GMT+7', 'yyyy-MM-dd') : plTxt_(lDue).trim();
        var lPb = 0, lPa = 0;
        for (var lpc = 0; lpc < lr.length; lpc++) {
          var lpv = lr[lpc];
          if (lpv && typeof lpv === 'object' && !(lpv instanceof Date) && typeof lpv.getContentUrl === 'function') { if (lpc + 1 === lsM2.after) lPa++; else lPb++; }
        }
        var lFo = lsFlt[li2 + 4] || { b: 0, a: 0 };
        if (lPb) lIt.pasteB = lPb;
        if (lPa) lIt.pasteA = lPa;
        if (lFo.b) lIt.floatB = lFo.b;
        if (lFo.a) lIt.floatA = lFo.a;
        var lBefN = lIds.filter(function (x) { return lAfterIds.indexOf(x) < 0; }).length;
        if (((lPb + lFo.b) > 0 && !lBefN) || ((lPa + lFo.a) > 0 && !lAfterIds.length)) { lIt.pendingImg = true; lsPendN++; }
      }
      return reply({ ok: true, items: lsOut2, pendingImgN: lsPendN });
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
        // No. = 1..n liên tục — V58: đúng cột "No." theo tiêu đề, ghi 1 lần (trước: cột A, từng ô dưới khoá)
        var pdN = pdLast - 3;
        if (pdN > 0) { var pdSeq = []; for (var pdk = 1; pdk <= pdN; pdk++) pdSeq.push([pdk]); pdQ.getRange(4, pdM.no, pdN, 1).setValues(pdSeq); }
      }
      if (/^PL/.test(pdId)) {
        var pvQ = pdSs.getSheetByName(PUNCH_TAB);
        if (pvQ && pvQ.getLastRow() >= 8) {
          var pvIds = pvQ.getRange(8, 15, pvQ.getLastRow() - 7, 1).getValues();
          for (var pvi = pvIds.length - 1; pvi >= 0; pvi--) {
            if (String(pvIds[pvi][0]) === pdId) { pvQ.deleteRow(8 + pvi); pdDel++; break; }
          }
          var pvLast = pvQ.getLastRow();
          // ITEM # = #1..#n — V58: ghi 1 lần
          var pvN = pvLast - 7;
          if (pvN > 0) { var pvSeq = []; for (var pvk = 1; pvk <= pvN; pvk++) pvSeq.push(['#' + pvk]); pvQ.getRange(8, 7, pvN, 1).setValues(pvSeq); }
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
      var fxName = String(data.sheet || 'Punch list đơn giản');
      // V58: chỉ nhận tab "Punch list đơn giản…" (so khớp bỏ dấu cách/gạch/hoa thường như plTuDongLayAnh) —
      //      không ghi link vào tab khác. Trả 'no sheet' để vòng lặp của app dừng như cũ.
      var fxNorm = function (s) { return String(s).toLowerCase().replace(/[\s\-–_]+/g, ''); };
      if (fxNorm(fxName).indexOf(fxNorm('Punch list đơn giản')) !== 0) return reply({ ok: false, error: 'no sheet' });
      var fxSh = fxSs.getSheetByName(fxName);
      if (!fxSh) return reply({ ok: false, error: 'no sheet' });
      // V55: KHÔNG giữ khoá ghi chung suốt mấy phút (trước đây QC khác bị báo "máy chủ đang ghi").
      //      Chỉ chặn 2 lần lấy ảnh cùng 1 tab chạy chồng nhau; khoá chung chỉ giữ lúc ghi từng dòng.
      var fxCache = CacheService.getScriptCache(), fxKey = 'pl_fiximg_' + fxSh.getSheetId();
      if (fxCache.get(fxKey)) return reply({ ok: false, busy: true, running: true,
        error: 'Lần bấm trước vẫn đang lấy ảnh cho tab này — đợi 2–3 phút rồi bấm lại (dòng đã xong sẽ bỏ qua).' });
      fxCache.put(fxKey, '1', 330);
      try {
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
      var fxOg = {}, fxX;   // V54: fxX = ảnh nổi lấy từ bản xuất Excel (tạo khi cần)
      try {
        var fxIms = fxSh.getImages();
        for (var fo = 0; fo < fxIms.length; fo++) {
          var fa = fxIms[fo].getAnchorCell().getRow();
          (fxOg[fa] = fxOg[fa] || []).push(fxIms[fo]);
        }
      } catch (eFo) {}
      var fxSaved = 0, fxFail = [], fxRows = 0, fxSkip = 0, fxStamp = Utilities.formatDate(new Date(), 'GMT+7', 'yyyyMMdd_HHmmss');
      var fxT0 = new Date().getTime(), fxPartial = false;   // V52: dừng an toàn trước giới hạn 6 phút của Google
      var fxIsImg = function (v) { return v && typeof v === 'object' && !(v instanceof Date) && typeof v.getContentUrl === 'function'; };
      for (var fr = 0; fr < fxN; fr++) {
        if (new Date().getTime() - fxT0 > 140000) { fxPartial = true; break; }   // V55: trả lời trước 3 phút, app tự bấm tiếp
        var fxCur = plTxt_(fxVals[fr][fxLink - 1]).trim();   // V53: ảnh dán vào chính cột link ≠ "đã có link"
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
        // V58: dòng cần dựng bản Excel tạm (ảnh nổi) mà đã quá 60 giây → dừng, đợt sau làm tiếp (không bắt đầu việc chậm ở ~139 giây)
        if (fxX === undefined && fxG.length && new Date().getTime() - fxT0 > 60000) {
          var fxNeedX = fxG.some(function (g) { try { return !fxAfterOnly || g.getAnchorCell().getColumn() === fxAfterCol; } catch (eNx) { return !fxAfterOnly; } });
          if (fxNeedX) { fxPartial = true; break; }
        }
        var fxFail0 = fxFail.length;   // V58: số lỗi trước dòng này
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
          var fgAfter = false, fgCol = 0;
          try { fgCol = fxG[fg].getAnchorCell().getColumn(); fgAfter = (fgCol === fxAfterCol); } catch (eAc2) {}
          if (fxAfterOnly && !fgAfter) continue;
          try {
            // V54: ảnh NỔI không có getBlob() → lấy ảnh từ bản xuất Excel của tab (theo hàng/cột)
            var fb2 = null;
            try { fb2 = fxG[fg].getBlob(); } catch (eGb) { fb2 = null; }
            if (!fb2) {
              if (fxX === undefined) fxX = plOgFromXlsx_(fxSh);
              fb2 = plOgPick_(fxX, fr + 4, fgCol);
            }
            if (!fb2) throw new Error('no blob');
            var ff2 = fxFol.createFile(fb2.setName('PLMIG_' + fxStamp + '_r' + (fr + 4) + 'og' + (fg + 1) + '.png'));
            fxIds.push(ff2.getId()); if (fgAfter) fxAfterIds.push(ff2.getId()); fxSaved++;
          } catch (eFg) { fxFail.push('dòng ' + (fr + 4) + ' ảnh nổi ' + (fg + 1)); }
        }
        // V58: dòng có ảnh lấy lỗi → KHÔNG ghi link (nếu ghi, lần sau dòng bị coi là "đã có link" và bỏ qua mãi);
        //      file đã lưu Drive của dòng này để nguyên, lần sau lấy lại cả dòng.
        // V59: …và bỏ vào thùng rác các file PLMIG_ vừa tạo cho dòng này (trước đây thành file mồ côi mới mỗi giờ)
        if (fxFail.length > fxFail0) { plTrashIds_(fxIds); fxSaved -= fxIds.length; continue; }
        if (fxIds.length) {
          var fxUrls = fxIds.map(function (id) { return 'https://drive.google.com/file/d/' + id + '/view'; });
          // V55: ghi từng dòng dưới khoá chung + kiểm tra dòng chưa bị dời (ai đó xoá dòng trong lúc chạy)
          var fxLk = LockService.getScriptLock();
          if (!fxLk.tryLock(20000)) { fxFail.push('dòng ' + (fr + 4) + ' (máy chủ bận — bấm lại)'); plTrashIds_(fxIds); fxSaved -= fxIds.length; continue; }   // V59: không để file mồ côi
          try {
            var fxNowDesc = plTxt_(fxSh.getRange(fr + 4, fxMap.desc).getValue()).trim();
            var fxNowLink = plTxt_(fxSh.getRange(fr + 4, fxLink).getValue()).trim();
            var fxNowSid = plTxt_(fxSh.getRange(fr + 4, fxMap.sid).getValue()).trim();   // V58: so thêm SID (mô tả hay trùng)
            if (fxNowDesc !== plTxt_(fxVals[fr][fxMap.desc - 1]).trim() || fxNowLink !== fxCur ||
                fxNowSid !== plTxt_(fxVals[fr][fxMap.sid - 1]).trim()) {
              fxFail.push('dòng ' + (fr + 4) + ' (dòng vừa bị đổi — bấm lại)');
              plTrashIds_(fxIds); fxSaved -= fxIds.length;   // V59: không để file mồ côi
              continue;
            }
            plAppendLinks_(fxSh.getRange(fr + 4, fxLink), fxUrls);   // V58: giữ link ẩn/công thức cũ
            if (fxAfterIds.length && fxAfterCol) { try { plAddAfterNote_(fxSh.getRange(fr + 4, fxAfterCol), fxAfterIds); } catch (eAn3) {} }   // V50
            SpreadsheetApp.flush();
            fxRows++;
          } finally { try { fxLk.releaseLock(); } catch (eRl) {} }
        }
      }
      if (!fxPartial) { try { plOgXlsxDone_(fxSh); } catch (eXd) {} }   // V57
      // V59: lưu kết quả lần lấy ảnh gần nhất của tab này (plDiag hiển thị cho Sếp)
      try { PropertiesService.getScriptProperties().setProperty('PL_FIX_LAST_' + fxSh.getSheetId(), JSON.stringify({ t: Utilities.formatDate(new Date(), 'GMT+7', 'yyyy-MM-dd HH:mm'),
        by: data._auto ? 'auto' : 'app', rows: fxRows, saved: fxSaved, skipped: fxSkip, failedN: fxFail.length,
        failed: fxFail.slice(0, 10).map(function (x) { return String(x).slice(0, 120); }), partial: fxPartial })); } catch (eP) {}
      return reply({ ok: true, rows: fxRows, saved: fxSaved, skipped: fxSkip, failed: fxFail, partial: fxPartial });
      } finally { try { fxCache.remove(fxKey); } catch (eCr) {} }
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
