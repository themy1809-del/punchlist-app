/**
 * XuatPDF_Sheet.gs — Menu "📄 XUẤT PDF" ngay trên Google Sheet Punch List
 * PHIEN_BAN_PDF: '15/10 menu-pdf'
 *
 * File ĐỘC LẬP: mọi hàm đều có tiền tố plPdf_ → không trùng tên với Code.gs.
 * KHÔNG dùng onOpen() đơn giản (tránh đè onOpen có sẵn) — dùng trigger cài đặt.
 *
 * CÁCH CÀI (1 lần):
 *   1. Mở Google Sheet Punch List → Tiện ích mở rộng → Apps Script.
 *   2. Bấm "+" → Tập lệnh → đặt tên XuatPDF_Sheet → dán toàn bộ file này → Lưu.
 *   3. Chọn hàm plPdf_caiDat → Chạy → cấp quyền.
 *   4. Tải lại Sheet → thấy menu "📄 XUẤT PDF".
 *   (KHÔNG cần Deploy lại web app /exec — file này không đụng doPost.)
 */

var PLPDF_PHIEN_BAN = '15/10 menu-pdf';
var PLPDF_THU_MUC = 'PUNCH LIST - PDF xuat tu Sheet';   // thư mục Drive chứa PDF

// ---------- Cài đặt / menu ----------
function plPdf_caiDat() {
  var ss = SpreadsheetApp.getActive();
  if (!ss) throw new Error('Hãy mở Apps Script TỪ trong Google Sheet (Tiện ích mở rộng → Apps Script).');
  var daCo = ScriptApp.getProjectTriggers().some(function (t) {
    return t.getHandlerFunction() === 'plPdf_onOpen';
  });
  if (!daCo) ScriptApp.newTrigger('plPdf_onOpen').forSpreadsheet(ss).onOpen().create();
  plPdf_onOpen();
  return 'Đã cài menu XUẤT PDF (' + PLPDF_PHIEN_BAN + ')';
}

function plPdf_onOpen() {
  SpreadsheetApp.getUi().createMenu('📄 XUẤT PDF')
    .addItem('Tab đang mở — A4 NGANG', 'plPdf_tabNgang')
    .addItem('Tab đang mở — A4 DỌC', 'plPdf_tabDoc')
    .addItem('Vùng đang chọn — A4 NGANG', 'plPdf_vungChon')
    .addSeparator()
    .addItem('Tất cả tab đang hiện — 1 file', 'plPdf_tatCa')
    .addItem('📂 Mở thư mục PDF', 'plPdf_moThuMuc')
    .addToUi();
}

// ---------- Các lệnh menu ----------
function plPdf_tabNgang() { plPdf_xuat_({ sheet: SpreadsheetApp.getActiveSheet(), ngang: true }); }
function plPdf_tabDoc()   { plPdf_xuat_({ sheet: SpreadsheetApp.getActiveSheet(), ngang: false }); }
function plPdf_tatCa()    { plPdf_xuat_({ sheet: null, ngang: true }); }

function plPdf_vungChon() {
  var sh = SpreadsheetApp.getActiveSheet();
  var rg = sh.getActiveRange();
  if (!rg || (rg.getNumRows() < 2 && rg.getNumColumns() < 2)) {
    SpreadsheetApp.getUi().alert('Hãy bôi đen vùng cần in (ít nhất 2 ô) rồi bấm lại.');
    return;
  }
  plPdf_xuat_({ sheet: sh, ngang: true, range: rg.getA1Notation() });
}

function plPdf_moThuMuc() {
  plPdf_hienLink_('📂 Thư mục PDF', plPdf_thuMuc_().getUrl(), '');
}

// ---------- Lõi xuất PDF ----------
function plPdf_xuat_(o) {
  var ss = SpreadsheetApp.getActive();
  SpreadsheetApp.flush();   // ghi hết công thức/định dạng trước khi xuất

  var p = {
    format: 'pdf', size: 'A4', portrait: o.ngang ? 'false' : 'true',
    fitw: 'true',              // vừa khổ ngang trang
    gridlines: 'false', printtitle: 'false', sheetnames: 'false',
    pagenum: 'CENTER', fzr: 'true',   // lặp lại hàng tiêu đề (hàng cố định) mỗi trang
    top_margin: '0.4', bottom_margin: '0.4', left_margin: '0.3', right_margin: '0.3',
    horizontal_alignment: 'CENTER'
  };
  if (o.sheet) p.gid = String(o.sheet.getSheetId());
  if (o.range) p.range = o.range;

  var qs = Object.keys(p).map(function (k) { return k + '=' + encodeURIComponent(p[k]); }).join('&');
  var url = 'https://docs.google.com/spreadsheets/d/' + ss.getId() + '/export?' + qs;

  var blob = plPdf_taiVe_(url);
  var tz = ss.getSpreadsheetTimeZone() || 'Asia/Ho_Chi_Minh';
  var ten = (o.sheet ? o.sheet.getName() : ss.getName())
          + (o.range ? ' (' + o.range + ')' : '')
          + ' - ' + Utilities.formatDate(new Date(), tz, 'yyyy-MM-dd HHmm') + '.pdf';
  blob.setName(ten.replace(/[\\\/:*?"<>|]/g, '-'));

  var f = plPdf_thuMuc_().createFile(blob);
  plPdf_hienLink_('✅ Đã xuất PDF', f.getUrl(), f.getName());
}

// Google giới hạn tốc độ export → thử lại khi gặp 429/5xx
function plPdf_taiVe_(url) {
  var loi = '';
  for (var i = 0; i < 4; i++) {
    var res = UrlFetchApp.fetch(url, {
      headers: { Authorization: 'Bearer ' + ScriptApp.getOAuthToken() },
      muteHttpExceptions: true
    });
    var code = res.getResponseCode();
    if (code === 200) {
      var b = res.getBlob();
      if (b.getContentType() === 'application/pdf') return b;
      loi = 'Google trả về không phải PDF (' + b.getContentType() + ')';
    } else {
      loi = 'HTTP ' + code;
      if (code !== 429 && code < 500) break;
    }
    Utilities.sleep(3000 * (i + 1));
  }
  throw new Error('Không xuất được PDF: ' + loi);
}

function plPdf_thuMuc_() {
  var it = DriveApp.getFoldersByName(PLPDF_THU_MUC);
  return it.hasNext() ? it.next() : DriveApp.createFolder(PLPDF_THU_MUC);
}

function plPdf_esc_(s) {
  return String(s || '').replace(/[&<>"']/g, function (c) {
    return { '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c];
  });
}

function plPdf_hienLink_(tieuDe, url, ten) {
  var u = plPdf_esc_(url);
  var html = '<div style="font-family:Arial,sans-serif;font-size:14px">'
    + (ten ? '<p style="word-break:break-all">' + plPdf_esc_(ten) + '</p>' : '')
    + '<p><a href="' + u + '" target="_blank" '
    + 'style="display:inline-block;background:#1a73e8;color:#fff;padding:10px 16px;'
    + 'border-radius:8px;text-decoration:none;font-weight:bold">🔗 Mở</a></p>'
    + '<p><input value="' + u + '" style="width:100%" onclick="this.select()" readonly></p></div>';
  SpreadsheetApp.getUi().showModelessDialog(
    HtmlService.createHtmlOutput(html).setWidth(420).setHeight(190), tieuDe);
}
