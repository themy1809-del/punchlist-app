/**
 * XuatPDF_Sheet.gs — Menu "📄 XUẤT PDF" ngay trên Google Sheet Punch List
 * PHIEN_BAN_PDF: '15/10 pdf-section'
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

var PLPDF_PHIEN_BAN = '15/10 pdf-section';
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
    .addItem('⭐ PUNCH LIST theo SECTION (giống app)', 'plPdf_theoSection')
    .addItem('⭐ PUNCH LIST theo SECTION — chỉ mục ĐANG MỞ', 'plPdf_theoSectionMo')
    .addItem('🧹 Làm gọn cột link ảnh (tab đang mở)', 'plPdf_gonLink')
    .addSeparator()
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


// =====================================================================
// PUNCH LIST THEO SECTION — mẫu gọn, ảnh TRƯỚC / SAU, giống nút xuất trên app
// Chạy trên tab "Punch list đơn giản - <dự án>" đang mở. Chỉ ĐỌC tab gốc;
// dựng sheet tạm → xuất PDF → XOÁ sheet tạm. Không sửa dữ liệu.
// =====================================================================
var PLPDF_MAX_ANH = 3;   // tối đa ảnh mỗi bên (trước / sau) trên 1 dòng

function plPdf_theoSection()   { plPdf_section_(false); }
function plPdf_theoSectionMo() { plPdf_section_(true); }

// Dò dòng tiêu đề + cột theo TÊN (tab BISON bị lệch cột → dò cả theo nội dung)
function plPdf_docTab_(sh) {
  var lastR = sh.getLastRow(), lastC = sh.getLastColumn();
  if (lastR < 2) throw new Error('Tab trống.');
  var top = sh.getRange(1, 1, Math.min(8, lastR), lastC).getDisplayValues();
  var hr = -1;
  for (var r = 0; r < top.length && hr < 0; r++) {
    var low = top[r].map(function (x) { return String(x).trim().toLowerCase(); });
    if (low.indexOf('no.') >= 0 && low.indexOf('description') >= 0) hr = r;
  }
  if (hr < 0) throw new Error('Không thấy dòng tiêu đề (No. / Description). Hãy mở tab "Punch list đơn giản - ..." rồi chạy lại.');
  var h = top[hr].map(function (x) { return String(x).trim().toLowerCase(); });
  var n = lastR - (hr + 1);
  if (n < 1) throw new Error('Tab chưa có mục nào.');
  var vals = sh.getRange(hr + 2, 1, n, lastC).getDisplayValues();
  var fmls = sh.getRange(hr + 2, 1, n, lastC).getFormulas();
  function col(test) { for (var c = 0; c < h.length; c++) if (test(h[c])) return c; return -1; }
  var C = {
    no: col(function (x) { return x === 'no.' || x === 'no'; }),
    sec: col(function (x) { return x === 'section' || x === 'sec'; }),
    part: col(function (x) { return x === 'part'; }),
    desc: col(function (x) { return x === 'description'; }),
    rem: col(function (x) { return x === 'remarks' || x === 'ghi chú'; }),
    st: col(function (x) { return x === 'status' || x === 'tình trạng'; }),
    qc: col(function (x) { return x === 'người lập' || x === 'inspected by'; }),
    link: col(function (x) { return x.indexOf('photo link') >= 0 || x.indexOf('link ảnh') >= 0; })
  };
  // cột link: nếu tiêu đề sai (vd BISON ghi "SID") → chọn cột có nhiều link Drive nhất
  if (C.link < 0) {
    var best = -1, bestN = 0;
    for (var c = 0; c < lastC; c++) {
      var k = 0;
      for (var i = 0; i < n; i++) if (vals[i][c].indexOf('drive.google.com') >= 0) k++;
      if (k > bestN) { bestN = k; best = c; }
    }
    C.link = best;
  }
  if (C.desc < 0 || C.sec < 0) throw new Error('Thiếu cột Section hoặc Description.');
  return { sh: sh, hr: hr, n: n, vals: vals, fmls: fmls, C: C };
}

function plPdf_ids_(s) { return String(s || '').match(/[-\w]{25,}/g) || []; }

// Tách ảnh trước / sau: ảnh đóng lỗi có tên PLSclose_ / PLclose_ (backend đặt khi đóng)
function plPdf_tachAnh_(ids, closed, cache) {
  var b = [], a = [];
  ids.forEach(function (id) {
    var nm = cache[id];
    if (nm === undefined) {
      try { nm = DriveApp.getFileById(id).getName(); } catch (e) { nm = ''; }
      cache[id] = nm;
    }
    (/close/i.test(nm) ? a : b).push(id);
  });
  return { b: b, a: a };
}

function plPdf_section_(chiMo) {
  var ui = SpreadsheetApp.getUi();
  var ss = SpreadsheetApp.getActive();
  var T = plPdf_docTab_(ss.getActiveSheet());
  var C = T.C;

  // danh sách section có trong tab
  var secs = {};
  T.vals.forEach(function (r) { var s = String(r[C.sec]).trim(); if (s && r[C.desc]) secs[s] = 1; });
  var secList = Object.keys(secs).sort(function (x, y) { return (parseFloat(x) || 0) - (parseFloat(y) || 0) || (x < y ? -1 : 1); });
  var resp = ui.prompt('Xuất PUNCH LIST theo SECTION',
    'Section có trong tab: ' + secList.join(', ') + '\n\nGõ section cần xuất (nhiều section: 14,16 · để TRỐNG = tất cả):',
    ui.ButtonSet.OK_CANCEL);
  if (resp.getSelectedButton() !== ui.Button.OK) return;
  var want = String(resp.getResponseText() || '').split(/[,;\s]+/).map(function (x) { return x.trim(); }).filter(String);

  // gom mục
  var cache = {}, items = [];
  for (var i = 0; i < T.n; i++) {
    var r = T.vals[i], f = T.fmls[i];
    var desc = String(r[C.desc] || '').trim();
    if (!desc) continue;
    var sec = String(r[C.sec] || '').trim();
    if (want.length && want.indexOf(sec) < 0) continue;
    var st = C.st >= 0 ? String(r[C.st] || '').trim().toLowerCase() : '';
    var closed = /^clos/.test(st);
    if (chiMo && closed) continue;
    var ids = C.link >= 0 ? plPdf_ids_(r[C.link]) : [];
    if (!ids.length) {   // không có link → lấy ảnh từ công thức IMAGE() trên dòng
      f.forEach(function (x) { plPdf_ids_(x).forEach(function (id) { if (ids.indexOf(id) < 0) ids.push(id); }); });
    }
    var sp = plPdf_tachAnh_(ids, closed, cache);
    items.push({ no: C.no >= 0 ? r[C.no] : '', sec: sec, part: C.part >= 0 ? r[C.part] : '', desc: desc,
                 rem: C.rem >= 0 ? r[C.rem] : '', qc: C.qc >= 0 ? r[C.qc] : '', closed: closed, b: sp.b, a: sp.a });
  }
  if (!items.length) { ui.alert('Không có mục nào khớp.'); return; }
  items.sort(function (x, y) { return (parseFloat(x.sec) || 0) - (parseFloat(y.sec) || 0) || (Number(x.no) || 0) - (Number(y.no) || 0); });

  var tab = T.sh.getName();
  var duAn = (tab.split('-').slice(1).join('-') || tab).trim();
  var tz = ss.getSpreadsheetTimeZone() || 'Asia/Ho_Chi_Minh';
  var ngay = Utilities.formatDate(new Date(), tz, 'dd/MM/yyyy');
  var nhan = (want.length ? 'Section ' + want.join(', ') : 'Tất cả section') + (chiMo ? ' · chỉ mục đang mở' : '');
  var tmp = null;
  try {
    tmp = plPdf_dungBaoCao_(ss, items, duAn, nhan, ngay);
    SpreadsheetApp.flush();
    Utilities.sleep(4000);   // chờ IMAGE() tải ảnh trước khi xuất
    var url = 'https://docs.google.com/spreadsheets/d/' + ss.getId() + '/export?format=pdf&gid=' + tmp.getSheetId()
      + '&size=A4&portrait=false&fitw=true&gridlines=false&printtitle=false&sheetnames=false&pagenum=CENTER&fzr=true'
      + '&top_margin=0.3&bottom_margin=0.35&left_margin=0.25&right_margin=0.25&horizontal_alignment=CENTER';
    var blob = plPdf_taiVe_(url);
    var ten = 'Punch list ' + duAn + ' - ' + (want.length ? 'Section ' + want.join('_') : 'Tat ca') + (chiMo ? ' (dang mo)' : '')
      + ' - ' + Utilities.formatDate(new Date(), tz, 'yyyy-MM-dd HHmm') + '.pdf';
    blob.setName(ten.replace(/[\\\/:*?"<>|]/g, '-'));
    var file = plPdf_thuMuc_().createFile(blob);
    plPdf_hienLink_('✅ Đã xuất ' + items.length + ' mục', file.getUrl(), file.getName());
  } finally {
    if (tmp) try { ss.deleteSheet(tmp); } catch (e) {}
  }
}

// Dựng sheet tạm: No | Part | Mô tả lỗi / Ghi chú | Người lập | Tình trạng | Trước ×3 | Sau ×3
function plPdf_dungBaoCao_(ss, items, duAn, nhan, ngay) {
  var M = PLPDF_MAX_ANH, NC = 5 + 2 * M;
  var NAVY = '#1c4587', LINE = '#b7c4d4';
  var sh = ss.insertSheet('PDF_TAM_' + Date.now());
  if (sh.getMaxColumns() < NC) sh.insertColumnsAfter(sh.getMaxColumns(), NC - sh.getMaxColumns());
  var W = [34, 90, 270, 70, 70];
  for (var k = 0; k < 2 * M; k++) W.push(86);
  W.forEach(function (w, i) { sh.setColumnWidth(i + 1, w); });

  var nMo = items.filter(function (x) { return !x.closed; }).length;
  sh.getRange(1, 1, 1, NC).merge().setValue('PUNCH LIST REPORT — ' + duAn + '   |   ' + nhan)
    .setBackground(NAVY).setFontColor('#fff').setFontWeight('bold').setFontSize(13).setVerticalAlignment('middle');
  sh.setRowHeight(1, 30);
  sh.getRange(2, 1, 1, NC).merge()
    .setValue('Ngày: ' + ngay + '     Tổng: ' + items.length + ' mục     Đang mở: ' + nMo + '     Đã đóng: ' + (items.length - nMo)
      + '     Người xuất: ' + (Session.getActiveUser().getEmail() || ''))
    .setBackground('#e8eef7').setFontColor('#1b3a5c').setFontSize(9).setVerticalAlignment('middle');
  sh.setRowHeight(2, 20);
  var head = ['No.', 'Part', 'Mô tả lỗi / Ghi chú', 'Người lập', 'Tình trạng'];
  for (k = 0; k < M; k++) head.push(k === 0 ? 'Ảnh TRƯỚC (lỗi)' : '');
  for (k = 0; k < M; k++) head.push(k === 0 ? 'Ảnh SAU (khắc phục)' : '');
  sh.getRange(3, 1, 1, NC).setValues([head]).setBackground(NAVY).setFontColor('#fff').setFontWeight('bold')
    .setFontSize(9).setHorizontalAlignment('center').setVerticalAlignment('middle');
  sh.getRange(3, 6, 1, M).merge(); sh.getRange(3, 6 + M, 1, M).merge();
  sh.setRowHeight(3, 22);
  sh.setFrozenRows(3);   // fzr=true → lặp tiêu đề mỗi trang

  var row = 4, cur = null, rows = [], fm = [], meta = [];
  items.forEach(function (it) {
    if (it.sec !== cur) {
      cur = it.sec;
      var g = items.filter(function (x) { return x.sec === cur; });
      var go = g.filter(function (x) { return !x.closed; }).length;
      meta.push({ r: row, g: true });
      rows.push(['SECTION ' + (cur || '(trống)') + ' — ' + g.length + ' mục · ' + go + ' đang mở · ' + (g.length - go) + ' đã đóng'].concat(new Array(NC - 1).join('.').split('.')));
      row++;
    }
    var txt = String(it.desc) + (it.rem ? '\n' + String(it.rem) : '');
    if (it.b.length > M || it.a.length > M) txt += '\n📷 Trước: ' + it.b.length + ' ảnh · Sau: ' + it.a.length + ' ảnh (in tối đa ' + M + '/bên)';
    var line = [it.no, it.part, txt, it.qc, it.closed ? 'ĐÃ ĐÓNG' : 'ĐANG MỞ'];
    function anh(list) {
      for (var j = 0; j < M; j++) {
        if (j < list.length) {
          var more = (j === M - 1 && list.length > M) ? ' (+' + (list.length - M) + ')' : '';
          line.push('=IFERROR(IMAGE("https://drive.google.com/thumbnail?id=' + list[j] + '&sz=w300",1),"ảnh' + more + '")');
        } else line.push('');
      }
    }
    anh(it.b); anh(it.a);
    meta.push({ r: row, closed: it.closed, nb: it.b.length, na: it.a.length });
    rows.push(line); row++;
  });
  sh.getRange(4, 1, rows.length, NC).setValues(rows);   // giá trị bắt đầu bằng "=" tự thành công thức
  var last = 3 + rows.length;
  sh.getRange(4, 1, rows.length, NC).setFontSize(9).setVerticalAlignment('middle').setWrap(true)
    .setBorder(true, true, true, true, true, true, LINE, SpreadsheetApp.BorderStyle.SOLID);
  sh.getRange(4, 1, rows.length, 2).setHorizontalAlignment('center');
  sh.getRange(4, 4, rows.length, 2).setHorizontalAlignment('center');
  meta.forEach(function (m) {
    if (m.g) {
      sh.getRange(m.r, 1, 1, NC).merge().setBackground('#e8eef7').setFontColor(NAVY).setFontWeight('bold').setHorizontalAlignment('left');
      sh.setRowHeight(m.r, 20);
      return;
    }
    sh.setRowHeight(m.r, (m.nb || m.na) ? 66 : 34);
    sh.getRange(m.r, 5).setBackground(m.closed ? '#d9ead3' : '#fce5cd').setFontColor(m.closed ? '#0b5d1e' : '#8a4b00').setFontWeight('bold');
    sh.getRange(m.r, 6 + M, 1, M).setBackground('#f3faf4');
    if (!m.na) sh.getRange(m.r, 6 + M).setValue(m.closed ? '—' : 'chưa khắc phục').setFontColor('#999999').setFontStyle('italic');
  });

  // khối ký
  var sg = last + 2, p = Math.floor(NC / 3);
  var parts = [[1, p, 'NGƯỜI LẬP / PREPARED BY'], [1 + p, p, 'NGƯỜI KIỂM TRA / CHECKED BY'], [1 + 2 * p, NC - 2 * p, 'ĐẠI DIỆN KHÁCH HÀNG / CLIENT']];
  parts.forEach(function (q) {
    sh.getRange(sg, q[0], 1, q[1]).merge().setValue(q[2]);
    sh.getRange(sg + 1, q[0], 1, q[1]).merge();
    sh.getRange(sg + 2, q[0], 1, q[1]).merge().setValue('(Ký, ghi rõ họ tên)');
  });
  sh.getRange(sg, 1, 3, NC).setBorder(true, true, true, true, true, false, LINE, SpreadsheetApp.BorderStyle.SOLID)
    .setHorizontalAlignment('center').setFontSize(9);
  sh.getRange(sg, 1, 1, NC).setBackground('#e8eef7').setFontWeight('bold');
  sh.getRange(sg + 2, 1, 1, NC).setFontStyle('italic').setFontColor('#666666');
  sh.setRowHeight(sg + 1, 50);

  // bỏ hàng/cột thừa để PDF không in trang trắng
  if (sh.getMaxRows() > sg + 2) sh.deleteRows(sg + 3, sh.getMaxRows() - (sg + 2));
  if (sh.getMaxColumns() > NC) sh.deleteColumns(NC + 1, sh.getMaxColumns() - NC);
  return sh;
}

// Làm gọn cột link ảnh: CHỈ đổi định dạng (cắt chữ, chữ nhỏ xám, cột hẹp). KHÔNG sửa nội dung ô
// — app/backend đọc ID ảnh từ chính các link này.
function plPdf_gonLink() {
  var T = plPdf_docTab_(SpreadsheetApp.getActiveSheet());
  if (T.C.link < 0) { SpreadsheetApp.getUi().alert('Không tìm thấy cột link ảnh.'); return; }
  var c = T.C.link + 1;
  T.sh.getRange(T.hr + 2, c, T.n, 1).setWrapStrategy(SpreadsheetApp.WrapStrategy.CLIP)
    .setFontSize(8).setFontColor('#8a8a8a').setVerticalAlignment('top');
  T.sh.setColumnWidth(c, 110);
  SpreadsheetApp.getUi().alert('Đã làm gọn cột ' + T.sh.getRange(1, c).getA1Notation().replace(/\d+/, '')
    + ' (link ảnh). Nội dung link giữ nguyên.');
}
