/**
 * XuatPDF_Sheet.gs — Menu "📄 XUẤT PDF" ngay trên Google Sheet Punch List
 * PHIEN_BAN_PDF: '16/10 qa2'
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

var PLPDF_PHIEN_BAN = '16/10 qa2';
var PLPDF_THU_MUC = 'PUNCH LIST - PDF xuat tu Sheet';   // thư mục Drive chứa PDF

// ---------- Cài đặt / menu ----------
function plPdf_caiDat() {
  var ss = SpreadsheetApp.getActive();
  if (!ss) throw new Error('Hãy mở Apps Script TỪ trong Google Sheet (Tiện ích mở rộng → Apps Script).');
  var daCo = ScriptApp.getProjectTriggers().some(function (t) {
    return t.getHandlerFunction() === 'plPdf_onOpen';
  });
  if (!daCo) ScriptApp.newTrigger('plPdf_onOpen').forSpreadsheet(ss).onOpen().create();
  // Run from the editor: getUi() is not allowed -> skip, menu appears after reloading the Sheet
  try { plPdf_onOpen(); } catch (e) { Logger.log('Đã cài trigger. Tải lại (F5) Google Sheet để thấy menu.'); }
  return 'Đã cài menu XUẤT PDF (' + PLPDF_PHIEN_BAN + ')';
}

function plPdf_onOpen() {
  SpreadsheetApp.getUi().createMenu('📄 XUẤT PDF')
    .addItem('⭐ PUNCH LIST theo SECTION (giống app)', 'plPdf_theoSection')
    .addItem('⭐ PUNCH LIST theo SECTION — chỉ mục ĐANG MỞ', 'plPdf_theoSectionMo')
    .addSeparator()
    .addItem('📊 Cập nhật DASHBOARD + nhận dự án mới', 'plPdf_dashboard')
    .addItem('🎨 Tô màu Section + Đóng/Mở (mọi tab dự án)', 'plPdf_toMau')
    .addItem('🧹 Làm gọn cột link ảnh (tab đang mở)', 'plPdf_gonLink')
    .addItem('⏱ Bật tự cập nhật Dashboard mỗi giờ', 'plPdf_batTuDong')
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
// opt.light = true: only display values (dashboard / colouring) - skip the heavy reads
function plPdf_docTab_(sh, opt) {
  var light = !!(opt && opt.light);
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
  var fmls = [], raw = [];   // raw 16/10: giá trị gốc — ảnh DÁN TAY trong ô là đối tượng CellImage
  if (!light) {
    fmls = sh.getRange(hr + 2, 1, n, lastC).getFormulas();
    try { raw = sh.getRange(hr + 2, 1, n, lastC).getValues(); } catch (eRaw) { raw = []; }
  }
  function col(test) { for (var c = 0; c < h.length; c++) if (test(h[c])) return c; return -1; }
  var C = {
    no: col(function (x) { return x === 'no.' || x === 'no'; }),
    sec: col(function (x) { return x === 'section' || x === 'sec'; }),
    part: col(function (x) { return x === 'part'; }),
    desc: col(function (x) { return x === 'description'; }),
    rem: col(function (x) { return x === 'remarks' || x === 'ghi chú'; }),
    st: col(function (x) { return x === 'status' || x === 'tình trạng'; }),
    qc: col(function (x) { return x === 'người lập' || x === 'inspected by'; }),
    link: col(function (x) { return x.indexOf('photo link') >= 0 || x.indexOf('link ảnh') >= 0; }),
    after: col(function (x) { return x.indexOf('photo after') === 0 || x.indexOf('ảnh khắc phục') === 0; })
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
  var rt = [];   // 16/10: ô link là CHỮ có gắn link (tên file) → lấy đường link ẩn
  var notes = [];   // 16/10: ghi chú "AFTER:id,..." trên ô Photo after (Code.gs V50+) = ảnh SAU
  if (!light) {
    try { if (C.link >= 0) rt = sh.getRange(hr + 2, C.link + 1, n, 1).getRichTextValues(); } catch (eRt) {}
    try { if (C.after >= 0) notes = sh.getRange(hr + 2, C.after + 1, n, 1).getNotes(); } catch (eN) {}
  }
  return { sh: sh, hr: hr, n: n, vals: vals, fmls: fmls, raw: raw, rt: rt, C: C, notes: notes };
}

function plPdf_linkIds_(v) {
  var out = [];
  if (!v) return out;
  try {
    var urls = [v.getLinkUrl()];
    v.getRuns().forEach(function (r) { urls.push(r.getLinkUrl()); });
    urls.forEach(function (u) { var m = String(u || '').match(/(?:\/d\/|[?&]id=)([-\w]{25,})/); if (m && out.indexOf(m[1]) < 0) out.push(m[1]); });
  } catch (e) {}
  return out;
}
// Drive IDs inside URLs only (/d/ID or ?id=ID / &id=ID) - used for formulas
function plPdf_urlIds_(s) {
  var out = [], m, re = /(?:\/d\/|[?&]id=)([-\w]{25,})/g, t = String(s || '');
  while ((m = re.exec(t))) if (out.indexOf(m[1]) < 0) out.push(m[1]);
  return out;
}
// Link column text: a line counts only if it is a bare ID or contains a Drive URL.
// File names like "PLSclose_20261015_101010_1.jpg" are NOT IDs (qa2).
function plPdf_ids_(s) {
  var out = [];
  String(s || '').split(/[\r\n]+/).forEach(function (ln) {
    ln = ln.trim();
    if (!ln) return;
    if (/^[-\w]{25,}$/.test(ln)) { if (out.indexOf(ln) < 0) out.push(ln); return; }
    plPdf_urlIds_(ln).forEach(function (x) { if (out.indexOf(x) < 0) out.push(x); });
  });
  return out;
}

// qa2: dòng có nội dung = có mô tả HOẶC part HOẶC ảnh (giống cleanItems của app).
// Bản đọc nhẹ (light) không có fmls/raw/rt → chỉ xét chữ trong cột link.
function plPdf_coNoiDung_(T, i) {
  var C = T.C, r = T.vals[i] || [];
  if (String(r[C.desc] || '').trim()) return true;
  if (C.part >= 0 && String(r[C.part] || '').trim()) return true;
  if (C.link >= 0 && String(r[C.link] || '').trim()) return true;
  if (plPdf_linkIds_(T.rt[i] && T.rt[i][0]).length) return true;
  var f = T.fmls[i] || [];
  for (var c = 0; c < f.length; c++) if (f[c] && plPdf_urlIds_(f[c]).length) return true;
  var v = T.raw[i] || [];
  for (var k = 0; k < v.length; k++) if (plPdf_laAnhDan_(v[k])) return true;
  return false;
}

// Ảnh dán tay trong ô của dòng i: cột Photo after → ảnh SAU, cột khác → ảnh TRƯỚC.
// Google tự in ảnh trong ô ra PDF, nên chỉ cần chép ảnh sang sheet tạm.
function plPdf_laAnhDan_(v) { return v && typeof v === 'object' && !(v instanceof Date) && typeof v.getContentUrl === 'function'; }
function plPdf_anhDan_(T, i) {
  var b = [], a = [], r = T.raw[i] || [];
  for (var c = 0; c < r.length; c++) {
    if (!plPdf_laAnhDan_(r[c])) continue;
    var o = { img: r[c], src: T.sh.getRange(T.hr + 2 + i, c + 1) };
    (c === T.C.after ? a : b).push(o);
  }
  return { b: b, a: a };
}

// Tách ảnh trước / sau: ảnh đóng lỗi có tên PLSclose_ / PLclose_ (backend đặt khi đóng)
function plPdf_tachAnh_(ids, closed, cache, afterSet) {
  var b = [], a = [];
  ids.forEach(function (id) {
    if (afterSet && afterSet[id]) { a.push(id); return; }   // 16/10: đánh dấu AFTER / công thức Photo after
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
  T.vals.forEach(function (r, i) { var s = String(r[C.sec]).trim(); if (s && plPdf_coNoiDung_(T, i)) secs[s] = 1; });
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
    if (!plPdf_coNoiDung_(T, i)) continue;
    var desc = String(r[C.desc] || '').trim() || '(chưa ghi mô tả)';
    var sec = String(r[C.sec] || '').trim();
    if (want.length && want.indexOf(sec) < 0) continue;
    var st = C.st >= 0 ? String(r[C.st] || '').trim().toLowerCase() : '';
    var closed = /^clos/.test(st);
    if (chiMo && closed) continue;
    var ids = C.link >= 0 ? plPdf_ids_(r[C.link]) : [];
    plPdf_linkIds_(T.rt[i] && T.rt[i][0]).forEach(function (x) { if (ids.indexOf(x) < 0) ids.push(x); });
    // qa2: luôn gom thêm ảnh từ công thức IMAGE()/HYPERLINK() trên dòng (giống listSimple của app)
    f.forEach(function (x) { plPdf_urlIds_(x).forEach(function (id) { if (ids.indexOf(id) < 0) ids.push(id); }); });
    var afterSet = {};
    var nt = T.notes[i] ? String(T.notes[i][0] || '').match(/AFTER:([-\w,]+)/) : null;
    if (nt) nt[1].split(',').forEach(function (x) { if (x) afterSet[x] = 1; });
    if (C.after >= 0) plPdf_urlIds_(f[C.after]).forEach(function (x) { afterSet[x] = 1; if (ids.indexOf(x) < 0) ids.push(x); });
    var sp = plPdf_tachAnh_(ids, closed, cache, afterSet);
    // 16/10: ảnh DÁN TAY trong ô (chưa có link Drive) → chép thẳng ảnh vào báo cáo
    var dan = plPdf_anhDan_(T, i);
    if (!sp.b.length) sp.b = dan.b;
    if (!sp.a.length) sp.a = dan.a;
    items.push({ no: C.no >= 0 ? r[C.no] : '', sec: sec, part: C.part >= 0 ? r[C.part] : '', desc: desc,
                 rem: C.rem >= 0 ? r[C.rem] : '', qc: C.qc >= 0 ? r[C.qc] : '', closed: closed, b: sp.b, a: sp.a });
  }
  if (!items.length) { ui.alert('Không có mục nào khớp.'); return; }
  items.sort(function (x, y) {   // 16/10: section chữ / số lẫn nhau không còn xen kẽ
    var a = parseFloat(x.sec), b = parseFloat(y.sec);
    var d = (isNaN(a) ? 1e9 : a) - (isNaN(b) ? 1e9 : b);
    return d || String(x.sec).localeCompare(String(y.sec)) || (Number(x.no) || 0) - (Number(y.no) || 0);
  });

  var tab = T.sh.getName();
  var duAn = (tab.split('-').slice(1).join('-') || tab).trim();
  var tz = ss.getSpreadsheetTimeZone() || 'Asia/Ho_Chi_Minh';
  var ngay = Utilities.formatDate(new Date(), tz, 'dd/MM/yyyy');
  var nhan = (want.length ? 'Section ' + want.join(', ') : 'Tất cả section') + (chiMo ? ' · chỉ mục đang mở' : '');
  ss.getSheets().forEach(function (x) { if (x.getName().indexOf('PDF_TAM_') === 0) { try { ss.deleteSheet(x); } catch (eD) {} } });   // dọn sheet tạm còn sót (lần trước bị ngắt)
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

  var row = 4, cur = null, rows = [], fm = [], meta = [], dan = [];
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
        if (j < list.length && typeof list[j] === 'object') {   // ảnh dán tay → chép sau khi ghi bảng
          dan.push({ r: row, c: line.length + 1, o: list[j] });
          line.push('');
        } else if (j < list.length) {
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
  dan.forEach(function (d) {   // ảnh dán tay: ghi ảnh trong ô; không được thì chép nguyên ô gốc
    var dst = sh.getRange(d.r, d.c);
    try { dst.setValue(d.o.img); }
    catch (eSv) { try { d.o.src.copyTo(dst, SpreadsheetApp.CopyPasteType.PASTE_VALUES, false); } catch (eCp) { dst.setValue('ảnh dán (lỗi chép)'); } }
  });
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
    if (!m.na) sh.getRange(m.r, 6 + M).setValue(m.closed ? 'chưa có ảnh sau' : 'chưa khắc phục').setFontColor('#999999').setFontStyle('italic');
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


// =====================================================================
// DASHBOARD + NHẬN DỰ ÁN MỚI + TÔ MÀU + SỬA ẢNH NỔI   ('15/10 dashboard')
// Dự án = mọi tab tên "Punch list đơn giản - <dự án>" (tự nhận tab mới).
// =====================================================================
var PLPDF_TAB_DB = 'DASHBOARD PUNCH';   // tên riêng — tránh đè tab của Dasboard.gs có sẵn
var PLPDF_TAB_CAT = 'Danh mục part';
var PLPDF_MARK = '__DU_AN__';   // part giả: chỉ để app biết có dự án (app tự ẩn)
var PLPDF_MAU_SEC = ['#e8f0fe', '#fef7e0', '#e6f4ea', '#fce8e6', '#e0f7fa'];   // xanh dương · vàng · xanh lá · hồng · xanh ngọc (nhạt)

function plPdf_tenDuAn_(tab) {
  var m = String(tab).match(/^punch\s*list\s*đơn\s*giản\s*-?\s*(.*)$/i);
  if (!m) return null;   // không phải tab dự án
  // Code.gs: dự án mặc định (PUNCH_PROJECT_NAME = BISON) dùng tab gốc "Punch list đơn giản" (không hậu tố)
  return m[1].trim() || (typeof PUNCH_PROJECT_NAME !== 'undefined' && PUNCH_PROJECT_NAME ? PUNCH_PROJECT_NAME : 'BISON');
}
function plPdf_cacTabDuAn_(ss) {
  return ss.getSheets().map(function (sh) { return { sh: sh, duAn: plPdf_tenDuAn_(sh.getName()) }; })
    .filter(function (x) { return x.duAn; });
}
// SID → ngày tạo: SPyyMMddHHmmssNN  hoặc  SP<epoch ms>_n
function plPdf_ngaySid_(sid) {
  var s = String(sid || ''), m;
  if ((m = s.match(/^SP(\d{13})(_|$)/))) return new Date(Number(m[1]));
  if ((m = s.match(/^SP(\d{2})(\d{2})(\d{2})(\d{2})(\d{2})(\d{2})/)))
    return new Date(2000 + Number(m[1]), Number(m[2]) - 1, Number(m[3]), Number(m[4]), Number(m[5]), Number(m[6]));
  return null;
}
function plPdf_cotSid_(T) {   // cột có nhiều giá trị SP... nhất (BISON có 2 cột "SID")
  var best = -1, bn = 0;
  for (var c = 0; c < (T.vals[0] || []).length; c++) {
    var k = 0;
    for (var i = 0; i < T.n; i++) if (/^SP\d/.test(T.vals[i][c])) k++;
    if (k > bn) { bn = k; best = c; }
  }
  return best;
}

// Ghi dự án chưa có vào "Danh mục part" (1 dòng: dự án | | __DU_AN__) → app mọi máy thấy dự án
function plPdf_dongBoDuAn_(ss, duAns) {
  var cat = ss.getSheetByName(PLPDF_TAB_CAT);
  if (!cat) return [];
  var last = cat.getLastRow(), co = {};
  if (last >= 2) {
    var v = cat.getRange(2, 1, last - 1, 3).getDisplayValues(), lastPj = '';
    v.forEach(function (r) { var pj = String(r[0]).trim() || lastPj; lastPj = pj; if (pj && String(r[2]).trim()) co[pj.toLowerCase()] = 1; });
  }
  var them = duAns.filter(function (d) { return !co[d.toLowerCase()]; });
  if (them.length) {
    // ghi ngay dưới dòng cuối có dữ liệu ở cột A:C (cột F có hướng dẫn — không đụng)
    var colA = cat.getRange(1, 1, Math.max(1, last), 3).getDisplayValues(), r = colA.length;
    while (r > 1 && !colA[r - 1].join('').trim()) r--;
    cat.getRange(r + 1, 1, them.length, 3).setValues(them.map(function (d) { return [d, '', PLPDF_MARK]; }));
    try { CacheService.getScriptCache().remove('catalog_v1'); } catch (eC) {}   // app thấy dự án mới ngay
  }
  return them;
}

function plPdf_dashboard() { plPdf_dashboard_(true); }
function plPdf_dashboardTuDong() { plPdf_dashboard_(false); }

function plPdf_dashboard_(coUi) {
  var ss = SpreadsheetApp.getActive() || SpreadsheetApp.openById(PropertiesService.getScriptProperties().getProperty('PLPDF_SS'));
  var tz = ss.getSpreadsheetTimeZone() || 'Asia/Ho_Chi_Minh';
  var now = new Date();
  var tabs = plPdf_cacTabDuAn_(ss), P = [], qcAll = {};
  tabs.forEach(function (t) {
    var T;
    try { T = plPdf_docTab_(t.sh, { light: true }); } catch (e) {
      if (/chưa có mục|trống/i.test(e.message)) P.push({ duAn: t.duAn, tab: t.sh.getName(), tong: 0, mo: 0, dong: 0, mo7: 0, cuNhat: 0, moi: null, sec: {} });
      else P.push({ duAn: t.duAn, tab: t.sh.getName(), loi: e.message });
      return;
    }
    var C = T.C, cSid = plPdf_cotSid_(T), d = { duAn: t.duAn, tab: t.sh.getName(), tong: 0, mo: 0, dong: 0, mo7: 0, cuNhat: 0, moi: null, sec: {} };
    for (var i = 0; i < T.n; i++) {
      var r = T.vals[i];
      if (!plPdf_coNoiDung_(T, i)) continue;
      var closed = C.st >= 0 && /^clos/i.test(String(r[C.st]).trim());
      var sec = String(r[C.sec] || '').trim() || '(trống)';
      var ng = cSid >= 0 ? plPdf_ngaySid_(r[cSid]) : null;
      d.tong++; closed ? d.dong++ : d.mo++;
      if (!d.sec[sec]) d.sec[sec] = { t: 0, m: 0 };
      d.sec[sec].t++; if (!closed) d.sec[sec].m++;
      if (ng && (!d.moi || ng > d.moi)) d.moi = ng;
      if (!closed && ng) {
        var tuoi = Math.floor((now - ng) / 86400000);
        if (tuoi > 7) d.mo7++;
        if (tuoi > d.cuNhat) d.cuNhat = tuoi;
      }
      var qc = C.qc >= 0 ? String(r[C.qc] || '').trim() : '';
      if (qc) { if (!qcAll[qc]) qcAll[qc] = { t: 0, m: 0 }; qcAll[qc].t++; if (!closed) qcAll[qc].m++; }
    }
    P.push(d);
  });
  var them = plPdf_dongBoDuAn_(ss, P.filter(function (x) { return !x.loi; }).map(function (x) { return x.duAn; }));

  // ---- ghi tab DASHBOARD (ghi đè tại chỗ, không xoá tab) ----
  var db = ss.getSheetByName(PLPDF_TAB_DB) || ss.insertSheet(PLPDF_TAB_DB, 0);
  try { db.getRange(1, 1, db.getMaxRows(), db.getMaxColumns()).breakApart(); } catch (eB) {}   // 16/10: gỡ ô gộp của lần trước
  db.clear(); db.clearConditionalFormatRules();
  try { db.getCharts().forEach(function (c) { db.removeChart(c); }); } catch (e) {}
  var NAVY = '#1c4587', LINE = '#b7c4d4', out = [], fmt = [];
  function row(a, f) { while (a.length < 8) a.push(''); out.push(a); fmt.push(f || ''); }
  row(['📊 DASHBOARD PUNCH LIST'], 'title');
  row(['Cập nhật: ' + Utilities.formatDate(now, tz, 'dd/MM/yyyy HH:mm') + '   ·   ' + P.length + ' dự án   ·   tự nhận mọi tab "Punch list đơn giản - <dự án>"'], 'sub');
  row([]);
  row(['Dự án', 'Tổng mục', 'Đang mở', 'Đã đóng', '% đã đóng', 'Tiến độ', 'Mở > 7 ngày', 'Mở lâu nhất (ngày)'], 'head');
  var T0 = { t: 0, m: 0, d: 0, m7: 0 };
  P.forEach(function (d) {
    if (d.loi) { row([d.duAn, 'Lỗi: ' + d.loi], 'err'); return; }
    var pct = d.tong ? d.dong / d.tong : 0;
    T0.t += d.tong; T0.m += d.mo; T0.d += d.dong; T0.m7 += d.mo7;
    row([d.duAn, d.tong, d.mo, d.dong, pct, '=SPARKLINE(' + pct.toFixed(4) + ',{"charttype","bar";"max",1;"color1","#16A34A";"empty","zero"})', d.mo7, d.cuNhat || ''], 'proj');
  });
  row(['TỔNG', T0.t, T0.m, T0.d, T0.t ? T0.d / T0.t : 0, '', T0.m7, ''], 'total');
  row([]);
  P.forEach(function (d) {
    if (d.loi) return;
    row(['DỰ ÁN ' + d.duAn + ' — theo Section' + (d.moi ? '   (mục mới nhất: ' + Utilities.formatDate(d.moi, tz, 'dd/MM/yyyy') + ')' : '')], 'band');
    row(['Section', 'Tổng mục', 'Đang mở', 'Đã đóng', '% đã đóng', 'Tiến độ'], 'head');
    Object.keys(d.sec).sort(function (a, b) { return (parseFloat(a) || 0) - (parseFloat(b) || 0); }).forEach(function (k) {
      var x = d.sec[k], pc = x.t ? (x.t - x.m) / x.t : 0;
      row([k, x.t, x.m, x.t - x.m, pc, '=SPARKLINE(' + pc.toFixed(4) + ',{"charttype","bar";"max",1;"color1","#16A34A";"empty","zero"})'], 'sec');
    });
    row([]);
  });
  var qcs = Object.keys(qcAll);
  if (qcs.length) {
    row(['THEO NGƯỜI LẬP (mọi dự án)'], 'band');
    row(['Người lập', 'Tổng mục', 'Đang mở', 'Đã đóng'], 'head');
    qcs.sort(function (a, b) { return qcAll[b].t - qcAll[a].t; }).forEach(function (q) { row([q, qcAll[q].t, qcAll[q].m, qcAll[q].t - qcAll[q].m], 'sec'); });
  }
  db.getRange(1, 1, out.length, 8).setValues(out);
  db.setColumnWidth(1, 170); for (var c = 2; c <= 8; c++) db.setColumnWidth(c, 105);
  db.getRange(1, 1, out.length, 8).setFontSize(10).setVerticalAlignment('middle');
  db.getRange(1, 2, out.length, 7).setHorizontalAlignment('center');
  fmt.forEach(function (f, i) {
    var rg = db.getRange(i + 1, 1, 1, 8), r = i + 1;
    if (f === 'title') { db.getRange(r, 1, 1, 8).merge().setFontSize(16).setFontWeight('bold').setFontColor('#fff').setBackground(NAVY).setHorizontalAlignment('left'); db.setRowHeight(r, 34); }
    else if (f === 'sub') { db.getRange(r, 1, 1, 8).merge().setFontColor('#555').setFontSize(9).setHorizontalAlignment('left'); }
    else if (f === 'head') rg.setBackground(NAVY).setFontColor('#fff').setFontWeight('bold');
    else if (f === 'band') { db.getRange(r, 1, 1, 8).merge().setBackground('#e8eef7').setFontColor(NAVY).setFontWeight('bold').setHorizontalAlignment('left'); }
    else if (f === 'total') rg.setFontWeight('bold').setBackground('#f1f3f4').setBorder(true, true, true, true, true, true, LINE, SpreadsheetApp.BorderStyle.SOLID);
    else if (f === 'err') rg.setFontColor('#c00');
    if (f === 'proj' || f === 'sec' || f === 'total') {
      rg.setBorder(true, true, true, true, true, true, LINE, SpreadsheetApp.BorderStyle.SOLID);
      db.getRange(r, 3).setFontColor('#b45309').setFontWeight('bold');   // Đang mở
      db.getRange(r, 4).setFontColor('#15803d').setFontWeight('bold');   // Đã đóng
      if (out[i][4] !== '') db.getRange(r, 5).setNumberFormat('0%');
      if (f === 'proj') db.getRange(r, 1).setFontWeight('bold');
      if (f === 'proj' && out[i][6] > 0) db.getRange(r, 7).setBackground('#f4cccc').setFontColor('#a50e0e').setFontWeight('bold');
    }
  });
  db.setFrozenRows(0);
  if (coUi) {
    SpreadsheetApp.getUi().alert('✅ Đã cập nhật DASHBOARD (' + P.length + ' dự án).'
      + (them.length ? '\n\nDự án MỚI đã ghi vào "Danh mục part" để app nhận: ' + them.join(', ')
        + '\n→ Trên app bấm 🔄 Tải lại danh mục.' : ''));
    db.activate();
  }
}

function plPdf_batTuDong() {
  var ss = SpreadsheetApp.getActive();
  PropertiesService.getScriptProperties().setProperty('PLPDF_SS', ss.getId());
  var co = ScriptApp.getProjectTriggers().some(function (t) { return t.getHandlerFunction() === 'plPdf_dashboardTuDong'; });
  if (!co) ScriptApp.newTrigger('plPdf_dashboardTuDong').timeBased().everyHours(1).create();
  plPdf_dashboard_(true);
}

// Tô màu: Section xen kẽ 5 màu nhạt (cột No.→Description) + Tình trạng cam/xanh. CHỈ định dạng có điều kiện.
function plPdf_toMau() {
  var ss = SpreadsheetApp.getActive(), n = 0;
  plPdf_cacTabDuAn_(ss).forEach(function (t) {
    var T; try { T = plPdf_docTab_(t.sh, { light: true }); } catch (e) { return; }
    var C = T.C, sh = t.sh, r0 = T.hr + 2, nr = sh.getMaxRows() - r0 + 1;
    if (nr < 1) return;
    var colL = function (c) { return sh.getRange(1, c + 1).getA1Notation().replace(/\d+/, ''); };
    var secA = '$' + colL(C.sec) + r0;
    var left = Math.min(C.no >= 0 ? C.no : 0, C.sec), right = Math.max(C.desc, C.part >= 0 ? C.part : 0);
    var rgL = sh.getRange(r0, left + 1, nr, right - left + 1);
    // giữ quy tắc người dùng tự đặt; bỏ quy tắc cũ do chính menu này tạo (nhận theo dấu "*1+0=" / "PLPDF")
    var keep = sh.getConditionalFormatRules().filter(function (rule) {
      var bc = rule.getBooleanCondition(); if (!bc) return true;
      var v = String((bc.getCriteriaValues() || [])[0] || '');
      return v.indexOf('/*PL*/') < 0 && v.indexOf('*1+0=') < 0;
    });
    var rules = [];
    PLPDF_MAU_SEC.forEach(function (mau, k) {
      rules.push(SpreadsheetApp.newConditionalFormatRule()
        .whenFormulaSatisfied('=AND(LEN(' + secA + ')>0,MOD(IFERROR(VALUE(' + secA + '),CODE(' + secA + ')),' + PLPDF_MAU_SEC.length + ')*1+0=' + k + ')')
        .setBackground(mau).setRanges([rgL]).build());
    });
    if (C.st >= 0) {
      var stA = '$' + colL(C.st) + r0, rgS = sh.getRange(r0, C.st + 1, nr, 1);
      rules.push(SpreadsheetApp.newConditionalFormatRule().whenFormulaSatisfied('=REGEXMATCH(LOWER(' + stA + '),"^clos")*1+0=1')
        .setBackground('#d9ead3').setFontColor('#0b5d1e').setBold(true).setRanges([rgS]).build());
      rules.push(SpreadsheetApp.newConditionalFormatRule().whenFormulaSatisfied('=AND(LEN(' + stA + ')>0,NOT(REGEXMATCH(LOWER(' + stA + '),"^clos")))*1+0=1')
        .setBackground('#fce5cd').setFontColor('#8a4b00').setBold(true).setRanges([rgS]).build());
    }
    sh.setConditionalFormatRules(keep.concat(rules));
    n++;
  });
  SpreadsheetApp.getUi().alert('🎨 Đã tô màu ' + n + ' tab dự án:\n• Mỗi Section 1 màu nhạt (xen kẽ 5 màu)\n• Tình trạng: CAM = đang mở · XANH = đã đóng\n(Chỉ là định dạng có điều kiện — dữ liệu không đổi.)');
}

// Ảnh NỔI (chèn đè lên ô) → khi lọc sẽ chồng lên nhau. Đổi thành ảnh TRONG Ô =IMAGE() theo link của dòng đó.
// 15/10: TẠM KHOÁ — chạy trên bản BACKUP thì ảnh không hiện lại. Chưa rõ nguyên nhân → không cho chạy.
function plPdf_suaAnhNoi() {
  SpreadsheetApp.getUi().alert('Lệnh này đang TẠM KHOÁ để kiểm tra lại (15/10). Không có gì bị thay đổi.');
}
function plPdf_suaAnhNoi_KHOA_() {
  var ui = SpreadsheetApp.getUi(), sh = SpreadsheetApp.getActiveSheet();
  var T = plPdf_docTab_(sh), C = T.C;
  var imgs = sh.getImages();
  if (!imgs.length) { ui.alert('Tab này không có ảnh nổi. Ảnh trong ô (=IMAGE) không bị đè khi lọc.'); return; }
  var h = sh.getRange(T.hr + 1, 1, 1, sh.getLastColumn()).getDisplayValues()[0].map(function (x) { return String(x).trim().toLowerCase(); });
  var cPhoto = h.indexOf('photo');
  if (cPhoto < 0) { ui.alert('Không thấy cột "Photo".'); return; }
  if (ui.alert('Sửa ảnh bị đè', 'Tab có ' + imgs.length + ' ảnh nổi.\nẢnh nằm ở dòng có link ảnh sẽ được đổi thành ảnh TRONG Ô (cột Photo) rồi gỡ ảnh nổi.\nẢnh ở dòng không có link: GIỮ NGUYÊN.\n\nTiếp tục?', ui.ButtonSet.YES_NO) !== ui.Button.YES) return;
  var doi = 0, giu = 0;
  imgs.forEach(function (im) {
    var r = im.getAnchorCell().getRow() - (T.hr + 2);
    var url = '';
    try { url = im.getUrl() || ''; } catch (e) {}
    var id = (url.match(/[-\w]{25,}/) || [''])[0];
    if (!id && r >= 0 && r < T.n && C.link >= 0) id = (plPdf_ids_(T.vals[r][C.link])[0] || '');
    if (!id || r < 0 || r >= T.n) { giu++; return; }
    var cell = sh.getRange(T.hr + 2 + r, cPhoto + 1);
    if (!cell.getFormula()) cell.setFormula('=IFERROR(IMAGE("https://drive.google.com/thumbnail?id=' + id + '&sz=w400",1),"")');
    if (sh.getRowHeight(cell.getRow()) < 90) sh.setRowHeight(cell.getRow(), 100);
    im.remove(); doi++;
  });
  ui.alert('Đã đổi ' + doi + ' ảnh nổi thành ảnh trong ô.' + (giu ? '\nGiữ nguyên ' + giu + ' ảnh (dòng không có link).' : ''));
}
