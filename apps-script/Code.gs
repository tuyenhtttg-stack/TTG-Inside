/**
 * KHÔNG GIAN SỐ NỘI BỘ (TTG Inside) – Máy chủ phân quyền
 * VietinBank Chi nhánh Tây Tiền Giang · phiên bản 1.6.0 · 29/10/2026
 *
 * Script này GẮN VỚI file Google Sheet phân quyền (Tiện ích mở rộng → Apps Script), làm 2 việc:
 *   1. Menu quản trị + tự tick quyền theo vai trò ngay trong Sheet.
 *   2. Web App cho trang Không Gian Số (GitHub): đăng nhập bằng ID, trả quyền, đếm lượt truy cập, trợ lý AI (giai đoạn 2).
 *
 * Triển khai:
 *   1. Mở Sheet → Tiện ích mở rộng → Apps Script → dán toàn bộ file này → Lưu.
 *   2. Chọn hàm caiDatBanDau → Chạy → cấp quyền (lần đầu Google hỏi quyền truy cập Sheet, gửi email).
 *   3. Triển khai → Triển khai mới → Loại: Ứng dụng web; Thực thi với tư cách: Tôi; Người có quyền truy cập: Bất kỳ ai.
 *      Copy URL dạng https://script.google.com/macros/s/.../exec → dán vào API_URL trong index.html.
 *   4. Sửa code sau này: Triển khai → Quản lý triển khai → Sửa (bút chì) → Phiên bản: Mới → Triển khai (URL giữ nguyên).
 *
 * Script Properties (Cài đặt dự án → Thuộc tính tập lệnh) – tự tạo khi cài đặt, chỉ thêm khi cần:
 *   TOKEN_SECRET       : khóa ký phiên đăng nhập (tự sinh; đổi giá trị = buộc mọi người đăng nhập lại)
 *   ADMIN_EMAIL        : email nhận cảnh báo (mặc định email chủ script)
 *   ANTHROPIC_API_KEY  : (giai đoạn 2) khóa Claude API cho trợ lý AI; để trống = tắt AI
 *   MODEL, DAILY_LIMIT, PER_MINUTE_LIMIT : (giai đoạn 2) như hướng dẫn cũ
 *   BT_FOLDER_ID       : (v1.5.0) ID thư mục BanTin chứa file BanTin_YYYY-MM-DD.json (tự đặt khi chạy Cài đặt Bản tin sáng)
 *   TA_SHEET_ID        : (v1.6.0) ID file cấu hình thẻ ATM KhongGianSoNoiBo_TheATM (tự tạo khi chạy Cài đặt Thẻ ATM máy thu)
 * Manifest appsscript.json (v1.6.0) khai báo quyền tối thiểu, Drive chỉ ĐỌC (drive.readonly).
 */

const TZ = 'Asia/Ho_Chi_Minh';
const SH = { CANBO: 'CanBo', VAITRO: 'VaiTro', UNGDUNG: 'UngDung', NHATKY: 'NhatKy', PHONGBAN: 'PhongBan' };
const PB_COL = { MA: 'Mã phòng/ban', TEN: 'Tên phòng/ban', LOAI: 'Loại', ON: 'Đang dùng', GHICHU: 'Ghi chú' };
// Danh mục phòng/ban ban đầu – tên trùng với tên phòng trong sổ công tác
const PHONGBAN_BAN_DAU = [
  ['76201', 'Ban Giám đốc', 'Ban'], ['76202', 'Phòng TCTH', 'Phòng'], ['76203', 'Phòng DVKH', 'Phòng'],
  ['76204', 'Phòng KHDN', 'Phòng'], ['76205', 'Phòng Bán lẻ', 'Phòng'], ['76206', 'Phòng HTTD', 'Phòng'],
  ['76230', 'PGD Cái Bè', 'PGD'], ['76231', 'PGD Long Định', 'PGD'], ['76232', 'PGD Hưng Long', 'PGD'],
  ['76233', 'PGD Ấp Bắc', 'PGD'], ['76240', 'PGD An Hữu', 'PGD']
];
const COL = { ID: 'ID', TEN: 'Họ và tên', GT: 'Giới tính', PB: 'Phòng/Ban', VT: 'Vai trò', TT: 'Trạng thái',
              NGAYCAP: 'Ngày cấp', LAST: 'Đăng nhập gần nhất', GHICHU: 'Ghi chú' };
const TT = { ON: 'Hoạt động', OFF: 'Khóa', NEW: 'Chưa cấp' };
const APP_RE = /^I\d{3}$/;

// Chặn đoán thử
const DEVICE_MAX_FAIL = 5, DEVICE_LOCK_MIN = 15;
const GLOBAL_MAX_FAIL = 30, GLOBAL_WINDOW_MIN = 10, GLOBAL_SUSPEND_MIN = 15;
const CACHE_KEY = 'kgs_data_v1', CACHE_SEC = 120;

/* =====================================================================
   PHẦN 1 – QUẢN TRỊ TRONG SHEET
   ===================================================================== */
function onOpen() {
  SpreadsheetApp.getUi().createMenu('Không Gian Số')
    .addItem('Cài đặt lần đầu', 'caiDatBanDau')
    .addSeparator()
    .addItem('Cài đặt Phòng/Ban', 'caiDatPhongBan')
    .addItem('Cài đặt / cập nhật Quản lý chi phí (I003, I004)', 'caiDatChiPhi')
    .addItem('Đồng bộ danh mục sang file chi phí', 'dongBoChiPhi')
    .addItem('Cài đặt Bản tin sáng (I005)', 'caiDatBanTin')
    .addItem('Cài đặt / cập nhật Thẻ ATM máy thu (I006, I007)', 'caiDatTheATM')
    .addItem('Gửi ngay email nhắc hủy thẻ ATM', 'nhacHuyTheATM')
    .addItem('Thêm ứng dụng mới…', 'themUngDung')
    .addItem('Sinh thêm ID dự phòng…', 'sinhThemID')
    .addItem('Áp dụng lại quyền theo vai trò (tất cả)', 'apDungLaiVaiTro')
    .addSeparator()
    .addItem('Thống kê truy cập', 'xemThongKe')
    .addItem('Kiểm tra dữ liệu', 'kiemTraDuLieu')
    .addItem('Làm mới dữ liệu cho ứng dụng', 'lamMoiBoNhoDem')
    .addToUi();
}

/** Cài đặt định dạng: ô tick, danh sách chọn, màu trạng thái, khóa phiên, email cảnh báo */
function caiDatBanDau() {
  const ss = SpreadsheetApp.getActive();
  const props = PropertiesService.getScriptProperties();
  if (!props.getProperty('TOKEN_SECRET')) props.setProperty('TOKEN_SECRET', Utilities.getUuid() + Utilities.getUuid());
  props.setProperty('SHEET_ID', ss.getId());
  if (ss.getSpreadsheetTimeZone() !== TZ) ss.setSpreadsheetTimeZone(TZ); // ngày giờ ghi vào Sheet theo giờ Việt Nam
  if (!props.getProperty('ADMIN_EMAIL')) props.setProperty('ADMIN_EMAIL', Session.getEffectiveUser().getEmail());

  const cb = ss.getSheetByName(SH.CANBO), vt = ss.getSheetByName(SH.VAITRO), ud = ss.getSheetByName(SH.UNGDUNG);
  ensureRows_(cb, 200); ensureRows_(vt, 50); ensureRows_(ud, 100);
  const last = cb.getMaxRows();
  const h = headerMap_(cb);

  cb.getRange(2, h[COL.ID] + 1, last - 1, 1).setNumberFormat('@');
  setList_(cb.getRange(2, h[COL.GT] + 1, last - 1, 1), ['Nam', 'Nữ']);
  setList_(cb.getRange(2, h[COL.TT] + 1, last - 1, 1), [TT.ON, TT.OFF, TT.NEW]);
  cb.getRange(2, h[COL.VT] + 1, last - 1, 1).setDataValidation(
    SpreadsheetApp.newDataValidation().requireValueInRange(vt.getRange('A2:A'), true).setAllowInvalid(false).build());
  cb.getRange(2, h[COL.NGAYCAP] + 1, last - 1, 1).setNumberFormat('dd/mm/yyyy');
  cb.getRange(2, h[COL.LAST] + 1, last - 1, 1).setNumberFormat('dd/mm/yyyy hh:mm');
  appCols_(cb).forEach(c => checkboxCol_(cb, c + 1, last));
  appCols_(vt).forEach(c => checkboxCol_(vt, c + 1, vt.getMaxRows()));
  checkboxCol_(ud, 4, 100); checkboxCol_(ud, 5, 100);

  // Màu theo trạng thái
  const ttCol = colLetter_(h[COL.TT] + 1), rng = cb.getRange('A2:' + colLetter_(cb.getLastColumn()) + last);
  const rule = (txt, bg, fg) => SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied('=$' + ttCol + '2="' + txt + '"').setBackground(bg).setFontColor(fg).setRanges([rng]).build();
  cb.setConditionalFormatRules([rule(TT.OFF, '#FDE7EC', '#9B0E35'), rule(TT.NEW, '#F3F5F7', '#8A97A3')]);

  [cb, vt, ud, ss.getSheetByName(SH.NHATKY)].forEach(s => { s.setFrozenRows(1); s.getRange(1, 1, 1, s.getLastColumn()).setFontWeight('bold').setFontColor('#FFFFFF').setBackground('#005993'); });
  caiDatPhongBan_(false);
  lamMoiBoNhoDem(true);
  const warn = [kiemTraVaiTro_(), kiemTraPhongBan_()].filter(Boolean).join('\n\n');
  SpreadsheetApp.getUi().alert((warn ? warn + '\n\n' : '') + 'Đã cài đặt xong.\n\nBước tiếp theo: Triển khai → Triển khai mới → Ứng dụng web (Thực thi: Tôi; Truy cập: Bất kỳ ai), rồi dán URL vào API_URL trong index.html.');
}

/** Tự tick quyền khi chọn vai trò; lan quyền khi sửa sheet VaiTro; xóa bộ nhớ đệm */
function onEdit(e) {
  try {
    const s = e.range.getSheet(), name = s.getName();
    if (![SH.CANBO, SH.VAITRO, SH.UNGDUNG, SH.PHONGBAN].includes(name)) return;
    CacheService.getScriptCache().remove(CACHE_KEY);
    if (e.range.getRow() < 2) return;

    if (name === SH.CANBO) {
      const h = headerMap_(s);
      const c1 = e.range.getColumn(), c2 = e.range.getLastColumn();
      // Chọn vai trò → chép quyền vai trò sang dòng cán bộ
      if (h[COL.VT] + 1 >= c1 && h[COL.VT] + 1 <= c2) {
        const roles = rolePerms_();
        const apps = appCols_(s);
        for (let r = e.range.getRow(); r <= e.range.getLastRow(); r++) {
          const role = String(s.getRange(r, h[COL.VT] + 1).getValue()).trim();
          const p = roles[role]; if (!p) continue;
          apps.forEach(c => s.getRange(r, c + 1).setValue(!!p[headerAt_(s, c)]));
        }
      }
      // Chuyển sang Hoạt động lần đầu → ghi Ngày cấp
      if (h[COL.TT] + 1 >= c1 && h[COL.TT] + 1 <= c2) {
        for (let r = e.range.getRow(); r <= e.range.getLastRow(); r++) {
          if (s.getRange(r, h[COL.TT] + 1).getValue() === TT.ON && !s.getRange(r, h[COL.NGAYCAP] + 1).getValue())
            s.getRange(r, h[COL.NGAYCAP] + 1).setValue(new Date());
        }
      }
    }

    if (name === SH.PHONGBAN) {
      // Đổi tên 1 phòng/ban (sửa 1 ô ở cột Tên) → cập nhật tên đó cho mọi cán bộ ở sheet CanBo
      const hp = headerMap_(s);
      if (e.range.getNumRows() === 1 && e.range.getNumColumns() === 1 && e.range.getColumn() === hp[PB_COL.TEN] + 1
          && e.oldValue !== undefined && e.value !== undefined && String(e.oldValue).trim() && String(e.value).trim()) {
        const oldName = String(e.oldValue).trim(), newName = String(e.value).trim();
        const cb = SpreadsheetApp.getActive().getSheetByName(SH.CANBO), h = headerMap_(cb), n = cb.getLastRow() - 1;
        if (n > 0) {
          const rng = cb.getRange(2, h[COL.PB] + 1, n, 1), vals = rng.getValues();
          let k = 0;
          const out = vals.map(v => { if (String(v[0]).trim() === oldName) { k++; return [newName]; } return v; });
          if (k) { rng.setValues(out); SpreadsheetApp.getActive().toast('Đã đổi "' + oldName + '" → "' + newName + '" cho ' + k + ' cán bộ.', 'Phòng/Ban'); }
        }
      }
      return;
    }

    if (name === SH.VAITRO) {
      // Đổi tick ứng dụng của 1 vai trò → áp dụng cho mọi cán bộ có vai trò đó (chỉ cột ứng dụng vừa đổi)
      const appId = headerAt_(s, e.range.getColumn() - 1);
      if (!APP_RE.test(appId)) return;
      const cb = SpreadsheetApp.getActive().getSheetByName(SH.CANBO);
      const h = headerMap_(cb), col = h[appId]; if (col === undefined) return;
      const roles = rolePerms_(), n = cb.getLastRow() - 1; if (n < 1) return;
      const vtVals = cb.getRange(2, h[COL.VT] + 1, n, 1).getValues();
      const target = cb.getRange(2, col + 1, n, 1), cur = target.getValues();
      const changedRoles = s.getRange(e.range.getRow(), 1, e.range.getNumRows(), 1).getValues().map(r => String(r[0]).trim());
      const out = cur.map((v, i) => { const role = String(vtVals[i][0]).trim(); return [changedRoles.includes(role) ? !!roles[role][appId] : v[0]]; });
      target.setValues(out);
    }
  } catch (err) { console.error(err); }
}

/** Menu: tạo sheet PhongBan (nếu chưa có) và gắn danh sách chọn cho cột Phòng/Ban ở CanBo */
function caiDatPhongBan() {
  const created = caiDatPhongBan_(true);
  lamMoiBoNhoDem(true);
  const warn = kiemTraPhongBan_();
  SpreadsheetApp.getUi().alert('Phòng/Ban',
    (created ? 'Đã tạo sheet PhongBan với ' + PHONGBAN_BAN_DAU.length + ' phòng/ban.\n' : 'Sheet PhongBan đã có sẵn, giữ nguyên dữ liệu.\n') +
    'Cột Phòng/Ban ở sheet CanBo đã có danh sách chọn.' + (warn ? '\n\n' + warn : ''), SpreadsheetApp.getUi().ButtonSet.OK);
}
/** Tạo sheet PhongBan nếu chưa có; luôn gắn lại định dạng + danh sách chọn. Trả về true nếu vừa tạo mới. */
function caiDatPhongBan_(activate) {
  const ss = SpreadsheetApp.getActive();
  let pb = ss.getSheetByName(SH.PHONGBAN), created = false;
  if (!pb) {
    const vt = ss.getSheetByName(SH.VAITRO);
    pb = ss.insertSheet(SH.PHONGBAN, vt ? vt.getIndex() : ss.getNumSheets());
    const heads = [PB_COL.MA, PB_COL.TEN, PB_COL.LOAI, PB_COL.ON, PB_COL.GHICHU];
    pb.getRange(1, 1, 1, heads.length).setValues([heads]);
    pb.getRange(2, 1, PHONGBAN_BAN_DAU.length, 1).setNumberFormat('@');
    pb.getRange(2, 1, PHONGBAN_BAN_DAU.length, 4).setValues(PHONGBAN_BAN_DAU.map(r => [r[0], r[1], r[2], true]));
    [12, 22, 10, 11, 36].forEach((w, i) => pb.setColumnWidth(i + 1, w * 8));
    created = true;
    const hd = ss.getSheetByName('HuongDan');
    if (hd) hd.getRange(hd.getLastRow() + 2, 1, 4, 1).setValues([['PHÒNG/BAN (sheet PhongBan)'],
      ['- Cột Phòng/Ban ở CanBo chọn từ danh sách; tên trùng với tên phòng trong sổ công tác.'],
      ['- Đổi tên phòng/ban: sửa trực tiếp ở sheet PhongBan, tên mới tự cập nhật cho mọi cán bộ của phòng đó. Không đổi Mã phòng/ban.'],
      ['- Phòng/ban không còn dùng: bỏ tick "Đang dùng" (không xóa dòng). Menu Kiểm tra dữ liệu sẽ báo cán bộ còn thuộc phòng đó.']]);
  }
  ensureRows_(pb, 50);
  const hp = headerMap_(pb);
  pb.getRange(2, hp[PB_COL.MA] + 1, pb.getMaxRows() - 1, 1).setNumberFormat('@');
  setList_(pb.getRange(2, hp[PB_COL.LOAI] + 1, pb.getMaxRows() - 1, 1), ['Ban', 'Phòng', 'PGD']);
  checkboxCol_(pb, hp[PB_COL.ON] + 1, pb.getMaxRows());
  pb.setFrozenRows(1);
  pb.getRange(1, 1, 1, pb.getLastColumn()).setFontWeight('bold').setFontColor('#FFFFFF').setBackground('#005993');

  const cb = ss.getSheetByName(SH.CANBO), h = headerMap_(cb);
  ensureRows_(cb, 200);
  const tenCol = colLetter_(hp[PB_COL.TEN] + 1);
  cb.getRange(2, h[COL.PB] + 1, cb.getMaxRows() - 1, 1).setDataValidation(
    SpreadsheetApp.newDataValidation().requireValueInRange(pb.getRange(tenCol + '2:' + tenCol), true).setAllowInvalid(false)
      .setHelpText('Chọn phòng/ban trong danh sách (sheet PhongBan).').build());
  if (created && activate) pb.activate();
  return created;
}
/** Danh mục phòng/ban: tên → {ma, loai, on} */
function phongBanMap_() {
  const pb = ss_().getSheetByName(SH.PHONGBAN); if (!pb || pb.getLastRow() < 2) return {};
  const v = pb.getDataRange().getValues(), hd = {}; v[0].forEach((x, i) => hd[String(x).trim()] = i);
  const out = {};
  v.slice(1).forEach(r => { const ten = String(r[hd[PB_COL.TEN]] || '').trim(); if (!ten) return;
    out[ten] = { ma: String(r[hd[PB_COL.MA]] || '').trim(), loai: String(r[hd[PB_COL.LOAI]] || '').trim(), on: r[hd[PB_COL.ON]] !== false }; });
  return out;
}
/** Liệt kê cán bộ có Phòng/Ban không có trong danh mục hoặc phòng đã ngưng dùng; mã trùng trong PhongBan */
function kiemTraPhongBan_() {
  const map = phongBanMap_(); if (!Object.keys(map).length) return '';
  const msg = [];
  const codes = Object.values(map).map(x => x.ma).filter(Boolean);
  const dup = codes.filter((c, i) => codes.indexOf(c) !== i);
  if (dup.length) msg.push('Mã phòng/ban bị trùng trong sheet PhongBan: ' + Array.from(new Set(dup)).join(', '));
  const cb = SpreadsheetApp.getActive().getSheetByName(SH.CANBO), h = headerMap_(cb), n = cb.getLastRow() - 1;
  if (n > 0) {
    const bad = cb.getRange(2, 1, n, cb.getLastColumn()).getValues()
      .map((r, i) => ({ row: i + 2, id: r[h[COL.ID]], pb: String(r[h[COL.PB]]).trim() }))
      .filter(x => x.pb && (!map[x.pb] || !map[x.pb].on));
    if (bad.length) msg.push('LƯU Ý: ' + bad.length + ' cán bộ có Phòng/Ban không có trong danh mục hoặc đã ngưng dùng, cần chọn lại:\n' +
      bad.slice(0, 10).map(x => '- Dòng ' + x.row + ' (ID ' + x.id + '): "' + x.pb + '"' + (map[x.pb] ? ' (ngưng dùng)' : '')).join('\n') + (bad.length > 10 ? '\n…' : ''));
  }
  return msg.join('\n\n');
}

/** Thêm ứng dụng mới: mã tự tăng (I003, I004…), thêm cột tick ở CanBo, VaiTro và 1 dòng ở UngDung */
function themUngDung() {
  const ui = SpreadsheetApp.getUi(), ss = SpreadsheetApp.getActive();
  const ud = ss.getSheetByName(SH.UNGDUNG);
  const codes = ud.getRange(2, 1, Math.max(ud.getLastRow() - 1, 1), 1).getValues().map(r => String(r[0])).filter(x => APP_RE.test(x));
  const next = 'I' + String(Math.max(0, ...codes.map(c => +c.slice(1))) + 1).padStart(3, '0');
  const r = ui.prompt('Thêm ứng dụng mới', 'Mã sẽ cấp: ' + next + '\nNhập tên ứng dụng:', ui.ButtonSet.OK_CANCEL);
  if (r.getSelectedButton() !== ui.Button.OK || !r.getResponseText().trim()) return;
  addApp_(next, r.getResponseText().trim(), '');
  lamMoiBoNhoDem(true);
  ui.alert('Đã thêm ' + next + '. Tick quyền ở sheet VaiTro hoặc từng cán bộ. Nhớ khai báo cùng mã ' + next + ' trong TOOLKIT của index.html.');
}

/** Đăng ký 1 ứng dụng: dòng ở UngDung + cột tick ở CanBo, VaiTro (bỏ qua nếu mã đã có) */
function addApp_(code, name, desc) {
  const ss = SpreadsheetApp.getActive(), ud = ss.getSheetByName(SH.UNGDUNG);
  const exists = ud.getRange(2, 1, Math.max(ud.getMaxRows() - 1, 1), 1).getValues().some(r => String(r[0]).trim() === code);
  if (!exists) {
    const row = firstEmptyRow_(ud, 1);
    ud.getRange(row, 1, 1, 6).setValues([[code, name, desc, true, false, Utilities.formatDate(new Date(), TZ, 'dd/MM/yyyy')]]);
    checkboxCol_(ud, 4, ud.getMaxRows()); checkboxCol_(ud, 5, ud.getMaxRows());
  }
  [SH.CANBO, SH.VAITRO].forEach(n => {
    const s = ss.getSheetByName(n);
    if (headerMap_(s)[code] !== undefined) return;
    const c = s.getLastColumn() + 1;
    if (c > s.getMaxColumns()) s.insertColumnAfter(s.getMaxColumns());
    s.getRange(1, c).setValue(code).setFontWeight('bold').setFontColor('#FFFFFF').setBackground('#005993');
    checkboxCol_(s, c, s.getMaxRows());
    s.setColumnWidth(c, 70);
  });
}
/** Dòng trống đầu tiên sau dữ liệu của 1 cột (không bị ô tick FALSE làm sai) */
function firstEmptyRow_(s, col) {
  const v = s.getRange(1, col, s.getMaxRows(), 1).getValues();
  for (let i = v.length - 1; i >= 0; i--) if (String(v[i][0]).trim() !== '') { const r = i + 2; if (r > s.getMaxRows()) s.insertRowsAfter(s.getMaxRows(), 20); return r; }
  return 2;
}

/** Sinh thêm ID ngẫu nhiên đầu 9, trạng thái Chưa cấp */
function sinhThemID() {
  const ui = SpreadsheetApp.getUi();
  const r = ui.prompt('Sinh thêm ID dự phòng', 'Số lượng ID cần thêm (1–50):', ui.ButtonSet.OK_CANCEL);
  if (r.getSelectedButton() !== ui.Button.OK) return;
  const n = Math.min(50, Math.max(1, parseInt(r.getResponseText(), 10) || 0));
  const cb = SpreadsheetApp.getActive().getSheetByName(SH.CANBO), h = headerMap_(cb);
  const existing = cb.getRange(2, h[COL.ID] + 1, Math.max(cb.getLastRow() - 1, 1), 1).getValues().map(x => String(x[0]));
  const ids = newIds_(n, existing);
  const width = cb.getLastColumn();
  const rows = ids.map(id => { const a = new Array(width).fill(''); a[h[COL.ID]] = id; a[h[COL.TT]] = TT.NEW; appCols_(cb).forEach(c => a[c] = false); return a; });
  const start = firstEmptyRow_(cb, h[COL.ID] + 1);
  ensureRows_(cb, start + rows.length);
  cb.getRange(start, 1, rows.length, width).setValues(rows);
  ui.alert('Đã thêm ' + ids.length + ' ID mới ở cuối sheet CanBo.');
}

/** Chép lại quyền vai trò cho mọi cán bộ (xóa các ngoại lệ tick riêng) */
function apDungLaiVaiTro() {
  const ui = SpreadsheetApp.getUi();
  if (ui.alert('Áp dụng lại quyền theo vai trò', 'Mọi ô tick ứng dụng ở CanBo sẽ được đặt lại theo sheet VaiTro (mất các ngoại lệ tick riêng). Tiếp tục?', ui.ButtonSet.YES_NO) !== ui.Button.YES) return;
  const cb = SpreadsheetApp.getActive().getSheetByName(SH.CANBO), h = headerMap_(cb), roles = rolePerms_();
  const n = cb.getLastRow() - 1; if (n < 1) return;
  const data = cb.getRange(2, 1, n, cb.getLastColumn()).getValues(), heads = cb.getRange(1, 1, 1, cb.getLastColumn()).getValues()[0];
  appCols_(cb).forEach(c => {
    const vals = data.map(r => { const p = roles[String(r[h[COL.VT]]).trim()]; return [p ? !!p[heads[c]] : r[c]]; });
    cb.getRange(2, c + 1, n, 1).setValues(vals);
  });
  lamMoiBoNhoDem(true);
  ui.alert('Đã áp dụng lại quyền theo vai trò.');
}

function xemThongKe() {
  const p = PropertiesService.getScriptProperties();
  const today = Utilities.formatDate(new Date(), TZ, 'yyyyMMdd');
  const d = Number(p.getProperty('stat_d_' + today) || 0), t = Number(p.getProperty('stat_total') || 0);
  const users = Object.values(loadData_().users).filter(u => u.tt === TT.ON).length;
  SpreadsheetApp.getUi().alert('Thống kê Không Gian Số Nội Bộ',
    'Lượt truy cập hôm nay: ' + fmt_(d) + '\nTổng lượt truy cập: ' + fmt_(t) + '\nCán bộ đang hoạt động: ' + users, SpreadsheetApp.getUi().ButtonSet.OK);
}

/** Kiểm tra nhanh: vai trò không hợp lệ, cán bộ Hoạt động thiếu họ tên/giới tính/vai trò */
function kiemTraDuLieu() {
  const cb = SpreadsheetApp.getActive().getSheetByName(SH.CANBO), h = headerMap_(cb), n = cb.getLastRow() - 1, msg = [];
  const w = kiemTraVaiTro_(); if (w) msg.push(w);
  const wp = kiemTraPhongBan_(); if (wp) msg.push(wp);
  if (n > 0) {
    const miss = cb.getRange(2, 1, n, cb.getLastColumn()).getValues().map((r, i) => ({ row: i + 2, r }))
      .filter(x => x.r[h[COL.TT]] === TT.ON && (!String(x.r[h[COL.TEN]]).trim() || !x.r[h[COL.GT]] || !x.r[h[COL.VT]]));
    if (miss.length) msg.push('Cán bộ "Hoạt động" còn thiếu Họ tên / Giới tính / Vai trò: dòng ' + miss.map(x => x.row).join(', '));
  }
  SpreadsheetApp.getUi().alert('Kiểm tra dữ liệu', msg.length ? msg.join('\n\n') : 'Dữ liệu hợp lệ.', SpreadsheetApp.getUi().ButtonSet.OK);
}

function lamMoiBoNhoDem(silent) {
  CacheService.getScriptCache().remove(CACHE_KEY);
  if (silent !== true) SpreadsheetApp.getActive().toast('Ứng dụng sẽ dùng dữ liệu mới ngay lần mở kế tiếp.', 'Không Gian Số');
}

/* =====================================================================
   PHẦN 2 – WEB APP (API cho trang Không Gian Số)
   ===================================================================== */
function doGet() { return json_({ ok: true, service: 'ttg-inside' }); }

function doPost(e) {
  let body = {};
  try { body = JSON.parse((e && e.postData && e.postData.contents) || '{}'); } catch (err) { return json_({ error: 'bad_request' }); }
  try {
    switch (body.action) {
      case 'login':  return json_(apiLogin_(body));
      case 'verify': return json_(apiVerify_(body));
      case 'open':   return json_(apiOpen_(body));
      case 'ai':     return json_(apiAi_(body));
      case 'cp_load':   return json_(apiCpLoad_(body));
      case 'cp_add':    return json_(apiCpAdd_(body));
      case 'cp_cancel': return json_(apiCpCancel_(body));
      case 'cp_report': return json_(apiCpReport_(body));
      case 'cp_ai':     return json_(apiCpAi_(body));
      case 'bt_list':   return json_(apiBtList_(body));
      case 'bt_get':    return json_(apiBtGet_(body));
      case 'ta_load':   return json_(apiTaLoad_(body));
      case 'ta_add':    return json_(apiTaAdd_(body));
      case 'ta_return': return json_(apiTaReturn_(body));
      case 'ta_undo':   return json_(apiTaUndo_(body));
      case 'ta_phieu_email': return json_(apiTaPhieuEmail_(body));
      case 'ta_cancel': return json_(apiTaCancel_(body));
      default:       return json_({ error: 'bad_request' });
    }
  } catch (err) {
    console.error(err);
    return json_({ error: 'server', message: 'Máy chủ đang bận, thử lại sau ít phút.' });
  }
}

function apiLogin_(b) {
  const id = String(b.id || '').trim(), dev = String(b.device || '').slice(0, 40) || 'khong-ro';
  const cache = CacheService.getScriptCache(), props = PropertiesService.getScriptProperties();

  const susp = Number(props.getProperty('LOGIN_SUSPENDED_UNTIL') || 0);
  if (Date.now() < susp) return { error: 'suspended', minutes: Math.ceil((susp - Date.now()) / 60000) };

  const devKey = 'fail_' + dev, devFails = Number(cache.get(devKey) || 0);
  if (devFails >= DEVICE_MAX_FAIL) return { error: 'device_locked', minutes: DEVICE_LOCK_MIN };

  const u = /^9\d{4}$/.test(id) ? loadData_().users[id] : null;
  if (!u || u.tt === TT.NEW) {
    const n = devFails + 1;
    cache.put(devKey, String(n), DEVICE_LOCK_MIN * 60);
    log_(id, '', 'Sai ID', '', 'Lần sai thứ ' + n + ' trên thiết bị', dev);
    globalFail_(props);
    if (n >= DEVICE_MAX_FAIL) { log_(id, '', 'Khóa thiết bị', '', DEVICE_LOCK_MIN + ' phút', dev); return { error: 'device_locked', minutes: DEVICE_LOCK_MIN }; }
    return { error: 'wrong', remaining: DEVICE_MAX_FAIL - n };
  }
  if (u.tt !== TT.ON) { log_(id, u.hoTen, 'Bị từ chối', '', 'ID đang ở trạng thái ' + u.tt, dev); return { error: 'locked_user' }; }

  cache.remove(devKey);
  touchLastLogin_(u.row);
  log_(id, u.hoTen, 'Đăng nhập', '', '', dev);
  return Object.assign({ ok: true, token: makeToken_(id) }, session_(u));
}

function apiVerify_(b) {
  const v = checkToken_(b.token); if (v.error) return v;
  return Object.assign({ ok: true }, session_(v.user));
}

/** Ghi nhật ký khi mở 1 ứng dụng; trả lỗi nếu không có quyền */
function apiOpen_(b) {
  const v = checkToken_(b.token); if (v.error) return v;
  const app = String(b.app || '');
  if (!userApps_(v.user).includes(app)) { log_(v.user.id, v.user.hoTen, 'Bị từ chối', app, 'Chưa được cấp quyền', ''); return { error: 'forbidden' }; }
  return { ok: true };
}

function session_(u) {
  return { user: { ten: firstName_(u.hoTen), xungHo: u.gt === 'Nữ' ? 'chị' : u.gt === 'Nam' ? 'anh' : 'anh/chị', phongBan: u.pb, maPhongBan: u.maPB || '', loaiPhongBan: u.loaiPB || '', vaiTro: u.vt },
           apps: taFilterApps_(u, userApps_(u)), stats: countVisit_(), ai: !!PropertiesService.getScriptProperties().getProperty('ANTHROPIC_API_KEY') && userApps_(u).includes('I002'),
           expires: endOfDayIso_() };
}

/* ---------- Phiên đăng nhập: có hiệu lực đến hết ngày (giờ Việt Nam) ---------- */
function makeToken_(id) {
  const payload = id + '.' + today_() + '.' + Utilities.getUuid().slice(0, 8);
  return payload + '.' + sign_(payload);
}
function checkToken_(tok) {
  const parts = String(tok || '').split('.');
  if (parts.length !== 4) return { error: 'expired' };
  const payload = parts.slice(0, 3).join('.');
  if (sign_(payload) !== parts[3]) return { error: 'expired' };
  if (parts[1] !== today_()) return { error: 'expired' };
  const u = loadData_().users[parts[0]];
  if (!u || u.tt !== TT.ON) return { error: 'revoked' };
  return { user: u };
}
function sign_(s) {
  const key = PropertiesService.getScriptProperties().getProperty('TOKEN_SECRET');
  if (!key) throw new Error('Chưa chạy caiDatBanDau');
  return Utilities.base64EncodeWebSafe(Utilities.computeHmacSha256Signature(s, key)).replace(/=+$/, '');
}

/* ---------- Đọc dữ liệu phân quyền (có bộ nhớ đệm 2 phút) ---------- */
function loadData_() {
  const cache = CacheService.getScriptCache(), hit = cache.get(CACHE_KEY);
  if (hit) return JSON.parse(hit);
  const ss = ss_();
  const ud = ss.getSheetByName(SH.UNGDUNG).getDataRange().getValues().slice(1)
    .filter(r => APP_RE.test(String(r[0]))).map(r => ({ id: String(r[0]), on: r[3] === true, all: r[4] === true }));
  const cb = ss.getSheetByName(SH.CANBO), vals = cb.getDataRange().getValues(), heads = vals[0].map(String);
  const h = {}; heads.forEach((x, i) => h[x.trim()] = i);
  const appIdx = heads.map((x, i) => APP_RE.test(x.trim()) ? i : -1).filter(i => i >= 0);
  const users = {};
  vals.slice(1).forEach((r, i) => {
    const id = String(r[h[COL.ID]]).trim(); if (!/^9\d{4}$/.test(id)) return;
    users[id] = { id, row: i + 2, hoTen: String(r[h[COL.TEN]] || '').trim(), gt: String(r[h[COL.GT]] || '').trim(), pb: String(r[h[COL.PB]] || '').trim(),
                  vt: String(r[h[COL.VT]] || '').trim(), tt: String(r[h[COL.TT]] || '').trim(),
                  perm: appIdx.filter(c => r[c] === true).map(c => heads[c].trim()) };
  });
  const pbMap = phongBanMap_();
  Object.values(users).forEach(u => { const p = pbMap[u.pb]; u.maPB = p ? p.ma : ''; u.loaiPB = p ? p.loai : ''; });
  const data = { apps: ud, users };
  try { cache.put(CACHE_KEY, JSON.stringify(data), CACHE_SEC); } catch (e) {}
  return data;
}
function userApps_(u) {
  return loadData_().apps.filter(a => a.on && (a.all || u.perm.includes(a.id))).map(a => a.id);
}
function rolePerms_() {
  const vt = ss_().getSheetByName(SH.VAITRO), v = vt.getDataRange().getValues(), heads = v[0].map(x => String(x).trim());
  const out = {};
  v.slice(1).forEach(r => { const name = String(r[0]).trim(); if (!name) return; out[name] = {}; heads.forEach((hd, i) => { if (APP_RE.test(hd)) out[name][hd] = r[i] === true; }); });
  return out;
}

/* ---------- Thống kê lượt truy cập ---------- */
function countVisit_() {
  const lock = LockService.getScriptLock();
  try { lock.waitLock(3000); } catch (e) { return readStats_(); }
  try {
    const p = PropertiesService.getScriptProperties(), dk = 'stat_d_' + today_();
    const t = Number(p.getProperty('stat_total') || 0) + 1, d = Number(p.getProperty(dk) || 0) + 1;
    p.setProperties({ stat_total: String(t), [dk]: String(d) });
    return { total: t, today: d };
  } finally { lock.releaseLock(); }
}
function readStats_() {
  const p = PropertiesService.getScriptProperties();
  return { total: Number(p.getProperty('stat_total') || 0), today: Number(p.getProperty('stat_d_' + today_()) || 0) };
}
/** Dọn bộ đếm ngày cũ (tùy chọn: đặt trình kích hoạt chạy hằng tháng) */
function donDepBoDem() {
  const p = PropertiesService.getScriptProperties(), keep = 'stat_d_' + today_();
  Object.keys(p.getProperties()).forEach(k => { if (/^stat_d_\d{8}$/.test(k) && k !== keep) p.deleteProperty(k); });
}

/* ---------- Chặn đoán thử toàn hệ thống ---------- */
function globalFail_(props) {
  const lock = LockService.getScriptLock();
  try { lock.waitLock(3000); } catch (e) { return; }
  try {
    const now = Date.now(), win = GLOBAL_WINDOW_MIN * 60000;
    const arr = JSON.parse(props.getProperty('FAIL_TIMES') || '[]').filter(t => now - t < win);
    arr.push(now);
    props.setProperty('FAIL_TIMES', JSON.stringify(arr.slice(-60)));
    if (arr.length > GLOBAL_MAX_FAIL) {
      props.setProperty('LOGIN_SUSPENDED_UNTIL', String(now + GLOBAL_SUSPEND_MIN * 60000));
      props.setProperty('FAIL_TIMES', '[]');
      log_('', '', 'Tạm ngưng đăng nhập', '', 'Quá ' + GLOBAL_MAX_FAIL + ' lần sai trong ' + GLOBAL_WINDOW_MIN + ' phút', '');
      const to = props.getProperty('ADMIN_EMAIL');
      if (to) MailApp.sendEmail(to, '[Không Gian Số] Tạm ngưng đăng nhập do nhập sai ID nhiều lần',
        'Hệ thống ghi nhận hơn ' + GLOBAL_MAX_FAIL + ' lần nhập sai ID trong ' + GLOBAL_WINDOW_MIN + ' phút.\n' +
        'Đăng nhập tạm ngưng ' + GLOBAL_SUSPEND_MIN + ' phút. Xem chi tiết tại sheet NhatKy.\n\nThời điểm: ' + Utilities.formatDate(new Date(), TZ, 'dd/MM/yyyy HH:mm'));
    }
  } finally { lock.releaseLock(); }
}

/* =====================================================================
   PHẦN 3 – TRỢ LÝ AI (giai đoạn 2; tắt khi chưa có ANTHROPIC_API_KEY)
   ===================================================================== */
const DEFAULT_MODEL = 'claude-haiku-5-5';
// Giữ đồng bộ với hằng RULES trong index.html
const SYSTEM_PROMPT = `Bạn là trợ lý tạo mã QR cho cán bộ VietinBank Chi nhánh Tây Tiền Giang (các điểm: Trụ sở Chi nhánh, PGD Cái Bè, An Hữu, Long Định, Hưng Long, Ấp Bắc).
Nhiệm vụ: đọc yêu cầu và soạn chính xác chuỗi sẽ mã hóa vào QR.
CHỈ trả về một đối tượng JSON, không thêm chữ nào khác:
{"reply":"câu trả lời tiếng Việt ngắn gọn cho cán bộ","qr":null hoặc {"kind":"url|text|tel|sms|email|wifi|vcard|geo|event","title":"tên ngắn của mã","payload":"chuỗi chính xác để mã hóa","caption":"dòng ngắn in dưới mã cho khách đọc"}}
Quy tắc:
- Không bịa link, số điện thoại, email, địa chỉ, tọa độ. Thiếu thông tin thì qr=null và hỏi lại đúng 1 câu.
- Thông tin công khai được dùng: tổng đài 1900 558 868, website https://www.vietinbank.vn.
- Định dạng payload: url đầy đủ https://; tel:+84xxxxxxxxx (bỏ số 0 đầu); SMSTO:số:nội dung; mailto:dia@chi?subject=..&body=.. (mã hóa URL); WIFI:T:WPA;S:tên;P:mật khẩu;; (escape \\ ; , : " bằng \\); vCard 3.0 (BEGIN:VCARD/VERSION:3.0/N/FN/ORG/TITLE/TEL/EMAIL/ADR/URL/END:VCARD, xuống dòng \\n); geo:vĩ độ,kinh độ hoặc link Google Maps cán bộ đưa; sự kiện BEGIN:VEVENT/SUMMARY/DTSTART:yyyymmddThhmmss/DTEND/LOCATION/END:VEVENT; còn lại là text.
- Payload tối đa 900 ký tự.
- Không tạo QR chuyển khoản/VietQR/thanh toán: trả qr=null, hướng dẫn dùng mã VietQR tạo từ VietinBank iPay hoặc hệ thống chính thức.
- Nếu yêu cầu chứa thông tin khách hàng, số tài khoản, CCCD: nhắc cán bộ không đưa dữ liệu này lên mã QR công khai; không tạo mã.
- Từ chối các yêu cầu không liên quan đến tạo mã QR.
- Khi cán bộ yêu cầu sửa mã vừa tạo, trả mã mới đầy đủ.`;

function apiAi_(b) {
  const v = checkToken_(b.token); if (v.error) return v;
  if (!userApps_(v.user).includes('I002')) return { error: 'forbidden' };
  const props = PropertiesService.getScriptProperties(), key = props.getProperty('ANTHROPIC_API_KEY');
  if (!key) return { error: 'not_configured' };
  const msgs = sanitize_(b.messages); if (!msgs.length) return { error: 'bad_request' };
  const lim = aiLimit_(props); if (lim) return lim;
  const res = UrlFetchApp.fetch('https://api.anthropic.com/v1/messages', {
    method: 'post', contentType: 'application/json', muteHttpExceptions: true,
    headers: { 'x-api-key': key, 'anthropic-version': '2023-06-01' },
    payload: JSON.stringify({ model: props.getProperty('MODEL') || DEFAULT_MODEL, max_tokens: 1200, system: SYSTEM_PROMPT, messages: msgs })
  });
  const code = res.getResponseCode();
  if (code === 429 || code === 529) return { error: 'rate_limited' };
  if (code !== 200) { console.error('Claude API ' + code + ': ' + res.getContentText().slice(0, 300)); return { error: 'upstream_error' }; }
  const text = (JSON.parse(res.getContentText()).content || []).filter(c => c.type === 'text').map(c => c.text).join('');
  return parseJson_(text) || { error: 'invalid_json' };
}
function sanitize_(arr) {
  if (!Array.isArray(arr)) return [];
  const merged = [];
  arr.filter(m => m && (m.role === 'user' || m.role === 'assistant') && typeof m.content === 'string' && m.content.trim())
     .slice(-12).forEach(m => { const c = m.content.slice(0, 2000), l = merged[merged.length - 1];
       if (l && l.role === m.role) l.content += '\n' + c; else merged.push({ role: m.role, content: c }); });
  while (merged.length && merged[0].role !== 'user') merged.shift();
  return merged.length && merged[merged.length - 1].role === 'user' ? merged : [];
}
function aiLimit_(props) {
  const lock = LockService.getScriptLock(); lock.waitLock(5000);
  try {
    const cache = CacheService.getScriptCache(), mk = 'ai_m_' + Utilities.formatDate(new Date(), TZ, 'yyyyMMddHHmm'), dk = 'ai_d_' + today_();
    const m = Number(cache.get(mk) || 0), d = Number(props.getProperty(dk) || 0);
    if (m >= Number(props.getProperty('PER_MINUTE_LIMIT') || 20)) return { error: 'rate_limited' };
    if (d >= Number(props.getProperty('DAILY_LIMIT') || 300)) return { error: 'daily_limit' };
    cache.put(mk, String(m + 1), 120); props.setProperty(dk, String(d + 1));
    return null;
  } finally { lock.releaseLock(); }
}
function parseJson_(text) {
  const tryP = s => { try { return JSON.parse(s); } catch (e) { return null; } };
  let v = tryP(text.trim());
  if (!v) { const f = text.match(/```(?:json)?\s*([\s\S]*?)```/); if (f) v = tryP(f[1]); }
  if (!v) { const a = text.indexOf('{'), z = text.lastIndexOf('}'); if (a >= 0 && z > a) v = tryP(text.slice(a, z + 1)); }
  return v && typeof v === 'object' ? { reply: String(v.reply || 'Đã xử lý.'), qr: v.qr && v.qr.payload ? v.qr : null } : null;
}

/* =====================================================================
   TIỆN ÍCH
   ===================================================================== */
/** File Sheet phân quyền: file chứa script; dự phòng theo SHEET_ID lưu khi cài đặt */
function ss_() {
  let ss = null; try { ss = SpreadsheetApp.getActiveSpreadsheet(); } catch (e) {}
  return ss || SpreadsheetApp.openById(PropertiesService.getScriptProperties().getProperty('SHEET_ID'));
}
function json_(o) { return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON); }
function today_() { return Utilities.formatDate(new Date(), TZ, 'yyyyMMdd'); }
function endOfDayIso_() { return Utilities.formatDate(new Date(), TZ, "yyyy-MM-dd'T'23:59:59+07:00"); }
function fmt_(n) { return String(n).replace(/\B(?=(\d{3})+(?!\d))/g, '.'); }
/** Tên gọi = chữ cuối của Họ và tên (Nguyễn Văn Tuyến → Tuyến) */
function firstName_(full) { const p = String(full || '').trim().split(/\s+/); return p[p.length - 1] || ''; }
function headerMap_(s) { const h = {}; s.getRange(1, 1, 1, s.getLastColumn()).getValues()[0].forEach((x, i) => h[String(x).trim()] = i); return h; }
function headerAt_(s, c0) { return String(s.getRange(1, c0 + 1).getValue()).trim(); }
function appCols_(s) { return s.getRange(1, 1, 1, s.getLastColumn()).getValues()[0].map((x, i) => APP_RE.test(String(x).trim()) ? i : -1).filter(i => i >= 0); }
/** Bảo đảm sheet có đủ số dòng (file chuyển từ Excel có thể chỉ có đúng số dòng dữ liệu) */
function ensureRows_(s, n) { const m = s.getMaxRows(); if (m < n) s.insertRowsAfter(m, n - m); }
/** Liệt kê cán bộ có Vai trò không có trong sheet VaiTro */
function kiemTraVaiTro_() {
  const ss = SpreadsheetApp.getActive(), cb = ss.getSheetByName(SH.CANBO), h = headerMap_(cb), roles = rolePerms_();
  const n = cb.getLastRow() - 1; if (n < 1) return '';
  const bad = cb.getRange(2, 1, n, cb.getLastColumn()).getValues()
    .map((r, i) => ({ row: i + 2, id: r[h[COL.ID]], vt: String(r[h[COL.VT]]).trim() }))
    .filter(x => x.vt && !roles[x.vt]);
  if (!bad.length) return '';
  return 'LƯU Ý: ' + bad.length + ' dòng ở sheet CanBo có Vai trò không còn trong sheet VaiTro, cần chọn lại:\n' +
    bad.slice(0, 10).map(x => '- Dòng ' + x.row + ' (ID ' + x.id + '): "' + x.vt + '"').join('\n') + (bad.length > 10 ? '\n…' : '');
}
function checkboxCol_(s, col1, lastRow) { s.getRange(2, col1, lastRow - 1, 1).insertCheckboxes(); }
function setList_(rng, list) { rng.setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(list, true).setAllowInvalid(false).build()); }
function colLetter_(n) { let s = ''; while (n > 0) { const m = (n - 1) % 26; s = String.fromCharCode(65 + m) + s; n = (n - m - 1) / 26; } return s; }
function touchLastLogin_(row) {
  try { const cb = ss_().getSheetByName(SH.CANBO), h = headerMap_(cb); cb.getRange(row, h[COL.LAST] + 1).setValue(new Date()); } catch (e) {}
}
function log_(id, ten, ev, app, detail, dev) {
  try { ss_().getSheetByName(SH.NHATKY).appendRow([Utilities.formatDate(new Date(), TZ, 'dd/MM/yyyy HH:mm:ss'), id, ten, ev, app, detail, String(dev || '').slice(0, 8)]); } catch (e) {}
}
/** ID 5 số đầu 9, 4 số sau ngẫu nhiên; loại dãy dễ đoán và ID chỉ khác ID khác 1 chữ số */
function newIds_(n, existing) {
  const bad = s => { const d = s.split('').map(Number), cnt = {}; s.split('').forEach(c => cnt[c] = (cnt[c] || 0) + 1);
    if (Math.max(...Object.values(cnt)) >= 3) return true;
    if ([0, 1, 2].every(i => d[i + 1] - d[i] === 1) || [0, 1, 2].every(i => d[i + 1] - d[i] === -1)) return true;
    if (s.slice(0, 2) === s.slice(2) || s === s.split('').reverse().join('') || s.endsWith('00') || s.startsWith('000')) return true;
    return false; };
  const near = (a, b) => a.split('').filter((c, i) => c !== b[i]).length <= 1;
  const all = existing.slice(), out = [];
  let guard = 0;
  while (out.length < n && guard++ < 20000) {
    const s = String(Math.floor(Math.random() * 10000)).padStart(4, '0'), id = '9' + s;
    if (bad(s) || all.some(x => near(x, id))) continue;
    all.push(id); out.push(id);
  }
  return out;
}

/* =====================================================================
   PHẦN 4 – NHẬT KÝ PHÊ DUYỆT CHI PHÍ (ứng dụng I003)
   Dữ liệu nằm ở file riêng KhongGianSoNoiBo_ChiPhi (id lưu ở thuộc tính CP_SHEET_ID).
   ===================================================================== */
const CP_APP = 'I003', CP_FILE_NAME = 'KhongGianSoNoiBo_ChiPhi';
const CP_SH = { DUYET: 'CP_NguoiDuyet', PHUTRACH: 'CP_PhuTrach', NHOM: 'CP_NhomChiPhi', CAUHINH: 'CP_CauHinh', NK: 'CP_NhatKy', KH: 'CP_KeHoach', HD: 'HuongDan' };
const CP_APP_BC = 'I004', CP_TEN = 'Quản lý chi phí', CP_TEN_BC = 'Báo cáo chi phí';
const CP_NK_HEAD = ['Mã', 'Thời gian nhập', 'Ngày phê duyệt', 'Tháng', 'Mã phòng/ban', 'Phòng/Ban', 'Nhóm chi phí', 'Nội dung', 'Số tiền',
  'Số tờ trình', 'ID người duyệt', 'Người duyệt', 'ID người nhập', 'Người nhập', 'Trạng thái', 'Lý do hủy', 'Người hủy', 'Thời gian hủy'];
const CP_ST = { OK: 'Hợp lệ', HUY: 'Đã hủy' };
const CP_NHOM_BAN_DAU = ['Văn phòng phẩm, in ấn', 'Tiếp khách, hội nghị', 'Quảng cáo, khuyến mại, sự kiện', 'Sửa chữa, bảo dưỡng tài sản',
  'Công cụ dụng cụ, mua sắm nhỏ', 'Điện, nước, viễn thông', 'Xăng dầu, vận chuyển', 'Công tác phí', 'Phúc lợi, hoạt động đoàn thể', 'Chi phí khác'];
const CP_CAUHINH_BAN_DAU = [
  ['VAI_TRO_NHAP', 'Trưởng phòng, Phó phòng', 'Vai trò được nhập khoản chi phí đã duyệt (cho phòng của mình)'],
  ['VAI_TRO_XEM_TAT_CA', 'Giám đốc, Quản trị hệ thống, Kế toán chi tiêu', 'Vai trò xem toàn bộ nhật ký và hạn mức'],
  ['PHONG_XEM_TAT_CA', 'Phòng DVKH', 'Lãnh đạo (vai trò được nhập) của các phòng này cũng xem toàn bộ (phòng kế toán)'],
  ['VAI_TRO_BAO_CAO_PHONG', 'Trưởng phòng', 'Vai trò xem Báo cáo chi phí của riêng phòng mình'],
  ['VAI_TRO_HUY', 'Giám đốc, Quản trị hệ thống', 'Vai trò được hủy khoản đã nhập bất kỳ lúc nào'],
  ['NGAY_NHAP_THANG_TRUOC', '5', 'Từ ngày 1 đến ngày này của tháng được nhập khoản phê duyệt tháng trước'],
  ['NGUONG_CANH_BAO', '20', 'Còn dưới % hạn mức này thì hiện màu cam']
];
// Vai trò được tick sẵn I003 ở sheet VaiTro khi cài đặt
const CP_VAI_TRO_MO = ['Giám đốc', 'Phó Giám đốc', 'Quản trị hệ thống', 'Trưởng phòng', 'Phó phòng', 'Kế toán chi tiêu'];
const CP_VAI_TRO_MO_BC = ['Giám đốc', 'Phó Giám đốc', 'Quản trị hệ thống', 'Trưởng phòng', 'Kế toán chi tiêu'];

/** Menu: tạo / cập nhật file chi phí, đăng ký I003 (Quản lý chi phí) và I004 (Báo cáo chi phí). Chạy lại an toàn. */
function caiDatChiPhi() {
  const ui = SpreadsheetApp.getUi(), props = PropertiesService.getScriptProperties();
  let cp = null;
  try { cp = props.getProperty('CP_SHEET_ID') ? SpreadsheetApp.openById(props.getProperty('CP_SHEET_ID')) : null; } catch (e) { cp = null; }
  const isNew = !cp;
  if (isNew) {
    cp = SpreadsheetApp.create(CP_FILE_NAME);
    cp.setSpreadsheetTimeZone(TZ);
    props.setProperty('CP_SHEET_ID', cp.getId());
    taoSheetChiPhi_(cp);
  }
  const changes = capNhatChiPhi_(cp);
  const n3 = dangKyUngDung_(CP_APP, CP_TEN, 'Hạn mức phê duyệt còn lại và nhật ký các khoản chi phí đã duyệt', CP_VAI_TRO_MO);
  const n4 = dangKyUngDung_(CP_APP_BC, CP_TEN_BC, 'Báo cáo chi phí theo phòng, nhóm, kế hoạch; nhận xét AI', CP_VAI_TRO_MO_BC);
  dongBoChiPhi_(cp);
  lamMoiBoNhoDem(true);
  const warn = kiemTraChiPhi_(cp);
  ui.alert(CP_TEN,
    (isNew ? 'Đã tạo file "' + CP_FILE_NAME + '" ở thư mục gốc Drive. Hãy kéo file vào thư mục KhongGianSoNoiBo (đường dẫn không đổi).\n\n' : 'File chi phí đã có, giữ nguyên dữ liệu.\n\n') +
    (changes.length ? 'Đã cập nhật: ' + changes.join('; ') + '.\n\n' : '') +
    (n3 ? 'Đã đăng ký ' + CP_APP + ' – ' + CP_TEN + ', tick quyền: ' + CP_VAI_TRO_MO.join(', ') + '.\n' : '') +
    (n4 ? 'Đã đăng ký ' + CP_APP_BC + ' – ' + CP_TEN_BC + ', tick quyền: ' + CP_VAI_TRO_MO_BC.join(', ') + '.\n' : '') +
    'Link file chi phí: ' + cp.getUrl() + (warn ? '\n\n' + warn : '') +
    '\n\nCuối cùng: Triển khai → Quản lý triển khai → Phiên bản mới.', ui.ButtonSet.OK);
}
/** Đăng ký ứng dụng; nếu cột quyền mới tạo thì tick theo danh sách vai trò rồi chép sang CanBo. Trả về true nếu mới. */
function dangKyUngDung_(code, name, desc, roles) {
  const ss = SpreadsheetApp.getActive(), vt = ss.getSheetByName(SH.VAITRO), isNew = headerMap_(vt)[code] === undefined;
  addApp_(code, name, desc);
  // Đồng bộ tên hiển thị trong UngDung
  const ud = ss.getSheetByName(SH.UNGDUNG), v = ud.getRange(2, 1, Math.max(ud.getMaxRows() - 1, 1), 2).getValues();
  const i = v.findIndex(r => String(r[0]).trim() === code); if (i >= 0 && v[i][1] !== name) ud.getRange(i + 2, 2, 1, 2).setValues([[name, desc]]);
  if (!isNew) return false;
  const hv = headerMap_(vt), nv = Math.max(firstEmptyRow_(vt, 1) - 2, 1);
  vt.getRange(2, hv[code] + 1, nv, 1).setValues(vt.getRange(2, 1, nv, 1).getValues().map(r => [roles.includes(String(r[0]).trim())]));
  const cb = ss.getSheetByName(SH.CANBO), h = headerMap_(cb), rp = rolePerms_(), n = Math.max(firstEmptyRow_(cb, h[COL.ID] + 1) - 2, 1);
  const vals = cb.getRange(2, h[COL.VT] + 1, n, 1).getValues(), tgt = cb.getRange(2, h[code] + 1, n, 1), cur = tgt.getValues();
  tgt.setValues(vals.map((r, k) => { const p = rp[String(r[0]).trim()]; return [p ? !!p[code] : cur[k][0] === true]; }));
  return true;
}
/** Nâng cấp file chi phí đã tạo bằng bản cũ: sheet kế hoạch, tham số mới, phòng kế toán = Phòng DVKH */
function capNhatChiPhi_(cp) {
  const out = [], ch = cp.getSheetByName(CP_SH.CAUHINH);
  if (ch) {
    const v = ch.getDataRange().getValues();
    const idx = k => v.findIndex(r => String(r[0]).trim() === k);
    const i = idx('PHONG_XEM_TAT_CA');
    if (i > 0 && String(v[i][1]).trim() === 'Phòng TCTH') { ch.getRange(i + 1, 2, 1, 2).setValues([[CP_CAUHINH_BAN_DAU[2][1], CP_CAUHINH_BAN_DAU[2][2]]]); out.push('phòng kế toán xem toàn bộ: Phòng DVKH'); }
    CP_CAUHINH_BAN_DAU.forEach(r => { if (idx(r[0]) < 0) { ch.appendRow(r); out.push('thêm tham số ' + r[0]); } });
  }
  if (!cp.getSheetByName(CP_SH.KH)) { taoSheetKeHoach_(cp); out.push('tạo sheet ' + CP_SH.KH); }
  CacheService.getScriptCache().remove('cp_cfg');
  return out;
}
/** Sheet kế hoạch chi phí năm: mỗi dòng 1 phòng (hoặc Toàn chi nhánh), nhóm để trống = tất cả nhóm */
function taoSheetKeHoach_(cp) {
  const kh = cp.insertSheet(CP_SH.KH), y = +Utilities.formatDate(new Date(), TZ, 'yyyy');
  kh.getRange(1, 1, 1, 5).setValues([['Năm', 'Phòng/Ban', 'Nhóm chi phí', 'Kế hoạch năm (đ)', 'Ghi chú']]);
  const pbs = ['Toàn chi nhánh'].concat(Object.keys(phongBanMap_()).filter(p => p !== 'Ban Giám đốc'));
  kh.getRange(2, 1, pbs.length, 5).setValues(pbs.map(p => [y, p, '', '', p === 'Toàn chi nhánh' ? 'Kế hoạch chi phí toàn chi nhánh' : '']));
  kh.getRange('D2:D').setNumberFormat('#,##0'); kh.getRange('A2:A').setNumberFormat('0');
  kh.setFrozenRows(1); kh.getRange(1, 1, 1, 5).setFontWeight('bold').setFontColor('#FFFFFF').setBackground('#005993');
  [70, 150, 220, 150, 260].forEach((w, i) => kh.setColumnWidth(i + 1, w));
  const hd = cp.getSheetByName(CP_SH.HD);
  if (hd) hd.getRange(hd.getLastRow() + 1, 1, 2, 1).setValues([['CP_KeHoach: kế hoạch chi phí năm. Mỗi dòng: Năm, Phòng/Ban (hoặc "Toàn chi nhánh"), Nhóm chi phí (để trống = tất cả nhóm), số tiền kế hoạch năm. Chưa có kế hoạch thì để trống.'],
    ['Báo cáo chi phí (I004) tự so sánh thực hiện lũy kế với kế hoạch năm và với tiến độ thời gian.']]);
}

function taoSheetChiPhi_(cp) {
  const style = (s, w) => { s.setFrozenRows(1); s.getRange(1, 1, 1, s.getLastColumn()).setFontWeight('bold').setFontColor('#FFFFFF').setBackground('#005993').setWrap(true);
    (w || []).forEach((x, i) => s.setColumnWidth(i + 1, x)); };
  const users = Object.values(loadData_().users);
  const idOf = name => (users.find(u => u.hoTen === name) || {}).id || '';
  const gd = idOf('Nguyễn Tiến Vĩnh'), linh = idOf('Phạm Minh Linh'), phuong = idOf('Nguyễn Kim Phương');

  const nk = cp.getSheets()[0]; nk.setName(CP_SH.NK);
  nk.getRange(1, 1, 1, CP_NK_HEAD.length).setValues([CP_NK_HEAD]);
  style(nk, [130, 130, 100, 70, 90, 120, 160, 260, 110, 100, 80, 140, 80, 140, 80, 180, 140, 130]);
  nk.getRange('B2:B').setNumberFormat('dd/MM/yyyy HH:mm'); nk.getRange('C2:C').setNumberFormat('dd/MM/yyyy');
  nk.getRange('D2:E').setNumberFormat('@'); nk.getRange('I2:I').setNumberFormat('#,##0'); nk.getRange('K2:K').setNumberFormat('@');
  nk.getRange('M2:M').setNumberFormat('@'); nk.getRange('R2:R').setNumberFormat('dd/MM/yyyy HH:mm');
  nk.setConditionalFormatRules([SpreadsheetApp.newConditionalFormatRule().whenFormulaSatisfied('=$O2="' + CP_ST.HUY + '"')
    .setFontColor('#9AA7B3').setStrikethrough(true).setRanges([nk.getRange('A2:R')]).build()]);
  nk.protect().setWarningOnly(true).setDescription('Dữ liệu do ứng dụng ghi. Không sửa tay; hủy khoản qua ứng dụng.');

  const du = cp.insertSheet(CP_SH.DUYET);
  du.getRange(1, 1, 1, 5).setValues([['ID người duyệt', 'Họ và tên', 'Hạn mức/tháng (đ)', 'Đang hiệu lực', 'Ghi chú']]);
  du.getRange(2, 1, 3, 5).setValues([[gd, 'Nguyễn Tiến Vĩnh', '', true, 'Giám đốc – không giới hạn hạn mức (để trống)'],
    [linh, 'Phạm Minh Linh', 30000000, true, 'Phó Giám đốc'], [phuong, 'Nguyễn Kim Phương', 30000000, true, 'Phó Giám đốc']]);
  du.getRange('A2:A').setNumberFormat('@'); du.getRange('C2:C').setNumberFormat('#,##0');
  style(du, [110, 170, 140, 100, 280]);

  const pt = cp.insertSheet(CP_SH.PHUTRACH);
  pt.getRange(1, 1, 1, 3).setValues([['Phòng/Ban', 'ID người duyệt', 'Ghi chú']]);
  const rows = [['Phòng DVKH', linh], ['Phòng HTTD', linh], ['PGD Long Định', linh], ['PGD An Hữu', linh],
    ['Phòng Bán lẻ', phuong], ['PGD Cái Bè', phuong], ['PGD Hưng Long', phuong], ['PGD Ấp Bắc', phuong],
    ['Phòng TCTH', gd], ['Phòng KHDN', gd]];
  pt.getRange(2, 1, rows.length, 3).setValues(rows.map(r => [r[0], r[1], '']));
  pt.getRange('B2:B').setNumberFormat('@');
  style(pt, [150, 120, 300]);

  const nh = cp.insertSheet(CP_SH.NHOM);
  nh.getRange(1, 1, 1, 3).setValues([['Nhóm chi phí', 'Đang dùng', 'Ghi chú']]);
  nh.getRange(2, 1, CP_NHOM_BAN_DAU.length, 2).setValues(CP_NHOM_BAN_DAU.map(x => [x, true]));
  style(nh, [240, 90, 280]);

  const ch = cp.insertSheet(CP_SH.CAUHINH);
  ch.getRange(1, 1, 1, 3).setValues([['Tham số', 'Giá trị', 'Ý nghĩa']]);
  ch.getRange(2, 1, CP_CAUHINH_BAN_DAU.length, 3).setValues(CP_CAUHINH_BAN_DAU);
  style(ch, [190, 300, 420]);

  const hd = cp.insertSheet(CP_SH.HD); hd.setColumnWidth(1, 900);
  // (sheet CP_KeHoach được tạo ở capNhatChiPhi_)
  const lines = ['HƯỚNG DẪN – QUẢN LÝ CHI PHÍ (I003) VÀ BÁO CÁO CHI PHÍ (I004) – Không Gian Số Nội Bộ', '',
    'CP_NguoiDuyet: mỗi người có thẩm quyền phê duyệt một dòng. Hạn mức để trống = không giới hạn (Giám đốc). Bỏ tick "Đang hiệu lực" khi thôi ủy quyền.',
    'CP_PhuTrach: mỗi phòng/ban một dòng, chọn ID người duyệt phụ trách. Đổi phân công chỉ cần đổi ID; khoản đã nhập vẫn tính cho người đã duyệt.',
    'CP_NhomChiPhi: danh mục nhóm chi phí để chọn khi nhập; bỏ tick "Đang dùng" thay vì xóa.',
    'CP_CauHinh: vai trò được nhập, được xem toàn bộ, được hủy; số ngày đầu tháng được nhập cho tháng trước; ngưỡng cảnh báo.',
    'CP_NhatKy: do ứng dụng ghi. Không sửa, không xóa tay. Hủy khoản qua ứng dụng (ghi lý do).',
    'Sau khi đổi tên cán bộ/phòng ở file phân quyền: mở file phân quyền → menu Không Gian Số → Đồng bộ danh mục sang file chi phí.',
    'Thay đổi cấu hình có hiệu lực sau tối đa 1 phút.'];
  hd.getRange(1, 1, lines.length, 1).setValues(lines.map(x => [x]));
  hd.getRange(1, 1).setFontWeight('bold').setFontSize(13).setFontColor('#005993');
}

/** Menu: cập nhật danh sách chọn (phòng/ban, người duyệt) và họ tên người duyệt trong file chi phí */
function dongBoChiPhi() {
  const id = PropertiesService.getScriptProperties().getProperty('CP_SHEET_ID');
  if (!id) return SpreadsheetApp.getUi().alert('Chưa cài đặt Quản lý chi phí.');
  const cp = SpreadsheetApp.openById(id);
  dongBoChiPhi_(cp);
  const warn = kiemTraChiPhi_(cp);
  SpreadsheetApp.getUi().alert('Đồng bộ danh mục', 'Đã cập nhật danh sách chọn và họ tên người duyệt.' + (warn ? '\n\n' + warn : ''), SpreadsheetApp.getUi().ButtonSet.OK);
}
function dongBoChiPhi_(cp) {
  CacheService.getScriptCache().remove('cp_cfg');
  const names = Object.keys(phongBanMap_()).filter(k => phongBanMap_()[k].on);
  const ids = Object.values(loadData_().users).filter(u => u.tt === TT.ON).map(u => u.id);
  const pt = cp.getSheetByName(CP_SH.PHUTRACH), du = cp.getSheetByName(CP_SH.DUYET), nh = cp.getSheetByName(CP_SH.NHOM);
  ensureRows_(pt, 50); ensureRows_(du, 30); ensureRows_(nh, 50);
  if (names.length) pt.getRange(2, 1, pt.getMaxRows() - 1, 1).setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(names, true).setAllowInvalid(false).build());
  pt.getRange(2, 2, pt.getMaxRows() - 1, 1).setDataValidation(SpreadsheetApp.newDataValidation().requireValueInRange(du.getRange('A2:A'), true).setAllowInvalid(false).build());
  if (ids.length) du.getRange(2, 1, du.getMaxRows() - 1, 1).setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(ids, false).setAllowInvalid(true)
    .setHelpText('ID cán bộ ở file phân quyền (trạng thái Hoạt động)').build());
  checkboxCol_(du, 4, du.getMaxRows()); checkboxCol_(nh, 2, nh.getMaxRows());
  // Điền họ tên theo ID
  const users = loadData_().users, n = du.getLastRow() - 1;
  if (n > 0) {
    const v = du.getRange(2, 1, n, 2).getValues();
    du.getRange(2, 2, n, 1).setValues(v.map(r => [users[String(r[0]).trim()] ? users[String(r[0]).trim()].hoTen : r[1]]));
  }
}
/** Cảnh báo cấu hình chi phí: người duyệt không tồn tại/không hoạt động, phòng chưa có người duyệt */
function kiemTraChiPhi_(cp) {
  try {
    const cfg = cpConfig_(true), users = loadData_().users, msg = [];
    Object.keys(cfg.duyet).forEach(id => { if (!users[id] || users[id].tt !== TT.ON) msg.push('- Người duyệt ID ' + id + ' không có hoặc không ở trạng thái Hoạt động trong CanBo.'); });
    Object.entries(cfg.phuTrach).forEach(([pb, id]) => { if (!cfg.duyet[id]) msg.push('- ' + pb + ': ID người duyệt ' + (id || '(trống)') + ' không có trong CP_NguoiDuyet hoặc không hiệu lực.'); });
    if (!Object.keys(cfg.duyet).length) msg.push('- Chưa có người duyệt nào đang hiệu lực.');
    return msg.length ? 'LƯU Ý cấu hình chi phí:\n' + msg.join('\n') : '';
  } catch (e) { return 'LƯU Ý: chưa đọc được file chi phí (' + e.message + ')'; }
}

/* ---------- Cấu hình (bộ nhớ đệm 60 giây) ---------- */
function cpFile_() {
  const id = PropertiesService.getScriptProperties().getProperty('CP_SHEET_ID');
  if (!id) throw new Error('cp_not_setup');
  return SpreadsheetApp.openById(id);
}
function cpConfig_(fresh) {
  const cache = CacheService.getScriptCache();
  if (!fresh) { const hit = cache.get('cp_cfg'); if (hit) return JSON.parse(hit); }
  const cp = cpFile_(), list = s => String(s || '').split(',').map(x => x.trim()).filter(Boolean);
  const rows = name => { const sh = cp.getSheetByName(name); return sh ? sh.getDataRange().getValues().slice(1) : []; };
  const duyet = {};
  rows(CP_SH.DUYET).forEach(r => { const id = String(r[0]).trim(); if (/^9\d{4}$/.test(id) && r[3] === true)
    duyet[id] = { id, ten: String(r[1] || '').trim(), limit: Number(r[2]) > 0 ? Number(r[2]) : 0 }; });
  const phuTrach = {};
  rows(CP_SH.PHUTRACH).forEach(r => { const pb = String(r[0]).trim(); if (pb) phuTrach[pb] = String(r[1]).trim(); });
  const nhom = rows(CP_SH.NHOM).filter(r => String(r[0]).trim() && r[1] !== false).map(r => String(r[0]).trim());
  const kv = {}; rows(CP_SH.CAUHINH).forEach(r => { if (String(r[0]).trim()) kv[String(r[0]).trim()] = String(r[1]).trim(); });
  const cfg = { duyet, phuTrach, nhom,
    vtNhap: list(kv.VAI_TRO_NHAP || 'Trưởng phòng, Phó phòng'), vtXem: list(kv.VAI_TRO_XEM_TAT_CA), pbXem: list(kv.PHONG_XEM_TAT_CA),
    vtHuy: list(kv.VAI_TRO_HUY), vtBcPhong: list(kv.VAI_TRO_BAO_CAO_PHONG || 'Trưởng phòng'), ngayThangTruoc: Math.max(0, Math.min(15, parseInt(kv.NGAY_NHAP_THANG_TRUOC || '5', 10) || 0)),
    nguong: Math.max(0, Math.min(90, parseInt(kv.NGUONG_CANH_BAO || '20', 10) || 20)) };
  try { cache.put('cp_cfg', JSON.stringify(cfg), 60); } catch (e) {}
  return cfg;
}
/** Quyền của 1 cán bộ trong ứng dụng chi phí */
function cpRole_(u, cfg) {
  const approverId = cfg.phuTrach[u.pb] || '';
  const canEnter = cfg.vtNhap.includes(u.vt) && !!approverId && !!cfg.duyet[approverId];
  const viewAll = cfg.vtXem.includes(u.vt) || (cfg.pbXem.includes(u.pb) && cfg.vtNhap.includes(u.vt));
  return { canEnter, viewAll, isApprover: !!cfg.duyet[u.id], cancelAny: cfg.vtHuy.includes(u.vt), approverId: canEnter ? approverId : '' };
}
function cpMonths_(cfg) {
  const t = Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd'), y = +t.slice(0, 4), m = +t.slice(5, 7), day = +t.slice(8, 10);
  const cur = t.slice(0, 7), prev = m === 1 ? (y - 1) + '-12' : y + '-' + String(m - 1).padStart(2, '0');
  return day <= cfg.ngayThangTruoc ? [cur, prev] : [cur];
}
/** Đọc nhật ký → mảng đối tượng */
function cpRows_(cp) {
  const sh = cp.getSheetByName(CP_SH.NK), n = sh.getLastRow() - 1; if (n < 1) return [];
  const fd = (d, f) => d instanceof Date ? Utilities.formatDate(d, TZ, f) : String(d || '');
  return sh.getRange(2, 1, n, CP_NK_HEAD.length).getValues().map((r, i) => ({
    row: i + 2, ma: String(r[0]), nhapLuc: fd(r[1], 'dd/MM/yyyy HH:mm'), nhapNgay: fd(r[1], 'yyyy-MM-dd'), ngayDuyet: fd(r[2], 'yyyy-MM-dd'),
    thang: cpThang_(r[2], r[3]), maPB: String(r[4]).trim(), pb: String(r[5]), nhom: String(r[6]), noiDung: String(r[7]), soTien: Number(r[8]) || 0,
    soToTrinh: String(r[9] || ''), duyetId: String(r[10]).trim(), duyetTen: String(r[11]), nhapId: String(r[12]).trim(), nhapTen: String(r[13]),
    tt: String(r[14]), lyDoHuy: String(r[15] || ''), huyBoi: String(r[16] || ''), huyLuc: fd(r[17], 'dd/MM/yyyy HH:mm')
  })).filter(x => x.ma);
}
/** Tháng yyyy-MM của 1 dòng nhật ký: ưu tiên Ngày phê duyệt; cột Tháng có thể bị Google Sheets tự đổi thành ngày */
function cpThang_(ngayDuyet, thang) {
  if (ngayDuyet instanceof Date) return Utilities.formatDate(ngayDuyet, TZ, 'yyyy-MM');
  if (thang instanceof Date) return Utilities.formatDate(thang, TZ, 'yyyy-MM');
  const m = String(thang || '').trim().match(/^(\d{4})-(\d{1,2})/);
  return m ? m[1] + '-' + m[2].padStart(2, '0') : String(thang || '').trim();
}
function cpUsed_(rows, approverId, month) {
  return rows.filter(x => x.tt === CP_ST.OK && x.duyetId === approverId && x.thang === month).reduce((s, x) => s + x.soTien, 0);
}
function cpGuard_(b) {
  const v = checkToken_(b.token); if (v.error) return v;
  if (!userApps_(v.user).includes(CP_APP)) return { error: 'forbidden' };
  let cfg; try { cfg = cpConfig_(); } catch (e) { return { error: 'not_configured', message: 'Chưa cài đặt Quản lý chi phí.' }; }
  const role = cpRole_(v.user, cfg);
  if (!role.canEnter && !role.viewAll && !role.isApprover) return { error: 'forbidden', message: 'Tài khoản chưa được phân công trong Quản lý chi phí.' };
  return { user: v.user, cfg, role };
}

/* ---------- API ---------- */
function apiCpLoad_(b) {
  const g = cpGuard_(b); if (g.error) return g;
  const { user, cfg, role } = g, cp = cpFile_(), all = cpRows_(cp);
  const month = /^\d{4}-\d{2}$/.test(String(b.month || '')) ? String(b.month) : Utilities.formatDate(new Date(), TZ, 'yyyy-MM');
  const today = Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd');
  // Phạm vi xem
  const see = x => role.viewAll || (role.isApprover && x.duyetId === user.id) || (role.canEnter && x.pb === user.pb);
  const entries = all.filter(x => x.thang === month && see(x))
    .sort((a, c) => (c.ngayDuyet + c.nhapLuc).localeCompare(a.ngayDuyet + a.nhapLuc) || c.row - a.row)
    .map(x => Object.assign({}, x, { row: undefined, nhapNgay: undefined,
      canCancel: x.tt === CP_ST.OK && (role.cancelAny || (x.nhapId === user.id && x.nhapNgay === today)) }));
  // Người duyệt hiển thị
  let ids = role.viewAll ? Object.keys(cfg.duyet) : [];
  if (role.isApprover) ids.push(user.id);
  if (role.canEnter) ids.push(role.approverId);
  if (role.viewAll) all.filter(x => x.thang === month && x.tt === CP_ST.OK).forEach(x => ids.push(x.duyetId));
  ids = Array.from(new Set(ids.filter(Boolean)));
  const users = loadData_().users;
  const approver = id => { const d = cfg.duyet[id], used = cpUsed_(all, id, month), lim = d ? d.limit : 0;
    return { id, ten: (users[id] && users[id].hoTen) || (d && d.ten) || ('ID ' + id), limit: lim, used, remaining: lim ? lim - used : null, active: !!d,
      phong: Object.keys(cfg.phuTrach).filter(p => cfg.phuTrach[p] === id) }; };
  const approvers = ids.map(approver).sort((a, c) => (c.limit ? 1 : 0) - (a.limit ? 1 : 0) || a.ten.localeCompare(c.ten));
  // Form nhập
  let form = null;
  if (role.canEnter) {
    const months = cpMonths_(cfg), d = cfg.duyet[role.approverId];
    form = { phong: user.pb, approver: approver(role.approverId).ten, approverId: role.approverId, limit: d.limit, nhom: cfg.nhom, months,
      used: months.reduce((o, m) => (o[m] = cpUsed_(all, role.approverId, m), o), {}), today };
  }
  return { ok: true, month, today, role: { canEnter: role.canEnter, viewAll: role.viewAll, isApprover: role.isApprover }, nguong: cfg.nguong,
    approvers, entries, form, nhom: cfg.nhom };
}

function apiCpAdd_(b) {
  const g = cpGuard_(b); if (g.error) return g;
  const { user, cfg, role } = g;
  if (!role.canEnter) return { error: 'forbidden', message: 'Tài khoản không có quyền nhập chi phí.' };
  const d = b.data || {};
  const noiDung = String(d.noiDung || '').trim().slice(0, 300), nhom = String(d.nhom || '').trim(), soToTrinh = String(d.soToTrinh || '').trim().slice(0, 50);
  const soTien = Math.round(Number(String(d.soTien || '').replace(/[^\d]/g, '')));
  const ngay = String(d.ngayDuyet || '');
  if (!noiDung) return { error: 'invalid', message: 'Nhập nội dung chi phí.' };
  if (!cfg.nhom.includes(nhom)) return { error: 'invalid', message: 'Chọn nhóm chi phí trong danh sách.' };
  if (!(soTien > 0) || soTien > 1e12) return { error: 'invalid', message: 'Số tiền không hợp lệ.' };
  if (!/^\d{4}-\d{2}-\d{2}$/.test(ngay)) return { error: 'invalid', message: 'Chọn ngày phê duyệt.' };
  const today = Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd'), month = ngay.slice(0, 7);
  if (ngay > today) return { error: 'invalid', message: 'Ngày phê duyệt không được sau hôm nay.' };
  if (!cpMonths_(cfg).includes(month)) return { error: 'invalid', message: 'Chỉ nhập khoản phê duyệt trong tháng này' + (cfg.ngayThangTruoc ? ' (hoặc tháng trước, từ ngày 1 đến ngày ' + cfg.ngayThangTruoc + ')' : '') + '.' };

  const lock = LockService.getScriptLock();
  try { lock.waitLock(10000); } catch (e) { return { error: 'busy', message: 'Hệ thống đang ghi dữ liệu khác, thử lại sau vài giây.' }; }
  try {
    const cp = cpFile_(), all = cpRows_(cp), ap = cfg.duyet[role.approverId];
    const used = cpUsed_(all, role.approverId, month);
    const dup = all.find(x => x.tt === CP_ST.OK && x.pb === user.pb && x.ngayDuyet === ngay && x.soTien === soTien &&
      x.noiDung.trim().toLowerCase() === noiDung.toLowerCase());
    if (dup && !d.xacNhanTrung)
      return { error: 'duplicate', message: 'Đã có khoản giống hệt (cùng ngày, số tiền, nội dung) nhập lúc ' + dup.nhapLuc + ' bởi ' + dup.nhapTen + '. Nếu đây là khoản khác, bấm "Vẫn lưu".' };
    if (ap.limit && used + soTien > ap.limit)
      return { error: 'over_limit', remaining: ap.limit - used, message: 'Vượt hạn mức phê duyệt tháng ' + month.slice(5) + '/' + month.slice(0, 4) + ' của ' + ap.ten + '. Còn lại ' + fmt_(ap.limit - used) + ' đ.' };
    const pb = phongBanMap_()[user.pb] || {};
    const ma = 'CP' + Utilities.formatDate(new Date(), TZ, 'yyMMddHHmmss') + Math.floor(Math.random() * 90 + 10);
    const sh = cp.getSheetByName(CP_SH.NK);
    const txt = v => v === '' || v === null || v === undefined ? '' : "'" + String(v);
    sh.appendRow([ma, new Date(), new Date(ngay + 'T12:00:00+07:00'), txt(month), txt(pb.ma || ''), user.pb, nhom, txt(noiDung), soTien, txt(soToTrinh),
      txt(role.approverId), ap.ten, txt(user.id), user.hoTen, CP_ST.OK, '', '', '']);
    log_(user.id, user.hoTen, 'Nhập chi phí', CP_APP, ma + ' · ' + fmt_(soTien) + ' đ · duyệt: ' + ap.ten, '');
  } finally { lock.releaseLock(); }
  return apiCpLoad_({ token: b.token, month });
}

function apiCpCancel_(b) {
  const g = cpGuard_(b); if (g.error) return g;
  const { user, cfg, role } = g, ma = String(b.ma || ''), lyDo = String(b.lyDo || '').trim().slice(0, 200);
  if (lyDo.length < 5) return { error: 'invalid', message: 'Ghi lý do hủy (tối thiểu 5 ký tự).' };
  const lock = LockService.getScriptLock();
  try { lock.waitLock(10000); } catch (e) { return { error: 'busy', message: 'Hệ thống đang bận, thử lại sau vài giây.' }; }
  let month;
  try {
    const cp = cpFile_(), x = cpRows_(cp).find(r => r.ma === ma);
    if (!x) return { error: 'invalid', message: 'Không tìm thấy khoản chi phí.' };
    if (x.tt !== CP_ST.OK) return { error: 'invalid', message: 'Khoản này đã được hủy trước đó.' };
    const today = Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd');
    if (!(role.cancelAny || (x.nhapId === user.id && x.nhapNgay === today)))
      return { error: 'forbidden', message: 'Chỉ người nhập được hủy trong ngày nhập. Sau đó liên hệ Giám đốc hoặc quản trị.' };
    cp.getSheetByName(CP_SH.NK).getRange(x.row, 15, 1, 4).setValues([[CP_ST.HUY, lyDo, user.hoTen + ' (' + user.id + ')', new Date()]]);
    log_(user.id, user.hoTen, 'Hủy chi phí', CP_APP, ma + ' · ' + fmt_(x.soTien) + ' đ · ' + lyDo, '');
    month = x.thang;
  } finally { lock.releaseLock(); }
  return apiCpLoad_({ token: b.token, month: b.month || month });
}

/* =====================================================================
   PHẦN 5 – BÁO CÁO CHI PHÍ (ứng dụng I004)
   Máy chủ trả số liệu đã tổng hợp theo tháng × phòng × nhóm trong phạm vi được xem.
   ===================================================================== */
const DEFAULT_MODEL_BC = 'claude-sonnet-5-5';

/** Phạm vi báo cáo: toàn chi nhánh / các phòng người duyệt phụ trách / phòng của mình */
function cpReportScope_(u, cfg) {
  const role = cpRole_(u, cfg);
  if (role.viewAll) return { type: 'all', label: 'Toàn chi nhánh' };
  if (role.isApprover) {
    const pbs = Object.keys(cfg.phuTrach).filter(p => cfg.phuTrach[p] === u.id);
    return { type: 'approver', id: u.id, label: 'Các phòng do ' + (u.gt === 'Nữ' ? 'chị ' : u.gt === 'Nam' ? 'anh ' : '') + firstName_(u.hoTen) + ' phê duyệt', pbs };
  }
  if (cfg.vtBcPhong.includes(u.vt) && u.pb) return { type: 'phong', pb: u.pb, label: u.pb };
  return null;
}
function cpReportGuard_(b) {
  const v = checkToken_(b.token); if (v.error) return v;
  if (!userApps_(v.user).includes(CP_APP_BC)) return { error: 'forbidden' };
  let cfg; try { cfg = cpConfig_(); } catch (e) { return { error: 'not_configured', message: 'Chưa cài đặt Quản lý chi phí.' }; }
  const scope = cpReportScope_(v.user, cfg);
  if (!scope) return { error: 'forbidden', message: 'Tài khoản chưa có phạm vi xem báo cáo chi phí.' };
  return { user: v.user, cfg, scope };
}
/** Số liệu báo cáo cho tháng chọn (yyyy-MM): khối tổng hợp 24 tháng, kế hoạch, hạn mức, khoản lớn, khoản lặp lại */
function cpReportData_(user, cfg, scope, month) {
  const cp = cpFile_(), all = cpRows_(cp), y = +month.slice(0, 4);
  const inScope = x => scope.type === 'all' || (scope.type === 'approver' && x.duyetId === scope.id) || (scope.type === 'phong' && x.pb === scope.pb);
  const rows = all.filter(x => x.tt === CP_ST.OK && inScope(x) && x.thang >= (y - 1) + '-01' && x.thang <= month);
  const cube = {};
  rows.forEach(x => { const k = x.thang + '|' + x.pb + '|' + x.nhom; const c = cube[k] || (cube[k] = { t: x.thang, pb: x.pb, nhom: x.nhom, sum: 0, n: 0 }); c.sum += x.soTien; c.n++; });
  const pick = x => ({ ngay: x.ngayDuyet, pb: x.pb, nhom: x.nhom, noiDung: x.noiDung, soTien: x.soTien });
  const top = rows.filter(x => x.thang === month).sort((a, c) => c.soTien - a.soTien).slice(0, 10).map(pick);
  const topNam = rows.filter(x => x.thang.slice(0, 4) === String(y)).sort((a, c) => c.soTien - a.soTien).slice(0, 10).map(pick);
  // Khoản lặp lại trong 3 tháng gần nhất (cùng nội dung)
  const m3 = cpShiftMonth_(month, -2), rep = {};
  rows.filter(x => x.thang >= m3).forEach(x => { const k = x.noiDung.toLowerCase().replace(/\s+/g, ' ').trim();
    const r = rep[k] || (rep[k] = { noiDung: x.noiDung, n: 0, sum: 0, pbs: {} }); r.n++; r.sum += x.soTien; r.pbs[x.pb] = 1; });
  const lapLai = Object.values(rep).filter(r => r.n >= 2).sort((a, c) => c.sum - a.sum).slice(0, 5).map(r => ({ noiDung: r.noiDung, n: r.n, sum: r.sum, pbs: Object.keys(r.pbs) }));
  // Hạn mức người duyệt trong phạm vi (tháng chọn)
  let ids = scope.type === 'all' ? Object.keys(cfg.duyet).filter(id => cfg.duyet[id].limit) : scope.type === 'approver' ? [scope.id] : [cfg.phuTrach[scope.pb]].filter(Boolean);
  ids = ids.filter(id => cfg.duyet[id] && cfg.duyet[id].limit);
  const hanMuc = ids.map(id => ({ label: 'Người duyệt phụ trách ' + Object.keys(cfg.phuTrach).filter(p => cfg.phuTrach[p] === id).join(', '),
    ten: (loadData_().users[id] || {}).hoTen || cfg.duyet[id].ten, limit: cfg.duyet[id].limit, used: cpUsed_(all, id, month) }));
  // Kế hoạch năm trong phạm vi
  const sh = cp.getSheetByName(CP_SH.KH), keHoach = [];
  if (sh && sh.getLastRow() > 1) sh.getRange(2, 1, sh.getLastRow() - 1, 4).getValues().forEach(r => {
    const nam = +r[0], pb = String(r[1]).trim(), nhom = String(r[2]).trim(), kh = Number(r[3]) || 0;
    if (nam !== y || !pb || !kh) return;
    const ok = scope.type === 'all' || (scope.type === 'phong' && pb === scope.pb) || (scope.type === 'approver' && (scope.pbs || []).includes(pb));
    if (ok) keHoach.push({ pb, nhom, kh });
  });
  const phongs = scope.type === 'all' ? Object.keys(cfg.phuTrach) : scope.type === 'approver' ? scope.pbs : [scope.pb];
  return { ok: true, month, today: Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd'), scope: { type: scope.type, label: scope.label },
    phongs, nhomList: cfg.nhom, cube: Object.values(cube), top, topNam, lapLai, hanMuc, keHoach, ai: !!PropertiesService.getScriptProperties().getProperty('ANTHROPIC_API_KEY') };
}
function cpShiftMonth_(m, k) { let y = +m.slice(0, 4), mo = +m.slice(5, 7) + k; while (mo < 1) { mo += 12; y--; } while (mo > 12) { mo -= 12; y++; } return y + '-' + String(mo).padStart(2, '0'); }

function apiCpReport_(b) {
  const g = cpReportGuard_(b); if (g.error) return g;
  const month = /^\d{4}-\d{2}$/.test(String(b.month || '')) ? String(b.month) : Utilities.formatDate(new Date(), TZ, 'yyyy-MM');
  return cpReportData_(g.user, g.cfg, g.scope, month);
}

/** Nhận xét AI: máy chủ tự lập bản tóm tắt số liệu (không gửi ID, tên cán bộ) rồi gọi Claude API */
function apiCpAi_(b) {
  const g = cpReportGuard_(b); if (g.error) return g;
  const props = PropertiesService.getScriptProperties(), key = props.getProperty('ANTHROPIC_API_KEY');
  if (!key) return { error: 'not_configured', message: 'Chưa bật trợ lý AI (thiếu ANTHROPIC_API_KEY trong thuộc tính tập lệnh).' };
  const month = /^\d{4}-\d{2}$/.test(String(b.month || '')) ? String(b.month) : Utilities.formatDate(new Date(), TZ, 'yyyy-MM');
  const d = cpReportData_(g.user, g.cfg, g.scope, month);
  const cacheKey = 'cp_ai_' + Utilities.base64EncodeWebSafe(Utilities.computeDigest(Utilities.DigestAlgorithm.MD5, JSON.stringify([d.scope, d.cube, d.hanMuc, d.keHoach, d.today])));
  const hit = CacheService.getScriptCache().get(cacheKey); if (hit) return { ok: true, text: hit, cached: true };
  const lim = aiLimit_(props); if (lim) return lim;
  const res = UrlFetchApp.fetch('https://api.anthropic.com/v1/messages', {
    method: 'post', contentType: 'application/json', muteHttpExceptions: true,
    headers: { 'x-api-key': key, 'anthropic-version': '2023-06-01' },
    payload: JSON.stringify({ model: props.getProperty('MODEL_BAO_CAO') || DEFAULT_MODEL_BC, max_tokens: 1800, system: CP_AI_PROMPT,
      messages: [{ role: 'user', content: cpSummary(d) }] })
  });
  const code = res.getResponseCode();
  if (code === 429 || code === 529) return { error: 'rate_limited', message: 'Dịch vụ AI đang quá tải, thử lại sau ít phút.' };
  if (code !== 200) { console.error('Claude API ' + code + ': ' + res.getContentText().slice(0, 300)); return { error: 'upstream_error', message: 'Dịch vụ AI tạm thời lỗi, thử lại sau.' }; }
  const text = (JSON.parse(res.getContentText()).content || []).filter(c => c.type === 'text').map(c => c.text).join('').trim();
  if (!text) return { error: 'upstream_error', message: 'AI không trả về nội dung.' };
  try { CacheService.getScriptCache().put(cacheKey, text, 1800); } catch (e) {}
  log_(g.user.id, g.user.hoTen, 'Phân tích AI', CP_APP_BC, d.scope.label + ' · tháng ' + month, '');
  return { ok: true, text };
}

/* ---- Dùng chung máy chủ (Code.gs) và trang (index.html): tính chỉ số báo cáo, tóm tắt cho AI ---- */
const CP_AI_PROMPT = 'Bạn là chuyên viên phân tích chi phí hoạt động của VietinBank Chi nhánh Tây Tiền Giang, viết nhận xét tham mưu cho Ban Giám đốc.\n' +
  'Dựa DUY NHẤT vào số liệu được cung cấp, viết nhận xét bằng tiếng Việt, văn phong hành chính – ngân hàng, ngắn gọn, cụ thể.\n' +
  'Cấu trúc bắt buộc, mỗi phần mở đầu bằng "## ":\n## Tình hình chung\n## Điểm cần lưu ý\n## Kết luận\n## Định hướng điều hành\n' +
  'Trong mỗi phần dùng gạch đầu dòng "- ". Phần Định hướng điều hành có 3–5 ý hành động cụ thể (ai, việc gì).\n' +
  'Quy tắc: không bịa số liệu; số tiền ghi dạng 1.000.000 đ; nêu rõ khi chưa đủ dữ liệu để so sánh (năm trước, kế hoạch); ' +
  'cảnh báo khi hạn mức người duyệt dự kiến hết trước cuối tháng hoặc thực hiện vượt tiến độ kế hoạch; không nêu tên cá nhân; tối đa 380 từ.';

function cpFmt(n) { return String(Math.round(n || 0)).replace(/\B(?=(\d{3})+(?!\d))/g, '.') + ' đ'; }
function cpShift(m, k) { let y = +m.slice(0, 4), mo = +m.slice(5, 7) + k; while (mo < 1) { mo += 12; y--; } while (mo > 12) { mo -= 12; y++; } return y + '-' + String(mo).padStart(2, '0'); }

/** Tính toàn bộ chỉ số từ dữ liệu báo cáo d (kết quả cp_report) */
function cpMetrics(d) {
  const m = d.month, y = +m.slice(0, 4), py = y - 1, mm = m.slice(5), prev = cpShift(m, -1);
  const sum = f => d.cube.filter(f).reduce((s, c) => s + c.sum, 0);
  const cnt = f => d.cube.filter(f).reduce((s, c) => s + c.n, 0);
  const inYtd = (c, yy) => c.t >= yy + '-01' && c.t <= yy + '-' + mm;
  const last3 = [cpShift(m, -1), cpShift(m, -2), cpShift(m, -3)];
  const months = {}; d.cube.forEach(c => months[c.t] = 1);
  const hasPrevYear = Object.keys(months).some(t => t.slice(0, 4) === String(py));
  const n3 = last3.filter(t => months[t]).length;
  const tot = { cur: sum(c => c.t === m), n: cnt(c => c.t === m), prev: sum(c => c.t === prev), hasPrev: !!months[prev],
    avg3: n3 ? sum(c => last3.includes(c.t)) / n3 : null, n3, ytd: sum(c => inYtd(c, y)), ytdPrev: sum(c => inYtd(c, py)), hasPrevYear,
    sameLastYear: sum(c => c.t === py + '-' + mm) };
  const pbs = Array.from(new Set((d.phongs || []).concat(d.cube.map(c => c.pb))));
  const byPhong = pbs.map(pb => ({ pb, cur: sum(c => c.pb === pb && c.t === m), prev: sum(c => c.pb === pb && c.t === prev),
    avg3: n3 ? sum(c => c.pb === pb && last3.includes(c.t)) / n3 : null, ytd: sum(c => c.pb === pb && inYtd(c, y)), ytdPrev: sum(c => c.pb === pb && inYtd(c, py)) }))
    .sort((a, b) => b.cur - a.cur || b.ytd - a.ytd);
  const nhoms = Array.from(new Set((d.nhomList || []).concat(d.cube.map(c => c.nhom))));
  const byNhom = nhoms.map(nh => ({ nhom: nh, cur: sum(c => c.nhom === nh && c.t === m), prev: sum(c => c.nhom === nh && c.t === prev),
    ytd: sum(c => c.nhom === nh && inYtd(c, y)) })).filter(x => x.cur || x.prev || x.ytd).sort((a, b) => b.cur - a.cur || b.ytd - a.ytd);
  byNhom.forEach(x => { x.share = tot.cur ? x.cur / tot.cur : 0; x.shareYtd = tot.ytd ? x.ytd / tot.ytd : 0; });
  const trend = [], trendPrev = [];
  for (let i = 1; i <= 12; i++) { const k = String(i).padStart(2, '0'); trend.push(sum(c => c.t === y + '-' + k)); trendPrev.push(sum(c => c.t === py + '-' + k)); }
  // Tiến độ thời gian của năm tính đến tháng chọn
  const today = d.today, isCur = today.slice(0, 7) === m, dim = new Date(y, +mm, 0).getDate(), day = isCur ? +today.slice(8, 10) : dim;
  const timePct = ((+mm - 1) + day / dim) / 12;
  // Kế hoạch
  const kh = d.keHoach || [];
  const planAll = kh.filter(k => k.pb === 'Toàn chi nhánh' && !k.nhom).reduce((s, k) => s + k.kh, 0);
  const planPb = {}; kh.filter(k => k.pb !== 'Toàn chi nhánh' && !k.nhom).forEach(k => planPb[k.pb] = (planPb[k.pb] || 0) + k.kh);
  const planNhom = {}; kh.filter(k => k.nhom).forEach(k => { const key = k.pb + '|' + k.nhom; planNhom[key] = (planNhom[key] || 0) + k.kh; });
  const planTotal = d.scope.type === 'all' && planAll ? planAll : Object.values(planPb).reduce((s, v) => s + v, 0);
  const plan = { total: planTotal, timePct, ytdPct: planTotal ? tot.ytd / planTotal : null,
    byPhong: byPhong.filter(p => planPb[p.pb]).map(p => ({ pb: p.pb, kh: planPb[p.pb], ytd: p.ytd, pct: p.ytd / planPb[p.pb] })),
    byNhom: Object.entries(planNhom).map(([k, v]) => { const [pb, nh] = k.split('|'); const ytd = sum(c => c.nhom === nh && (pb === 'Toàn chi nhánh' || c.pb === pb) && inYtd(c, y)); return { pb, nhom: nh, kh: v, ytd, pct: ytd / v }; }) };
  // Hạn mức và dự báo ngày hết hạn mức
  const hanMuc = (d.hanMuc || []).map(h => { const rate = isCur && day ? h.used / day : 0, proj = isCur ? rate * dim : h.used;
    const dayOut = isCur && rate > 0 && h.used < h.limit ? Math.ceil(h.limit / rate) : null;
    return Object.assign({}, h, { pct: h.used / h.limit, remaining: h.limit - h.used, proj, dayOut: dayOut && dayOut <= dim ? dayOut : null, isCur }); });
  // Ma trận phòng × nhóm (tháng chọn và lũy kế năm)
  const matrix = (f) => { const cells = {}; let max = 0; d.cube.filter(f).forEach(c => { const k = c.pb + '|' + c.nhom; cells[k] = (cells[k] || 0) + c.sum; max = Math.max(max, cells[k]); }); return { cells, max }; };
  return { m, y, py, prev, tot, byPhong, byNhom, trend, trendPrev, plan, hanMuc, isCur, day, dim,
    matrixCur: matrix(c => c.t === m), matrixYtd: matrix(c => inYtd(c, y)) };
}

/** Bản tóm tắt số liệu gửi AI (không chứa ID, tên cán bộ) */
function cpSummary(d) {
  const k = cpMetrics(d), L = [], ml = k.m.slice(5) + '/' + k.y, pct = v => (v * 100).toFixed(1).replace('.', ',') + '%';
  const chg = (a, b) => b ? (a >= b ? '+' : '') + pct((a - b) / b) : 'không có số liệu so sánh';
  L.push('PHẠM VI: ' + d.scope.label + '. THÁNG BÁO CÁO: ' + ml + (k.isCur ? ' (đang diễn ra, số liệu đến ngày ' + k.day + '/' + k.dim + ')' : ' (đã kết thúc)') + '.');
  L.push('Tổng chi phí đã phê duyệt tháng: ' + cpFmt(k.tot.cur) + ' (' + k.tot.n + ' khoản). Tháng trước: ' + (k.tot.hasPrev ? cpFmt(k.tot.prev) + ', thay đổi ' + chg(k.tot.cur, k.tot.prev) : 'chưa có dữ liệu') +
    '. Bình quân 3 tháng trước: ' + (k.tot.avg3 !== null ? cpFmt(k.tot.avg3) + ' (tính trên ' + k.tot.n3 + ' tháng có dữ liệu)' : 'chưa có dữ liệu') + '.');
  L.push('Lũy kế từ đầu năm ' + k.y + ': ' + cpFmt(k.tot.ytd) + '. Cùng kỳ năm ' + k.py + ': ' + (k.tot.hasPrevYear ? cpFmt(k.tot.ytdPrev) + ' (' + chg(k.tot.ytd, k.tot.ytdPrev) + ')' : 'chưa có dữ liệu năm trước') + '.');
  L.push('Kế hoạch năm: ' + (k.plan.total ? cpFmt(k.plan.total) + ', đã thực hiện ' + pct(k.plan.ytdPct) + ' trong khi tiến độ thời gian là ' + pct(k.plan.timePct) : 'chưa giao kế hoạch') + '.');
  k.plan.byPhong.forEach(p => L.push('  Kế hoạch ' + p.pb + ': ' + cpFmt(p.kh) + ', thực hiện ' + pct(p.pct) + '.'));
  L.push('THEO PHÒNG/BAN (tháng | tháng trước | bình quân 3 tháng | lũy kế năm):');
  k.byPhong.filter(p => p.cur || p.prev || p.ytd).forEach(p => L.push('- ' + p.pb + ': ' + cpFmt(p.cur) + ' | ' + cpFmt(p.prev) + ' | ' + (p.avg3 !== null ? cpFmt(p.avg3) : '—') + ' | ' + cpFmt(p.ytd)));
  L.push('THEO NHÓM CHI PHÍ (tháng, tỷ trọng | tháng trước | lũy kế năm):');
  k.byNhom.forEach(x => L.push('- ' + x.nhom + ': ' + cpFmt(x.cur) + ', ' + pct(x.share) + ' | ' + cpFmt(x.prev) + ' | ' + cpFmt(x.ytd)));
  if (k.hanMuc.length) { L.push('HẠN MỨC PHÊ DUYỆT ỦY QUYỀN (tháng ' + ml + '):');
    k.hanMuc.forEach(h => L.push('- ' + h.label + ': hạn mức ' + cpFmt(h.limit) + ', đã dùng ' + cpFmt(h.used) + ' (' + pct(h.pct) + '), còn ' + cpFmt(h.remaining) +
      (h.isCur ? '; theo tốc độ hiện tại dự kiến cả tháng ' + cpFmt(h.proj) + (h.dayOut ? ', hết hạn mức khoảng ngày ' + h.dayOut : '') : ''))); }
  if (d.top.length) { L.push('CÁC KHOẢN LỚN NHẤT TRONG THÁNG:'); d.top.slice(0, 5).forEach(t => L.push('- ' + t.noiDung + ' (' + t.pb + ', ' + t.nhom + '): ' + cpFmt(t.soTien))); }
  if (d.lapLai.length) { L.push('KHOẢN LẶP LẠI TRONG 3 THÁNG GẦN NHẤT:'); d.lapLai.forEach(r => L.push('- ' + r.noiDung + ': ' + r.n + ' lần, tổng ' + cpFmt(r.sum) + ' (' + r.pbs.join(', ') + ')')); }
  return L.join('\n');
}

/* =====================================================================
   PHẦN 6 – BẢN TIN SÁNG (ứng dụng I005) – từ v1.5.0
   Bản tin do tác vụ định kỳ của Claude tạo lúc 06:19 hằng ngày, lưu file BanTin_YYYY-MM-DD.json
   vào thư mục BanTin (trong KhongGianSoNoiBo). Script chỉ ĐỌC thư mục này (quyền drive.readonly).
   ===================================================================== */
const BT_APP = 'I005', BT_TEN = 'Bản tin sáng';
const BT_FOLDER_DEFAULT = '1WpMZsPOTvPcJOaSfDqKlTZbuKbNK9tKB';   // thư mục KhongGianSoNoiBo/BanTin
const BT_VAI_TRO_MO = ['Giám đốc', 'Phó Giám đốc', 'Quản trị hệ thống', 'Trưởng phòng', 'Phó phòng'];
const BT_SO_NGAY = 30, BT_FILE_RE = /^BanTin_(\d{4}-\d{2}-\d{2})\.json$/;

/** Menu: đăng ký I005, kiểm tra thư mục bản tin. Chạy lại an toàn. */
function caiDatBanTin() {
  const ui = SpreadsheetApp.getUi(), props = PropertiesService.getScriptProperties();
  if (!props.getProperty('BT_FOLDER_ID')) props.setProperty('BT_FOLDER_ID', BT_FOLDER_DEFAULT);
  let folder;
  try { folder = DriveApp.getFolderById(props.getProperty('BT_FOLDER_ID')); }
  catch (e) { ui.alert(BT_TEN, 'Không mở được thư mục bản tin (BT_FOLDER_ID = ' + props.getProperty('BT_FOLDER_ID') + ').\nKiểm tra lại ID thư mục trong Thuộc tính tập lệnh.', ui.ButtonSet.OK); return; }
  const isNew = dangKyUngDung_(BT_APP, BT_TEN, 'Bản tin sáng phục vụ điều hành: địa bàn, chính sách, thị trường, gợi ý điều hành', BT_VAI_TRO_MO);
  CacheService.getScriptCache().remove('bt_idx');
  lamMoiBoNhoDem(true);
  const idx = btIndex_(), ds = Object.keys(idx).sort();
  ui.alert(BT_TEN,
    (isNew ? 'Đã đăng ký ' + BT_APP + ' – ' + BT_TEN + ', tick quyền: ' + BT_VAI_TRO_MO.join(', ') + '.\n\n' : BT_APP + ' đã có, giữ nguyên quyền.\n\n') +
    'Thư mục bản tin: ' + folder.getName() + '\nSố bản tin trong ' + BT_SO_NGAY + ' ngày: ' + ds.length +
    (ds.length ? ' (mới nhất ' + ds[ds.length - 1].split('-').reverse().join('/') + ')' : '') +
    '\n\nCuối cùng: Triển khai → Quản lý triển khai → Phiên bản mới.', ui.ButtonSet.OK);
}

/** Danh mục bản tin 30 ngày: { 'YYYY-MM-DD': {id, t} }. Trùng ngày thì lấy file sửa sau cùng. Đệm 10 phút. */
function btIndex_() {
  const cache = CacheService.getScriptCache(), hit = cache.get('bt_idx');
  if (hit) return JSON.parse(hit);
  const from = Utilities.formatDate(new Date(Date.now() - BT_SO_NGAY * 864e5), TZ, 'yyyy-MM-dd'), out = {};
  const it = DriveApp.getFolderById(PropertiesService.getScriptProperties().getProperty('BT_FOLDER_ID') || BT_FOLDER_DEFAULT).getFiles();
  while (it.hasNext()) {
    const f = it.next(), m = f.getName().match(BT_FILE_RE);
    if (!m || m[1] < from || f.isTrashed()) continue;
    const t = f.getLastUpdated().getTime();
    if (!out[m[1]] || out[m[1]].t < t) out[m[1]] = { id: f.getId(), t: t };
  }
  cache.put('bt_idx', JSON.stringify(out), 600);
  return out;
}
/** Đọc 1 bản tin (đệm 6 giờ theo phiên bản file) */
function btDoc_(ngay, meta) {
  const cache = CacheService.getScriptCache(), key = 'bt_' + ngay + '_' + meta.t, hit = cache.get(key);
  if (hit) return JSON.parse(hit);
  let d;
  try { d = JSON.parse(DriveApp.getFileById(meta.id).getBlob().getDataAsString('UTF-8')); } catch (e) { return null; }
  if (!d || d.phienBan !== 2 || !Array.isArray(d.diemNhan)) return null;
  const s = JSON.stringify(d);
  if (Utilities.newBlob(s).getBytes().length < 95000) cache.put(key, s, 21600); // CacheService tối đa 100 KB/khóa
  return d;
}
function btGuard_(b) {
  const v = checkToken_(b.token); if (v.error) return v;
  if (!userApps_(v.user).includes(BT_APP)) return { error: 'forbidden' };
  return v;
}
/** Danh sách ngày có bản tin + nội dung bản tin mới nhất (thẻ trang chủ dùng diemNhan) */
function apiBtList_(b) {
  const v = btGuard_(b); if (v.error) return v;
  const idx = btIndex_(), ds = Object.keys(idx).sort().reverse();
  for (const n of ds) { const d = btDoc_(n, idx[n]); if (d) return { ok: true, ngay: ds, latest: d }; }
  return { ok: true, ngay: ds, latest: null };
}
function apiBtGet_(b) {
  const v = btGuard_(b); if (v.error) return v;
  const ngay = String(b.ngay || ''), idx = btIndex_();
  if (!/^\d{4}-\d{2}-\d{2}$/.test(ngay) || !idx[ngay]) return { error: 'not_found' };
  const d = btDoc_(ngay, idx[ngay]);
  return d ? { ok: true, doc: d } : { error: 'not_found' };
}


/* =====================================================================
   PHẦN 7 – THẺ ATM BỊ MÁY THU (I006 Quản lý, I007 Báo cáo) – từ v1.6.0
   Dữ liệu: 3 sổ Google Sheet sẵn có (mỗi phòng nhận thẻ 1 file, sheet đầu tiên).
   Cấu hình: file KhongGianSoNoiBo_TheATM (TA_CauHinh, TA_NoiNhan, TA_NganHang, TA_ATM) – sửa trên sheet, không sửa code.
   ===================================================================== */
const TA_APP = 'I006', TA_APP_BC = 'I007', TA_TEN = 'Quản lý thẻ ATM bị máy thu', TA_TEN_BC = 'Báo cáo thẻ ATM máy thu';
const TA_FILE_NAME = 'KhongGianSoNoiBo_TheATM';
const TA_SH = { CAUHINH: 'TA_CauHinh', NOINHAN: 'TA_NoiNhan', NGANHANG: 'TA_NganHang', ATM: 'TA_ATM', HD: 'HuongDan' };
// Cột trong sổ (0-based): A–I theo Google Form cũ, J–S do ứng dụng quản lý
const TC = { TS: 0, GIAO_BOI: 1, NGAY_GIAO: 2, NGAY_TRA: 3, TEN: 4, SO_GOC: 5, NH_GOC: 6, NH_KHAC: 7, GHI_CHU: 8,
             MA: 9, SO: 10, NH: 11, TT: 12, XL_BOI: 13, XL_LUC: 14, XL_GHI_CHU: 15, CANH_BAO: 16, ATM: 17, ID_NHAP: 18 };
const TA_NCOL = 19;
const TA_HEAD_APP = ['Mã thẻ', 'Số thẻ (chuẩn)', 'Ngân hàng', 'Trạng thái', 'Cán bộ trả/hủy', 'Thời gian xử lý', 'Ghi chú xử lý', 'Cảnh báo dữ liệu', 'ATM thu thẻ', 'ID cán bộ nhập'];
const TA_VT_MO = ['Quản trị hệ thống', 'Điện toán', 'Trưởng phòng', 'Phó phòng', 'GDV', 'Cán bộ DVKH'];
const TA_VT_MO_BC = TA_VT_MO.concat(['Giám đốc', 'Phó Giám đốc']);
const TA_CAUHINH_BAN_DAU = [
  ['SO_NGAY_HUY', '90', 'Thẻ còn giữ từ số ngày này trở lên (tính từ ngày vào sổ) vào danh sách cần hủy; cũng là mốc gửi email nhắc hủy'],
  ['NHAC_HUY_THU', 'Thứ Hai', 'Ngày gửi email nhắc thẻ đến hạn hủy: Thứ Hai … Thứ Bảy, Chủ nhật, Hằng ngày; để trống = tắt'],
  ['NHAC_HUY_GIO', '8', 'Giờ gửi email nhắc (0–23). Đổi giờ xong chạy lại menu Cài đặt thẻ ATM'],
  ['NHOM_TUOI_TON', '15, 30, 60, 90', 'Các mốc số ngày giữ thẻ để chia nhóm trong báo cáo tồn'],
  ['SO_CHU_SO_DAU', '6', 'Số chữ số đầu của số thẻ được lưu: 6 (theo mẫu BM09-GNT) hoặc 4 (như sổ cũ)'],
  ['MIEN_EMAIL', 'vietinbank.vn', 'Chỉ gửi email tới địa chỉ thuộc tên miền này'],
  ['GUI_EMAIL_KHI_NHAP', 'Có', 'Có/Không: gửi email cho cán bộ phụ trách phòng nhận khi nhập thẻ'],
  ['VAI_TRO_NHAP', 'Quản trị hệ thống, Điện toán', 'Vai trò được nhập thẻ, xem danh sách (không được trả thẻ)'],
  ['VAI_TRO_TRA', 'Trưởng phòng, Phó phòng, GDV, Cán bộ DVKH', 'Vai trò được trả thẻ – chỉ sổ của phòng mình (phòng có trong TA_NoiNhan)'],
  ['VAI_TRO_HUY', 'Trưởng phòng, Phó phòng', 'Vai trò được thực hiện hủy thẻ quá hạn của phòng mình'],
  ['VAI_TRO_XEM_TAT_CA', 'Giám đốc, Phó Giám đốc, Quản trị hệ thống, Điện toán', 'Vai trò xem báo cáo của tất cả phòng nhận thẻ'],
  ['HOAN_TAC_TRONG_NGAY', 'Có', 'Có/Không: người trả thẻ được hoàn tác trong ngày'],
  ['DIA_DANH', 'Tiền Giang', 'Địa danh ở dòng ngày tháng trên phiếu'],
  ['TEN_CHI_NHANH', 'Chi nhánh Tây Tiền Giang', 'Tên chi nhánh trên email, báo cáo'],
  ['MAU_PHIEU_MA', 'BM09-GNT', 'Mã biểu mẫu in ở chân phiếu'],
  ['MAU_PHIEU_BAN_HANH', '02', 'Lần ban hành biểu mẫu'],
  ['MAU_PHIEU_SUA_DOI', '00', 'Lần sửa đổi biểu mẫu']
];
// Danh mục ngân hàng ban đầu. BIN nội địa 9704xx theo danh sách mã BIN của NHNN (31/12/2023), tên theo hiện hành; sửa/thêm trên sheet TA_NganHang.
const TA_NGANHANG_BAN_DAU = [
  ['VietinBank', 'Ngân hàng TMCP Công Thương Việt Nam', '970415', 'Vietinbank, CTG'],
  ['Vietcombank', 'Ngân hàng TMCP Ngoại thương Việt Nam', '970436', 'VCB'],
  ['Agribank', 'Ngân hàng Nông nghiệp và Phát triển Nông thôn Việt Nam', '970405', 'AgriBank, VBARD'],
  ['BIDV', 'Ngân hàng TMCP Đầu tư và Phát triển Việt Nam', '970418', ''],
  ['Sacombank', 'Ngân hàng TMCP Sài Gòn Thương Tín', '970403', 'STB'],
  ['Techcombank', 'Ngân hàng TMCP Kỹ Thương Việt Nam', '970407', 'TCB'],
  ['ACB', 'Ngân hàng TMCP Á Châu', '970416', ''],
  ['MB', 'Ngân hàng TMCP Quân đội', '970422', 'MB Bank, MBBank'],
  ['VPBank', 'Ngân hàng TMCP Việt Nam Thịnh Vượng', '970432', 'VP Bank'],
  ['TPBank', 'Ngân hàng TMCP Tiên Phong', '970423', 'TP Bank, TienPhongBank'],
  ['HDBank', 'Ngân hàng TMCP Phát triển TP. Hồ Chí Minh', '970437', 'HD Bank'],
  ['VIB', 'Ngân hàng TMCP Quốc tế Việt Nam', '970441', ''],
  ['SHB', 'Ngân hàng TMCP Sài Gòn – Hà Nội', '970443', ''],
  ['Eximbank', 'Ngân hàng TMCP Xuất Nhập khẩu Việt Nam', '970431', 'EIB'],
  ['MSB', 'Ngân hàng TMCP Hàng Hải Việt Nam', '970426', 'Maritime Bank'],
  ['SeABank', 'Ngân hàng TMCP Đông Nam Á', '970440', ''],
  ['OCB', 'Ngân hàng TMCP Phương Đông', '970448', ''],
  ['Nam A Bank', 'Ngân hàng TMCP Nam Á', '970428', 'Nam A, NamABank'],
  ['LPBank', 'Ngân hàng TMCP Lộc Phát Việt Nam', '970449', 'LienVietPostBank'],
  ['ABBANK', 'Ngân hàng TMCP An Bình', '970425', 'ABB'],
  ['Bac A Bank', 'Ngân hàng TMCP Bắc Á', '970409', 'BacABank'],
  ['PVcomBank', 'Ngân hàng TMCP Đại Chúng Việt Nam', '970412', ''],
  ['KienlongBank', 'Ngân hàng TMCP Kiên Long', '970452', ''],
  ['Shinhan Bank', 'Ngân hàng TNHH MTV Shinhan Việt Nam', '970424', 'Shinhan'],
  ['Vikki Bank', 'Ngân hàng TNHH MTV Số Vikki', '970406', 'Vikki, DongA Bank'],
  ['BaoViet Bank', 'Ngân hàng TMCP Bảo Việt', '970438', 'BaoVietBank'],
  ['SCB', 'Ngân hàng TMCP Sài Gòn', '970429', ''],
  ['NCB', 'Ngân hàng TMCP Quốc Dân', '970419', ''],
  ['Saigonbank', 'Ngân hàng TMCP Sài Gòn Công Thương', '970400', ''],
  ['VietABank', 'Ngân hàng TMCP Việt Á', '970427', 'Viet A Bank'],
  ['VietBank', 'Ngân hàng TMCP Việt Nam Thương Tín', '970433', ''],
  ['BVBank', 'Ngân hàng TMCP Bản Việt', '970454', 'Viet Capital Bank'],
  ['PGBank', 'Ngân hàng TMCP Thịnh vượng và Phát triển', '970430', 'PG Bank'],
  ['MBV', 'Ngân hàng TNHH MTV Việt Nam Hiện Đại', '970414', 'OceanBank'],
  ['VCBNeo', 'Ngân hàng TNHH MTV Ngoại thương Công nghệ số', '970444', 'CBBank'],
  ['GPBank', 'Ngân hàng TNHH MTV Dầu khí Toàn Cầu', '970408', ''],
  ['Co-opBank', 'Ngân hàng Hợp tác xã Việt Nam', '970446', ''],
  ['Woori Bank', 'Ngân hàng TNHH MTV Woori Việt Nam', '970457', 'Woori'],
  ['UOB', 'Ngân hàng TNHH MTV UOB Việt Nam', '970458', ''],
  ['Public Bank', 'Ngân hàng TNHH MTV Public Việt Nam', '970439', ''],
  ['HSBC', 'Ngân hàng TNHH MTV HSBC (Việt Nam)', '', ''],
  ['Home Credit', 'Công ty Tài chính TNHH MTV Home Credit Việt Nam', '', 'Hone credit']
];

/** Menu: tạo/cập nhật file cấu hình, chuẩn bị 3 sổ, đăng ký I006/I007, cài lịch email nhắc hủy. Chạy lại an toàn. */
function caiDatTheATM() {
  const ui = SpreadsheetApp.getUi(), props = PropertiesService.getScriptProperties();
  let f = null;
  try { f = props.getProperty('TA_SHEET_ID') ? SpreadsheetApp.openById(props.getProperty('TA_SHEET_ID')) : null; } catch (e) { f = null; }
  const isNew = !f;
  if (isNew) { f = SpreadsheetApp.create(TA_FILE_NAME); f.setSpreadsheetTimeZone(TZ); props.setProperty('TA_SHEET_ID', f.getId()); taoFileTheATM_(f); }
  const changes = capNhatTheATM_(f);
  CacheService.getScriptCache().remove('ta_cfg');
  const cfg = taConfig_(true);
  const so = cfg.noiNhan.filter(n => n.on).map(n => { try { return n.pb + ': ' + chuanBiSoThe_(n, cfg); } catch (e) { return n.pb + ': LỖI ' + e.message; } });
  const n6 = dangKyUngDung_(TA_APP, TA_TEN, 'Nhập thẻ ATM bị máy thu (camera), sổ thẻ theo phòng, trả thẻ và in phiếu BM09-GNT', TA_VT_MO);
  const n7 = dangKyUngDung_(TA_APP_BC, TA_TEN_BC, 'Báo cáo thẻ tồn, thẻ đã trả, thẻ quá hạn cần hủy; hủy thẻ quá hạn', TA_VT_MO_BC);
  const trig = caiLichNhacHuy_(cfg);
  lamMoiBoNhoDem(true);
  const warn = taKiemTra_(cfg);
  ui.alert(TA_TEN,
    (isNew ? 'Đã tạo file "' + TA_FILE_NAME + '" ở thư mục gốc Drive. Hãy kéo file vào thư mục KhongGianSoNoiBo.\n\n' : 'File cấu hình đã có, giữ nguyên dữ liệu.\n\n') +
    (changes.length ? 'Đã cập nhật cấu hình: ' + changes.join('; ') + '.\n\n' : '') +
    'Sổ thẻ:\n' + so.join('\n') + '\n\n' +
    (n6 ? 'Đã đăng ký ' + TA_APP + ' – ' + TA_TEN + '.\n' : '') + (n7 ? 'Đã đăng ký ' + TA_APP_BC + ' – ' + TA_TEN_BC + '.\n' : '') +
    trig + '\nLink file cấu hình: ' + f.getUrl() + (warn ? '\n\n' + warn : '') +
    '\n\nTiếp theo: tắt nhận câu trả lời ở các Google Form sổ thẻ cũ; Triển khai → Quản lý triển khai → Phiên bản mới.', ui.ButtonSet.OK);
}

function taoFileTheATM_(f) {
  const style = (s, w) => { s.setFrozenRows(1); s.getRange(1, 1, 1, s.getLastColumn()).setFontWeight('bold').setFontColor('#FFFFFF').setBackground('#005993').setWrap(true);
    (w || []).forEach((x, i) => s.setColumnWidth(i + 1, x)); };
  const ch = f.getSheets()[0]; ch.setName(TA_SH.CAUHINH);
  ch.getRange(1, 1, 1, 3).setValues([['Tham số', 'Giá trị', 'Ý nghĩa']]);
  ch.getRange(2, 2, TA_CAUHINH_BAN_DAU.length, 1).setNumberFormat('@');
  ch.getRange(2, 1, TA_CAUHINH_BAN_DAU.length, 3).setValues(TA_CAUHINH_BAN_DAU);
  style(ch, [190, 330, 520]);

  const nn = f.insertSheet(TA_SH.NOINHAN);
  nn.getRange(1, 1, 1, 8).setValues([['Phòng nhận thẻ', 'ID file sổ thẻ', 'Tiền tố mã thẻ', 'Tên ghi trên phiếu (Kính gửi)', 'Email nhận', 'Email CC', 'Đang dùng', 'Ghi chú']]);
  nn.getRange(2, 1, 3, 8).setValues([
    ['Phòng DVKH', '14GcehdRyF-CsqM3JYAzj5TdsSCgHPlbW9fhojlBgs-g', 'DV', 'Chi nhánh Tây Tiền Giang – Phòng DVKH', '', '', true, 'Nhiều email cách nhau dấu phẩy'],
    ['PGD An Hữu', '10FMygZUyuCuPgybTA5bljB3me2OhvN--SPdasYDh8rg', 'AH', 'PGD An Hữu', '', '', true, ''],
    ['PGD Ấp Bắc', '1jYMen_rpNnItPby2Zf9OyJqCfxnyg7jro3_EO0I-5M0', 'AB', 'PGD Ấp Bắc', '', '', true, '']]);
  nn.getRange('B2:C').setNumberFormat('@'); nn.getRange(2, 7, 30, 1).insertCheckboxes();
  nn.getRange(2, 7, 3, 1).setValues([[true], [true], [true]]);
  style(nn, [130, 320, 90, 260, 260, 200, 80, 220]);

  const nh = f.insertSheet(TA_SH.NGANHANG);
  nh.getRange(1, 1, 1, 5).setValues([['Tên ngắn', 'Tên đầy đủ', 'Đầu số BIN (cách nhau dấu phẩy)', 'Tên gọi khác (cách nhau dấu phẩy)', 'Đang dùng']]);
  nh.getRange(2, 3, TA_NGANHANG_BAN_DAU.length, 1).setNumberFormat('@');
  nh.getRange(2, 1, TA_NGANHANG_BAN_DAU.length, 4).setValues(TA_NGANHANG_BAN_DAU);
  nh.getRange(2, 5, 80, 1).insertCheckboxes(); nh.getRange(2, 5, TA_NGANHANG_BAN_DAU.length, 1).setValue(true);
  style(nh, [130, 360, 220, 220, 80]);

  const atm = f.insertSheet(TA_SH.ATM);
  atm.getRange(1, 1, 1, 4).setValues([['Mã ATM', 'Địa điểm đặt ATM (in trên phiếu)', 'Phòng nhận thẻ mặc định', 'Đang dùng']]);
  atm.getRange('A2:A').setNumberFormat('@'); atm.getRange(2, 4, 50, 1).insertCheckboxes();
  style(atm, [110, 360, 160, 80]);

  const hd = f.insertSheet(TA_SH.HD); hd.setColumnWidth(1, 900);
  const lines = ['HƯỚNG DẪN – THẺ ATM BỊ MÁY THU (I006) VÀ BÁO CÁO (I007) – Không Gian Số Nội Bộ', '',
    'TA_CauHinh: tham số chung (số ngày hủy, lịch email nhắc, số chữ số đầu của thẻ, vai trò, địa danh trên phiếu…). Sửa cột Giá trị; có hiệu lực sau tối đa 1 phút.',
    'TA_NoiNhan: mỗi phòng nhận thẻ một dòng – tên phòng (trùng PhongBan), ID file sổ thẻ, tiền tố mã thẻ, tên ghi trên phiếu, email nhận/CC (chỉ @vietinbank.vn). Thêm phòng mới: tạo file sổ cùng cấu trúc rồi thêm dòng.',
    'TA_NganHang: danh mục ngân hàng; đầu số BIN 6 số để camera tự nhận ngân hàng; tên gọi khác để chuẩn hóa tên gõ tay. Bỏ tick Đang dùng thay vì xóa.',
    'TA_ATM: danh sách ATM của chi nhánh (địa điểm in vào mục "Giao dịch ở ATM đặt tại" trên phiếu). Có thể để trống.',
    'Sổ thẻ: cột A–I giữ như sổ cũ; cột J–S do ứng dụng ghi. Không sửa tay cột J–Q; sửa sai sót ở A–I.',
    'Sau khi đổi giờ gửi email nhắc: chạy lại menu Không Gian Số → Cài đặt / cập nhật Thẻ ATM máy thu.'];
  hd.getRange(1, 1, lines.length, 1).setValues(lines.map(x => [x]));
  hd.getRange(1, 1).setFontWeight('bold').setFontSize(13).setFontColor('#005993');
}
/** Bổ sung tham số mới cho file cấu hình tạo bằng bản cũ */
function capNhatTheATM_(f) {
  const out = [], ch = f.getSheetByName(TA_SH.CAUHINH);
  if (ch) { const keys = ch.getDataRange().getValues().map(r => String(r[0]).trim());
    TA_CAUHINH_BAN_DAU.forEach(r => { if (!keys.includes(r[0])) { ch.appendRow(r); out.push('thêm ' + r[0]); } }); }
  return out;
}

/** Chuẩn bị 1 sổ: thêm cột, chuyển cột J–M (công thức tạm) thành giá trị, công thức cảnh báo mới, bổ sung dòng thiếu */
function chuanBiSoThe_(n, cfg) {
  const sh = SpreadsheetApp.openById(n.fileId).getSheets()[0];
  if (sh.getMaxColumns() < TA_NCOL) sh.insertColumnsAfter(sh.getMaxColumns(), TA_NCOL - sh.getMaxColumns());
  const last = taLastRow_(sh);
  // J–M: nếu còn là công thức mảng thì chốt thành giá trị (để ứng dụng ghi được "Đã hủy", mã thẻ cố định)
  const fJ = sh.getRange(1, TC.MA + 1, 1, 4).getFormulas()[0];
  if (fJ.some(Boolean) && last > 1) {
    const v = sh.getRange(2, TC.MA + 1, last - 1, 4).getValues();
    sh.getRange(1, TC.MA + 1, 1, 4).clearContent();
    sh.getRange(2, TC.MA + 1, last - 1, 4).setValues(v);
  }
  sh.getRange(1, TC.MA + 1, 1, 7).setValues([TA_HEAD_APP.slice(0, 7)]);
  sh.getRange(1, TC.ATM + 1, 1, 2).setValues([TA_HEAD_APP.slice(8)]);
  sh.getRange(1, TC.CANH_BAO + 1).setFormula(TA_Q_FORMULA);
  sh.getRange(2, TC.SO_GOC + 1, sh.getMaxRows() - 1, 1).setNumberFormat('@');
  sh.getRange(2, TC.MA + 1, sh.getMaxRows() - 1, 2).setNumberFormat('@');
  sh.getRange(2, TC.ID_NHAP + 1, sh.getMaxRows() - 1, 1).setNumberFormat('@');
  sh.getRange(2, TC.XL_LUC + 1, sh.getMaxRows() - 1, 1).setNumberFormat('dd/mm/yyyy hh:mm');
  sh.getRange(2, TC.NGAY_TRA + 1, sh.getMaxRows() - 1, 1).setNumberFormat('dd/mm/yyyy');
  sh.setColumnWidth(TC.ATM + 1, 180); sh.setColumnWidth(TC.ID_NHAP + 1, 110);
  const rows = taReadRows_(sh, n, cfg, true);
  return rows.length + ' thẻ (' + rows.filter(x => x.tt === TA_ST.GIU).length + ' đang giữ)';
}
// Cảnh báo dữ liệu (công thức mảng ở Q1, cú pháp tiếng Anh của Apps Script)
const TA_Q_FORMULA = '=ARRAYFORMULA({"Cảnh báo dữ liệu";IF(A2:A="","",REGEXREPLACE(' +
  'IF(NOT(REGEXMATCH(F2:F&"","^\\s*\\d{4,6}\\s*[xX*]{2,}\\s*\\d{4}\\s*$")),"Số thẻ không đúng dạng số đầu + xxxx + 4 số cuối; ","")&' +
  'IF(REGEXMATCH(F2:F&"","^\\s*(9740|9794)"),"Đầu số "&LEFT(TRIM(F2:F&""),4)&" nghi gõ nhầm 9704; ","")&' +
  'IF(NOT(ISNUMBER(C2:C)),"Ngày giao không đúng dạng ngày; ",IF(C2:C>INT(A2:A),"Ngày giao sau ngày ghi sổ; ",""))&' +
  'IF((D2:D<>"")*NOT(ISNUMBER(D2:D)),"Ngày trả không đúng dạng ngày; ","")&' +
  'IF(ISNUMBER(D2:D)*ISNUMBER(C2:C)*(D2:D<C2:C),"Ngày trả trước ngày giao; ","")&' +
  'IF(REGEXMATCH(E2:E&"","\\d"),"Họ tên là mã/số, cần kiểm tra; ","")&' +
  'IF((G2:G="Khác")*(TRIM(H2:H&"")=""),"Chọn Khác nhưng chưa ghi tên ngân hàng; ","")&' +
  'IF((M2:M="' + 'Đang giữ' + '")*(K2:K<>"")*(COUNTIFS(K2:K,K2:K,E2:E,E2:E,M2:M,"Đang giữ")>1),"Trùng thẻ đang giữ với dòng khác; ",""),";\\s*$",""))})';

/** Cài lịch chạy nhacHuyTheATM mỗi ngày theo giờ cấu hình (hàm tự kiểm tra thứ trong tuần) */
function caiLichNhacHuy_(cfg) {
  try {
    ScriptApp.getProjectTriggers().filter(t => t.getHandlerFunction() === 'nhacHuyTheATM').forEach(t => ScriptApp.deleteTrigger(t));
    if (cfg.nhacThu === null) return 'Email nhắc hủy: đang tắt (NHAC_HUY_THU để trống).\n';
    ScriptApp.newTrigger('nhacHuyTheATM').timeBased().everyDays(1).atHour(cfg.nhacGio).inTimezone(TZ).create();
    return 'Email nhắc hủy: ' + (cfg.nhacThu === 'all' ? 'hằng ngày' : TA_THU[cfg.nhacThu]) + ', khoảng ' + cfg.nhacGio + ' giờ.\n';
  } catch (e) { return 'LƯU Ý: chưa cài được lịch email nhắc (' + e.message + '). Kiểm tra quyền script.scriptapp trong appsscript.json.\n'; }
}
const TA_THU = ['Chủ nhật', 'Thứ Hai', 'Thứ Ba', 'Thứ Tư', 'Thứ Năm', 'Thứ Sáu', 'Thứ Bảy'];

/** Kiểm tra cấu hình: email, vai trò, phòng, file sổ */
function taKiemTra_(cfg) {
  const msg = [], roles = Object.keys(rolePerms_()), pbs = phongBanMap_();
  [['VAI_TRO_NHAP', cfg.vtNhap], ['VAI_TRO_TRA', cfg.vtTra], ['VAI_TRO_HUY', cfg.vtHuy], ['VAI_TRO_XEM_TAT_CA', cfg.vtXem]].forEach(([k, l]) =>
    l.filter(r => !roles.includes(r)).forEach(r => msg.push('- ' + k + ': vai trò "' + r + '" không có trong sheet VaiTro.')));
  cfg.noiNhan.filter(n => n.on).forEach(n => {
    if (!pbs[n.pb]) msg.push('- TA_NoiNhan: "' + n.pb + '" không có trong sheet PhongBan.');
    if (!n.emails.length) msg.push('- ' + n.pb + ': chưa có Email nhận (thông báo nhập thẻ, nhắc hủy).');
    n.emailBad.forEach(x => msg.push('- ' + n.pb + ': email "' + x + '" không thuộc @' + cfg.mien + ', sẽ bị bỏ qua.'));
  });
  return msg.length ? 'LƯU Ý cấu hình thẻ ATM:\n' + msg.join('\n') : '';
}

/* ---------- Cấu hình (bộ nhớ đệm 60 giây) ---------- */
function taFile_() {
  const id = PropertiesService.getScriptProperties().getProperty('TA_SHEET_ID');
  if (!id) throw new Error('ta_not_setup');
  return SpreadsheetApp.openById(id);
}
function taConfig_(fresh) {
  const cache = CacheService.getScriptCache();
  if (!fresh) { const hit = cache.get('ta_cfg'); if (hit) return JSON.parse(hit); }
  const f = taFile_(), list = s => String(s || '').split(/[,;]/).map(x => x.trim()).filter(Boolean);
  const rows = name => { const sh = f.getSheetByName(name); return sh && sh.getLastRow() > 1 ? sh.getDataRange().getValues().slice(1) : []; };
  const kv = {}; rows(TA_SH.CAUHINH).forEach(r => { if (String(r[0]).trim()) kv[String(r[0]).trim()] = String(r[1]).trim(); });
  const yes = v => /^(c[oó]|yes|true|1|x)$/i.test(String(v || '').trim());
  const mien = (kv.MIEN_EMAIL || 'vietinbank.vn').replace(/^@/, '').toLowerCase();
  const okMail = e => new RegExp('^[^@\\s]+@' + mien.replace(/\./g, '\\.') + '$', 'i').test(e);
  const noiNhan = rows(TA_SH.NOINHAN).filter(r => String(r[0]).trim() && String(r[1]).trim()).map(r => {
    const em = list(r[4]), cc = list(r[5]);
    return { pb: String(r[0]).trim(), fileId: String(r[1]).trim(), prefix: String(r[2] || '').trim().toUpperCase() || 'TH', kinhGui: String(r[3] || '').trim() || String(r[0]).trim(),
      emails: em.filter(okMail), cc: cc.filter(okMail), emailBad: em.concat(cc).filter(e => !okMail(e)), on: r[6] !== false };
  });
  const banks = rows(TA_SH.NGANHANG).filter(r => String(r[0]).trim() && r[4] !== false).map(r => ({
    ten: String(r[0]).trim(), full: String(r[1] || '').trim(), bins: list(r[2]).map(x => x.replace(/\D/g, '')).filter(x => x.length >= 4), alias: list(r[3]) }));
  const atms = rows(TA_SH.ATM).filter(r => String(r[1]).trim() && r[3] !== false).map(r => ({ ma: String(r[0]).trim(), diaDiem: String(r[1]).trim(), pb: String(r[2] || '').trim() }));
  const thuKey = taKhongDau(kv.NHAC_HUY_THU || '').replace(/\s+/g, '');
  const thuMap = { chunhat: 0, cn: 0, thuhai: 1, t2: 1, '2': 1, thuba: 2, t3: 2, '3': 2, thutu: 3, t4: 3, '4': 3, thunam: 4, t5: 4, '5': 4, thusau: 5, t6: 5, '6': 5, thubay: 6, t7: 6, '7': 6 };
  const nhacThu = !thuKey ? null : /hangngay|moingay/.test(thuKey) ? 'all' : (thuKey in thuMap ? thuMap[thuKey] : 1);
  const num = (v, d, lo, hi) => { const x = parseInt(v, 10); return isNaN(x) ? d : Math.max(lo, Math.min(hi, x)); };
  const cfg = {
    soNgayHuy: num(kv.SO_NGAY_HUY, 90, 1, 3650), nhacThu, nhacGio: num(kv.NHAC_HUY_GIO, 8, 0, 23),
    nhomTon: list(kv.NHOM_TUOI_TON || '15,30,60,90').map(Number).filter(x => x > 0).sort((a, b) => a - b),
    soDau: num(kv.SO_CHU_SO_DAU, 6, 4, 6) >= 6 ? 6 : 4, mien, guiEmailNhap: kv.GUI_EMAIL_KHI_NHAP === undefined ? true : yes(kv.GUI_EMAIL_KHI_NHAP),
    vtNhap: list(kv.VAI_TRO_NHAP || 'Quản trị hệ thống, Điện toán'), vtTra: list(kv.VAI_TRO_TRA || 'Trưởng phòng, Phó phòng'),
    vtHuy: list(kv.VAI_TRO_HUY || 'Trưởng phòng, Phó phòng'), vtXem: list(kv.VAI_TRO_XEM_TAT_CA),
    hoanTac: kv.HOAN_TAC_TRONG_NGAY === undefined ? true : yes(kv.HOAN_TAC_TRONG_NGAY),
    diaDanh: kv.DIA_DANH || '', tenCN: kv.TEN_CHI_NHANH || 'Chi nhánh Tây Tiền Giang',
    bm: { ma: kv.MAU_PHIEU_MA || 'BM09-GNT', banHanh: kv.MAU_PHIEU_BAN_HANH || '02', suaDoi: kv.MAU_PHIEU_SUA_DOI || '00' },
    noiNhan, banks, atms };
  try { cache.put('ta_cfg', JSON.stringify(cfg), 60); } catch (e) {}
  return cfg;
}
/** Quyền trong ứng dụng thẻ ATM */
function taRole_(u, cfg) {
  const nn = cfg.noiNhan.find(x => x.on && x.pb === u.pb);
  const canEnter = cfg.vtNhap.includes(u.vt), canReturn = !!nn && cfg.vtTra.includes(u.vt), canCancel = !!nn && cfg.vtHuy.includes(u.vt);
  const viewAll = canEnter || cfg.vtXem.includes(u.vt);
  const scope = viewAll ? cfg.noiNhan.filter(x => x.on).map(x => x.pb) : (nn ? [u.pb] : []);
  return { canEnter, canReturn, canCancel, viewAll, scope, myPb: nn ? u.pb : '' };
}
/** Ẩn I006/I007 với cán bộ không thuộc phòng nhận thẻ và không có vai trò liên quan */
function taFilterApps_(u, apps) {
  if (!apps.includes(TA_APP) && !apps.includes(TA_APP_BC)) return apps;
  let r; try { r = taRole_(u, taConfig_()); } catch (e) { return apps.filter(a => a !== TA_APP && a !== TA_APP_BC); }
  return apps.filter(a => a === TA_APP ? (r.canEnter || r.canReturn) : a === TA_APP_BC ? r.scope.length > 0 : true);
}
function taGuard_(b, app) {
  const v = checkToken_(b.token); if (v.error) return v;
  if (!userApps_(v.user).includes(app)) return { error: 'forbidden' };
  let cfg; try { cfg = taConfig_(); } catch (e) { return { error: 'not_configured', message: 'Chưa cài đặt ứng dụng thẻ ATM.' }; }
  const role = taRole_(v.user, cfg);
  if (!role.scope.length || (app === TA_APP && !role.canEnter && !role.canReturn))
    return { error: 'forbidden', message: 'Tài khoản không thuộc phòng nhận thẻ hoặc chưa có vai trò trong ứng dụng thẻ ATM.' };
  return { user: v.user, cfg, role };
}

/* ---------- Đọc / ghi sổ ---------- */
function taLastRow_(sh) {
  const v = sh.getRange(1, 1, sh.getMaxRows(), 1).getValues();
  for (let i = v.length - 1; i >= 0; i--) if (v[i][0] !== '' && v[i][0] !== null) return i + 1;
  return 1;
}
function taYmd_(d) {
  if (d instanceof Date) return Utilities.formatDate(d, TZ, 'yyyy-MM-dd');
  const m = String(d || '').trim().match(/^(\d{1,2})\/(\d{1,2})\/(\d{2,4})/);
  if (!m) return '';
  const y = m[3].length === 2 ? '20' + m[3] : m[3];
  return y + '-' + m[2].padStart(2, '0') + '-' + m[1].padStart(2, '0');
}
function taBankOf_(name, cfg) {
  const k = taBankKey(name); if (!k) return '';
  const b = cfg.banks.find(x => taBankKey(x.ten) === k || taBankKey(x.full) === k || x.alias.some(a => taBankKey(a) === k));
  return b ? b.ten : String(name).trim();
}
/** Đọc sổ của 1 phòng → mảng thẻ. sync=true: bổ sung mã/số/ngân hàng/trạng thái cho dòng mới từ Google Form */
function taReadRows_(sh, n, cfg, sync) {
  const last = taLastRow_(sh); if (last < 2) return [];
  const v = sh.getRange(2, 1, last - 1, TA_NCOL).getValues();
  const today = Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd'), out = [];
  const used = new Set(v.map(r => String(r[TC.MA]).trim()).filter(Boolean));
  v.forEach((r, i) => {
    if (r[TC.TS] === '' || r[TC.TS] === null) return;
    if (!String(r[TC.MA]).trim() && sync) {
      const ts = r[TC.TS] instanceof Date ? r[TC.TS] : new Date();
      const pm = taParseMask(r[TC.SO_GOC]);
      const nh = taBankOf_(String(r[TC.NH_GOC]).trim() === 'Khác' ? r[TC.NH_KHAC] : r[TC.NH_GOC], cfg);
      const tt = String(r[TC.TT]).trim() || (r[TC.NGAY_TRA] !== '' ? TA_ST.TRA : TA_ST.GIU);
      const fill = [taNewMa_(n.prefix, ts, used), pm ? taMask(pm.dau, pm.cuoi) : '', nh, tt];
      sh.getRange(i + 2, TC.MA + 1, 1, 4).setValues([fill]);
      [r[TC.MA], r[TC.SO], r[TC.NH], r[TC.TT]] = fill;
    }
    const vao = taYmd_(r[TC.NGAY_GIAO]) || taYmd_(r[TC.TS]), tra = taYmd_(r[TC.NGAY_TRA]);
    const xl = String(r[TC.XL_BOI] || ''), xlId = (xl.match(/\((9\d{4})\)/) || [])[1] || '';
    out.push({ row: i + 2, pb: n.pb, ma: String(r[TC.MA]).trim(), nhapLuc: r[TC.TS] instanceof Date ? Utilities.formatDate(r[TC.TS], TZ, 'dd/MM/yyyy HH:mm') : String(r[TC.TS]),
      ngayVao: vao, ngayTra: tra, hoTen: String(r[TC.TEN] || '').trim(), soThe: String(r[TC.SO] || '').trim() || String(r[TC.SO_GOC] || '').trim(),
      nganHang: String(r[TC.NH] || '').trim() || String(r[TC.NH_GOC] || '').trim(), tt: String(r[TC.TT] || '').trim() || TA_ST.GIU,
      giaoBoi: String(r[TC.GIAO_BOI] || '').trim(), ghiChu: String(r[TC.GHI_CHU] || '').trim(), atm: String(r[TC.ATM] || '').trim(),
      xlBoi: xl.replace(/\s*\(9\d{4}\)\s*$/, ''), xlId, xlLuc: r[TC.XL_LUC] instanceof Date ? Utilities.formatDate(r[TC.XL_LUC], TZ, 'dd/MM/yyyy HH:mm') : String(r[TC.XL_LUC] || ''),
      xlNgay: r[TC.XL_LUC] instanceof Date ? Utilities.formatDate(r[TC.XL_LUC], TZ, 'yyyy-MM-dd') : '', xlGhiChu: String(r[TC.XL_GHI_CHU] || '').trim(),
      canhBao: String(r[TC.CANH_BAO] || '').trim(), soNgay: taDays(vao, tra || today) });
  });
  return out;
}
/** Mã thẻ = tiền tố + thời điểm vào sổ (yyMMdd-HHmmss); trùng thì cộng thêm giây */
function taNewMa_(prefix, ts, used) {
  let t = ts.getTime(), ma;
  do { ma = prefix + '-' + Utilities.formatDate(new Date(t), TZ, 'yyMMdd-HHmmss'); t += 1000; } while (used.has(ma));
  used.add(ma); return ma;
}
function taSheet_(n) {
  const sh = SpreadsheetApp.openById(n.fileId).getSheets()[0];
  if (sh.getMaxColumns() < TA_NCOL) throw new Error('so_not_ready');
  return sh;
}
/** Đọc thẻ trong phạm vi (có khóa ghi để bổ sung dòng mới an toàn) */
function taCards_(cfg, pbs) {
  const lock = LockService.getScriptLock(); let locked = false;
  try { lock.waitLock(8000); locked = true; } catch (e) {}
  try {
    return cfg.noiNhan.filter(n => n.on && pbs.includes(n.pb)).reduce((a, n) => a.concat(taReadRows_(taSheet_(n), n, cfg, locked)), []);
  } finally { if (locked) lock.releaseLock(); }
}
function taPublicCfg_(cfg) {
  return { soNgayHuy: cfg.soNgayHuy, soDau: cfg.soDau, nhomTon: cfg.nhomTon, mien: cfg.mien, hoanTac: cfg.hoanTac, diaDanh: cfg.diaDanh, tenCN: cfg.tenCN, bm: cfg.bm,
    banks: cfg.banks, atms: cfg.atms, phong: cfg.noiNhan.filter(n => n.on).map(n => ({ pb: n.pb, kinhGui: n.kinhGui, coEmail: n.emails.length > 0 })) };
}
function taOut_(x, user, cfg, today) {
  const o = Object.assign({}, x); delete o.row; delete o.xlId;
  o.canUndo = cfg.hoanTac && x.tt === TA_ST.TRA && x.xlId === user.id && x.xlNgay === today;
  o.cuaToi = x.xlId === user.id;
  return o;
}

/* ---------- API ---------- */
/** view: 'so' (I006: đang giữ + đã trả 60 ngày), 'huy' (thẻ đã hủy), 'bc' (I007: tất cả) */
function apiTaLoad_(b) {
  const view = ['so', 'huy', 'bc'].includes(b.view) ? b.view : 'so';
  const g = taGuard_(b, view === 'bc' ? TA_APP_BC : (b.app === TA_APP_BC ? TA_APP_BC : TA_APP)); if (g.error) return g;
  const { user, cfg, role } = g, today = Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd');
  let cards;
  try { cards = taCards_(cfg, role.scope); } catch (e) { if (String(e.message) === 'so_not_ready') return { error: 'not_configured', message: 'Sổ thẻ chưa được chuẩn bị. Quản trị chạy menu Cài đặt thẻ ATM.' }; throw e; }
  if (view === 'so') cards = cards.filter(x => x.tt === TA_ST.GIU || (x.tt === TA_ST.TRA && taDays(x.ngayTra || x.xlNgay, today) <= 60));
  if (view === 'huy') cards = cards.filter(x => x.tt === TA_ST.HUY);
  cards.sort((a, c) => (c.ngayVao + c.nhapLuc).localeCompare(a.ngayVao + a.nhapLuc));
  return { ok: true, today, view, me: user.hoTen, role: { canEnter: role.canEnter, canReturn: role.canReturn, canCancel: role.canCancel, viewAll: role.viewAll, scope: role.scope, myPb: role.myPb },
    cfg: taPublicCfg_(cfg), cards: cards.map(x => taOut_(x, user, cfg, today)) };
}

/** Nhập 1 lô thẻ cho 1 phòng nhận; gửi 1 email cho cán bộ phụ trách */
function apiTaAdd_(b) {
  const g = taGuard_(b, TA_APP); if (g.error) return g;
  const { user, cfg, role } = g;
  if (!role.canEnter) return { error: 'forbidden', message: 'Vai trò của tài khoản không được nhập thẻ.' };
  const n = cfg.noiNhan.find(x => x.on && x.pb === String(b.pb || ''));
  if (!n) return { error: 'invalid', message: 'Chọn phòng nhận thẻ.' };
  const list = Array.isArray(b.cards) ? b.cards.slice(0, 30) : [];
  if (!list.length) return { error: 'invalid', message: 'Chưa có thẻ nào trong lô.' };
  const clean = [];
  for (let i = 0; i < list.length; i++) {
    const c = list[i] || {}, dau = String(c.dau || '').replace(/\D/g, ''), cuoi = String(c.cuoi || '').replace(/\D/g, '');
    if (dau.length !== cfg.soDau || cuoi.length !== 4) return { error: 'invalid', message: 'Thẻ thứ ' + (i + 1) + ': cần ' + cfg.soDau + ' số đầu và 4 số cuối.' };
    const nh = taBankOf_(c.nganHang, cfg);
    if (!nh) return { error: 'invalid', message: 'Thẻ thứ ' + (i + 1) + ': chọn ngân hàng chủ thẻ.' };
    const ten = taKhongDau(c.hoTen || '').toUpperCase().replace(/[^A-Z0-9 .'-]/g, ' ').replace(/\s+/g, ' ').trim().slice(0, 60);
    const atm = String(c.atm || '').trim().slice(0, 120), ghiChu = String(c.ghiChu || '').trim().slice(0, 200);
    clean.push({ dau, cuoi, so: taMask(dau, cuoi), nh, ten, atm, ghiChu });
  }
  const lock = LockService.getScriptLock();
  try { lock.waitLock(10000); } catch (e) { return { error: 'busy', message: 'Hệ thống đang ghi dữ liệu khác, thử lại sau vài giây.' }; }
  let added = [];
  try {
    const sh = taSheet_(n), all = taReadRows_(sh, n, cfg, true), cur = all.filter(x => x.tt === TA_ST.GIU), used = new Set(all.map(x => x.ma));
    // Trùng: cùng 4 số đầu, 4 số cuối, họ tên với thẻ đang giữ ở phòng này
    const k = (so, ten) => { const p = taParseMask(so); return p ? p.dau.slice(0, 4) + p.cuoi + '|' + taKhongDau(ten) : ''; };
    const curKeys = cur.reduce((o, x) => (o[k(x.soThe, x.hoTen)] = x, o), {});
    const dups = clean.map((c, i) => ({ i, x: curKeys[k(c.so, c.ten)] })).filter(d => d.x);
    if (dups.length && !b.xacNhanTrung)
      return { error: 'duplicate', message: dups.map(d => 'Thẻ thứ ' + (d.i + 1) + ' trùng thẻ đang giữ ' + d.x.ma + ' (vào sổ ' + d.x.ngayVao.split('-').reverse().join('/') + ').').join(' ') + ' Nếu đúng là thẻ bị thu lần nữa, bấm "Vẫn lưu".' };
    const now = new Date(), day = new Date(Utilities.formatDate(now, TZ, 'yyyy-MM-dd') + 'T12:00:00+07:00');
    const row0 = taLastRow_(sh) + 1;
    if (row0 + clean.length > sh.getMaxRows()) sh.insertRowsAfter(sh.getMaxRows(), clean.length + 20);
    const txt = v => v === '' ? '' : "'" + v;
    const left = [], right = [];
    clean.forEach((c, i) => {
      const ts = new Date(now.getTime() + i * 1000), ma = taNewMa_(n.prefix, ts, used);
      left.push([ts, user.hoTen, day, '', c.ten, txt(c.so), c.nh, '', c.ghiChu, txt(ma), txt(c.so), c.nh, TA_ST.GIU, '', '', '']);
      right.push([c.atm, txt(user.id)]);
      added.push(Object.assign({ ma }, c));
    });
    sh.getRange(row0, 1, clean.length, 16).setValues(left);           // A–P (bỏ qua Q: công thức cảnh báo)
    sh.getRange(row0, TC.ATM + 1, clean.length, 2).setValues(right);  // R–S
    log_(user.id, user.hoTen, 'Nhập thẻ ATM', TA_APP, n.pb + ' · ' + added.length + ' thẻ · ' + added.map(a => a.ma).join(', '), '');
  } finally { lock.releaseLock(); }
  // Email cho cán bộ phụ trách phòng nhận (ngoài khóa ghi)
  let mail = '';
  if (cfg.guiEmailNhap) {
    if (!n.emails.length) mail = 'Chưa gửi email: ' + n.pb + ' chưa có email nhận trong cấu hình.';
    else {
      try {
        const rows = added.map((a, i) => '<tr><td>' + (i + 1) + '</td><td>' + taEsc(a.ten || '(không tên)') + '</td><td>' + taEsc(taMaskView(a.so)) + '</td><td>' + taEsc(a.nh) + '</td><td>' +
          taEsc(a.atm) + '</td><td>' + taEsc(a.ghiChu) + '</td></tr>').join('');
        MailApp.sendEmail({ to: n.emails.join(','), cc: n.cc.join(','), name: 'Không Gian Số Nội Bộ',
          subject: '[Thẻ ATM máy thu] ' + n.pb + ' nhận ' + added.length + ' thẻ ngày ' + Utilities.formatDate(new Date(), TZ, 'dd/MM/yyyy'),
          htmlBody: '<p>Kính gửi ' + taEsc(n.pb) + ',</p><p>Cán bộ ' + taEsc(user.hoTen) + ' đã bàn giao <b>' + added.length + ' thẻ ATM bị máy thu</b> vào sổ của phòng lúc ' +
            Utilities.formatDate(new Date(), TZ, 'HH:mm dd/MM/yyyy') + ':</p><table border="1" cellpadding="5" style="border-collapse:collapse;font-family:Arial;font-size:13px">' +
            '<tr style="background:#005993;color:#fff"><th>STT</th><th>Họ tên trên thẻ</th><th>Số thẻ</th><th>Ngân hàng</th><th>ATM thu thẻ</th><th>Ghi chú</th></tr>' + rows + '</table>' +
            '<p>Thẻ đã có trong sổ trên Không Gian Số Nội Bộ → Thẻ ATM → Quản lý thẻ ATM bị máy thu.</p><p style="color:#888;font-size:12px">Email gửi tự động – ' + taEsc(cfg.tenCN) + '.</p>' });
        mail = 'Đã gửi email cho ' + n.emails.join(', ') + '.';
      } catch (e) { mail = 'Chưa gửi được email: ' + e.message; }
    }
  }
  return Object.assign(apiTaLoad_({ token: b.token, view: 'so' }), { added: added.map(a => a.ma), mail });
}

function taFind_(b, g, pbOverride) {
  const pb = pbOverride || g.role.myPb, n = g.cfg.noiNhan.find(x => x.on && x.pb === pb);
  if (!n) return { error: 'forbidden', message: 'Tài khoản không thuộc phòng nhận thẻ.' };
  const sh = taSheet_(n), x = taReadRows_(sh, n, g.cfg, true).find(r => r.ma === String(b.ma || ''));
  if (!x) return { error: 'invalid', message: 'Không tìm thấy thẻ trong sổ của phòng.' };
  return { n, sh, x };
}
/** Trả thẻ cho khách */
function apiTaReturn_(b) {
  const g = taGuard_(b, TA_APP); if (g.error) return g;
  const { user } = g;
  if (!g.role.canReturn) return { error: 'forbidden', message: 'Vai trò của tài khoản không được trả thẻ.' };
  const lock = LockService.getScriptLock();
  try { lock.waitLock(10000); } catch (e) { return { error: 'busy', message: 'Hệ thống đang bận, thử lại sau vài giây.' }; }
  let ma;
  try {
    const f = taFind_(b, g); if (f.error) return f;
    if (f.x.tt !== TA_ST.GIU) return { error: 'invalid', message: 'Thẻ này đang ở trạng thái "' + f.x.tt + '".' };
    const now = new Date(), day = new Date(Utilities.formatDate(now, TZ, 'yyyy-MM-dd') + 'T12:00:00+07:00');
    f.sh.getRange(f.x.row, TC.NGAY_TRA + 1).setValue(day);
    f.sh.getRange(f.x.row, TC.TT + 1, 1, 4).setValues([[TA_ST.TRA, user.hoTen + ' (' + user.id + ')', now, String(b.ghiChu || '').trim().slice(0, 200)]]);
    log_(user.id, user.hoTen, 'Trả thẻ ATM', TA_APP, f.x.ma + ' · ' + f.n.pb, '');
    ma = f.x.ma;
  } finally { lock.releaseLock(); }
  return Object.assign(apiTaLoad_({ token: b.token, view: 'so' }), { returned: ma });
}
/** Hoàn tác trả thẻ (người trả, trong ngày) */
function apiTaUndo_(b) {
  const g = taGuard_(b, TA_APP); if (g.error) return g;
  const { user, cfg } = g;
  const lock = LockService.getScriptLock();
  try { lock.waitLock(10000); } catch (e) { return { error: 'busy', message: 'Hệ thống đang bận, thử lại sau vài giây.' }; }
  try {
    const f = taFind_(b, g); if (f.error) return f;
    const today = Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd');
    if (!(cfg.hoanTac && f.x.tt === TA_ST.TRA && f.x.xlId === user.id && f.x.xlNgay === today))
      return { error: 'forbidden', message: 'Chỉ người trả thẻ được hoàn tác, trong ngày trả.' };
    f.sh.getRange(f.x.row, TC.NGAY_TRA + 1).setValue('');
    f.sh.getRange(f.x.row, TC.TT + 1, 1, 4).setValues([[TA_ST.GIU, '', '', '']]);
    log_(user.id, user.hoTen, 'Hoàn tác trả thẻ ATM', TA_APP, f.x.ma, '');
  } finally { lock.releaseLock(); }
  return apiTaLoad_({ token: b.token, view: 'so' });
}
/** Dữ liệu điền phiếu BM09 từ 1 thẻ đã trả */
function taPhieuData_(x, n, cfg) {
  const bank = cfg.banks.find(k => k.ten === x.nganHang), ngay = (x.xlLuc.match(/(\d{2})\/(\d{2})\/(\d{4})/) || []).slice(1);
  return { kinhGui: n.kinhGui, hoTen: x.hoTen, soThe: x.soThe, nganHang: bank && bank.full ? bank.full + ' (' + bank.ten + ')' : x.nganHang,
    atm: x.atm, diaDanh: cfg.diaDanh, ngay: ngay.length ? ngay : ['', '', ''], gioTra: x.xlLuc ? x.xlLuc.slice(11) + ' ' + x.xlLuc.slice(0, 10) : '', gdv: x.xlBoi, maThe: x.ma, bm: cfg.bm };
}
/** Gửi phiếu BM09 dạng PDF qua email (dự phòng khi máy in WiFi lỗi) */
function apiTaPhieuEmail_(b) {
  const g = taGuard_(b, TA_APP); if (g.error) return g;
  const { user, cfg } = g, email = String(b.email || '').trim().toLowerCase();
  if (!g.role.canReturn) return { error: 'forbidden', message: 'Vai trò của tài khoản không được trả thẻ.' };
  if (!new RegExp('^[^@\\s]+@' + cfg.mien.replace(/\./g, '\\.') + '$').test(email)) return { error: 'invalid', message: 'Chỉ gửi được tới email @' + cfg.mien + '.' };
  const f = taFind_(b, g); if (f.error) return f;
  if (f.x.tt !== TA_ST.TRA) return { error: 'invalid', message: 'Thẻ chưa ở trạng thái đã trả.' };
  const html = '<html><head><meta charset="utf-8"><style>@page{size:A4;margin:12mm}body{margin:0}' + TA_PHIEU_CSS + '</style></head><body>' + taPhieuHtml(taPhieuData_(f.x, f.n, cfg)) + '</body></html>';
  const name = 'Phieu_tra_the_' + f.x.ma + '.pdf';
  const pdf = Utilities.newBlob(html, 'text/html', 'phieu.html').getAs('application/pdf').setName(name);
  MailApp.sendEmail({ to: email, name: 'Không Gian Số Nội Bộ', subject: '[Thẻ ATM máy thu] Phiếu trả thẻ ' + f.x.ma,
    htmlBody: '<p>Đính kèm Giấy đề nghị trả thẻ do ATM thu (' + taEsc(cfg.bm.ma) + ') của thẻ ' + taEsc(f.x.ma) + ', in để khách hàng ký.</p><p style="color:#888;font-size:12px">Email gửi tự động – ' + taEsc(cfg.tenCN) + '.</p>',
    attachments: [pdf] });
  log_(user.id, user.hoTen, 'Gửi phiếu trả thẻ', TA_APP, f.x.ma + ' → ' + email, '');
  return { ok: true, message: 'Đã gửi phiếu PDF tới ' + email + '.' };
}
/** Hủy các thẻ quá hạn (sau khi đã rà soát, kiểm kê) */
function apiTaCancel_(b) {
  const g = taGuard_(b, TA_APP_BC); if (g.error) return g;
  const { user, cfg, role } = g;
  if (!role.canCancel) return { error: 'forbidden', message: 'Vai trò của tài khoản không được hủy thẻ.' };
  const mas = Array.isArray(b.mas) ? b.mas.map(String).slice(0, 300) : [], ghiChu = String(b.ghiChu || '').trim().slice(0, 200);
  if (!mas.length) return { error: 'invalid', message: 'Chưa chọn thẻ cần hủy.' };
  if (ghiChu.length < 5) return { error: 'invalid', message: 'Ghi số biên bản hoặc căn cứ hủy (tối thiểu 5 ký tự).' };
  const lock = LockService.getScriptLock();
  try { lock.waitLock(15000); } catch (e) { return { error: 'busy', message: 'Hệ thống đang bận, thử lại sau vài giây.' }; }
  let ok = 0, skip = 0;
  try {
    const n = cfg.noiNhan.find(x => x.on && x.pb === role.myPb), sh = taSheet_(n), rows = taReadRows_(sh, n, cfg, true), now = new Date();
    rows.filter(x => mas.includes(x.ma)).forEach(x => {
      if (x.tt !== TA_ST.GIU || x.soNgay < cfg.soNgayHuy) { skip++; return; }
      sh.getRange(x.row, TC.TT + 1, 1, 4).setValues([[TA_ST.HUY, user.hoTen + ' (' + user.id + ')', now, ghiChu]]); ok++;
    });
    log_(user.id, user.hoTen, 'Hủy thẻ ATM quá hạn', TA_APP_BC, role.myPb + ' · ' + ok + ' thẻ · ' + ghiChu, '');
  } finally { lock.releaseLock(); }
  return Object.assign(apiTaLoad_({ token: b.token, view: 'bc' }), { cancelled: ok, skipped: skip });
}

/** Lịch hằng ngày: đúng thứ cấu hình thì gửi mỗi phòng danh sách thẻ đến hạn hủy. Chạy tay từ menu để gửi ngay. */
function nhacHuyTheATM(e) {
  const cfg = taConfig_(true), manual = !(e && e.triggerUid);
  const dow = +Utilities.formatDate(new Date(), TZ, 'u') % 7; // u: 1=Thứ Hai … 7=Chủ nhật
  if (!manual && (cfg.nhacThu === null || (cfg.nhacThu !== 'all' && cfg.nhacThu !== dow))) return;
  const report = [];
  cfg.noiNhan.filter(n => n.on).forEach(n => {
    try {
      const due = taReadRows_(taSheet_(n), n, cfg, true).filter(x => x.tt === TA_ST.GIU && x.soNgay >= cfg.soNgayHuy).sort((a, b) => b.soNgay - a.soNgay);
      if (!due.length) { report.push(n.pb + ': không có thẻ đến hạn hủy'); return; }
      if (!n.emails.length) { report.push(n.pb + ': ' + due.length + ' thẻ đến hạn, chưa có email nhận'); return; }
      const rows = due.map((x, i) => '<tr><td>' + (i + 1) + '</td><td>' + taEsc(x.ma) + '</td><td>' + taEsc(x.hoTen) + '</td><td>' + taEsc(taMaskView(x.soThe)) + '</td><td>' + taEsc(x.nganHang) +
        '</td><td>' + x.ngayVao.split('-').reverse().join('/') + '</td><td style="text-align:right">' + x.soNgay + '</td></tr>').join('');
      MailApp.sendEmail({ to: n.emails.join(','), cc: n.cc.join(','), name: 'Không Gian Số Nội Bộ',
        subject: '[Thẻ ATM máy thu] ' + n.pb + ': ' + due.length + ' thẻ quá ' + cfg.soNgayHuy + ' ngày cần rà soát hủy',
        htmlBody: '<p>Kính gửi ' + taEsc(n.pb) + ',</p><p>Sổ thẻ của phòng có <b>' + due.length + ' thẻ</b> khách chưa nhận lại đã giữ từ ' + cfg.soNgayHuy +
          ' ngày trở lên. Đề nghị rà soát, kiểm kê và thực hiện hủy theo quy định tại Không Gian Số Nội Bộ → Thẻ ATM → Báo cáo thẻ ATM máy thu.</p>' +
          '<table border="1" cellpadding="5" style="border-collapse:collapse;font-family:Arial;font-size:13px"><tr style="background:#005993;color:#fff"><th>STT</th><th>Mã thẻ</th><th>Họ tên trên thẻ</th><th>Số thẻ</th><th>Ngân hàng</th><th>Ngày vào sổ</th><th>Số ngày giữ</th></tr>' +
          rows + '</table><p style="color:#888;font-size:12px">Email gửi tự động – ' + taEsc(cfg.tenCN) + '.</p>' });
      report.push(n.pb + ': đã gửi ' + due.length + ' thẻ tới ' + n.emails.join(', '));
    } catch (err) { report.push(n.pb + ': lỗi ' + err.message); }
  });
  log_('', 'Hệ thống', 'Nhắc hủy thẻ ATM', TA_APP_BC, report.join(' | '), '');
  if (manual) { try { SpreadsheetApp.getUi().alert('Email nhắc hủy thẻ', report.join('\n'), SpreadsheetApp.getUi().ButtonSet.OK); } catch (err) {} }
}

/* ===== Dùng chung máy chủ (Code.gs) và trang (index.html) – Thẻ ATM bị máy thu (I006, I007) ===== */
const TA_ST = { GIU: 'Đang giữ', TRA: 'Đã trả khách', HUY: 'Đã hủy' };
const TA_LOGO = 'data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAWgAAABbCAMAAABpou9QAAAASFBMVEX////+///+/v78/f76+/zi6vG8zuCfss1olbw3baEUWZQRVpIPWZkPWJcPVpEOVpPjG1HcD0fZD0YwSYQOVJEOVJAOVI8OU4+FcLcfAAAPKElEQVR42u1d2bKjug4lxkw2Nuku4/3/f3oleQaTgZ101bmFzsPpnTCYhaxhSXaa5pL3hbUt/q/rh3GcSMZx6Dv8jLfswudT0hLI4yQUyEKC/xLTSGC3F0KfgrkFlNWslh8pjSUxUv4s8BFgzS+oP2M1Gt5Py7xIudq1EG20XADrgTeX/fiAOveTUqDCa1WMNUpNfXNB/Tt1Zk03AsxyfSDSKjV2F9K/U+dBzOsLMovhstS/wLkbZ21fAdqaGZX6wuwkztMszfqSWDlPXcMv1M7gDF5Qri+LRJ94If228KYXyxs4A9KL6C/rcUKfxbK+KRfS78d1YJ/fxhmR7poryntDboxPSr8PNNjprrld+L2sz5yNs1xPiJ3H9spc3jDQwzmcQafn4TLTbxhoIfU5oLUUV+LyOsMxqpMKDaLGy3a8HNmp9bzI+YrxXtRniDjMeaDNMnF26fRXPaGTH/CHVyr+AtB8Ur8CelXTVbF9JeTol/WXoi4r/YrlOMhVjKyIqbvD8UrEXyI5aly/XeaKqPqhoruQfB7bVRXaLtNYkfpLwfTwkueWQ7+OXT1AQdtxyRPh02LqQN8gPC6F3+pAYyh9IfmU5jD6VbLoKOTWRvQXlo/1+Sj9fgtoCPAuI/0M6APo3gNaqstIPzMd4yeAXi+gnzJK4ydMBwSDF5aPFfqIuXsTaHMB/RToagryrukw4gLz3wAtLzD/CdDrBfRlOv6fnKG+gH4K9XSFd/+hhMVcQD+R9jNAYwreJqmXW9jjr7+nTO1O2JfK9i0/qp4ecB1WiiOghbQ1oP9zpNL33nf1wtgWbbcgW7kodQS0UgscsD1H9E2fpKsoDGu6+D37p8aR9zvpOt58fgUfPOLQ43pXVv1O5PGdkRLXx4ppHCr9uDe4Ei5cngHsvE6rsWg4zbSOGV7EPFaK4qz1B9S//qJx7GcpyhKzEO4BPzwONA8AHa53rSGdujpAUZd5nqah7x4UTHjXD9Ok5iXTawW+sFcAPoqUlabpFhsp6XsJ2s/+KdBqtRtZF1AnMcIw2YeBRginWnd+8m9yXWipd+fNOmuweuVch3cg8BF3r6CjxeLLKlOBkcdAseZIo9eVauL/0nSgRuudWCO1WyzJPgy0NtSdz/Ya3QsNk19b5RfUo5d2JcKK/jMHPn7FAetZWa1XbamSNcw2KjjbmZ1Qm7RvN1SnmuVZjbYHHcfrPH6ybTBEFnKu6BK1OEq4pZzQaAHK6ZF414EDG1DIgXiDQocQ1t0wyXmVXkk7oV31UevdIqLYsWr/9bKXY6BpseTIPucSYwhXbbBF2wGTaOxbCnkcyC3Y4TvKn0zog6HrWq9lrdtvQsDpqKRt6rKWO38HlkPJk11NXRB+5g2VQFsyHJlSw9C/AXTF37esE+gWSFEJw66PCP8tJOLt7Hjr9LobhcB4Dp/oJ2bkJSgY3Lgns++umIPcVVHQIOZT60cLoI2TdZUxXpo+N8Ei0Lvn94K4eV1uu8GB/PdQHNgDKjbpNVqQYIQWryw/quzkTe9aTW9O1UgSyA8ALZxIFQke/cGe4wzoqn1kDmayufeHGJdg4/vxULNNOi/L9Rb0Dsy552LJ6PwaaNl7IzRM4bNPNmg+AxozJIIZLcbflwVtCLnPhrWlfUBAizslV6jfdoWfA1qvMvZi8tGbOXMwy78CNB3Uv6jMpV7f++zZb210h7pwB7HBTx/rD/hWzisExLtAMyJ2WB3onrk2Nx55eLsz0vAtPyaHiiN5eVgC+lCfGML894Qg1KzCUIGT4THrwqxQ0xdG9c7hbhD127/54LFEv41AD23+JjCL2pKCrL1tmZ2NRrdbTEyJCWsLcigxf8Wt8nGmAVeBZgVpyO/3vyflfo/Bed5rLbNVADH79OuKtjkCc9u/YcTutnsrv8+cYRWQwtvgXn2O2WkfAd0y/2k5y3FkSDHgSB5vhtb6A/uONfu3ly7KttTS/c85nP/cs3Hm7jA3EmGeuty8JTQS4cFclinUPKtFTInsQTYMJPjRZcS/3O3cNQZOPJcXeGQ+uMvQVdhjoOd9JIpXdQOZkY3omnTxnofhoHC8bbxVIJEqQGfn0NhhNOdMx98/OUMEwxdWhypifFT41KQZBUn/CEMck8qxYZqRowIB40JkDyk9HDnDs8QlNvhcLoHHTYmmZRadY8wIGFD4fprVSpdR8zS4i1SBBiMd58nIshkJ6Ze7ApJDYoABuIUOkFQ3eI6iP+A6eKufcCs/efdA06Y+KlwD833WDOdsdElc3NIKr8QstflDgVkYlPqBHCl+C2NB1tW9IAv/IUhN63gYfORUX5CS8h2GzwnneKAd96l6eC3SLbXWxhIby/bO0NtYuPI2juZ4VWtN0BQYyNh5glUQ0Fj2oL+6AW7vDtRILLtH3QGNYe0cqFlsT8TPmuF96/HnPpQGt03Mkla+OT2jvFXvgdExSaUNhspN9rQ08PLBreGCsS1NgSwCOlclLT0OXs+hKMZiBx0jg/JXwjs+hCGl1ahwJVVuwWMhHXBzkYLAgCM4+rHYcwPGS7faAn1reW5K25sjL5r3HeL9vnFseR3BLn35nskVIm1tExvQpg2GAIewd6S1pPsHQHPm5k0BNF6+3HnSOn+cJyw6LsWZjd34bDxO76v7et0BjVDbclySSIgN0CxzWTbxhCeQvt931OXGHbaBHgzGhHvmKXyJL8Zz2gZtmXGGwmn81nTggrzZTQq97oBe0TtYY2y2awvHp81S8GAuF3fRnyVySjgSve9w0WsNaPrYmnS4yw82QGfH25yPJTv95z37vCNz26SE2s2fIiukx0lvgcWNb0CHBS4DE7NJ6ghGE6tf0Rniv93LogfYAo23wEPijhg/iuxMzt4VyyW1nIIfSergiGoNvsJsFukUhWyDozHpTywblUCjKfJDQYI6ixlwqr5upyFTaWo12NT6ZMnPZKWVkVZyZERi/M4a5YoO3SA83wMZG/o8EOHmL5pG+KtvQhvbDmgMFODweU08xkajN9pK21GyfSMFqDzML5X2MNkDvSgc2Wxt5voLoMlI6uCuykQRBvVyPI3xcy1LzUZDU7dlnY+DyXIUjC1jnWteAGUfHMndxE20aOiU5I6Z5eE4A70f2JmORQwd5/Cu/G6UFCM/Iv4hmOh9KJD0Qy5ixEIHxJw/9gBoqya6FThFHaumG43ugymyldJWC/74Fagh3OD1vClf5qW9QQ1j7VgJNCq0DWuN3FuDI7pJO/Dd8HZcx+0AaKrrkAxCB5XtaBId7gJqFSo9K9ZM0V7CLjRRwT9sgMZn8bdSGQOfA827sLsd4dxWktjuCe3xBwmO7rCqyUp3yPPSStuWNQhPniKpijwOSpp/VLfB2kICGimhQ6BJ42/EGMUR6A3QqeMgfoRujMW3SSaLKk0wEIBRV6MOgw6EyCnmLXscSyD+u2DyibXitUpo0zwk8oiye9B8glM1TEJQSrAOfiKLfgM0D0ei4yv6Lh38ITB5TaMxFXU+g/JTr4l9CbSKi9qDy9SCLFroXtaYEfEmpY81oO0Siq885A1boEVoCZDL4Z5TZCjvdaypuNI3zYPScSL5nZLlXQZsA7QbpTViKPqInLk0jsDcAX1goxOxnFmvDdAyW9SuvMs00gVDq07jij4rkGQl0NRpGHoIlgrQeEJwE49KdwR17+uGRcHw7ssqj5n25JwB3YC655MKoOMUL7dQWJKFfQNoNe2zpgJoumCs+CaXiU0pwUQjj8v2WcEOaPYE6NUc9QPsoW58JTyVwB2J+awRgqW+gxXiAJGzkRsbnYiRQkJUtThe45NAcy+4Jbm3cDjXDkojLRsOgH6m0VmXLX/atRJaOzrq6XBNHa90sWSeZdWmSBNLoNvHm2a59PjDGs1SKOuvS4XwAKgqMos2hP3vAx2URbxQFtr2Br3aK1Rs7xaID8cpHAFtavJVoIn6l/7j/uNABybg5X6h0AL2TkMW229NYTyVd2A6jP9VnY3M3wQ6OUAM1rK47BOmw4jgjuU3d/baLyOQnvXdOkOvL2KqyxecIcsbh6IDjMfYZesM9Qmg4X0Nyhx1bX0S6NR3EJxv51nRTXinvbrXZVMFH9gHgQ53XzWatT4L7yIsN3YyvMPMMCS939wmF4xNuZVe7DLYJSwy0HqsqLDW2w3GoNHTaaBZVpsukjqPqM32tM1SzBNAR9bVfnGb3F2cE+j1jY32iBXtS9kSDO65DneKmxZo8H4DdBs7gIfoIMCKpkBJ+RSc8RvyQvok0F2ZTXxpcdJmj7e0e2lbJZUIwli445OiBjnfOF+oVf9bjU7LFmI1KrdiRGInUinkNGeAZmlXsCLd/KY7tFFjS6BjrmyISITZzCFiAfBRgmnDU37iTwqMY9+cBRruGreYm0KFwe0HhZyhzGhSSBz6AYLPOk36mkaz7JJf2721jVXm0Ih3qwCdvQ9QIze56LdgcMIiq3rzHJLOaM3hF0CvKkps3HUNKFviX+plVlkR7QTQ2SXt99bt5FslZ700G6DZLf5WxqImbDEZhQpFwz7rVrD5ZtW/AHqf5uMPbyD3yrJf7cD2NWOlPCplvajR+cC/tvVzriE6LTTYAO1W0LjmGuvaYKzcRp8Y8C55FfwXQFd+tGfum9BnU1mqan8DNEueSn8rmqa+RhsbPFhTB5oqmL7ESZvOumgW15bwtjjJfANoK5VIEwfXru3aDX4BdLEd6dfMdJw3xQ897YAma67KTXy13ayWwrMWf8zHgNa4dnVOP/flZros8yw1usl0Dujc6ln5pV9hoqmIa5mxg63NIRO4XCeNlbv2rkB3aWPplE0DFB5jJMYiBDTYWLiGDEV9ilJEDrSkz7QkoIXcLujEuYNrV/PKJ1YlhNJRq62G992HNbcBaDo7B1r6C3qg/fG+3nmjoZJPoK6yLyBNbWfUlFl0QYaWxDlvcuTDNKuFPJUx4O33K1p9UyMKRh3CVwj8w3nJgNbhM9ccUhP/k9BtQcL3bkWwb3JUI4/n+ibHuRw8SxeXxVhkaHJMN1efXdSYrxSldtfyWZo+RLJF226Hi55V+EnsprZIr+G9b32FF+OvQQqZ+mJTZlo0yuaNs1HcKmy+fZ/M9Q/jknUaSDqXhydyt2LNtiuXOhuyv/wBQ9G2+8UVli+stvaN6PH5a7Ri+/lZV1nHQcv58Ffn3c/Nf0ABP4/s/wDKirsv8AbQ0wAAAABJRU5ErkJggg==';
/** Bỏ dấu tiếng Việt, chữ thường – dùng để tìm kiếm */
function taKhongDau(s) {
  return String(s || '').normalize('NFD').replace(/[\u0300-\u036f]/g, '').replace(/\u0111/g, 'd').replace(/\u0110/g, 'D').toLowerCase();
}
/** Khóa so khớp tên ngân hàng: in hoa, bỏ dấu, bỏ khoảng trắng và . - */
function taBankKey(s) { return taKhongDau(s).toUpperCase().replace(/[\s.\-_]/g, ''); }
/** Số thẻ che: 6 số đầu → 970415xxxxxx2187 (16 ký tự, theo BM09); 4 số đầu → 9704xxxx2187 (như sổ cũ) */
function taMask(dau, cuoi) { dau = String(dau || ''); return dau + (dau.length >= 6 ? 'xxxxxx' : 'xxxx') + String(cuoi || ''); }
/** Tách số thẻ che (mọi kiểu gõ cũ: xxx, XXX, ***) → {dau, cuoi} hoặc null */
function taParseMask(s) { const m = String(s || '').match(/^\s*(\d{4,6})\s*[xX*]{2,}\s*(\d{4})\s*$/); return m ? { dau: m[1], cuoi: m[2] } : null; }
/** Hiển thị số thẻ che có nhóm: 970415 xxxxxx 2187 */
function taMaskView(m) { const r = String(m || '').match(/^(\d{4,6})(x+)(\d{4})$/i); return r ? r[1] + ' ' + r[2] + ' ' + r[3] : String(m || ''); }
/** Số ngày giữa 2 ngày yyyy-MM-dd */
function taDays(from, to) { if (!from || !to) return null; return Math.round((Date.parse(to + 'T00:00:00Z') - Date.parse(from + 'T00:00:00Z')) / 864e5); }
function taEsc(s) { return String(s == null ? '' : s).replace(/[&<>"']/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c])); }

/** Giấy đề nghị trả thẻ do ATM thu (mẫu BM09-GNT) – HTML A4 dùng để in và xuất PDF.
 *  d: {kinhGui, hoTen, soThe, nganHang, atm, diaDanh, ngay:[dd,mm,yyyy], gioTra, gdv, maThe, bm:{ma,banHanh,suaDoi,quyTrinh}} */
function taPhieuHtml(d) {
  const e = taEsc, dots = n => '<span class="dt">' + '.'.repeat(n) + '</span>';
  const val = (v, n) => v ? '<b class="fill">' + e(v) + '</b>' : dots(n);
  const bm = d.bm || {}, ng = d.ngay || ['', '', ''];
  const box = '&#9744;', tick = '&#9746;';
  return '<div class="bm09">' +
  '<table class="hd"><tr><td><div class="bn">Ngân hàng TMCP Công thương Việt Nam</div><div class="bn-en">Vietnam Joint Stock Commercial Bank for Industry and Trade</div></td>' +
  '<td class="lg"><img src="' + TA_LOGO + '" alt="VietinBank"></td></tr></table>' +
  '<h1>GIẤY ĐỀ NGHỊ TRẢ THẺ DO ATM THU</h1>' +
  '<p class="kg"><i><b>Kính gửi</b>: <b>Ngân hàng TMCP Công thương Việt Nam Chi nhánh/PGD/</b>To: VietinBank- Branch:</i> ' + val(d.kinhGui, 40) + '</p>' +
  '<p class="nt"><i>(Đề nghị khách hàng/Please: (i) khi điền số thẻ chỉ điền 06 số đầu và 04 số cuối/ fill in card number fields <b>only</b> the first 6 digits and the last 4 digits of the card number); (ii) các trường thông tin có dấu (*) là bắt buộc/ Information fields marked with an asterisk (*) are required; (iii) các nội dung không liên quan đề nghị gạch chéo/the unrelated contents should be crossed out)</i></p>' +
  '<div class="sec">THÔNG TIN CHỦ THẺ <i>(Cardholder information)</i></div>' +
  '<table class="ln"><tr><td style="width:66%">Họ tên chủ thẻ *' + val(d.hoTen, 60) + '<div class="en">(Cardholder\'s name)</div></td>' +
  '<td>Giới tính * ' + box + ' Nam &nbsp; ' + box + ' Nữ<div class="en">(Sex) &nbsp; (Male) &nbsp; (Female)</div></td></tr></table>' +
  '<table class="ln"><tr><td style="width:40%">Số CCCD/CMND/HC*' + dots(34) + '<div class="en">(ID Number /PP Number)</div></td><td style="width:28%">Ngày cấp' + dots(26) + '<div class="en">(Date of issue)</div></td><td>Nơi cấp' + dots(28) + '<div class="en">(Place of issue)</div></td></tr></table>' +
  '<table class="ln"><tr><td style="width:40%">Điện thoại' + dots(40) + '<div class="en">(Mobile phone)</div></td><td>Địa chỉ' + dots(70) + '<div class="en">(Address)</div></td></tr></table>' +
  '<div class="sec">TRẢ LẠI THẺ DO ATM THU <i>(Return of card retained by ATM)</i></div>' +
  '<div class="sub"><b><i>Khách hàng đề nghị Ngân hàng trả lại thẻ do ATM thu theo thông tin dưới đây</i></b> <i>(The customer requests the Bank to return the card retained by the ATM with the following information):</i></div>' +
  '<div class="rw">Số thẻ<sup>*</sup> <i>(Card number)</i> ' + val(taMaskView(d.soThe), 80) + '</div>' +
  '<div class="rw">Loại thẻ <i>(Card type)</i> ' + dots(100) + '</div>' +
  '<div class="rw">Ngân hàng phát hành <i>(bank issue)</i>* ' + val(d.nganHang, 80) + '</div>' +
  '<div class="rw"><b>Giao dịch ở ATM đặt tại</b> <i>(Transaction at ATM located at)</i> <b>:</b> ' + val(d.atm, 60) + '</div>' +
  '<div class="rw"><b>vào lúc</b><i>(at)</i> ' + dots(8) + ' <b>giờ</b> <i>(hour)</i> ' + dots(8) + ' <b>ngày</b> <i>(date)</i> ' + dots(6) + ' <b>tháng</b> <i>(month)</i> ' + dots(6) + ' <b>năm</b> <i>(year)</i>' + dots(8) + '</div>' +
  '<div class="sec">ĐỀ NGHỊ MỞ KHÓA THẺ ( Vietinbank issue only)</div>' +
  '<div class="sub"><b><i>Khách hàng đề nghị Ngân hàng mở khóa thẻ theo thông tin dưới đây</i></b> <i>(The customer requests the Bank to open the card retained by the ATM with the following information):</i></div>' +
  '<div class="rw">Số thẻ<sup>*</sup> <i>(Card number)</i> ' + dots(110) + '</div>' +
  '<table class="sg"><tr><td style="width:50%"></td><td class="c"><i>' + e(d.diaDanh || '') + (d.diaDanh ? '' : dots(14)) + ', ngày <sub>date</sub> ' + e(ng[0]) + ' tháng <sub>month</sub> ' + e(ng[1]) + ' năm <sub>year</sub> ' + e(ng[2]) + '</i>' +
  '<div><b>KHÁCH HÀNG (</b><i>Customer)</i></div><div><i>Ký, họ tên/Sign, full name</i></div><div class="sp"></div></td></tr>' +
  '<tr><td colspan="2" class="secrow">PHẦN DÀNH CHO NGÂN HÀNG <i>(For Bank use only)</i></td></tr>' +
  '<tr><td>Xác nhận đã xác thực KH: ' + box + '</td><td></td></tr>' +
  '<tr><td>Đã trả thẻ cho KH: ' + tick + '<br>Đã mở khóa thẻ cho KH: ' + box + '<br>Ngày giờ thực hiện: <b class="fill">' + e(d.gioTra || '') + '</b></td><td class="mut">' + (d.maThe ? 'Mã thẻ trong sổ: ' + e(d.maThe) : '') + '</td></tr>' +
  '<tr><td class="c"><b>Giao dịch viên</b><div><i>Teller</i></div><div class="sp2"></div><div class="fill">' + e(d.gdv || '') + '</div></td><td class="c"><b>Kiểm soát viên</b><div><i>Supervisor</i></div><div class="sp2"></div></td></tr></table>' +
  '<table class="ft"><tr><td class="c" style="width:60%"><b>QUY TRÌNH</b></td><td class="c">Lần ban hành: ' + e(bm.banHanh || '02') + '</td><td class="c">Trang: 1/1</td></tr>' +
  '<tr><td class="c">' + e(bm.quyTrinh || 'NGHIỆP VỤ PHÁT HÀNH, SỬ DỤNG VÀ QUẢN LÝ THẺ TRONG HỆ THỐNG NGÂN HÀNG THƯƠNG MẠI CỔ PHẦN CÔNG THƯƠNG VIỆT NAM') + '</td><td class="c">Lần sửa đổi: ' + e(bm.suaDoi || '00') + '</td><td class="c">' + e(bm.ma || 'BM09-GNT') + '</td></tr></table>' +
  '</div>';
}
/** CSS của phiếu (A4) */
const TA_PHIEU_CSS = '.bm09,.bm09 *{-webkit-print-color-adjust:exact;print-color-adjust:exact}.bm09{font-family:"Times New Roman",Times,serif;color:#000;font-size:12.5px;line-height:1.35;width:100%;max-width:190mm;margin:0 auto;background:#fff}' +
  '.bm09 table{width:100%;border-collapse:collapse;font-size:inherit}.bm09 sub{font-size:8px}.bm09 td{vertical-align:top;padding:2px 4px}' +
  '.bm09 .hd td{padding:0 0 4px}.bm09 .bn{color:#1F4E79;font-weight:bold;font-size:14px}.bm09 .bn-en{color:#1F4E79;font-style:italic;font-weight:bold;font-size:11px}' +
  '.bm09 .lg{text-align:right;width:40%}.bm09 .lg img{height:34px}' +
  '.bm09 h1{text-align:center;color:#C00000;font-size:16px;margin:10px 0 14px}' +
  '.bm09 .kg{margin:0 0 2px}.bm09 .nt{margin:0 0 6px;font-size:11.5px}' +
  '.bm09 .sec,.bm09 .secrow{background:#5B9BD5;color:#fff;font-weight:bold;padding:2px 4px;border:1px solid #5B9BD5}' +
  '.bm09 .sub{background:#DEEAF6;padding:2px 4px;margin-bottom:4px}' +
  '.bm09 .ln td{padding:3px 4px 1px}.bm09 .en{font-size:10px;font-style:italic}' +
  '.bm09 .rw{padding:3px 4px}.bm09 .dt{color:#555;letter-spacing:-.5px;word-break:break-all}' +
  '.bm09 .fill{font-family:Arial,Helvetica,sans-serif;font-size:13px;color:#00007A}' +
  '.bm09 .sg{margin-top:6px;border:1px dashed #000}.bm09 .sg td{border:1px dashed #000}.bm09 .c{text-align:center}' +
  '.bm09 .sp{height:70px}.bm09 .sp2{height:60px}.bm09 .mut{color:#555;font-size:11px}' +
  '.bm09 .ft{margin-top:14px;font-size:10px}.bm09 .ft td{border:1px solid #000;padding:4px;vertical-align:middle}';
