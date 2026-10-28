/**
 * KHÔNG GIAN SỐ NỘI BỘ (TTG Inside) – Máy chủ phân quyền
 * VietinBank Chi nhánh Tây Tiền Giang · phiên bản 1.5.0 · 28/10/2026
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
 * Manifest appsscript.json (v1.5.0) khai báo quyền tối thiểu, Drive chỉ ĐỌC (drive.readonly).
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
           apps: userApps_(u), stats: countVisit_(), ai: !!PropertiesService.getScriptProperties().getProperty('ANTHROPIC_API_KEY') && userApps_(u).includes('I002'),
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
