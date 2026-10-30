/**
 * KHÔNG GIAN SỐ NỘI BỘ (TTG Inside) – Máy chủ phân quyền
 * VietinBank Chi nhánh Tây Tiền Giang · phiên bản 1.6.1 · 30/10/2026
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
 *   ANTHROPIC_API_KEY  : khóa Claude API (sk-ant-…) cho trợ lý AI, nhận xét báo cáo, đọc ảnh thẻ; để trống = tắt AI
 *   ANTHROPIC_WORKSPACE_ID : chỉ cần khi khóa API không gắn với workspace nào (wrkspc_…)
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
      case 'ta_ocr':    return json_(apiTaOcr_(body));
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
    headers: aiHeaders_(key),
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
/** Header gọi Claude API; thêm anthropic-workspace-id nếu khóa API không gắn workspace (thuộc tính ANTHROPIC_WORKSPACE_ID) */
function aiHeaders_(key) {
  const h = { 'x-api-key': key, 'anthropic-version': '2023-06-01' }, ws = PropertiesService.getScriptProperties().getProperty('ANTHROPIC_WORKSPACE_ID');
  if (ws) h['anthropic-workspace-id'] = ws;
  return h;
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
    headers: aiHeaders_(key),
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
  ['MAU_PHIEU_SUA_DOI', '00', 'Lần sửa đổi biểu mẫu'],
  ['DOC_THE_BANG', 'AI', 'Cách đọc ảnh thẻ khi chụp: AI (gửi ảnh cho Claude đọc, chính xác với số in nổi) hoặc Máy (đọc trên điện thoại, không gửi ảnh đi)'],
  ['MODEL_DOC_THE', 'claude-sonnet-5-5', 'Mô hình Claude dùng để đọc ảnh thẻ (cần ANTHROPIC_API_KEY trong thuộc tính tập lệnh)']
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
    docThe: /^m[aá]y$/i.test(kv.DOC_THE_BANG || '') ? 'may' : 'ai', modelDocThe: kv.MODEL_DOC_THE || 'claude-sonnet-5-5',
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
  const aiOn = !!PropertiesService.getScriptProperties().getProperty('ANTHROPIC_API_KEY');
  return { soNgayHuy: cfg.soNgayHuy, soDau: cfg.soDau, nhomTon: cfg.nhomTon, mien: cfg.mien, hoanTac: cfg.hoanTac, diaDanh: cfg.diaDanh, tenCN: cfg.tenCN, bm: cfg.bm,
    docThe: cfg.docThe === 'ai' && aiOn ? 'ai' : 'may', aiOn,
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
  const { user, cfg } = g, email = String(b.email || '').trim().toLowerCase(), loai = b.loai === 'docx' ? 'docx' : 'pdf';
  if (!g.role.canReturn) return { error: 'forbidden', message: 'Vai trò của tài khoản không được trả thẻ.' };
  if (!new RegExp('^[^@\\s]+@' + cfg.mien.replace(/\./g, '\\.') + '$').test(email)) return { error: 'invalid', message: 'Chỉ gửi được tới email @' + cfg.mien + '.' };
  const f = taFind_(b, g); if (f.error) return f;
  if (f.x.tt !== TA_ST.TRA) return { error: 'invalid', message: 'Thẻ chưa ở trạng thái đã trả.' };
  const d = taPhieuData_(f.x, f.n, cfg);
  let file;
  if (loai === 'docx') file = taPhieuDocx_(d).setName('Phieu_tra_the_' + f.x.ma + '.docx');
  else {
    const html = '<html><head><meta charset="utf-8"><style>@page{size:A4;margin:12mm}body{margin:0}' + TA_PHIEU_CSS + '</style></head><body>' + taPhieuHtml(d) + '</body></html>';
    file = Utilities.newBlob(html, 'text/html', 'phieu.html').getAs('application/pdf').setName('Phieu_tra_the_' + f.x.ma + '.pdf');
  }
  const ten = loai === 'docx' ? 'Word' : 'PDF';
  MailApp.sendEmail({ to: email, name: 'Không Gian Số Nội Bộ', subject: '[Thẻ ATM máy thu] Phiếu trả thẻ ' + f.x.ma + ' (' + ten + ')',
    htmlBody: '<p>Đính kèm Giấy đề nghị trả thẻ do ATM thu (' + taEsc(cfg.bm.ma) + ') của thẻ ' + taEsc(f.x.ma) + ', file ' + ten + '. ' +
      (loai === 'docx' ? 'Có thể mở bằng Word để bổ sung nội dung trước khi in cho khách hàng ký.' : 'In để khách hàng ký.') +
      '</p><p style="color:#888;font-size:12px">Email gửi tự động – ' + taEsc(cfg.tenCN) + '.</p>',
    attachments: [file] });
  log_(user.id, user.hoTen, 'Gửi phiếu trả thẻ', TA_APP, f.x.ma + ' (' + ten + ') → ' + email, '');
  return { ok: true, message: 'Đã gửi phiếu ' + ten + ' tới ' + email + '.' };
}
/** Phiếu BM09-GNT dạng Word (.docx): điền vào mẫu gốc của anh (logo đổi sang PNG cho nhẹ). Ô chưa có dữ liệu giữ dòng chấm để ghi tay. */
function taPhieuDocx_(d) {
  const x = s => String(s || '').replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
  const ng = d.ngay || [];
  const v = { KINH_GUI: d.kinhGui || '................................', HO_TEN: d.hoTen ? ' ' + d.hoTen : '', SO_THE: taMaskView(d.soThe), NGAN_HANG: d.nganHang || '',
    ATM: d.atm || '...............................……………………………', DIA_DANH: d.diaDanh || '...........', DD: ng[0] || '....', MM: ng[1] || '........', YYYY: ng[2] || '.......',
    GIO_TRA: d.gioTra || '', GDV: d.gdv || '' };
  const files = Utilities.unzip(Utilities.newBlob(Utilities.base64Decode(TA_BM09_DOCX), 'application/zip', 'bm09.zip'));
  const out = files.map(f => {
    if (f.getName() !== 'word/document.xml') return f;
    const xml = f.getDataAsString('UTF-8').replace(/\{\{(\w+)\}\}/g, (m, k) => x(v[k]));
    return Utilities.newBlob(xml, 'application/xml', 'word/document.xml');
  });
  return Utilities.zip(out, 'phieu.docx').setContentType('application/vnd.openxmlformats-officedocument.wordprocessingml.document');
}
// Mẫu BM09-GNT (Word) đã gắn ô điền {{...}} – tạo từ file anh gửi ngày 29/10/2026
const TA_BM09_DOCX = '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';
/** Đọc ảnh thẻ bằng Claude (phương án B). Ảnh không được lưu; chỉ trả về số đầu, 4 số cuối, họ tên, ngân hàng. */
const TA_OCR_PROMPT = 'Bạn đọc ảnh chụp thẻ ngân hàng (thẻ ATM/ghi nợ/tín dụng) bị máy ATM thu giữ, để cán bộ ngân hàng Việt Nam nhập sổ bàn giao thẻ.\n' +
  'Số thẻ và họ tên thường in nổi (dập nổi), có thể bị lóa sáng, nghiêng, mòn; hãy đọc thật cẩn thận từng chữ số. Ảnh có thể là mặt trước hoặc mặt sau.\n' +
  'TUYỆT ĐỐI không trả về số thẻ đầy đủ. Chỉ trả về một đối tượng JSON, không kèm giải thích, theo mẫu:\n' +
  '{"dau":"<SO_DAU chữ số đầu tiên của số thẻ>","cuoi":"<4 chữ số cuối>","hoTen":"<họ tên chủ thẻ in trên thẻ, chữ in hoa không dấu; rỗng nếu thẻ không in tên>",' +
  '"nganHang":"<tên ngân hàng phát hành, chọn đúng một tên trong danh sách; rỗng nếu không xác định>","doTinCay":"cao|trung bình|thấp","ghiChu":"<ngắn gọn: chữ số nào mờ/không chắc; rỗng nếu rõ>"}\n' +
  'Nếu không thấy số thẻ thì để dau và cuoi rỗng. Không đoán bừa: chữ số không đọc được thì để doTinCay là "thấp" và nêu trong ghiChu.';
function apiTaOcr_(b) {
  const g = taGuard_(b, TA_APP); if (g.error) return g;
  const { user, cfg } = g;
  if (!g.role.canEnter) return { error: 'forbidden', message: 'Vai trò của tài khoản không được nhập thẻ.' };
  const props = PropertiesService.getScriptProperties(), key = props.getProperty('ANTHROPIC_API_KEY');
  if (!key) return { error: 'not_configured', message: 'Chưa bật AI đọc thẻ (thiếu ANTHROPIC_API_KEY trong thuộc tính tập lệnh).' };
  const img = String(b.image || '').replace(/^data:image\/\w+;base64,/, ''), mt = /^image\/(jpeg|png|webp)$/.test(String(b.mediaType)) ? String(b.mediaType) : 'image/jpeg';
  if (img.length < 1000 || img.length > 4.5e6 || !/^[A-Za-z0-9+/=]+$/.test(img.slice(0, 200))) return { error: 'invalid', message: 'Ảnh không hợp lệ hoặc quá lớn. Chụp lại.' };
  const lim = aiLimit_(props); if (lim) return Object.assign(lim, { message: lim.error === 'daily_limit' ? 'Đã hết lượt AI trong ngày. Nhập tay.' : 'AI đang bận, thử lại sau vài giây.' });
  const banks = cfg.banks.map(x => x.ten).join(', ');
  const res = UrlFetchApp.fetch('https://api.anthropic.com/v1/messages', {
    method: 'post', contentType: 'application/json', muteHttpExceptions: true,
    headers: aiHeaders_(key),
    payload: JSON.stringify({ model: cfg.modelDocThe, max_tokens: 300, temperature: 0, system: TA_OCR_PROMPT.replace('SO_DAU', String(cfg.soDau)),
      messages: [{ role: 'user', content: [{ type: 'image', source: { type: 'base64', media_type: mt, data: img } },
        { type: 'text', text: 'Danh sách ngân hàng: ' + banks + '. Số chữ số đầu cần lấy: ' + cfg.soDau + '.' }] }] })
  });
  const code = res.getResponseCode();
  if (code === 429 || code === 529) return { error: 'rate_limited', message: 'Dịch vụ AI đang quá tải, thử lại sau ít phút.' };
  if (code !== 200) { console.error('Claude API ' + code + ': ' + res.getContentText().slice(0, 300)); return { error: 'upstream_error', message: 'Dịch vụ AI tạm thời lỗi. Nhập tay hoặc thử lại.' }; }
  const text = (JSON.parse(res.getContentText()).content || []).filter(c => c.type === 'text').map(c => c.text).join('');
  let v = null; try { const a = text.indexOf('{'), z = text.lastIndexOf('}'); v = JSON.parse(text.slice(a, z + 1)); } catch (e) { v = null; }
  if (!v) return { error: 'upstream_error', message: 'AI không đọc được ảnh. Chụp lại hoặc nhập tay.' };
  let dau = String(v.dau || '').replace(/\D/g, ''), cuoi = String(v.cuoi || '').replace(/\D/g, '');
  dau = dau.length >= cfg.soDau ? dau.slice(0, cfg.soDau) : ''; cuoi = cuoi.length >= 4 ? cuoi.slice(-4) : '';
  const binBank = dau.length >= 6 ? (cfg.banks.find(k => k.bins.some(x => x.length >= 6 && dau.startsWith(x))) || {}).ten : '';
  const nganHang = binBank || (v.nganHang ? taBankOf_(v.nganHang, cfg) : '');
  const hoTen = taKhongDau(v.hoTen || '').toUpperCase().replace(/[^A-Z .'-]/g, ' ').replace(/\s+/g, ' ').trim().slice(0, 60);
  log_(user.id, user.hoTen, 'Đọc ảnh thẻ bằng AI', TA_APP, (dau && cuoi ? 'đọc được số thẻ' : 'chưa đọc được số thẻ') + ' · ' + String(v.doTinCay || ''), '');
  return { ok: true, dau, cuoi, hoTen, nganHang: cfg.banks.some(k => k.ten === nganHang) ? nganHang : '', nganHangDoc: nganHang,
    doTinCay: String(v.doTinCay || '').slice(0, 20), ghiChu: String(v.ghiChu || '').slice(0, 200) };
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
const TA_LOGO = 'data:image/png;base64,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';
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
