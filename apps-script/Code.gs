/**
 * KHÔNG GIAN SỐ NỘI BỘ (TTG Inside) – Máy chủ phân quyền
 * VietinBank Chi nhánh Tây Tiền Giang · phiên bản 1.1.1 · 26/10/2026
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
 */

const TZ = 'Asia/Ho_Chi_Minh';
const SH = { CANBO: 'CanBo', VAITRO: 'VaiTro', UNGDUNG: 'UngDung', NHATKY: 'NhatKy' };
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
  lamMoiBoNhoDem(true);
  const warn = kiemTraVaiTro_();
  SpreadsheetApp.getUi().alert((warn ? warn + '\n\n' : '') + 'Đã cài đặt xong.\n\nBước tiếp theo: Triển khai → Triển khai mới → Ứng dụng web (Thực thi: Tôi; Truy cập: Bất kỳ ai), rồi dán URL vào API_URL trong index.html.');
}

/** Tự tick quyền khi chọn vai trò; lan quyền khi sửa sheet VaiTro; xóa bộ nhớ đệm */
function onEdit(e) {
  try {
    const s = e.range.getSheet(), name = s.getName();
    if (![SH.CANBO, SH.VAITRO, SH.UNGDUNG].includes(name)) return;
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

/** Thêm ứng dụng mới: mã tự tăng (I003, I004…), thêm cột tick ở CanBo, VaiTro và 1 dòng ở UngDung */
function themUngDung() {
  const ui = SpreadsheetApp.getUi(), ss = SpreadsheetApp.getActive();
  const ud = ss.getSheetByName(SH.UNGDUNG);
  const codes = ud.getRange(2, 1, Math.max(ud.getLastRow() - 1, 1), 1).getValues().map(r => String(r[0])).filter(x => APP_RE.test(x));
  const next = 'I' + String(Math.max(0, ...codes.map(c => +c.slice(1))) + 1).padStart(3, '0');
  const r = ui.prompt('Thêm ứng dụng mới', 'Mã sẽ cấp: ' + next + '\nNhập tên ứng dụng:', ui.ButtonSet.OK_CANCEL);
  if (r.getSelectedButton() !== ui.Button.OK || !r.getResponseText().trim()) return;
  ud.appendRow([next, r.getResponseText().trim(), '', true, false, Utilities.formatDate(new Date(), TZ, 'dd/MM/yyyy')]);
  checkboxCol_(ud, 4, 100); checkboxCol_(ud, 5, 100);
  [SH.CANBO, SH.VAITRO].forEach(n => {
    const s = ss.getSheetByName(n), c = s.getLastColumn() + 1;
    if (c > s.getMaxColumns()) s.insertColumnAfter(s.getMaxColumns());
    s.getRange(1, c).setValue(next).setFontWeight('bold').setFontColor('#FFFFFF').setBackground('#005993');
    checkboxCol_(s, c, s.getMaxRows());
    s.setColumnWidth(c, 70);
  });
  lamMoiBoNhoDem(true);
  ui.alert('Đã thêm ' + next + '. Tick quyền ở sheet VaiTro hoặc từng cán bộ. Nhớ khai báo cùng mã ' + next + ' trong TOOLKIT của index.html.');
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
  cb.getRange(cb.getLastRow() + 1, 1, rows.length, width).setValues(rows);
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
  return { user: { ten: firstName_(u.hoTen), xungHo: u.gt === 'Nữ' ? 'chị' : u.gt === 'Nam' ? 'anh' : 'anh/chị', phongBan: u.pb, vaiTro: u.vt },
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
