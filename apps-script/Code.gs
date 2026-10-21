/**
 * Trung gian gọi Claude API cho "Không Gian Số Nội Bộ – TTG Inside" (tiện ích Tạo QR theo yêu cầu).
 * Trang HTML (GitHub Pages) gửi hội thoại tới Web App này; API key nằm trong Script Properties, không lộ ra trình duyệt.
 *
 * Script Properties cần đặt (Project Settings → Script Properties):
 *   ANTHROPIC_API_KEY  : khóa API Claude (bắt buộc)
 *   MODEL              : tùy chọn, mặc định claude-haiku-5-5
 *   DAILY_LIMIT        : tùy chọn, tổng lượt gọi/ngày toàn hệ thống, mặc định 300
 *   PER_MINUTE_LIMIT   : tùy chọn, lượt gọi/phút toàn hệ thống, mặc định 20
 *
 * Triển khai:
 *   1. script.google.com → New project → dán file này → Lưu.
 *   2. Project Settings → Script Properties → thêm các khóa ở trên.
 *   3. Chạy hàm testCallClaude() một lần để cấp quyền, xem Execution log.
 *   4. Deploy → New deployment → Web app; Execute as: Me; Who has access: Anyone → copy URL .../exec.
 *   5. Dán URL vào APPS_SCRIPT_URL trong index.html. Sửa code sau này: Manage deployments → Edit → New version (giữ nguyên URL).
 */

const DEFAULT_MODEL = 'claude-haiku-5-5';
const MAX_TURNS = 12;          // chỉ giữ 12 lượt gần nhất
const MAX_CHARS = 2000;        // giới hạn độ dài mỗi lượt
const TZ = 'Asia/Ho_Chi_Minh';

// Giữ đồng bộ với hằng RULES trong file HTML
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

/** Kiểm tra nhanh trên trình duyệt: mở URL Web App sẽ thấy {"ok":true} */
function doGet() {
  return json_({ ok: true, service: 'qr-assistant' });
}

function doPost(e) {
  try {
    const props = PropertiesService.getScriptProperties();
    const apiKey = props.getProperty('ANTHROPIC_API_KEY');
    if (!apiKey) return json_({ error: 'not_configured', message: 'Chưa cấu hình ANTHROPIC_API_KEY.' });

    // 1) Đọc & làm sạch hội thoại
    const body = JSON.parse((e && e.postData && e.postData.contents) || '{}');
    const messages = sanitize_(body.messages);
    if (!messages.length) return json_({ error: 'bad_request', message: 'Thiếu nội dung yêu cầu.' });

    // 2) Giới hạn lượt dùng
    const limitErr = checkLimits_(props);
    if (limitErr) return json_(limitErr);

    // 3) Gọi Claude API
    const res = UrlFetchApp.fetch('https://api.anthropic.com/v1/messages', {
      method: 'post',
      contentType: 'application/json',
      headers: { 'x-api-key': apiKey, 'anthropic-version': '2023-06-01' },
      payload: JSON.stringify({
        model: props.getProperty('MODEL') || DEFAULT_MODEL,
        max_tokens: 1200,
        system: SYSTEM_PROMPT,
        messages: messages
      }),
      muteHttpExceptions: true
    });

    const code = res.getResponseCode();
    if (code === 429 || code === 529) return json_({ error: 'rate_limited', message: 'Claude API đang quá tải.' });
    if (code !== 200) {
      console.error('Claude API ' + code + ': ' + res.getContentText().slice(0, 500));
      return json_({ error: 'upstream_error', message: 'Claude API lỗi ' + code });
    }

    const data = JSON.parse(res.getContentText());
    const text = (data.content || []).filter(c => c.type === 'text').map(c => c.text).join('');
    const out = parseJson_(text);
    if (!out) return json_({ error: 'invalid_json', message: 'Trợ lý trả lời sai định dạng.' });
    return json_(out);
  } catch (err) {
    console.error(err);
    return json_({ error: 'upstream_error', message: String(err) });
  }
}

/** Chỉ nhận role user/assistant, cắt độ dài, bắt đầu & kết thúc bằng user */
function sanitize_(arr) {
  if (!Array.isArray(arr)) return [];
  let msgs = arr
    .filter(m => m && (m.role === 'user' || m.role === 'assistant') && typeof m.content === 'string' && m.content.trim())
    .map(m => ({ role: m.role, content: m.content.slice(0, MAX_CHARS) }))
    .slice(-MAX_TURNS);
  // gộp các lượt liền nhau cùng role
  const merged = [];
  msgs.forEach(m => {
    const last = merged[merged.length - 1];
    if (last && last.role === m.role) last.content += '\n' + m.content; else merged.push(m);
  });
  while (merged.length && merged[0].role !== 'user') merged.shift();
  if (!merged.length || merged[merged.length - 1].role !== 'user') return [];
  return merged;
}

/** Giới hạn theo phút (CacheService) và theo ngày (Script Properties) */
function checkLimits_(props) {
  const lock = LockService.getScriptLock();
  lock.waitLock(5000);
  try {
    const perMin = Number(props.getProperty('PER_MINUTE_LIMIT') || 20);
    const daily = Number(props.getProperty('DAILY_LIMIT') || 300);
    const cache = CacheService.getScriptCache();
    const minKey = 'm_' + Utilities.formatDate(new Date(), TZ, 'yyyyMMddHHmm');
    const m = Number(cache.get(minKey) || 0);
    if (m >= perMin) return { error: 'rate_limited', message: 'Quá nhiều yêu cầu, thử lại sau 1 phút.' };

    const dayKey = 'd_' + Utilities.formatDate(new Date(), TZ, 'yyyyMMdd');
    const d = Number(props.getProperty(dayKey) || 0);
    if (d >= daily) return { error: 'daily_limit', message: 'Đã hết lượt dùng trợ lý trong ngày.' };

    cache.put(minKey, String(m + 1), 120);
    props.setProperty(dayKey, String(d + 1));
    return null;
  } finally {
    lock.releaseLock();
  }
}

/** Đọc JSON từ câu trả lời (chấp nhận khối ```json hoặc văn bản thừa đầu/cuối) */
function parseJson_(text) {
  const tryParse = s => { try { return JSON.parse(s); } catch (e) { return null; } };
  let v = tryParse(text.trim());
  if (!v) { const f = text.match(/```(?:json)?\s*([\s\S]*?)```/); if (f) v = tryParse(f[1]); }
  if (!v) { const a = text.indexOf('{'), b = text.lastIndexOf('}'); if (a >= 0 && b > a) v = tryParse(text.slice(a, b + 1)); }
  if (!v || typeof v !== 'object') return null;
  return { reply: String(v.reply || 'Đã xử lý.'), qr: v.qr && v.qr.payload ? v.qr : null };
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

/** Chạy tay 1 lần trong trình soạn thảo để cấp quyền & thử kết nối */
function testCallClaude() {
  const fake = { postData: { contents: JSON.stringify({ messages: [{ role: 'user', content: 'Tạo mã gọi tổng đài 1900 558 868' }] }) } };
  console.log(doPost(fake).getContent());
}

/** Xóa bộ đếm lượt dùng của các ngày cũ (có thể đặt trigger chạy hằng tuần) */
function cleanupCounters() {
  const props = PropertiesService.getScriptProperties();
  const today = 'd_' + Utilities.formatDate(new Date(), TZ, 'yyyyMMdd');
  Object.keys(props.getProperties()).forEach(k => { if (/^d_\d{8}$/.test(k) && k !== today) props.deleteProperty(k); });
}
