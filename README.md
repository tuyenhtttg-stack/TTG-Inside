# Không Gian Số Nội Bộ · TTG Inside

Bộ tiện ích nội bộ dành cho cán bộ **VietinBank Chi nhánh Tây Tiền Giang**, thiết kế cho điện thoại.
Phiên bản 1.0.0 – 21/10/2026.

| Nhóm | Tiện ích | Mô tả |
|---|---|---|
| Mã QR | Tạo QR kết nối WiFi | 4 bước: chọn điểm giao dịch → chọn mạng WiFi → quét thử → in thẻ niêm yết A6 |
| Mã QR | Tạo QR theo yêu cầu | Biểu mẫu tạo mã: đường link, gọi điện, tin nhắn, email, danh thiếp, văn bản; ô nhập tự nhận dạng |

## Cấu trúc thư mục

```
index.html              Trang chính (toàn bộ giao diện và chương trình; font, logo nhúng sẵn)
manifest.webmanifest    Khai báo để cài lên màn hình chính điện thoại
sw.js                   Service worker: mở nhanh, dùng được khi mạng yếu
lib/qrcode.js           Thư viện tạo mã QR (qrcode-generator 1.4.4, MIT)
lib/jsQR.js             Thư viện đọc mã QR từ ảnh (jsQR 1.4.0, Apache-2.0)
icons/                  Biểu tượng ứng dụng
apps-script/Code.gs     Backend trợ lý AI cho giai đoạn 2 (chưa dùng)
```

Không cần cài đặt hay build. Chỉ cần đưa nguyên thư mục lên GitHub.

## Đưa lên GitHub Pages

1. Tạo repository mới (ví dụ `ttg-inside`), tải toàn bộ nội dung thư mục này lên nhánh `main`.
2. Vào **Settings → Pages** → Source: *Deploy from a branch* → Branch: `main` / `(root)` → Save.
3. Sau 1–2 phút, trang chạy tại `https://<tài-khoản>.github.io/ttg-inside/`.
4. Gửi link cho cán bộ; hướng dẫn **Thêm vào màn hình chính** để mở như ứng dụng.

Trang đặt `noindex` để Google không đưa vào kết quả tìm kiếm. Link vẫn công khai với ai có link.

## Cấu hình (sửa trong `index.html`, phần đầu thẻ `<script>` cuối file)

| Hằng số | Ý nghĩa |
|---|---|
| `SUITE` | Tên bộ công cụ (tiếng Việt / tiếng Anh) |
| `TOOLKIT` | Danh mục nhóm và tiện ích hiển thị ở trang chủ |
| `WIFI_LIST` | Mạng WiFi khách của từng điểm giao dịch: `ssid`, `sec`, `hidden`. **Không ghi mật khẩu.** |
| `WIFI_LIST_IS_SAMPLE` | `true` khi còn dữ liệu mẫu (trang hiện cảnh báo vàng). Đặt `false` sau khi nhập tên mạng thật. |
| `APPS_SCRIPT_URL` | Để trống ở giai đoạn 1. Giai đoạn 2 dán URL Web App Apps Script để bật trợ lý AI. |
| `APP_VERSION` | Số phiên bản hiển thị ở tiện ích Tạo QR theo yêu cầu |

Giá trị `sec`: `WPA2`, `WPA2/WPA3`, `WPA3`, `WPA` (không rõ phiên bản), `WEP`, `nopass` (mạng mở).

Ví dụ:

```js
const WIFI_LIST_IS_SAMPLE=false;
const WIFI_LIST={
  'Trụ sở Chi nhánh':[{ssid:'Ten_Mang_That',sec:'WPA2'}],
  'PGD Cái Bè':[{ssid:'Ten_Mang_That',sec:'WPA2/WPA3'}],
  ...
};
```

## Cập nhật phiên bản

Mỗi lần sửa `index.html` hoặc thư viện:
1. Mở `sw.js`, tăng `VERSION` (ví dụ `1.0.0` → `1.0.1`).
2. Cập nhật `APP_VERSION` trong `index.html`.
3. Commit lên GitHub. Máy cán bộ nhận bản mới ở lần mở kế tiếp (có mạng).

## Thêm tiện ích mới

1. Thêm một mục vào `TOOLKIT` (id, tên, mô tả, biểu tượng).
2. Thêm khối `<section class="view" id="v-<id>" hidden>…</section>` trong `<main>`.
3. Trang chủ, thanh điều hướng và tab cùng nhóm tự cập nhật.

## Giai đoạn 2 – trợ lý AI (tùy chọn)

`apps-script/Code.gs` là backend gọi Claude API. Hướng dẫn triển khai nằm ở đầu file.
Khi có URL Web App, dán vào `APPS_SCRIPT_URL`. Nội dung hướng dẫn trợ lý có ở 2 nơi phải giữ đồng bộ:
`RULES` (index.html) và `SYSTEM_PROMPT` (Code.gs).

## Bảo mật dữ liệu

- Không lưu mật khẩu WiFi trong mã nguồn. Mật khẩu do cán bộ nhập, chỉ dùng trên máy để vẽ mã.
- Ảnh chụp mã chia sẻ WiFi được đọc ngay trên điện thoại, không gửi đi đâu.
- Trang chỉ nhớ điểm giao dịch chọn gần nhất trên từng máy (bộ nhớ trình duyệt).
- Không nhập thông tin khách hàng, số tài khoản, CCCD vào các mã QR.
