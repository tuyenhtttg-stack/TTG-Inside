# Không Gian Số Nội Bộ · TTG Inside

Bộ tiện ích nội bộ dành cho cán bộ **VietinBank Chi nhánh Tây Tiền Giang**, thiết kế cho điện thoại.
Phiên bản 1.4.1 – 27/10/2026: thêm đăng nhập bằng ID và phân quyền theo Google Sheet.

| Nhóm | Tiện ích | Mô tả |
|---|---|---|
| Mã QR | Tạo QR kết nối WiFi | 4 bước: chọn điểm giao dịch → chọn mạng WiFi → quét thử → in thẻ niêm yết A6 |
| Chi phí | Quản lý chi phí (I003) | Hạn mức phê duyệt còn lại theo thời gian thực; trưởng/phó phòng nhập khoản đã duyệt; PGĐ, Giám đốc, kế toán xem theo quyền |
| Chi phí | Báo cáo chi phí (I004) | Tổng quan, dự báo hạn mức, theo phòng/nhóm, ma trận, xu hướng 12 tháng, kế hoạch năm, khoản lớn/lặp lại, nhận xét AI |
| Mã QR | Tạo QR theo yêu cầu | Biểu mẫu tạo mã: đường link, gọi điện, tin nhắn, email, danh thiếp, văn bản; ô nhập tự nhận dạng |

## Cấu trúc thư mục

```
index.html              Trang chính (toàn bộ giao diện và chương trình; font, logo nhúng sẵn)
manifest.webmanifest    Khai báo để cài lên màn hình chính điện thoại
sw.js                   Service worker: mở nhanh, dùng được khi mạng yếu
lib/qrcode.js           Thư viện tạo mã QR (qrcode-generator 1.4.4, MIT)
lib/jsQR.js             Thư viện đọc mã QR từ ảnh (jsQR 1.4.0, Apache-2.0)
icons/                  Biểu tượng ứng dụng
apps-script/Code.gs     Máy chủ phân quyền (gắn với Google Sheet) + trợ lý AI giai đoạn 2
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
| `TOOLKIT` | Danh mục nhóm và tiện ích; mỗi tiện ích có `app` = mã ứng dụng trong Sheet (I001, I002…) |
| `WIFI_LIST` | Mạng WiFi khách của từng điểm giao dịch: `ssid`, `sec`, `hidden`. **Không ghi mật khẩu.** |
| `WIFI_LIST_IS_SAMPLE` | `true` khi còn dữ liệu mẫu (trang hiện cảnh báo vàng). Đặt `false` sau khi nhập tên mạng thật. |
| `API_URL` | **Bắt buộc.** URL Web App Apps Script của file Sheet phân quyền. Để trống thì không ai đăng nhập được. |
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

## Phân quyền (Google Sheet + Apps Script)

1. Tải file `KhongGianSoNoiBo_PhanQuyen.xlsx` lên thư mục `KhongGianSoNoiBo` trên Drive, mở bằng Google Trang tính
   (Tệp → Lưu dưới dạng Google Trang tính nếu Drive chưa tự chuyển).
2. Trong file Sheet: Tiện ích mở rộng → Apps Script → dán `apps-script/Code.gs` → Lưu → chạy hàm `caiDatBanDau`, cấp quyền.
3. Triển khai → Triển khai mới → Ứng dụng web; Thực thi với tư cách: **Tôi**; Người có quyền truy cập: **Bất kỳ ai**.
4. Copy URL `.../exec`, dán vào `API_URL` trong `index.html`, tăng `VERSION` trong `sw.js`, commit lên GitHub.
5. Cấp ID cho cán bộ theo sheet `HuongDan` trong file.
6. (v1.2.0) Menu **Không Gian Số → Cài đặt Phòng/Ban**: tạo sheet `PhongBan` (Mã phòng/ban, Tên, Loại, Đang dùng) và danh sách chọn cho cột Phòng/Ban ở `CanBo`. Máy chủ trả thêm `maPhongBan`, `loaiPhongBan` cho ứng dụng phân quyền theo phòng.

Cách hoạt động:
- Cán bộ nhập ID 5 số (che như mật khẩu, bấm biểu tượng con mắt để hiện). Đăng nhập có hiệu lực đến hết ngày.
- Trang chủ chỉ hiện tiện ích được cấp quyền; mở thẳng đường dẫn tiện ích chưa được cấp sẽ bị chặn.
- Khóa cán bộ hoặc bỏ quyền trong Sheet có hiệu lực ở lần mở kế tiếp (tối đa khoảng 2 phút).
- Sai 5 lần trên một máy → khóa máy 15 phút; toàn hệ thống sai quá 30 lần/10 phút → tạm ngưng đăng nhập 15 phút và email báo quản trị.
- Lượt truy cập (hôm nay / tổng) hiện nhỏ ở cuối trang; quản trị xem thêm qua menu **Không Gian Số → Thống kê truy cập**.
- Lưu ý: tiện ích chạy hoàn toàn trên máy (như tạo QR) chỉ được ẩn với người không có quyền. Tiện ích có dữ liệu nội bộ phải lấy dữ liệu qua Apps Script, kiểm tra phiên đăng nhập ở mỗi lần gọi.

## Quản lý chi phí (I003) và Báo cáo chi phí (I004) – từ v1.3.0 / v1.4.0

1. Dán `apps-script/Code.gs` v1.3.0 vào Apps Script của file phân quyền → Lưu.
2. Trong file phân quyền: menu **Không Gian Số → Cài đặt / cập nhật Quản lý chi phí (I003, I004)** (chạy lại an toàn, tự nâng cấp file cũ). Script tạo file `KhongGianSoNoiBo_ChiPhi`
   (ở thư mục gốc Drive – kéo vào thư mục KhongGianSoNoiBo), đăng ký I003 (tick: Giám đốc, Phó Giám đốc, Quản trị hệ thống, Trưởng phòng, Phó phòng, Kế toán chi tiêu) và I004 (tick: Giám đốc, Phó Giám đốc, Quản trị hệ thống, Trưởng phòng, Kế toán chi tiêu).
3. Triển khai → Quản lý triển khai → Phiên bản mới (URL giữ nguyên). Tải `index.html` + `sw.js` mới lên GitHub.
4. Cấu hình trong file chi phí (không cần sửa code): `CP_NguoiDuyet` (người duyệt, hạn mức; để trống = không giới hạn), `CP_PhuTrach` (phòng → người duyệt),
   `CP_NhomChiPhi`, `CP_KeHoach` (kế hoạch năm theo phòng/nhóm, để trống nếu chưa có), `CP_CauHinh` (vai trò nhập / xem toàn bộ / được hủy / xem báo cáo phòng mình, phòng kế toán = Phòng DVKH, số ngày đầu tháng nhập cho tháng trước, ngưỡng cảnh báo).
6. Nhận xét AI (I004): thêm thuộc tính tập lệnh `ANTHROPIC_API_KEY`; tùy chọn `MODEL_BAO_CAO` (mặc định claude-sonnet-5-5). Chỉ gửi số liệu tổng hợp, không gửi ID/tên cán bộ; kết quả cùng số liệu được dùng lại trong 30 phút.
5. Đổi tên cán bộ hoặc phòng ở file phân quyền xong: menu **Đồng bộ danh mục sang file chi phí**.

Quy tắc: vượt hạn mức không cho lưu; người nhập tự hủy trong ngày (ghi lý do), sau đó Giám đốc/Quản trị hủy; không sửa số tiền (hủy và nhập lại);
khoản đã nhập luôn tính cho người duyệt tại thời điểm nhập (phân công lại không làm lệch hạn mức).

## Cập nhật phiên bản

Mỗi lần sửa `index.html` hoặc thư viện:
1. Mở `sw.js`, tăng `VERSION` (ví dụ `1.0.0` → `1.0.1`).
2. Cập nhật `APP_VERSION` trong `index.html`.
3. Commit lên GitHub. Máy cán bộ nhận bản mới ở lần mở kế tiếp (có mạng).

## Thêm tiện ích mới

1. Trong Sheet: menu **Không Gian Số → Thêm ứng dụng mới** (cấp mã I003, I004…).
1. Thêm một mục vào `TOOLKIT` (id, app = mã vừa cấp, tên, mô tả, biểu tượng).
2. Thêm khối `<section class="view" id="v-<id>" hidden>…</section>` trong `<main>`.
3. Trang chủ, thanh điều hướng và tab cùng nhóm tự cập nhật.

## Giai đoạn 2 – trợ lý AI (tùy chọn)

Thêm thuộc tính tập lệnh `ANTHROPIC_API_KEY` trong Apps Script (cùng script phân quyền). Cán bộ có quyền I002 sẽ dùng được trợ lý AI, không cần sửa `index.html`. Nội dung hướng dẫn trợ lý có ở 2 nơi phải giữ đồng bộ:
`RULES` (index.html) và `SYSTEM_PROMPT` (Code.gs).

## Bảo mật dữ liệu

- Không lưu mật khẩu WiFi trong mã nguồn. Mật khẩu do cán bộ nhập, chỉ dùng trên máy để vẽ mã.
- Ảnh chụp mã chia sẻ WiFi được đọc ngay trên điện thoại, không gửi đi đâu.
- Trang chỉ nhớ phiên đăng nhập trong ngày và điểm giao dịch chọn gần nhất trên từng máy (bộ nhớ trình duyệt).
- File Sheet phân quyền không chia sẻ cho ai; trang chỉ nhận về tên gọi, phòng/ban, vai trò và danh sách ứng dụng của người đang đăng nhập.
- Không nhập thông tin khách hàng, số tài khoản, CCCD vào các mã QR.
