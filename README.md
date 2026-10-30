# Không Gian Số Nội Bộ · TTG Inside

Bộ tiện ích nội bộ dành cho cán bộ **VietinBank Chi nhánh Tây Tiền Giang**, thiết kế cho điện thoại.
Phiên bản 1.7.0 – 30/10/2026: Thẻ ATM máy thu mở rộng 6 sổ (thêm PGD Cái Bè, Long Định, Hưng Long), phân quyền nhập/trả/hủy/báo cáo theo từng sổ.

| Nhóm | Tiện ích | Mô tả |
|---|---|---|
| Thông tin | Bản tin sáng (I005) | Bản tin 06:30 mỗi sáng: điểm nhấn, lưu ý và gợi ý điều hành, chỉ số thị trường, văn bản sắp hiệu lực, tin địa bàn, chính sách, cạnh tranh; xem lại 30 ngày |
| Thẻ ATM | Quản lý thẻ ATM bị máy thu (I006) | Chụp thẻ để đọc tự động số thẻ, họ tên, ngân hàng (đọc ngay trên điện thoại); nhập theo lô và email phòng nhận; sổ thẻ nhóm theo ngân hàng; trả thẻ, in phiếu BM09-GNT hoặc gửi qua email file PDF / Word (Word dùng đúng mẫu BM09 gốc, sửa được) |
| Thẻ ATM | Báo cáo thẻ ATM máy thu (I007) | Thẻ tồn theo thời gian giữ, theo ngân hàng, theo phòng; thẻ đã trả trong kỳ; thẻ quá hạn cần hủy: xuất danh sách rà soát, tải CSV, thực hiện hủy; tra cứu thẻ đã hủy |
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
lib/tesseract/          Thư viện đọc chữ trên ảnh thẻ (Tesseract.js 5.1.1, Apache-2.0) – chỉ tải khi bấm chụp thẻ
icons/                  Biểu tượng ứng dụng
apps-script/Code.gs     Máy chủ phân quyền (gắn với Google Sheet) + trợ lý AI giai đoạn 2
apps-script/appsscript.json  Manifest khai báo quyền tối thiểu (Drive chỉ đọc)
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

## Bản tin sáng (I005) – từ v1.5.0

Luồng dữ liệu: tác vụ định kỳ của Claude (06:19 hằng ngày) nghiên cứu, soạn bản tin, lưu file `BanTin_YYYY-MM-DD.json`
vào thư mục Drive `KhongGianSoNoiBo/BanTin`, tự xóa file quá 30 ngày → Apps Script đọc thư mục (chỉ đọc) → trang hiển thị.

1. Apps Script: dán `apps-script/Code.gs` v1.5.0 → Lưu.
2. Cài đặt dự án (bánh răng) → tick **Hiển thị tệp kê khai "appsscript.json"** → mở `appsscript.json`, dán nội dung `apps-script/appsscript.json` → Lưu.
3. Trong file phân quyền: menu **Không Gian Số → Cài đặt Bản tin sáng (I005)** → cấp quyền (Google hỏi thêm quyền *xem tệp trên Drive*).
   Script đăng ký I005, tick quyền cho Giám đốc, Phó Giám đốc, Quản trị hệ thống, Trưởng phòng, Phó phòng; đặt `BT_FOLDER_ID`.
4. Triển khai → Quản lý triển khai → Phiên bản mới. Tải `index.html` + `sw.js` mới lên GitHub.

Người xem: theo cột I005 ở sheet CanBo (bỏ tick để thu hồi). Trang chủ hiện thẻ 3 điểm nhấn; bấm vào xem toàn bộ, chọn ngày trong 30 ngày.
Bản tin mới được đọc lại sau tối đa 10 phút (bộ nhớ đệm). Lưu ý điều hành là gợi ý từ tin công khai, cần đối chiếu văn bản gốc.

## Thẻ ATM bị máy thu (I006) và Báo cáo (I007) – từ v1.6.0

Dữ liệu dùng tiếp 3 sổ trong thư mục SoBanGiaoThe (Phòng DVKH, PGD An Hữu, PGD Ấp Bắc); cột A–I giữ như sổ cũ, cột J–S do ứng dụng ghi.

1. Apps Script: dán `apps-script/Code.gs` v1.6.0 và `apps-script/appsscript.json` (thêm quyền `script.scriptapp` để cài lịch email nhắc hủy) → Lưu.
2. Trong file phân quyền: menu **Không Gian Số → Cài đặt / cập nhật Thẻ ATM máy thu (I006, I007)** → cấp quyền.
   Script tạo file cấu hình `KhongGianSoNoiBo_TheATM` (kéo vào thư mục KhongGianSoNoiBo), chuẩn bị 3 sổ, đăng ký I006/I007, cài lịch email nhắc hủy.
3. Mở file cấu hình, điền **Email nhận** (và CC) cho từng phòng ở sheet `TA_NoiNhan` (chỉ @vietinbank.vn); điền danh sách ATM ở `TA_ATM` nếu muốn in địa điểm ATM trên phiếu.
4. Tắt nhận câu trả lời ở các Google Form sổ thẻ cũ (Form → Câu trả lời → Chấp nhận câu trả lời: tắt).
5. Triển khai → Quản lý triển khai → Phiên bản mới. Tải toàn bộ thư mục (gồm `lib/tesseract/`) lên GitHub.

Tham số đổi trên sheet `TA_CauHinh`, không sửa code: số ngày hủy (mặc định 90), thứ và giờ gửi email nhắc hủy, mốc nhóm tuổi tồn, số chữ số đầu thẻ (6 theo BM09 / 4), tên miền email, vai trò nhập/trả/hủy/xem tất cả, hoàn tác trong ngày, địa danh và mã biểu mẫu in trên phiếu.
Danh mục ngân hàng và đầu số BIN ở `TA_NganHang` (BIN nội địa 9704xx theo danh sách mã BIN của NHNN).

Quy tắc: Quản trị hệ thống, Điện toán nhập thẻ (không trả); Trưởng/Phó phòng, GDV, Cán bộ DVKH của phòng nhận thẻ trả thẻ sổ phòng mình; người trả hoàn tác trong ngày;
thẻ giữ từ 90 ngày vào danh sách cần hủy và email nhắc; Trưởng/Phó phòng nhận thẻ thực hiện hủy sau khi rà soát (ghi số biên bản); thẻ đã hủy ẩn khỏi sổ, tra cứu lại được.

### Từ v1.7.0 – 6 sổ, phân quyền theo sổ

Sổ: Phòng DVKH, PGD An Hữu, PGD Ấp Bắc, PGD Cái Bè, PGD Long Định, PGD Hưng Long. Sổ mới có thêm cột T "Nguồn thẻ".
Phân quyền từng sổ ở `TA_NoiNhan` cột I–L: **Nhận bàn giao từ chi nhánh** (ô tích), **Vai trò nhập tại phòng**, **Vai trò trả thẻ**, **ID PGĐ phụ trách**.

| Việc | Ai làm |
|---|---|
| Nhập bàn giao (có email) | `VAI_TRO_NHAP` (QT, Điện toán) vào các sổ có tích "Nhận bàn giao" |
| Nhập tại phòng (không email) | Vai trò ở cột J, chỉ sổ của phòng mình |
| Trả thẻ | Vai trò ở cột K (trống thì dùng `VAI_TRO_TRA`), chỉ sổ phòng mình |
| Hủy thẻ | `VAI_TRO_HUY` (Phó phòng), chỉ sổ phòng mình |
| Báo cáo | `VAI_TRO_XEM_TAT_CA`: tất cả sổ; `VAI_TRO_BAO_CAO`: sổ phòng mình; `VAI_TRO_PGD`: các sổ có ID mình ở cột L |

Cập nhật từ v1.6.x: dán Code.gs → chạy lại menu **Cài đặt / cập nhật Thẻ ATM máy thu** (tự thêm 3 sổ, cột I–L, tham số vai trò) → Triển khai phiên bản mới → tải `index.html`, `sw.js` lên GitHub.
Đọc ảnh thẻ (v1.6.1): mặc định gửi ảnh qua Apps Script tới Claude (tham số `DOC_THE_BANG` = AI, mô hình `MODEL_DOC_THE`, cần `ANTHROPIC_API_KEY`); đặt `DOC_THE_BANG` = Máy để đọc trên điện thoại.
Ảnh không được lưu ở đâu; ứng dụng chỉ lưu số đầu và 4 số cuối, không lưu số thẻ đầy đủ. Lượt AI tính chung giới hạn `DAILY_LIMIT`/`PER_MINUTE_LIMIT`.

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
