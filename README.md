# Không Gian Số Nội Bộ · TTG Inside

Bộ tiện ích nội bộ dành cho cán bộ **VietinBank Chi nhánh Tây Tiền Giang**, thiết kế cho điện thoại.
Phiên bản 2.1.0 – 11/11/2026: thông báo popup và thông báo đẩy cho sổ công tác; thêm I011 Tính lãi tiết kiệm và I012 ATM và điểm giao dịch. Bản 2.0.0: Sổ công tác (I009, I010). Bản 1.9.x: Theo dõi giấy ủy quyền.

| Nhóm | Tiện ích | Mô tả |
|---|---|---|
| Thông tin | Bản tin sáng (I005) | Bản tin 06:30 mỗi sáng: điểm nhấn, lưu ý và gợi ý điều hành, chỉ số thị trường, văn bản sắp hiệu lực, tin địa bàn, chính sách, cạnh tranh; xem lại 30 ngày |
| Thẻ ATM | Quản lý thẻ ATM bị máy thu (I006) | Chụp thẻ để đọc tự động số thẻ, họ tên, ngân hàng (đọc ngay trên điện thoại); nhập theo lô và email phòng nhận; sổ thẻ nhóm theo ngân hàng; trả thẻ, in phiếu BM09-GNT hoặc gửi qua email file PDF / Word (Word dùng đúng mẫu BM09 gốc, sửa được) |
| Thẻ ATM | Báo cáo thẻ ATM máy thu (I007) | Thẻ tồn theo thời gian giữ, theo ngân hàng, theo phòng; thẻ đã trả trong kỳ; thẻ quá hạn cần hủy: xuất danh sách rà soát, tải CSV, thực hiện hủy; tra cứu thẻ đã hủy |
| Mã QR | Tạo QR kết nối WiFi | 4 bước: chọn điểm giao dịch → chọn mạng WiFi → quét thử → in thẻ niêm yết A6 |
| Chi phí | Quản lý chi phí (I003) | Hạn mức phê duyệt còn lại theo thời gian thực; trưởng/phó phòng nhập khoản đã duyệt; PGĐ, Giám đốc, kế toán xem theo quyền |
| Chi phí | Báo cáo chi phí (I004) | Tổng quan, dự báo hạn mức, theo phòng/nhóm, ma trận, xu hướng 12 tháng, kế hoạch năm, khoản lớn/lặp lại, nhận xét AI |
| Phục vụ khách hàng | Tính lãi tiết kiệm (I011) | Tiền lãi theo 5 phương thức lĩnh lãi, lịch lĩnh lãi; so sánh cầm cố sổ với rút trước hạn. Tham số ở sheet TL_CauHinh |
| Phục vụ khách hàng | ATM và điểm giao dịch (I012) | Danh sách ATM/CDM và điểm giao dịch đọc từ file quản lý ATM; tìm kiếm, lọc, tìm điểm gần nhất, bản đồ, chỉ đường Google Maps |
| Mã QR | Tạo QR theo yêu cầu | Biểu mẫu tạo mã: đường link, gọi điện, tin nhắn, email, danh thiếp, văn bản; ô nhập tự nhận dạng |

## Cấu trúc thư mục

```
index.html              Trang chính (toàn bộ giao diện và chương trình; font, logo nhúng sẵn)
manifest.webmanifest    Khai báo để cài lên màn hình chính điện thoại
sw.js                   Service worker: mở nhanh, dùng được khi mạng yếu; nhận thông báo đẩy
lib/qrcode.js           Thư viện tạo mã QR (qrcode-generator 1.4.4, MIT)
lib/jsQR.js             Thư viện đọc mã QR từ ảnh (jsQR 1.4.0, Apache-2.0)
lib/leaflet/            Thư viện bản đồ (Leaflet 1.9.4, BSD-2) – chỉ tải khi mở màn hình ATM; nền bản đồ OpenStreetMap
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

### Từ v1.8.0 – Lô chờ hủy

Quy trình hủy 2 bước tại Báo cáo thẻ ATM máy thu (I007):
1. **Lập lô chờ hủy** (`VAI_TRO_CHO_HUY`, mặc định Trưởng phòng, Phó phòng): chọn thẻ quá hạn → "Chuyển sang chờ hủy". Mỗi lần chuyển là một lô riêng, mã `LH-<tiền tố>-yyMMdd-NN` (VD: LH-CB-261104-01), không gộp.
2. **Hủy theo lô** (`VAI_TRO_HUY`, Phó phòng): mở lô → ghi số biên bản → hủy toàn bộ thẻ còn "Chờ hủy" trong lô.

Trong thời gian chờ: thẻ vẫn hiện trong sổ I006 (nhãn Chờ hủy), khách đến nhận vẫn trả được và thẻ tự rời lô; có thể "Rút khỏi lô" để thẻ về Đang giữ. Mỗi lô in được danh sách kiểm kê và tải CSV.
Sổ thêm cột U–W (Lô chờ hủy, Cán bộ lập lô, Thời gian lập lô) – ứng dụng tự thêm khi đọc sổ, chạy lại menu cài đặt để định dạng.

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

## Theo dõi giấy ủy quyền (I008) – từ v1.9.0

- **Phòng áp dụng:** Phòng DVKH và 5 PGD (sheet `UQ_Phong`). Mỗi phòng chỉ thấy giấy của phòng mình.
- **Nhập (Trưởng/Phó phòng):** tích xác nhận đã che thông tin nhạy cảm → chọn file PDF → AI (Claude) kiểm tra theo 15 tiêu chí ở `UQ_TieuChi` (BLDS 2015, Luật Doanh nghiệp 2020 sửa đổi 2025, TT 17/2024) → cán bộ sửa thông tin nếu cần → **Tiếp tục nhập** hoặc **Từ chối nhập**. Kết quả "Cần xem lại/Chưa phù hợp" phải ghi lý do vẫn nhập.
- **Bảng theo dõi:** giấy gần hết hạn nhất ở trên; lọc Sắp hết hạn/Đã hết hạn; "Bảng đầy đủ" xem đủ cột và tải CSV; GDV được xem bảng và file PDF.
- **Đã thay ủy quyền khác:** chỉ Trưởng/Phó phòng; giấy chuyển sang Đã thay thế, tra cứu ở tab Đã thay thế. Khi nhập giấy mới cùng doanh nghiệp, ứng dụng gợi ý chuyển giấy cũ.
- **Email nhắc:** hằng ngày từ 10 ngày trước hạn đến 30 ngày sau hạn, gửi email Trưởng/Phó phòng ở `UQ_Phong`.
- **Lưu trữ:** sổ ở file `KhongGianSoNoiBo_UyQuyen`, file PDF ở thư mục `KhongGianSoNoiBo_UyQuyen_File` theo phòng (không chia sẻ). Không lưu số CCCD; dãy số dài (số tài khoản…) chỉ giữ 4 số cuối.
- **Quyền Drive:** `appsscript.json` đổi `drive.readonly` → `drive` để lưu file PDF; lần chạy đầu Google hỏi cấp lại quyền.

## Sổ công tác (I009) và Quản lý sổ công tác (I010) – từ v2.0.0

- **I009 – mọi cán bộ 10 phòng (trừ Ban Giám đốc):** đăng ký công tác (ngày kết thúc không chọn được ngày trước ngày bắt đầu, hỏi lại khi khác ngày; giờ/phút chọn bằng combo, bước 15 phút; nút Sáng/Chiều/Cả ngày; "Lặp lại" chuyến cũ), báo ngay cho Trưởng/Phó phòng qua email, sổ của tôi theo tháng (từ 10/2026), lịch tháng, số ngày lũy kế công tác phí.
- **I010 – Trưởng/Phó phòng (phòng mình), PGĐ (phòng phụ trách theo `CP_PhuTrach`), Quản trị hệ thống (toàn quyền):** sổ thời gian thực (mới nhất ở trên, tự cập nhật 60 giây), **Từ chối phê duyệt** / khôi phục, lọc theo cán bộ/trạng thái, bảng chấm công tác phí, xuất 2 PDF (Báo cáo sổ công tác và chấm công tác phí; Bảng chấm công công tác phí khổ ngang).
- **Quy tắc (như skill chấm công tác phí):** tính theo ngày bắt đầu, nhiều lượt cùng ngày tính 1; không tính: mục đích tại chi nhánh, nội dung có "chi nhánh"/"trình ký hồ sơ", xe cơ quan, lượt bị từ chối; đủ điều kiện khi **> 10 ngày** → 1.200.000 đ. Chỉ áp dụng vai trò ở `VAI_TRO_CTP` và cán bộ `PHONG_CTP`.
- **Hạn khai báo:** đăng ký trước tối đa 7 ngày, khai bù tối đa 3 ngày; sổ tháng khóa sau hết ngày 1 tháng sau.
- **Email:** gửi ngay; khi hạn mức email trong ngày còn ít hơn `EMAIL_DU_PHONG` (hoặc đặt `CHE_DO_EMAIL` = Theo ca) thì gom gửi 2 ca 11:30 và 17:00. Ngày 1 hằng tháng 18 giờ tự gửi 2 PDF tháng trước cho Trưởng phòng.
- **Dữ liệu cũ:** menu "Nhập dữ liệu sổ công tác từ Google Form (từ 10/2026)" – ghép tên với CanBo, tên không ghép được hiện ở `SCT_BiDanh` để điền ID rồi chạy lại (không nhập trùng).

## Thông báo popup – từ v2.1.0

- **Khi nào có thông báo:** cán bộ đăng ký công tác → người thuộc `VAI_TRO_POPUP` (sheet SCT_CauHinh) và có phòng đó trong phạm vi xem sổ nhận thông báo. Lãnh đạo phòng nhận của phòng mình; vai trò xem tất cả/toàn quyền (Quản trị hệ thống) nhận của mọi phòng. Muốn ngừng nhận: xóa tên vai trò khỏi `VAI_TRO_POPUP`.
- **Hai dạng:** (1) popup trong ứng dụng khi đang mở – trang hỏi máy chủ mỗi `POPUP_CHU_KY_GIAY` giây; (2) thông báo đẩy ra màn hình điện thoại khi đã đóng ứng dụng – cán bộ bấm chuông → **Bật thông báo trên máy này** (iPhone phải cài TTG Inside ra màn hình chính, iOS 16.4 trở lên).
- **Tham số:** `THONG_BAO_POPUP` (Bật/Tắt), `VAI_TRO_POPUP`, `POPUP_CHU_KY_GIAY`, `GUI_EMAIL_DANG_KY` (Tắt = không gửi email mỗi lượt đăng ký nữa; báo cáo PDF ngày 1 vẫn gửi).
- **Dữ liệu:** sheet `ThongBao` (sổ thông báo, giữ 7 ngày gần nhất trên màn hình) và `ThongBaoMay` (các máy đã bật thông báo) trong file phân quyền. Không sửa tay cột Khóa máy.
- **Bảo mật:** tin đẩy không mang nội dung; điện thoại tự hỏi lại máy chủ để lấy nội dung, nên dịch vụ đẩy của Google/Apple không đọc được. Màn hình khóa chỉ hiện tên cán bộ, phòng, thời gian, mục đích – không hiện nội dung chi tiết. Khóa ký tự sinh, lưu ở Thuộc tính tập lệnh `TB_VAPID_PRIV`, `TB_VAPID_PUB` (không xóa: xóa thì mọi máy phải bật lại thông báo).
- **Thử:** menu **Không Gian Số → Gửi thử thông báo popup cho tôi…**

## Tính lãi tiết kiệm (I011) và ATM, điểm giao dịch (I012) – từ v2.1.0

- Cài đặt: menu **Không Gian Số → Cài đặt / cập nhật Tính lãi, ATM và điểm giao dịch (I011, I012)**. Tạo sheet `TL_CauHinh`, `ATM_CauHinh` trong file phân quyền, đăng ký 2 mã và tick sẵn cho mọi vai trò.
- `TL_CauHinh`: kỳ hạn, lãi suất không kỳ hạn, biên độ và lãi suất vay cầm cố tối thiểu, tỷ lệ vay tối đa, cơ sở ngày. Công thức quy đổi lãi suất theo CV 4277/CV-NHCT3 nằm trong `index.html`.
- `ATM_CauHinh`: `FILE_ID` của file quản lý ATM (sheet `DS_ATM`, `DS_PGD`), có hiện máy chưa công khai cho cán bộ hay không, bật/tắt bản đồ. Ứng dụng tìm cột theo tên tiêu đề; không hiển thị cột "Ghi chú nội bộ". Sửa file ATM xong bấm **Làm mới** trên màn hình (dữ liệu đệm 10 phút).
- Ảnh trên Drive phải chia sẻ "Bất kỳ ai có đường liên kết". Vị trí "gần tôi" chỉ tính trên điện thoại, không gửi về máy chủ.
