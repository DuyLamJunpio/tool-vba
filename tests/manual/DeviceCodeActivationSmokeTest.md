# Test case: kích hoạt và phục hồi cache Excel

Áp dụng cho `license-sample/api-v1-one-file/modLicenseApiV1.bas`. Tổng **49 ca thủ công**; bộ mock **71 ca**, bộ hồi quy có chữ ký **104 ca**. Ngày 08/11/2026: mock 64/64 đạt trong Excel thật và kích hoạt thật sản phẩm thử `test-01` trên `shop.sunprime.vn` đạt. Ngày 11/11/2026 (phụ lục DEC-107: thu hồi, lỗi máy chủ, nhiều bản sao): mock 71/71 và hồi quy 104/104 đạt trong Excel thật với HTTP/cache/UI giả lập; DC-45 đến DC-49 chưa thử với máy chủ thật. Dùng workbook/tài khoản thử nghiệm; không ghi/chụp token, fingerprint, private key hoặc device_code. Chỉ ghi tên ca, mã lỗi và kết quả.

API chạy tại `https://shop.sunprime.vn/api/v1`. “Cần môi trường” nghĩa là cần API/Store thật hoặc VM, không phải chưa có code API. Hướng dẫn module: [README](../../license-sample/api-v1-one-file/README.md).

## Điều kiện

- Import module vào workbook `.xlsm` thử nghiệm riêng, không import cùng module cũ. Cấu hình đúng public key/`kid`, product/version và URL HTTPS nếu cần gọi API thật.
- SHA-256 khớp manifest; Debug → Compile VBAProject không báo lỗi. Dùng profile/VM thử nghiệm riêng cho các ca xóa cache thật.
- Cache là ciphertext DPAPI trong mục `Activation_<productKey>` của sản phẩm đang thử; không xóa cache khách hàng thật để chuẩn bị thử.
- Workbook có sheet hướng dẫn "Huong dan" visible; mọi sheet khác nằm trong `Cfg_ProtectedSheetNames`. Ghép `ThisWorkbook-events.txt` vào sự kiện hiện có. Lưu bản thử/phân phối với các sheet đó very hidden. Các ca thay giờ/ổ hệ thống dùng VM riêng.
- Hồi quy có chữ ký dùng token do bộ ký API local tạo với claim license/device/JTI giả; HTTP/cache/UI là stub. Cách chạy: [Kiểm tra trước phát hành](../../license-sample/api-v1-one-file/README.md#kiểm-tra-trước-phát-hành).

| ID | Tình huống / thao tác | Kết quả mong đợi | Cách kiểm tra |
| --- | --- | --- | --- |
| DC-01 | Mẫu chưa cấu hình; mở rồi OK/X trên hướng dẫn | Chỉ workbook này đóng, không lưu; workbook khác và Excel vẫn mở | Excel thật |
| DC-02 | Không bật macro | Auto_Open không chạy; ghi nhận giới hạn, không báo nhầm là đã khóa file | Excel thật |
| DC-03 | Chưa có cache, cấu hình đủ; chọn No/X ở mã thiết bị | Không redeem, không lưu token mới, đóng riêng workbook | Cần môi trường |
| DC-04 | Chọn Yes ở lời mời | Mở URL HTTPS cố định không mang mã/token/fingerprint; chỉ hiển thị user_code | Cần môi trường |
| DC-05 | Chưa duyệt trên Store; chọn OK kiểm tra | AUTHORIZATION_PENDING, chưa cấp quyền | Cần môi trường |
| DC-06 | Duyệt mã đúng với license ACTIVE rồi redeem | Kiểm tra chữ ký/claim, lưu DPAPI, chỉ cho dùng quyền trong token mới | Cần môi trường |
| DC-07 | Không có license ACTIVE hoặc license không dùng được khi redeem | Không lưu token, không cấp quyền; hướng dẫn rồi đóng | Cần môi trường |
| DC-08 | Mã sai/hết hạn/đã dùng hoặc sai fingerprint khi redeem | Không cấp token, báo lỗi cố định an toàn | Cần môi trường |
| DC-09 | Kích hoạt lại cùng máy ACTIVE | Không tăng số thiết bị ACTIVE | Cần môi trường |
| DC-10 | Hết suất thiết bị | DEVICE_LIMIT_EXCEEDED; hướng dẫn gỡ máy cũ trong tài khoản; OK/X đóng riêng workbook | Cần môi trường |
| DC-11 | Token cache sai chữ ký/kid/prd/ver | Không validate token sai, xóa cache; đề nghị mã mới nếu cấu hình đủ, quyền vẫn False | Hồi quy có chữ ký + môi trường |
| DC-12 | Đến rvi; validate-device trả token mới hợp lệ | Lưu token mới, dùng quyền mới, không yêu cầu JWT tài khoản | Hồi quy có chữ ký + môi trường |
| DC-13 | Đến rvi, mất mạng trong/sau ogh | Giữ cache; chỉ trong grace và chưa hết exp được dùng; hết grace yêu cầu mạng | Hồi quy có chữ ký + môi trường |
| DC-14 | Server từ chối rõ ràng dù cache còn grace | Không dùng grace để vượt từ chối; chỉ ba lỗi token/device chỉ định được xóa và đề nghị mã mới. Từ 11/11/2026: năm mã quyền (revoked, suspended, expired, hết suất, không có license) còn ghi dấu từ chối vào cache | Hồi quy có chữ ký + môi trường |
| DC-15 | Workbook đã có macro khác | Không thay macro gốc; cổng giấy phép chỉ đóng ThisWorkbook không lưu | Excel thật |
| DC-16 | Chạy License_RunMockTests | **71/71**; không gọi API, không ghi/xóa Registry và không cấp quyền; DPAPI dùng dữ liệu test trong bộ nhớ | Excel thật |
| DC-17 | Token chỉ hết exp, có mạng | Gia hạn bằng validate-device; không xóa cache vì riêng exp, không bắt mã mới | Hồi quy có chữ ký + môi trường |
| DC-18 | Token chỉ hết exp, mất mạng | Giữ cache, quyền False, hướng dẫn cần mạng | Hồi quy có chữ ký + môi trường |
| DC-19 | ogh = 0, chưa đến rvi, có mạng | Vẫn gọi validate-device, chỉ token mới hợp lệ được cấp quyền | Hồi quy có chữ ký + môi trường |
| DC-20 | ogh = 0, mất mạng ngay cùng giây xác minh trước | Giữ cache, quyền False, không đề nghị mã mới | Hồi quy có chữ ký + môi trường |
| DC-21 | Server DEVICE_REVOKED sau gỡ máy hoặc license từng rời ACTIVE | Xóa cache; đề nghị mã mới ngay cùng lần mở, tối đa một lần; hủy thì đóng riêng workbook | Hồi quy Auto_Open + môi trường |
| DC-22 | Fingerprint cache khác máy sau đổi phần cứng | Xóa cache trước validate; đề nghị mã mới cùng lần mở; hướng dẫn hết suất thì gỡ máy cũ | Hồi quy Auto_Open + môi trường |
| DC-23 | Server TOKEN_INVALID khi kid bị gỡ | Xóa cache, đề nghị mã mới cùng lần mở; cache không qua kid local cũng phục hồi tương tự | Hồi quy Auto_Open + môi trường |
| DC-24 | Server DEVICE_MISMATCH qua 4xx trừ 429 | Xóa cache; chỉ đề nghị mã mới sau khi xóa thành công, quyền cũ False | Hồi quy có chữ ký + môi trường |
| DC-25 | TOKEN_INVALID do token mới server trả về hết hạn/sai chữ ký | Không lưu token mới, giữ credential cache cũ, không tự lặp kích hoạt | Hồi quy có chữ ký |
| DC-26 | NETWORK_UNAVAILABLE, RATE_LIMITED, REQUEST_TIMESTAMP_INVALID, 5xx; thêm body TOKEN_INVALID trong 5xx/429 | Không xóa credential hoặc dời thời điểm validate; không đề nghị mã mới. Mốc giờ có thể tăng. Từ 11/11/2026: 429 và 5xx được dùng `ogh` như mất mạng; REQUEST_TIMESTAMP_INVALID vẫn từ chối | Hồi quy có chữ ký + môi trường |
| DC-27 | LICENSE_EXPIRED/LICENSE_SUSPENDED rồi gia hạn/mở khóa | Lần từ chối giữ cache; lần mở sau validate-device dùng credential cũ và nhận token mới, không cần mã mới | Hồi quy có chữ ký + môi trường |
| DC-28 | CacheClear bị lỗi quyền Registry | CACHE_UNAVAILABLE, quyền False, không đề nghị mã mới, không tự lặp | Hồi quy stub lỗi xóa; VM riêng nếu thử thật |
| DC-29 | Hủy mã mới sau phục hồi, rồi mở lại | Mỗi lần mở đề nghị tối đa một lần; cache chết không bị gửi validate lại; No/X không redeem, không cấp quyền | Hồi quy Auto_Open + môi trường |
| DC-30 | Chạy License_RunSignedRegressionTests từ module test sinh ra | **104/104**: 43 ca bảng/cổng/khôi phục + 61 ca clock/refresh/DPAPI/key ring/sheet/dấu từ chối/phản hồi lạ/bản sao; không chạm Registry/browser/worksheet thật | Excel thật, HTTP/cache/UI giả lập |
| DC-31 | C1: tắt macro khi mở bản đã lưu ẩn; sau đó bật macro/kích hoạt hợp lệ | Tắt macro chỉ thấy sheet hướng dẫn; premium very hidden. Chỉ khi có quyền mới visible | Excel thật, workbook đã ghép events |
| DC-32 | C1: từ chối, đóng bằng UI/code, Save/SaveAs khi có quyền; mở bản lưu với macro tắt | Trước lưu/đóng premium very hidden; sau lưu phiên hợp lệ hiện lại. Bản lưu không lộ sheet; workbook khác không bị tác động | Excel thật, cả BeforeClose/BeforeSave/AfterSave |
| DC-33 | C1: tên sheet sai, không có guide visible, workbook khóa cấu trúc | SHEET_CONFIG_INVALID, quyền False; không lưu nếu không ẩn được, không dùng mật khẩu sheet | Excel thật |
| DC-34 | C2: giờ hiện tại thấp hơn max đã lưu | CLOCK_ROLLBACK, không validate hoặc xóa cache, ẩn premium rồi hướng dẫn/đóng. Đồng bộ giờ rồi mở lại | Signed regression + VM |
| DC-35 | C2: token có iat đã ký lớn hơn giờ máy; thử giờ bằng mốc | Từ 11/11/2026: iat đi trước giờ máy tối đa 300 giây được chấp nhận (máy chậm hơn máy chủ vẫn kích hoạt và validate được); quá 300 giây bị CLOCK_ROLLBACK. Giờ máy thấp hơn mốc đã lưu vẫn bị CLOCK_ROLLBACK, không dung sai; không tin iat từ token sai chữ ký | Signed regression + VM |
| DC-36 | C2: tăng giờ, kiểm tra quyền, mở lại rồi lùi giờ | Max không giảm; từ chối giờ thấp hơn mốc mới. Cập nhật max không dời thời điểm validate/rvi/ogh | Signed regression + VM |
| DC-37 | C3: workbook đang dùng, gọi PremiumAllowed/FeatureAllowed sau rvi, server thành công | Một validate nhận token mới, tiếp tục dùng không thông báo; hạn và quyền theo token mới | Signed regression + môi trường |
| DC-38 | C3: quá rvi, mất mạng; gọi kiểm tra nhiều lần trong cùng chu kỳ và sang chu kỳ sau | Tối đa một request/chu kỳ; trong grace còn exp thì tiếp tục. Chu kỳ sau được thử lại. Không có timer tự chạy | Signed regression + môi trường |
| DC-39 | C3: gọi lồng lúc validate, hoặc server từ chối khi còn grace; thêm token chỉ hết exp | Chặn request lồng; không vượt từ chối bằng grace. Exp-only được gia hạn online. Hết quyền thì hướng dẫn, đóng riêng workbook một lần | Signed regression + môi trường |
| DC-40 | C4: round-trip cache mới, cache cũ không entropy, cache khác product | Mới có đủ bốn trường (năm khi có dấu từ chối) và max không giảm; cache bốn trường của bản trước 11/11/2026 vẫn dùng được. Cache không entropy/khác product coi chưa kích hoạt, không hiện lỗi giải mã; mã mới một lần nếu cần | Mock/signed regression + VM |
| DC-41 | C9: token bằng khóa đầu và thứ hai, sau đó bỏ cặp thứ hai | Hai khóa được nhận khi cấu hình đủ; bỏ khóa hai thì chỉ khóa đầu dùng được. Cache kid lạ bị TOKEN_INVALID và phục hồi theo quy tắc cũ | Signed regression + môi trường xoay khóa |
| DC-42 | C9: header/payload kid khác nhau, ký bằng sai khóa của kid, kid lạ; cấu hình nửa cặp/trùng/sai kích thước | Token invalid không cấp quyền; cấu hình sai không cấp quyền; không thử khóa còn lại để vượt chữ ký | Mock/signed regression |
| DC-43 | C5/C6: TLS/chứng chỉ sai/redirect; mạng trực tiếp, proxy Windows tĩnh/PAC và override | Chỉ TLS 1.2/1.3, không bỏ kiểm tra chứng chỉ/redirect; dùng proxy theo cấu hình, lỗi có timeout, không credentials trong VBA | Mạng/VM thử nghiệm |
| DC-44 | C7: SystemDrive C: so với bản trước; VM hệ thống trên ổ khác C | C: giữ fingerprint; ổ khác dùng đúng serial hệ thống, có thể cần kích hoạt lại/gỡ máy cũ khi hết suất. Không in hash | VM và đối chiếu nội bộ không ghi hash |
| DC-45 | Phụ lục 11/11/2026, thay đổi 1: file đã kích hoạt đang mở, admin thu hồi giấy phép; sau đó tắt mạng và mở lại (các bước ở mục dưới) | Lần kiểm tra có mạng kế tiếp: báo giấy phép không còn được dùng, đóng riêng workbook. Mở lại khi tắt mạng, còn trong `ogh`: vẫn bị từ chối với "Can ket noi mang...", không hỏi mã mới | Cần môi trường |
| DC-46 | Thay đổi 2: máy chủ trả 429 khi còn trong `ogh`, rồi khi đã hết `ogh` (các bước ở mục dưới) | Trong `ogh`: dùng tiếp, không thông báo. Hết `ogh`: "Can ket noi mang...", đóng riêng workbook. Cache không bị xóa trong cả hai trường hợp | Cần môi trường |
| DC-47 | Thay đổi 3: hai bản sao của một file đã kích hoạt mở cùng lúc qua vài chu kỳ `rvi` (các bước ở mục dưới) | Cả hai chạy bình thường, không hỏi mã. Admin thấy một máy. Mỗi chu kỳ khoảng một lần validate, không phải hai | Cần môi trường |
| DC-48 | Thay đổi 1: dấu từ chối đang có, admin gia hạn giấy phép hết hạn; mở file khi có mạng | Validate thành công, dùng được ngay, không hỏi mã; sau đó tắt mạng vẫn dùng được trong `ogh` | Hồi quy có chữ ký + môi trường |
| DC-49 | Thay đổi 2: mở file sau proxy công ty trả trang chặn 403 HTML, hoặc Wi-Fi có trang đăng nhập | Như mất mạng: dùng tiếp trong `ogh`, không xóa cache, không hỏi mã mới | Hồi quy có chữ ký + mạng thử nghiệm |

## Thử tay phụ lục 11/11/2026 (DC-45, DC-46, DC-47)

Chuẩn bị chung: một sản phẩm thử trên `shop.sunprime.vn` có gói `rvi = 900` (giây) và `ogh = 72` (giờ); file đóng gói bằng tool-vba từ module bản 11/11/2026; tài khoản thử đã mua và đã kích hoạt file trên máy thử. Trong file có một macro gọi `License_PremiumAllowed()` (ví dụ một nút bấm), vì module chỉ kiểm tra khi mở file và khi macro gọi hàm này. Không ghi hay chụp token, fingerprint, device_code.

### DC-45: thu hồi rồi tắt mạng

1. Mở file khi có mạng, xác nhận các sheet hiện.
2. Trên Admin, thu hồi giấy phép của tài khoản thử.
3. Chờ quá 15 phút kể từ lúc mở file, bấm nút gọi `License_PremiumAllowed()`. Mong đợi: hộp "Giay phep khong con duoc phep su dung...", workbook đóng, Excel và file khác vẫn mở.
4. Tắt mạng (rút dây hoặc tắt Wi-Fi). Mở lại file. Mong đợi: hộp "Can ket noi mang de xac minh lai giay phep.", workbook đóng, **không** hiện mã kích hoạt mới.
5. Bật mạng, mở lại. Mong đợi: lại báo giấy phép không còn được dùng.
6. Thu hồi là vĩnh viễn (DEC-128 viết lại ngày 11/11/2026), nên dùng một giấy phép thử có thể bỏ. Cho tài khoản thử mua lại đúng sản phẩm và phiên bản đó để có giấy phép mới, rồi mở file khi có mạng. Mong đợi: hiện mã kích hoạt mới đúng một lần (máy chủ trả `DEVICE_REVOKED` thay cho `LICENSE_REVOKED` vì tài khoản đã có giấy phép mới; cần bản API có quy tắc này); xác nhận mã thì dùng được, số máy của tài khoản không tăng. Chưa mua lại thì file vẫn báo giấy phép không còn được dùng.

Biến thể không mất giấy phép: ở bước 2 bấm "Tạm khóa" thay cho "Thu hồi" (bước 3 đến 5 cho kết quả như nhau), ở bước 6 bấm "Mở lại" rồi mở file: hiện mã kích hoạt mới đúng một lần. Biến thể này đã thử với máy chủ thật ngày 11/11/2026.

### DC-46: 429 trong và ngoài `ogh`

Không chỉnh đồng hồ máy ở ca này: giờ lệch quá 5 phút làm máy chủ trả `REQUEST_TIMESTAMP_INVALID`, và chỉnh tiến rồi chỉnh về làm module báo `CLOCK_ROLLBACK`.

1. Mở file khi có mạng, để mở.
2. Làm tài khoản thử chạm giới hạn 60 lần validate mỗi giờ (ví dụ nhờ phía API gửi lặp validate bằng token thử, hoặc tạm hạ giới hạn trên môi trường thử) và giữ trạng thái đó.
3. Chờ quá 15 phút, bấm nút gọi `License_PremiumAllowed()`. Mong đợi: không có hộp thoại nào, file dùng tiếp.
4. Đóng file, mở lại khi vẫn đang bị 429. Mong đợi: mở bình thường, không thông báo.
5. Hết `ogh`: dùng một gói thử có `ogh = 1` (giờ) cho bước này. Kích hoạt, mở file một lần khi có mạng, rồi giữ trạng thái 429 quá 1 giờ kể từ lần mở đó và mở lại file. Mong đợi: "Can ket noi mang de xac minh lai giay phep.", workbook đóng, không hỏi mã mới.
6. Bỏ trạng thái 429, mở file. Mong đợi: dùng được ngay, không hỏi mã.

### DC-47: hai bản sao

1. Sao chép file đã kích hoạt thành hai file tên khác nhau, mở cả hai. Mong đợi: cả hai mở bình thường, không hỏi mã.
2. Trong 35 phút, cứ vài phút bấm nút gọi `License_PremiumAllowed()` ở cả hai file.
3. Nhờ phía API đếm số lần `validate-device` của tài khoản thử trong khoảng đó. Mong đợi: khoảng 2 đến 3 lần (một lần mỗi chu kỳ 15 phút), không phải 4 đến 6.
4. Trên Admin: tài khoản vẫn chỉ có một máy cho sản phẩm này.
5. Thu hồi giấy phép, chờ quá 15 phút, bấm nút ở file thứ nhất (bị từ chối, đóng). Tắt mạng, chờ tới khi file thứ hai đến hạn kiểm tra (chậm nhất 15 phút), bấm nút ở file thứ hai. Mong đợi: file thứ hai cũng bị từ chối với "Can ket noi mang..." và đóng.

## Chạy nhanh

1. Chạy `powershell -NoProfile -File .\contract-check.ps1` tại `license-sample/api-v1-one-file`.
2. Import `modLicenseApiV1.bas` vào workbook thử nghiệm, Compile, chạy mock 71/71.
3. Chạy `node .\tests\build-excel-regression.mjs D:\Code\digital-license-api`, import `TestResults/modLicenseRegression.bas` vào workbook thử nghiệm, Compile, chạy hồi quy 104/104. Không nhập module test vào file khách hàng.
4. Kiểm tra đầu cuối trên URL HTTPS phát hành, đặc biệt DC-21/22/23, DC-26/27/28/29, DC-31–44 và DC-45–47. Ghi kết quả an toàn, không ghi token hoặc định danh máy.
