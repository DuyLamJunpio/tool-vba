# Một module VBA cho API v1

Import `modLicenseApiV1.bas` vào workbook `.xlsm` trên Windows, VBA7. Module chứa cổng mở file, device-code/redeem/validate-device, xác minh Ed25519, cache DPAPI và các hàm kiểm tra quyền. Không import cùng các module cũ trong `license-sample` vì có thể trùng tên thủ tục. Không có quyền thì hiện hướng dẫn rồi **chỉ đóng ThisWorkbook, không lưu**, sau OK/X; Excel và workbook khác vẫn mở.

Bản C1–C7/C9 ngày 29/10/2026 **chưa được Compile hoặc chạy trong Excel**. Phải Debug → Compile VBAProject và chạy hai bộ test trước khi phát hành; kiểm tra tĩnh không thay thế bước này. Xem [RELEASE-CHECKS.md](RELEASE-CHECKS.md).

## Cấu hình

- Điền `Cfg_ApiBaseUrl`, `Cfg_ProductKey`, `Cfg_ProductVersion`, `Cfg_ActivationWebUrl` từ cấu hình phát hành được duyệt. Bốn giá trị đang để trống. API là gốc HTTPS kết thúc bằng `/api/v1`; trang kích hoạt là URL HTTPS cố định, không chứa mã/token/fingerprint.
- `Cfg_ActivationKeyIds` và `Cfg_ActivationPublicKeysBase64` là hai mảng song song, mỗi mảng đúng hai phần tử. Cặp đầu giữ `k2026-10` và public key hiện có; **cả hai phần tử của cặp thứ hai mặc định trống**. Khi xoay khóa, điền đồng thời `kid` duy nhất và public key Ed25519 32 byte dạng Base64 vào cặp thứ hai. Header chọn đúng một khóa; payload `kid` phải bằng header. Không thử khóa khác khi chữ ký sai. Khóa lạ bị từ chối. Cặp thiếu một nửa, khóa sai kích thước hoặc `kid` trùng làm cấu hình không hợp lệ.
- `Cfg_ProtectedSheetNames` mặc định `Array()`. Chủ workbook điền tên worksheet premium, ví dụ `Array("Premium", "Reports")`, và tự tạo một sheet hướng dẫn “bật macro” luôn nhìn thấy, nằm ngoài danh sách này. Không dùng mật khẩu sheet.
- `Cfg_ProxyServer` mặc định trống: đọc cấu hình proxy Windows của người dùng (proxy tĩnh/PAC/tự phát hiện), rồi dùng mặc định WinHTTP nếu không lấy được proxy người dùng. Có thể điền máy chủ/cổng được quản trị viên cung cấp; không điền thông tin đăng nhập. Timeout dùng bốn hàm `Cfg_*TimeoutMs` hiện có. Phải thử trên mạng công ty thực tế.

**Excel không nhận hoặc giữ JWT tài khoản, private key, refresh token hay shared secret.** Public key không phải bí mật. Không sửa `VerifyEd25519Detached` thành `True`; mọi lỗi xác minh đều từ chối quyền.

## Tích hợp sheet và sự kiện workbook

1. Tạo workbook thử nghiệm và sheet hướng dẫn riêng. Import module, điền cấu hình được duyệt và danh sách sheet premium.
2. Ghép [ThisWorkbook-events.txt](ThisWorkbook-events.txt) vào module **ThisWorkbook** hiện có, giữ các xử lý nghiệp vụ. `Workbook_BeforeClose` gọi `License_OnClose`; `BeforeSave` ẩn sheet và hủy lưu nếu không ẩn được; `AfterSave` hiện lại nếu phiên còn quyền. Không import tệp sự kiện như standard module.
3. Lưu bản phân phối khi mọi sheet premium đã `xlSheetVeryHidden` và sheet hướng dẫn đang visible. Sau khi được cấp quyền, module đặt premium thành `xlSheetVisible`. Khi không có quyền, đóng hoặc trước khi lưu, module ẩn lại. Tên sheet sai, workbook khóa cấu trúc hoặc không có sheet hướng dẫn visible gây `SHEET_CONFIG_INVALID` và không cấp quyền.
4. `Auto_Open` là cổng mở file trực tiếp. Với `Workbooks.Open`, chương trình mở file phải gọi `RunAutoMacros xlAutoOpen` hoặc tích hợp một cổng `Workbook_Open` phù hợp. Đảm bảo cổng chỉ chạy **một lần**, không gọi cả hai khi mở thông thường. `BeforeClose` cần thiết cho đóng bằng chương trình vì `Auto_Close` có thể không chạy.
5. Trong thủ tục premium, gọi `If Not License_FeatureAllowed("feature_key") Then Exit Sub`, hoặc `License_PremiumAllowed()` nếu chỉ kiểm tra quyền chung.

Việc tự gia hạn giữa phiên diễn ra khi một trong các hàm kiểm tra quyền được gọi; module không có timer chạy nền. Khi đến `rvi` hoặc quyền hiện tại không còn dùng được, module thử online một lần mỗi chu kỳ `rvi`, có cờ chống gọi lại. Thành công thì khách tiếp tục làm việc; thất bại chỉ được xét offline nếu là lỗi mạng và cache còn `exp`/`ogh`. Hết quyền trong phiên thì ẩn sheet, hiện hướng dẫn và đóng riêng workbook không lưu.

Ẩn sheet chỉ hạn chế chia sẻ thông thường, không mã hóa nội dung. Macro bị tắt thì VBA không thể đóng file; bản phân phối đã lưu ẩn chỉ hiện sheet hướng dẫn. Người có kỹ thuật vẫn có thể sửa VBA hoặc trích xuất dữ liệu.

## Token, đồng hồ và cache

Token được kiểm tra chữ ký trước khi tin payload, rồi kiểm tra `kid`, `prd`, `ver`, `dev`, `lic`, `jti`, `exp`, `rvi`, `ogh`, `iat`, `feat` và fingerprint cache. `rvi`/`ogh` chỉ lấy từ token đã ký.

Giữ nguyên DEC-104: token chỉ quá `exp` vẫn là credential để gọi validate-device, không dùng offline và không bắt nhập mã lại. Token mới bắt buộc chưa hết hạn. `ogh = 0` buộc validate online khi mở kể cả chưa đến `rvi`. Server từ chối rõ ràng không được dùng grace để vượt qua.

Cache duy nhất là ciphertext DPAPI trong HKCU, gồm activation token, fingerprint, thời điểm validate thành công và **thời điểm lớn nhất đã thấy**. Entropy là nhãn công khai cố định gắn với product key, dùng giống nhau khi mã hóa/giải mã. Cache cũ không có entropy hoặc không giải mã được được coi như chưa kích hoạt, không hiện lỗi giải mã. Vì vậy nâng cấp này có thể cần kích hoạt lại bằng mã mới một lần. Không có token/fingerprint/device_code trong sheet, log, URL hay hộp thoại; chỉ `user_code` được hiển thị.

Mốc giờ chỉ tăng và được lưu khi kiểm tra quyền; việc cập nhật mốc này **không thay đổi thời điểm validate**, nên không kéo dài `rvi` hoặc `ogh`. Giờ UTC hiện tại nhỏ hơn mốc đã lưu hoặc `iat` đã ký bị từ chối với `CLOCK_ROLLBACK`. Khách cần bật đồng bộ ngày giờ Windows rồi mở lại; cache vẫn giữ. Đây là kiểm tra cục bộ, không ngăn người có quyền trên máy xóa/khôi phục cả cache.

Giữ nguyên quy tắc phục hồi: xóa cache khi fingerprint khác máy, token cache sai chữ ký/`kid`/`prd`/`ver`, hoặc HTTP 4xx (trừ 429) trả `TOKEN_INVALID`, `DEVICE_REVOKED`, `DEVICE_MISMATCH`. Xóa thành công mới được đề nghị mã mới ngay trong cùng lần mở, tối đa một lần; khách có thể No/X để đóng. Xóa thất bại là `CACHE_UNAVAILABLE`. Không xóa vì mất mạng, rate limit, timestamp sai, HTTP 5xx, `LICENSE_EXPIRED`/`LICENSE_SUSPENDED`; sau gia hạn/mở khóa dùng lại credential cũ. Cập nhật mốc giờ có thể ghi lại ciphertext nhưng giữ nguyên token và thời điểm validate.

Fingerprint dùng serial ổ lấy từ `%SystemDrive%`. Nếu là `C:`, đầu vào hash giữ nguyên như bản cũ; máy dùng ổ hệ thống khác C có thể đổi fingerprint và cần kích hoạt lại. HTTPS dùng TLS 1.2/1.3, tắt redirect và giữ nguyên kiểm tra chứng chỉ.

## API và kiểm tra

- `POST /auth/device-code`: product/version/fp_hash/request_id/timestamp.
- `POST /auth/device-code/redeem`: device_code/fp_hash/request_id/timestamp.
- `POST /activations/validate-device`: Bearer activation token; body cùng token/fp_hash/request_id/timestamp.

Store đăng nhập và duyệt mã bằng phiên tài khoản riêng; JWT đó không đi qua Excel. Các route đã có trong source API local, chưa xác nhận hostname production. Xem [DEVICE-CODE-CONTRACT.md](DEVICE-CODE-CONTRACT.md).

Chạy `powershell -NoProfile -File .\contract-check.ps1` để kiểm tra hợp đồng, ASCII và SHA-256. Trong Excel chạy `License_RunMockTests`, kỳ vọng **64/64**. Sinh module hồi quy bằng `node .\tests\build-excel-regression.mjs D:\Code\digital-license-api`, import `TestResults/modLicenseRegression.bas` vào workbook thử nghiệm riêng, Compile rồi chạy `License_RunSignedRegressionTests`, kỳ vọng **57/57**. Bộ sinh dùng API signer với khóa test tạm trong bộ nhớ, không đọc `.env`; HTTP/cache/UI được giả lập, các ca DPAPI dùng mã hóa Windows thật với nơi lưu thay bằng bộ nhớ. Không đưa module hồi quy vào file khách hàng. Cần chạy thêm [44 ca thủ công](tests/manual/DeviceCodeActivationSmokeTest.md) theo môi trường thích hợp.
