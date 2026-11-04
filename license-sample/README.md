# Mẫu kiểm tra giấy phép VBA cho Excel trên Windows

Đây là bộ mã mẫu nhiều module ban đầu. Chưa có API contract, URL, khóa công khai hoặc token thật, nên các tính năng cao cấp mặc định **không được phép**. Bộ nhiều module này vẫn giữ workbook mở và hiện hướng dẫn. Bản một file mới ở [`api-v1-one-file/`](api-v1-one-file/) có chính sách đóng workbook sau OK/X và client mã thiết bị riêng; không nhập hai bộ mã cùng lúc.

## Giả định và giới hạn

- Excel trên Windows 10/11, VBA7, Office 32 hoặc 64 bit.
- `modLicenseDevice` băm `ApplicationId` và Windows `MachineGuid` bằng SHA-256. Đây là cách nhận diện đơn giản cho mẫu. Tài liệu `GIOI-HAN-CHIA-SE-FILE-VBA.md` còn nêu những thành phần khác; chính sách dung sai khi thay phần cứng chưa được chốt.
- Token được bảo vệ bằng Windows DPAPI theo tài khoản hiện tại và lưu trong HKCU qua `SaveSetting`. Không ghi token hoặc fingerprint lên sheet, log hoặc hộp thoại.
- **Khác biệt cần giải quyết:** mã mẫu đã cung cấp trước đây minh họa RSA PKCS#1 v1.5/SHA-256, còn tài liệu thiết kế trong repo dự kiến **Ed25519**. Không dùng token Ed25519 với hàm RSA. Mặc định khóa công khai để trống và bộ giải mã token trả `False`, nên mẫu không cấp quyền.
- Chu kỳ xác minh và thời gian offline trong config chỉ là giá trị mẫu. API contract và chính sách giấy phép phải xác định giá trị thật. Token hết hạn cứng hoặc máy chủ từ chối thì không áp dụng thời gian offline.
- VBA không ngăn được người có quyền sửa file gỡ/sửa macro, và không mã hóa dữ liệu trên sheet.

## Các file và nơi đặt mã

| File | Nơi đặt trong VBA Editor |
| --- | --- |
| `ThisWorkbook.txt` | Dán nội dung vào đối tượng `ThisWorkbook` |
| `modLicenseConfig.bas` | Import standard module |
| `modLicenseContract.bas` | Import standard module |
| `modLicenseCrypto.bas` | Import standard module |
| `modLicenseDevice.bas` | Import standard module |
| `modLicenseStore.bas` | Import standard module |
| `modLicenseApi.bas` | Import standard module |
| `modLicenseCore.bas` | Import standard module |
| `modLicenseTests.bas` | Import standard module |

## Import và kiểm thử khi chưa có API

1. Tạo workbook `.xlsm` mới để thử. Nhấn Alt+F11, chọn **File → Import File** và import tám file `.bas`.
2. Mở `ThisWorkbook` trong Project Explorer, dán nội dung `ThisWorkbook.txt` vào cửa sổ mã. Không import file này như standard module.
3. Chọn **Debug → Compile VBAProject**. Bộ mã chưa được biên dịch trong Excel ở workspace này, nên cần làm bước này trước khi dùng.
4. Chạy `License_RunMockTests` từ VBA Editor. Tám trường hợp giả lập kiểm tra chính sách thời gian và vector SHA-256. Chúng không gọi API và không cấp quyền thật.
5. Đóng rồi mở workbook khi macro được bật. Khi config và contract còn trống, workbook vẫn mở, hiện hướng dẫn; `License_PremiumAllowed()` trả `False`.
6. Trong tác vụ cao cấp, gọi `License_PremiumAllowed()` ngay trước khi chạy. Xem `Sample_PremiumAction`.

## Phải thay khi có API contract

- Điền `Cfg_ApplicationId`, `Cfg_VerifyUrl`, timeout, chu kỳ xác minh, thời gian offline và khóa công khai phù hợp.
- Triển khai `Contract_DecodeToken`: lấy **đúng byte gốc đã được ký** và các claim từ signed payload. Không serialize lại JSON để kiểm tra chữ ký.
- Triển khai `Contract_BuildVerifyJson` và `Contract_ReadVerifyResponse`: cấu trúc request/response, HTTP status và kết quả chấp nhận/từ chối.
- Nếu contract theo thiết kế Ed25519, thay `Crypto_VerifyRsaSha256` bằng bộ kiểm tra Ed25519 đã được kiểm thử. Không nhúng private key hoặc shared secret.
- Bổ sung luồng kích hoạt ban đầu và token fixture được ký thật để thử chữ ký đúng/sai, sai thiết bị, sai ứng dụng, hết hạn, máy chủ từ chối và offline grace.

Không tự đặt cờ cấp quyền trước khi có token/contract hợp lệ. Không dùng mật khẩu bảo vệ từng sheet và không đóng workbook khi xác minh thất bại.
