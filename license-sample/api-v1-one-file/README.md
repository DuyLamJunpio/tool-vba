# Module bản quyền VBA (API v1)

`modLicenseApiV1.bas` và [ThisWorkbook-events.txt](ThisWorkbook-events.txt) là code bảo vệ bản quyền mà tool-vba gắn vào sản phẩm Excel trước khi bán trên `shop.sunprime.vn`. Chạy trên Excel Windows, VBA7, Office 32/64 bit.

Module có cổng mở file, kích hoạt bằng mã thiết bị (device-code/redeem/validate-device), xác minh Ed25519, cache DPAPI và các hàm kiểm tra quyền. Không có quyền thì hiện hướng dẫn rồi **chỉ đóng ThisWorkbook, không lưu**, sau OK/X. Excel và workbook khác vẫn mở.

Không import cùng các module cũ trong [`license-sample/`](../README.md) vì trùng tên thủ tục. File `.bas` thuần ASCII; thông báo tiếng Việt không dấu.

## Trạng thái (08/11/2026)

- Đã kiểm tra: **64/64** mock test (`License_RunMockTests`) trong Excel thật, và kích hoạt thật sản phẩm thử `test-01` trên `shop.sunprime.vn`.
- Cấu hình chung đã điền (xem [Cấu hình](#cấu-hình)). Giá trị theo sản phẩm vẫn để trống trong source.
- Cache kích hoạt tách theo sản phẩm: mục Registry `Activation_<productKey>`. Hai sản phẩm trên cùng máy không còn ghi đè kích hoạt của nhau.
- Hộp thoại kích hoạt chỉ còn Yes/No (bỏ nút Cancel thừa). Mã 8 ký tự được tự chép vào clipboard; khách chỉ cần Ctrl+V trên trang web.
- **Đang thêm:** gửi tên máy `device_name` khi redeem (xem [Kích hoạt máy](#kích-hoạt-máy)).
- Script kiểm tra ngoài Excel chưa theo kịp các thay đổi trên (xem [Kiểm tra trước phát hành](#kiểm-tra-trước-phát-hành)).
- File thử: `out/test-01/test-01_1.0.0_ban-quyen.xlsm` (sheet "Huong dan" hiện, sheet "Du lieu" VeryHidden, có module và sự kiện ThisWorkbook).

## Cấu hình

### Chung (đã điền trong source)

| Hàm | Giá trị |
| --- | --- |
| `Cfg_ApiBaseUrl` | `https://shop.sunprime.vn/api/v1` |
| `Cfg_ActivationKeyIds` | `Array("k-d7b49d36a6111497", "")` |
| `Cfg_ActivationPublicKeysBase64` | Public key Ed25519 32 byte Base64 của `k-d7b49d36a6111497`; phần tử thứ hai trống |
| `Cfg_ActivationWebUrl` | `https://shop.sunprime.vn/account/activate` |
| `Cfg_ProxyServer` | Trống: dùng proxy Windows của người dùng (tĩnh/PAC/tự phát hiện), rồi mặc định WinHTTP |
| `Cfg_*TimeoutMs` | Resolve/connect/send 5 s, receive 10 s |

- Hai mảng khóa song song, mỗi mảng đúng hai phần tử. Cặp thứ hai chỉ điền khi xoay khóa: `kid` duy nhất và public key tương ứng.
- Header token chọn đúng một khóa; `kid` trong payload phải bằng header. Không thử khóa khác khi chữ ký sai. Khóa lạ bị từ chối. Cặp thiếu một nửa, sai kích thước hoặc trùng `kid` làm cấu hình không hợp lệ.
- Trang kích hoạt là URL HTTPS cố định, không chứa mã, token hay fingerprint.
- Proxy: có thể điền máy chủ:cổng do quản trị viên cấp. Không điền thông tin đăng nhập.

### Theo từng sản phẩm (để trống trong source, điền khi đóng gói)

| Hàm | Giá trị |
| --- | --- |
| `Cfg_ProductKey` | `product_key` của sản phẩm trên cửa hàng, ví dụ `test-01` |
| `Cfg_ProductVersion` | Phiên bản đúng của file, ví dụ `1.0.0` |
| `Cfg_ProtectedSheetNames` | Mọi sheet trừ sheet hướng dẫn "Huong dan", ví dụ `Array("Du lieu")` |

Hiện điền tay trong VBA Editor. Việc tiếp theo của tool-vba là tự điền khi bấm Tạo (xem [README gốc](../../README.md#việc-tiếp-theo-để-app-phục-vụ-đúng-mục-đích-bản-quyền)).

**Excel không nhận hoặc giữ JWT tài khoản, private key, refresh token hay shared secret.** Public key không phải bí mật. Không sửa `VerifyEd25519Detached` thành `True`; mọi lỗi xác minh đều từ chối quyền.

## Gắn vào sản phẩm

1. Import module, điền ba giá trị theo sản phẩm.
2. Ghép [ThisWorkbook-events.txt](ThisWorkbook-events.txt) vào module **ThisWorkbook** hiện có, giữ xử lý nghiệp vụ. Không import tệp này như standard module.
   - `Workbook_BeforeClose` gọi `License_OnClose`.
   - `Workbook_BeforeSave` ẩn sheet; hủy lưu nếu không ẩn được.
   - `Workbook_AfterSave` hiện lại sheet nếu phiên còn quyền.
3. Sheet hướng dẫn "Huong dan" luôn hiện. **Mọi sheet khác** nằm trong `Cfg_ProtectedSheetNames` và lưu ở `xlSheetVeryHidden`.
   - Không chọn “sheet premium” riêng (quyết định 08/11/2026): chưa kích hoạt thì khách chỉ thấy sheet hướng dẫn.
   - Có quyền thì module hiện các sheet này. Khi không có quyền, khi đóng hoặc trước khi lưu thì ẩn lại.
   - Sai tên sheet, workbook khóa cấu trúc hoặc không có sheet hướng dẫn visible: `SHEET_CONFIG_INVALID`, không cấp quyền.
   - Không dùng mật khẩu sheet.
4. Sheet hướng dẫn nêu cách Unblock file tải về và link `https://shop.sunprime.vn/guides/macros` (xem [File tải từ internet](#file-tải-từ-internet)).
5. `Auto_Open` là cổng khi mở file trực tiếp. Nếu chương trình khác mở bằng `Workbooks.Open`, nó phải gọi `RunAutoMacros xlAutoOpen` hoặc có cổng `Workbook_Open` phù hợp. Cổng chỉ chạy **một lần**.
6. Trong thủ tục cần quyền: `If Not License_FeatureAllowed("feature_key") Then Exit Sub`, hoặc `License_PremiumAllowed()` nếu chỉ kiểm tra quyền chung.

Tự gia hạn giữa phiên chỉ xảy ra khi một hàm kiểm tra quyền được gọi; không có timer nền. Khi đến `rvi` hoặc quyền hiện tại hết dùng được, module thử online tối đa một lần mỗi chu kỳ `rvi`, có cờ chống gọi lồng. Thành công thì khách làm tiếp. Lỗi mạng chỉ được xét offline khi cache còn `exp`/`ogh`. Hết quyền giữa phiên thì ẩn sheet, hiện hướng dẫn và đóng riêng workbook không lưu.

Ẩn sheet chỉ hạn chế chia sẻ thông thường, không mã hóa nội dung. Macro bị tắt thì VBA không đóng được file; bản phân phối đã lưu ẩn chỉ hiện sheet hướng dẫn. Người có kỹ thuật vẫn có thể sửa VBA hoặc trích dữ liệu.

## File tải từ internet

Windows chặn macro trong file tải từ internet (Mark of the Web). Quyết định 08/11/2026:

- Khách Unblock từng file tải về **một lần**: chuột phải file → Properties → General → tick **Unblock** → OK. Mở lại file, bấm **Enable Content**.
- Không làm script thêm trusted location: khách để file trong thư mục riêng hoặc thư mục công ty.
- Sheet hướng dẫn giải thích các bước trên và link `https://shop.sunprime.vn/guides/macros`.
- Ký VBA project: hoãn.

## Kích hoạt máy

1. Chưa có quyền: module gọi `POST /auth/device-code`, nhận mã 8 ký tự (`XXXX-XXXX`) và **tự chép mã vào clipboard**. Chỉ chép mã ngắn này, không bao giờ chép token hay fingerprint.
2. Hộp thoại Yes/No hiện mã. Yes mở trang kích hoạt; khách đăng nhập, Ctrl+V và xác nhận. No đóng workbook.
3. Khách bấm OK để kiểm tra; module gọi `POST /auth/device-code/redeem` và nhận token Ed25519. Cancel/X đóng workbook.
4. Khi redeem, module gửi tên máy Windows làm `device_name` (tùy chọn; chỉ chữ, số, `-`, `_`, `.`; tối đa 63 ký tự). Trang "Thiết bị" của khách và tab "Khách hàng" của admin dùng tên này để phân biệt máy. File cũ không gửi vẫn chạy; máy đó hiện "Máy chưa có tên".
5. Khi đến `rvi` hoặc token hết `exp`: `POST /activations/validate-device` với Bearer activation token.

Store duyệt mã bằng phiên tài khoản của khách (`/auth/device-code/lookup|approve|deny`); JWT đó không đi qua Excel.

Quản lý thiết bị trên `digital-license-api` (xong 08/11/2026, DEC-113, chờ deploy): khách xem máy, file trên từng máy và gỡ máy; tự gỡ tối đa bằng số thiết bị của gói trong 30 ngày, quá thì liên hệ hỗ trợ. Admin gỡ không giới hạn.

## Token, đồng hồ và cache

- Token được kiểm tra chữ ký trước khi tin payload, rồi kiểm tra `kid`, `prd`, `ver`, `dev`, `lic`, `jti`, `exp`, `rvi`, `ogh`, `iat`, `feat` và fingerprint cache. `rvi`/`ogh` chỉ lấy từ token đã ký.
- DEC-104: token chỉ quá `exp` vẫn dùng làm credential để gọi validate-device. Không dùng offline, không bắt nhập mã lại. Token mới nhận về phải chưa hết hạn. `ogh = 0` buộc validate online khi mở, kể cả chưa đến `rvi`. Server từ chối rõ ràng thì không dùng grace.
- Cache duy nhất là ciphertext DPAPI trong HKCU (`SaveSetting "DigitalLicenseVbaV1", "Activation_<productKey>"`). Gồm token, fingerprint, thời điểm validate thành công và **thời điểm lớn nhất đã thấy**. Entropy là nhãn công khai cố định gắn với product key.
- Cache không giải mã được coi như chưa kích hoạt, không hiện lỗi. Mục chung `Activation` của bản module trước không còn được đọc. Không có token/fingerprint/device_code trong sheet, log, URL hay hộp thoại.
- Đồng hồ: mốc giờ chỉ tăng, cập nhật khi kiểm tra quyền, **không dời thời điểm validate** nên không kéo dài `rvi`/`ogh`. Giờ UTC nhỏ hơn mốc đã lưu hoặc `iat` đã ký: `CLOCK_ROLLBACK`, giữ cache; khách bật đồng bộ giờ Windows rồi mở lại. Không chống được việc xóa/khôi phục cả cache.
- Phục hồi: xóa cache khi fingerprint khác máy, token cache sai chữ ký/`kid`/`prd`/`ver`, hoặc HTTP 4xx (trừ 429) trả `TOKEN_INVALID`, `DEVICE_REVOKED`, `DEVICE_MISMATCH`. Xóa thành công mới đề nghị mã mới, tối đa một lần mỗi lần mở. Xóa thất bại: `CACHE_UNAVAILABLE`.
- Không xóa cache vì mất mạng, 429, timestamp sai, 5xx, `LICENSE_EXPIRED`, `LICENSE_SUSPENDED`. Sau gia hạn/mở khóa dùng lại credential cũ.
- Fingerprint dùng serial ổ `%SystemDrive%`. Ổ `C:` giữ đầu vào hash như bản cũ; máy dùng ổ hệ thống khác có thể phải kích hoạt lại.
- HTTPS chỉ TLS 1.2/1.3, tắt redirect, giữ kiểm tra chứng chỉ.

## Kiểm tra trước phát hành

Ngoài Excel (tại thư mục này):

```powershell
powershell -NoProfile -File .\contract-check.ps1
node .\tests\build-excel-regression.mjs D:\Code\digital-license-api
node .\tests\vba-static-check.mjs
```

- `contract-check.ps1`: ràng buộc hợp đồng/bảo mật, ASCII, SHA-256 manifest (chữ hoa). Không nới kiểm tra để qua.
- Hai script đã cập nhật theo 08/11/2026: URL production, kid `k-d7b49d36a6111497`, cache theo sản phẩm (`CacheSection()`), hộp thoại Yes/No, clipboard, `device_name`. Bộ hồi quy thay URL production bằng `regression.invalid` nên không gọi được `shop.sunprime.vn`. Kết quả 08/11/2026: contract-check PASS, 64/64 mock, 57/57 signed.
- **Thứ tự phát hành:** file có `device_name` chỉ được tải lên cửa hàng **sau khi** API đã deploy bản nhận trường này. API cũ từ chối trường lạ nên redeem sẽ lỗi.
- `vba-static-check.mjs`: ASCII, cấu trúc khối, số tham số. **Không phải trình biên dịch VBA.**
- Tại `D:\Code\tool-vba`: `dotnet test ToolVba.sln --no-restore`. Nếu chỉ được ghi thư mục này, thêm `-p:DirectoryBuildPropsPath=<thư mục này>\tests\workspace-output.props --results-directory <thư mục này>\TestResults\dotnet-results`.

Trong Excel Windows:

1. Tạo `.xlsm` thử riêng, import `modLicenseApiV1.bas`, Debug → Compile VBAProject.
2. Chạy `License_RunMockTests`: kỳ vọng **64/64** (8 gốc/KAT RFC 8032, 8 policy, 16 phục hồi cache, 32 clock/cycle/key ring/DPAPI/SystemDrive). Không gọi API, không ghi Registry, không cấp quyền.
3. Sau lệnh `build-excel-regression.mjs`, import `TestResults/modLicenseRegression.bas` vào workbook thử riêng, Compile, chạy `License_RunSignedRegressionTests`: kỳ vọng **57/57**. Bộ sinh dùng API signer với khóa test tạm trong bộ nhớ, không đọc `.env`. **Không đưa module hồi quy vào file khách.**
4. Kích hoạt thật một sản phẩm thử trên `shop.sunprime.vn` (như `test-01`). Không ghi/chụp token, fingerprint, device_code.
5. Chạy các ca thủ công cần thiết trong [tests/manual/DeviceCodeActivationSmokeTest.md](../../tests/manual/DeviceCodeActivationSmokeTest.md) (44 ca).

Kết quả bắt buộc:

| Nhóm | Cần xác nhận |
| --- | --- |
| Kích hoạt, phục hồi cache | Chỉ hiện user_code; duyệt Store rồi redeem. Cache invalid/revoked/mismatch chỉ đề nghị mã mới một lần sau khi xóa. Xóa lỗi thì khóa, không lặp. |
| DEC-104, `ogh = 0` | Chỉ quá exp thì gia hạn online, offline khóa. `ogh = 0` vẫn validate khi mở. |
| Lỗi tạm thời | Mất mạng/429/timestamp/5xx/expired/suspended giữ credential. Server từ chối không được grace. |
| Sheet | Bản lưu: chỉ "Huong dan" hiện. Có quyền mới hiện sheet khác. Close bằng code, Save/SaveAs, mở lại với macro tắt đều không lộ sheet. Cấu hình sai thì fail-closed. |
| Đồng hồ | Giờ lùi: `CLOCK_ROLLBACK`, giữ cache. Mốc không giảm; tăng mốc không kéo dài `rvi`/`ogh`. Dùng VM. |
| Giữa phiên | Một request mỗi chu kỳ `rvi`, chặn gọi lồng, không timer nền. |
| DPAPI | Round-trip đủ bốn trường; cache cũ/khác sản phẩm coi như chưa kích hoạt. |
| HTTPS, proxy | TLS 1.2/1.3, không bỏ qua chứng chỉ, không redirect. Thử trực tiếp, proxy tĩnh, PAC, override. |
| Ổ hệ thống, hai khóa | `C:` giữ fingerprint. Cả hai khóa hợp lệ được nhận khi cấu hình đủ; kid lạ, header/payload lệch, sai khóa đều bị từ chối. |

## Hợp đồng nằm ở đâu

Nguồn chuẩn nằm trong `D:\Code\digital-license-api`:

- `docs/api/v1.md`, mục **Excel device-code activation**: request/response của device-code, redeem, validate-device.
- `docs/DECISIONS.md`: DEC-107 (cổng bản quyền Excel, sheet ẩn, gia cố client), DEC-104 (gia hạn token hết hạn).
- `docs/GIOI-HAN-CHIA-SE-FILE-VBA.md`: VBA hạn chế chia sẻ file được gì, không được gì.

`contract-check.ps1` kiểm tra source này theo hợp đồng đó.
