# Module bản quyền VBA (API v1)

`modLicenseApiV1.bas` và [ThisWorkbook-events.txt](ThisWorkbook-events.txt) là code bảo vệ bản quyền mà tool-vba gắn vào sản phẩm Excel trước khi bán trên `shop.sunprime.vn`. Chạy trên Excel Windows, VBA7, Office 32/64 bit.

Module có cổng mở file, kích hoạt bằng mã thiết bị (device-code/redeem/validate-device), xác minh Ed25519, cache DPAPI và các hàm kiểm tra quyền. Không có quyền thì hiện hướng dẫn rồi **chỉ đóng ThisWorkbook, không lưu**, sau OK/X. Excel và workbook khác vẫn mở.

Không import cùng các module cũ trong [`license-sample/`](../README.md) vì trùng tên thủ tục. File `.bas` thuần ASCII; thông báo tiếng Việt không dấu.

## Trạng thái (08/11/2026)

- Đã kiểm tra: **64/64** mock test (`License_RunMockTests`) trong Excel thật, và kích hoạt thật sản phẩm thử `test-01` trên `shop.sunprime.vn`.
- Cấu hình chung đã điền (xem [Cấu hình](#cấu-hình)). Giá trị theo sản phẩm vẫn để trống trong source.
- Cache kích hoạt tách theo sản phẩm: mục Registry `Activation_<productKey>`. Hai sản phẩm trên cùng máy không còn ghi đè kích hoạt của nhau.
- Hộp thoại kích hoạt chỉ còn Yes/No (bỏ nút Cancel thừa). Mã 8 ký tự được tự chép vào clipboard; khách chỉ cần Ctrl+V trên trang web.
- Đã thêm: gửi tên máy `device_name` khi redeem (xem [Kích hoạt máy](#kích-hoạt-máy)). API bản nhận trường này đã deploy ngày 09/11/2026.
- Script kiểm tra ngoài Excel chưa theo kịp các thay đổi trên (xem [Kiểm tra trước phát hành](#kiểm-tra-trước-phát-hành)).
- **Gắn vào sản phẩm do tool-vba làm** (DEC-115): app điền ba giá trị theo sản phẩm, thêm sheet hướng dẫn, ẩn các sheet khác và ghép sự kiện vào thủ tục có sẵn. File mẫu để xuất bản gói dựng bằng [`scripts/build-license-template.ps1`](../../scripts/build-license-template.ps1).
- Kết quả các ca đóng file và `Auto_Open` (Excel thật): xem [Ca đóng file và `Auto_Open`](#ca-đóng-file-và-auto_open).

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

tool-vba tự điền khi bấm Tạo, từ mã đóng gói lấy trên Admin web (xem [README gốc](../../README.md#dành-cho-người-đóng-gói)). Trong file mẫu xuất bản lên Admin, ba giá trị này **phải trống** (`vbNullString`, `vbNullString`, `Array()`); script dựng mẫu kiểm lại sau khi lưu. Nếu đóng gói tay không qua app thì điền trong VBA Editor.

**Excel không nhận hoặc giữ JWT tài khoản, private key, refresh token hay shared secret.** Public key không phải bí mật. Không sửa `VerifyEd25519Detached` thành `True`; mọi lỗi xác minh đều từ chối quyền.

## Gắn vào sản phẩm

Từ DEC-115, **tool-vba làm các bước 1 đến 3 dưới đây khi bấm Tạo** ([README gốc](../../README.md#dành-cho-người-đóng-gói) nêu cách dùng và các trường hợp app từ chối). Danh sách bên dưới là đặc tả những gì kết quả phải có, dùng khi kiểm hoặc khi đóng gói tay.

1. Import module, điền ba giá trị theo sản phẩm.
2. Ghép [ThisWorkbook-events.txt](ThisWorkbook-events.txt) vào module **ThisWorkbook** hiện có, giữ xử lý nghiệp vụ. Không import tệp này như standard module.
   - `Workbook_BeforeClose` gọi `License_OnClose` **chỉ khi file không còn thay đổi chưa lưu**: `If ThisWorkbook.Saved Then modLicenseApiV1.License_OnClose`. Khi còn thay đổi, Excel sắp hỏi "Save changes?": chọn Save thì `Workbook_BeforeSave` ẩn sheet trước khi lưu, chọn Don't Save thì file trên đĩa vẫn là bản đã lưu (đã ẩn), chọn Cancel thì file vẫn mở và khách làm tiếp được (ca CL-1).
   - Khi sản phẩm đã có `Workbook_BeforeClose`, lời gọi nằm ở **cuối** thủ tục: `If Not Cancel And (ThisWorkbook.Saved) Then modLicenseApiV1.License_OnClose`, để sản phẩm hỏi "đóng file?" mà người dùng trả lời Không thì không mất quyền. Thủ tục có `Exit Sub`, `Exit Function` hoặc `End` bị từ chối.
   - `Workbook_BeforeSave` ẩn sheet; hủy lưu nếu không ẩn được. `Workbook_BeforeSave` và `Workbook_AfterSave` có sẵn thì lời gọi nằm ở **đầu** thủ tục.
   - `Workbook_AfterSave` hiện lại sheet nếu phiên còn quyền.
3. Sheet hướng dẫn "Huong dan" luôn hiện. **Mọi sheet khác** nằm trong `Cfg_ProtectedSheetNames` và lưu ở `xlSheetVeryHidden`.
   - Không chọn “sheet premium” riêng (quyết định 08/11/2026): chưa kích hoạt thì khách chỉ thấy sheet hướng dẫn.
   - Có quyền thì module hiện các sheet này. Khi không có quyền, khi đóng hoặc trước khi lưu thì ẩn lại.
   - Sai tên sheet, workbook khóa cấu trúc hoặc không có sheet hướng dẫn visible: `SHEET_CONFIG_INVALID`, không cấp quyền.
   - Không dùng mật khẩu sheet.
4. Sheet hướng dẫn nêu cách Unblock file tải về và link `https://shop.sunprime.vn/guides/macros` (xem [File tải từ internet](#file-tải-từ-internet)).
5. `Auto_Open` là cổng khi mở file trực tiếp. Nếu chương trình khác mở bằng `Workbooks.Open`, nó phải gọi `RunAutoMacros xlAutoOpen` hoặc có cổng `Workbook_Open` phù hợp. Cổng chỉ chạy **một lần**. File **không được có `Auto_Open` thứ hai** ở module khác: Excel khi đó không chạy cái nào (xem [Ca đóng file và `Auto_Open`](#ca-đóng-file-và-auto_open)); tool-vba từ chối file như vậy.
6. Trong thủ tục cần quyền: `If Not License_FeatureAllowed("feature_key") Then Exit Sub`, hoặc `License_PremiumAllowed()` nếu chỉ kiểm tra quyền chung.

Tự gia hạn giữa phiên chỉ xảy ra khi một hàm kiểm tra quyền được gọi; không có timer nền. Khi đến `rvi` hoặc quyền hiện tại hết dùng được, module thử online tối đa một lần mỗi chu kỳ `rvi`, có cờ chống gọi lồng. Thành công thì khách làm tiếp. Lỗi mạng chỉ được xét offline khi cache còn `exp`/`ogh`. Hết quyền giữa phiên thì ẩn sheet, hiện hướng dẫn và đóng riêng workbook không lưu.

Ẩn sheet chỉ hạn chế chia sẻ thông thường, không mã hóa nội dung. Macro bị tắt thì VBA không đóng được file; bản phân phối đã lưu ẩn chỉ hiện sheet hướng dẫn. Người có kỹ thuật vẫn có thể sửa VBA hoặc trích dữ liệu.

## Ca đóng file và `Auto_Open`

Kết quả với Excel thật (08/11/2026). Kịch bản và các bước làm tay: [tests/manual/LicenseGateSmokeTest.md](../../tests/manual/LicenseGateSmokeTest.md).

| Ca | Kết quả |
| --- | --- |
| Sản phẩm có `Workbook_BeforeClose` đặt `Cancel = True` (người dùng trả lời Không), đóng bằng lệnh giao diện `FileClose` trong Excel ẩn, macro bật | **ĐẠT.** File vẫn mở, sheet bảo vệ vẫn hiện, `License_OnClose` không chạy, cả qua dòng bản quyền cuối `BeforeClose` (`If Not Cancel Then …`) lẫn qua `Auto_Close`. |
| Đối chứng: `BeforeClose` không hủy, cùng cách đóng, workbook không có thay đổi chưa lưu | **ĐẠT.** File đóng; ngay sau dòng bản quyền sheet bảo vệ ở `xlSheetVeryHidden`, tức `License_OnClose` đã chạy. |
| Bấm **Cancel** ở hộp "Save changes?" của Excel (CL-1, chủ dự án làm tay, Excel 16.0.20527 x64) | **Bản trước: KHÔNG ĐẠT.** Với `If Not Cancel Then …`, file vẫn mở nhưng sheet "Du lieu" đã bị ẩn: Excel chạy `BeforeClose` trước hộp này nên `License_OnClose` đã chạy. **Đã sửa** (chủ dự án chọn, 08/11/2026): chỉ gọi khi `ThisWorkbook.Saved`. **Lần 2 (sau khi sửa dòng hook): vẫn KHÔNG ĐẠT.** "Du lieu" bị ẩn ngay khi bấm đóng, trước khi hộp hiện ra, do `Auto_Close` của module. **Đã sửa `Auto_Close`** cùng điều kiện. **Lần 3 (09/11/2026, file thử dựng lại, Excel hiện do script điều khiển, đóng như nút X): ĐẠT.** Khi hộp hiện, "Du lieu" vẫn còn; bấm Hủy bỏ thì file vẫn mở, "Du lieu" hiện, B1 có chữ. |
| CL-2: bấm **Don't Save** ở hộp đó (cùng script, 09/11/2026) | **ĐẠT.** Mở lại thì file đúng bản đã lưu: "Du lieu" hiện như lúc dựng, B1 trống. |
| CL-3: bấm **Save** ở hộp đó (chủ dự án làm tay) | **ĐẠT.** File đã lưu chỉ còn "Huong dan" hiện, "Du lieu" ở `veryHidden` (đọc trực tiếp từ file). |
| Một module có `Auto_Open`, `RunAutoMacros(xlAutoOpen)` | Chạy, mỗi lần gọi chạy một lần. |
| Hai module cùng có `Public Sub Auto_Open` (module bản quyền và một module khác), `RunAutoMacros(xlAutoOpen)`, thử cả hai thứ tự tạo module | **Không module nào chạy, không có lỗi.** `Application.Run "'file'!Auto_Open"` báo không có macro; gọi kèm tên module (`modLicenseApiV1.Auto_Open`) thì chạy. Cổng bản quyền khi đó **im lặng không chạy** lúc mở file. |
| Hai `Auto_Open`, mở bằng giao diện (double-click), cả hai thứ tự (AO-3, chủ dự án làm tay) | **Không module nào chạy**: cột A chỉ có tiêu đề "log" đặt sẵn, không thêm dòng `license` hay `other`, kể cả sau Enable Content. Có một hộp thông báo hiện ra (chưa ghi nội dung). Khớp với ca `RunAutoMacros`. Các ca dùng module giả vì `Auto_Open` thật gọi mạng và hiện hộp thoại. |

- Đóng bằng `Workbook.Close` (COM hoặc VBA) không chạy `Auto_Close`; `Workbook_BeforeClose` vẫn chạy. Khi đóng bằng giao diện, Excel chạy cả `Workbook_BeforeClose` lẫn `Auto_Close` của module **trước** hộp "Save changes?" (CL-1 lần 2). Vì vậy `Auto_Close` cũng chỉ gọi `License_OnClose` khi `ThisWorkbook.Saved` (phụ lục 3 của DEC-115).
- Hạn chế đã chấp nhận: nếu add-in khác để `Application.EnableEvents` tắt thì `Workbook_BeforeSave` không chạy, nên chọn Save ở hộp hỏi có thể lưu sheet đang hiện. Đây cũng là hạn chế sẵn có khi bấm Ctrl+S trong tình huống đó.
- Vì kết quả hai `Auto_Open`, tool-vba từ chối file có `Auto_Open` riêng thay vì ghép chung.

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

Chạy tự động trong Excel ẩn (ví dụ trên máy không có người ngồi): các `MsgBox` báo kết quả làm Excel ẩn treo, nên chạy trên **bản sao** nơi mỗi `MsgBox` kết quả được thay bằng ghi một ô (ví dụ `Sheet1!A1 = "PASS 64/64"`); không sửa file trong repo. Ghi PID Excel khi chạy, nếu treo chỉ dừng đúng PID đó. Kết quả 08/11/2026 theo cách này: 64/64 (module thật) và 57/57.
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
