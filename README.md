# Công cụ đóng gói VBA bản quyền cho Excel

Ứng dụng WPF chạy trên Windows. Người đóng gói **thả file Excel vào và bấm Tạo**: app dán gói code VBA mới nhất vào một bản sao `.xlsm`, file gốc không bị sửa. Gói code và bản thân ứng dụng đều tự cập nhật từ `digital-license-api`.

Mục lục tài liệu: [docs/README.md](docs/README.md).

## Mục đích

Chủ dự án chốt lại ngày 08/11/2026:

- tool-vba gắn **code VBA bảo vệ bản quyền** vào các sản phẩm Excel của công ty trước khi tải lên cửa hàng `shop.sunprime.vn` (`digital-license-api`).
- Code bản quyền gồm [`modLicenseApiV1.bas`](license-sample/api-v1-one-file/modLicenseApiV1.bas) và [`ThisWorkbook-events.txt`](license-sample/api-v1-one-file/ThisWorkbook-events.txt). Hướng dẫn module: [license-sample/api-v1-one-file/README.md](license-sample/api-v1-one-file/README.md).
- Khách mua file trên cửa hàng và kích hoạt máy một lần. Chưa kích hoạt thì chỉ thấy sheet hướng dẫn.
- Người đóng gói: chủ dự án và nhân viên được cấp tài khoản ADMIN trên `digital-license-api`. Sau này có thể cấp thêm.
- App không dùng để phát code nghiệp vụ nội bộ cho nhân viên.

Hiện app **chưa đóng gói được module bản quyền**: máy chủ còn bỏ qua module tên `modLicense*`. Xem [Trạng thái và việc tiếp theo](#trạng-thái-và-việc-tiếp-theo).

## Dành cho người đóng gói

1. Cài một lần bằng `ToolVba-win-Setup.exe` (link `…/api/v1/tool/app/ToolVba-win-Setup.exe`). Không cần quyền admin Windows.
2. Mở app, kéo thả một hoặc nhiều file `.xlsx`/`.xlsm` vào cửa sổ (hoặc bấm **Chọn file…**), rồi bấm **Tạo file có VBA**.
3. Kết quả lưu cạnh file gốc với tên `<tên>_vba.xlsm`. Nếu tên đã có thì tự đánh số `(2)`, `(3)`… Bấm **Đổi nơi lưu** nếu muốn lưu chỗ khác.

Các thông báo trên màn hình:

- **Chip phiên bản** ở góc trên: bản gói code đang dùng. App tải bản mới khi mở và trước mỗi lần bấm Tạo. Mất mạng thì dùng bản đã lưu gần nhất.
- **Có phiên bản ứng dụng mới**: bấm **Cập nhật ngay**, app tự tải, cài và mở lại.
- **Excel chưa cho phép công cụ dán code VBA**: bấm **Bật giúp tôi**. Sau khi bạn xác nhận, app bật “Trust access to the VBA project object model” cho riêng tài khoản Windows đó (HKCU `AccessVBOM`). Tùy chọn này **giữ nguyên, app không tắt lại** (quyết định của chủ dự án). Nếu Group Policy đang chặn, app chỉ báo liên hệ IT, không ghi gì.

### File đã có VBA và file chưa có VBA

| File nguồn | App làm gì |
| --- | --- |
| `.xlsx` (chưa có VBA) | Dán toàn bộ gói code, lưu thành `.xlsm`. |
| `.xlsm` đã có VBA | Module trùng tên với gói được **thay bằng bản mới**. Module do nhân viên tự viết được giữ nguyên. Module do bản gói cũ thêm vào nhưng nay đã bỏ thì bị gỡ. |
| ThisWorkbook / sheet đã có code | Code của gói nằm trong vùng đánh dấu `'=== tool-vba BEGIN … ===`. Mỗi lần cập nhật chỉ thay vùng này. Nếu ngoài vùng đã có thủ tục trùng tên (ví dụ `Workbook_Open`), app dừng và nêu tên thủ tục trùng. |

Sau mỗi lần dán, app ghi thuộc tính tài liệu `ToolVba.PackageVersion` (phiên bản gói) và `ToolVba.Managed.N` (danh sách module do gói quản lý).

## Xuất bản gói code VBA (Admin web)

Việc xuất bản nằm trên **Admin web của digital-license-api**, mục **Tải công cụ → Phát hành tool-vba → Xuất bản code VBA**. Chỉ tài khoản ADMIN. App desktop không có phần quản trị.

1. Sửa code trong file Excel mẫu `.xlsm` (module `.bas`, class `.cls`, ThisWorkbook, code của sheet).
   - Máy chủ **hiện vẫn bỏ qua** module tên bắt đầu bằng `modLicense`. Quy tắc này đang được gỡ để gói chứa được module bản quyền.
2. Tải file `.xlsm` lên (hoặc các file VBA lẻ `.bas`, `.cls`, `.frm` + `.frx`) → **Kiểm tra** để xem danh sách → nhập ghi chú → **Xuất bản**.
3. Máy chủ tự tách code, cấp phiên bản (`YYYY.M.D.N`) và ký gói. Người đóng gói nhận bản mới ở lần bấm Tạo tiếp theo.
4. Lỡ xuất bản bản lỗi: ở danh sách bên dưới bấm **Kích hoạt lại** cạnh bản trước.

UserForm: máy chủ không có Excel nên không dựng lại được phần giao diện của form từ `.xlsm`. Nếu file mẫu có UserForm, trang sẽ yêu cầu tải thêm `.frm` + `.frx` (Alt+F11 → chuột phải vào form → Export File).

Ô **Bản ứng dụng tối thiểu** chỉ điền khi gói mới cần bản app mới hơn. Khi đó app cũ chặn nút Tạo và yêu cầu cập nhật.

## Phát hành và cấu hình

Hợp đồng API giữa app và máy chủ: [docs/tool-release-api-contract.md](docs/tool-release-api-contract.md).

**Một lần duy nhất:**

1. Chạy `node scripts/new-release-key.mjs`.
   - Đặt dòng private key vào biến `TOOL_RELEASE_SIGNING_PRIVATE_KEY` của máy chủ API. Chỉ đặt ở đó, không commit.
   - Đặt public key vào `ToolVbaReleasePublicKey` trong [src/ToolVba.App/release.props](src/ToolVba.App/release.props), cùng với `ToolVbaApiBaseUrl`.
2. Áp migration `tool_releases` trên API: `pnpm db:dev:deploy` (xem DEC-111 trong repo API).
3. Cài vpk: `dotnet tool install -g vpk --version 1.2.161`.

**Mỗi lần ra bản app mới:**

```powershell
.\scripts\build-app.ps1 -Version 1.0.1
```

Script build bản self-contained win-x64, đóng gói bằng Velopack (có gói delta) và mở thư mục `out\release\upload`. Tải toàn bộ file trong thư mục đó lên Admin web: **Tải công cụ → Phát hành tool-vba → Phát hành ứng dụng**. Người dùng thấy banner "Có phiên bản ứng dụng mới" và bấm **Cập nhật ngay**.

**Chuyển đổi:** ai đang dùng bản zip cũ cần chạy `ToolVba-win-Setup.exe` một lần; từ đó app tự cập nhật. Các file zip/artifact phân phối cũ sẽ bị xóa sau 1–2 tuần, khi kênh API đã chạy ổn.

**Chưa ký mã Authenticode:** dùng nội bộ, ít người, nên chưa mua chứng chỉ. Windows SmartScreen sẽ cảnh báo ở lần cài đầu.

Bản build không có `release.props` đã điền (chạy dev) vẫn mở được, nhưng hiển thị “chưa được cấu hình máy chủ cập nhật” và không dán code.

## Build, test và chạy

```powershell
cd D:\Code\tool-vba
dotnet test ToolVba.sln
dotnet run --project src\ToolVba.App\ToolVba.App.csproj
```

Self-test với Excel thật. Chỉ dùng file tổng hợp; cần Excel và quyền VBA Project đã bật:

```powershell
dotnet run --project tests\ToolVba.IntegrationRunner\ToolVba.IntegrationRunner.csproj -- --selftest D:\Temp\tool-vba-selftest
```

Self-test chạy lần lượt:

1. Cài gói v1 vào `.xlsx`.
2. Thêm một module của nhân viên, rồi cập nhật lên v2.
3. Kiểm tra: module của nhân viên còn nguyên, module đã bỏ bị gỡ, không có tên kiểu `modReport1`, ThisWorkbook có đúng một vùng code.
4. Xác nhận app từ chối khi file có sẵn `Workbook_Open` trùng tên.

Integration runner còn giữ hai chế độ cũ là `--bas` (một module + manifest) và chèn `vbaProject.bin`.

Kiểm thử tay: [tests/manual/WorkbookSmokeTest.md](tests/manual/WorkbookSmokeTest.md) (app) và [tests/manual/DeviceCodeActivationSmokeTest.md](tests/manual/DeviceCodeActivationSmokeTest.md) (module bản quyền).

## An toàn dữ liệu

- File nguồn luôn được giữ nguyên.
  - Excel chỉ làm việc trên bản tạm, với macro, sự kiện và cập nhật liên kết đều tắt.
  - App mở lại bản đã lưu để xác minh, so lại hash nguồn, rồi mới chuyển bản tạm thành file kết quả.
- Gói code được kiểm tra nhiều lớp:
  - Chữ ký ECDSA P-256 của máy chủ.
  - SHA-256 của file zip và của từng file bên trong.
  - Chặn zip-slip và zip bomb.
- App copy gói sang thư mục riêng và kiểm tra lại ngay trước khi import, nên cache bị sửa trên đĩa không ảnh hưởng tới code được dán.
- Log cục bộ `%LocalAppData%\ToolVba\logs\operations.jsonl` chỉ ghi tên file, hash, phiên bản gói và kết quả. Không ghi nội dung workbook hay code VBA.
- Rủi ro đã chấp nhận: máy chủ ký gói, nên **tài khoản ADMIN bị lộ có thể phát tán macro**. Hiện chỉ có một tài khoản ADMIN do chủ dự án quản lý. Dùng mật khẩu mạnh, cấp quyền ADMIN cho ít người.
- Chữ ký VBA hoặc chữ ký package cũ có thể mất hiệu lực sau khi dán code.
- Excel tiếng Việt đặt tên module workbook là `SổlàmviệcNày`. Excel làm mất tên này khi lưu `.xlsm` (module mất tên, sự kiện workbook không chạy). Vì vậy app đổi tên module này thành `ThisWorkbook` trước khi dán; từ khóa `ThisWorkbook` trong VBA không bị ảnh hưởng. File nào đã bị Excel làm mất tên module thì app dừng và hướng dẫn sửa.
- Không hỗ trợ `.xls`, `.xlsb`, `.xltx` hoặc định dạng khác.

## Thành phần

| Thư mục | Nội dung |
| --- | --- |
| `src/ToolVba.App` | WPF một màn hình; tự cập nhật bằng Velopack |
| `src/ToolVba.Core/Packages` | Đọc/tạo gói, gộp ThisWorkbook/sheet, kiểm chữ ký |
| `src/ToolVba.Core/Updates` | Tải, kiểm chữ ký và cache gói từ API |
| `src/ToolVba.Core/Excel` | Thao tác VBA project qua Excel COM |
| `license-sample/api-v1-one-file` | Module bản quyền VBA và bộ kiểm tra của nó |
| `tests/ToolVba.Tests` | Unit test, không cần Excel |
| `tests/ToolVba.IntegrationRunner` | Self-test và runner thủ công với Excel thật |
| `tests/manual` | Kịch bản kiểm thử tay |
| `scripts/` | Tạo khóa ký, build bản app để tải lên Admin web |
| `docs/` | Mục lục, hợp đồng API phát hành, lưu trữ |

## Trạng thái và việc tiếp theo

Cập nhật 08/11/2026, nhánh `feat/one-screen-auto-update`.

### Đã có

- App một màn hình như mô tả ở trên: thả file, Tạo, lưu `<tên>_vba.xlsm`; gói ký ECDSA P-256 tải từ API, có cache và dùng offline; tự cập nhật bằng Velopack; nút **Bật giúp tôi**; xuất bản và **Kích hoạt lại** trên Admin web.
- Module bản quyền đã chạy thật: **64/64** mock test trong Excel thật và kích hoạt thật sản phẩm thử `test-01` trên `shop.sunprime.vn` đều đạt (08/11/2026). Chi tiết thay đổi: [module README](license-sample/api-v1-one-file/README.md#trạng-thái-08112026).
- File thử: `out/test-01/test-01_1.0.0_ban-quyen.xlsm` (sheet "Huong dan" hiện, sheet "Du lieu" VeryHidden, có module và sự kiện ThisWorkbook). Thư mục `out/` không commit.
- Đã làm xong trên `digital-license-api` (08/11/2026, DEC-113, chờ deploy): trang **Thiết bị** cho khách (xem máy, file trên từng máy, gỡ máy; tự gỡ tối đa bằng số thiết bị của gói trong 30 ngày, quá thì liên hệ hỗ trợ). Admin tách tab **Nhân viên** và **Khách hàng** (chi tiết khách có giấy phép và máy; admin gỡ máy không giới hạn).

### Quyết định đã chốt

- Không chọn “sheet premium”. Mọi sheet trừ sheet hướng dẫn đều ẩn cho tới khi máy được kích hoạt. Khách đã mua file và chỉ cần kích hoạt máy một lần.
- File tải từ internet (Mark of the Web): khách tự Unblock từng file một lần. Không làm script thêm trusted location (khách để file trong thư mục riêng hoặc của công ty). Sheet hướng dẫn nêu cách Unblock và link `https://shop.sunprime.vn/guides/macros`.
- Ký VBA project: hoãn.
- Không mua chứng chỉ ký mã Authenticode lúc này.
- `AccessVBOM` bật bằng **Bật giúp tôi** thì giữ nguyên.
- Một tài khoản ADMIN do chủ dự án quản lý.
- Xóa zip/artifact phân phối cũ sau 1–2 tuần khi kênh API chạy ổn.

### Việc tiếp theo (để app phục vụ đúng mục đích bản quyền)

1. Gỡ quy tắc máy chủ bỏ qua module tên `modLicense*` (`digital-license-api`: `apps/api/src/modules/tool-releases/vba-source/vba-source-package.ts`).
2. Khi bấm Tạo: chọn sản phẩm và phiên bản trên cửa hàng, rồi tự điền `Cfg_ProductKey`, `Cfg_ProductVersion`, `Cfg_ProtectedSheetNames`.
3. Ghép sự kiện bản quyền (`Workbook_BeforeClose`, `Workbook_BeforeSave`, `Workbook_AfterSave`) vào thủ tục sẵn có của sản phẩm, thay vì dừng vì trùng tên.
4. Tự thêm sheet hướng dẫn "Huong dan" và ẩn mọi sheet khác (`xlSheetVeryHidden`).
5. Xóa zip/artifact phân phối cũ (sau 1–2 tuần).
