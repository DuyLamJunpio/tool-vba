# Công cụ đóng gói VBA bản quyền cho Excel

Ứng dụng WPF chạy trên Windows. Người đóng gói **thả một file Excel của một sản phẩm, dán mã đóng gói lấy từ Admin web rồi bấm Tạo**: app dán gói code VBA mới nhất (có module bản quyền) vào một bản sao `.xlsm`, điền `productKey` và phiên bản, thêm sheet hướng dẫn và ẩn các sheet khác. File gốc không bị sửa. Gói code và bản thân ứng dụng đều tự cập nhật từ `digital-license-api`.

Mục lục tài liệu: [docs/README.md](docs/README.md).

## Mục đích

Chủ dự án chốt lại ngày 08/11/2026:

- tool-vba gắn **code VBA bảo vệ bản quyền** vào các sản phẩm Excel của công ty trước khi tải lên cửa hàng `shop.sunprime.vn` (`digital-license-api`).
- Code bản quyền gồm [`modLicenseApiV1.bas`](license-sample/api-v1-one-file/modLicenseApiV1.bas) và [`ThisWorkbook-events.txt`](license-sample/api-v1-one-file/ThisWorkbook-events.txt). Hướng dẫn module: [license-sample/api-v1-one-file/README.md](license-sample/api-v1-one-file/README.md).
- Khách mua file trên cửa hàng và kích hoạt máy một lần. Chưa kích hoạt thì chỉ thấy sheet hướng dẫn.
- Người đóng gói: chủ dự án và nhân viên được cấp tài khoản ADMIN trên `digital-license-api`. Sau này có thể cấp thêm.
- App không dùng để phát code nghiệp vụ nội bộ cho nhân viên.

Gói VBA trên máy chủ giờ chứa cả module `modLicense*` (DEC-115). Việc còn lại trước khi dùng thật là deploy và kiểm thử production: xem [Thứ tự deploy và kiểm thử production](#thứ-tự-deploy-và-kiểm-thử-production).

## Dành cho người đóng gói

1. Cài một lần bằng `ToolVba-win-Setup.exe` (link `…/api/v1/tool/app/ToolVba-win-Setup.exe`). Không cần quyền admin Windows.
2. Trên Admin web, mở sản phẩm → khu **Phiên bản** → **Sao chép mã đóng gói**. Nhập số phiên bản sắp đóng gói (gợi ý số kế tiếp; chưa có thì `1.0.0`). Admin chép vào clipboard đúng một dòng `{"productKey":"<mã>","version":"<x.y.z>"}`, không kèm token hay id.
3. Mở app, kéo thả **một** file `.xlsx`/`.xlsm` của **một** sản phẩm vào cửa sổ (hoặc bấm **Chọn file…**). Dán mã vào ô **Mã đóng gói**, hoặc nhập tay **Mã sản phẩm (productKey)** và **Phiên bản (x.y.z)**. Bấm **Tạo file có bản quyền**. App không đăng nhập vào máy chủ và không gọi route cần tài khoản.
4. Kết quả lưu cạnh file gốc với tên `<productKey>_<phiên bản>_ban-quyen.xlsm`. Nếu tên đã có thì tự đánh số `(2)`, `(3)`… Bấm **Đổi nơi lưu** nếu muốn lưu chỗ khác.
5. Tải file kết quả lên đúng sản phẩm và phiên bản trên Admin web. Máy chủ đọc `Cfg_ProductKey` và `Cfg_ProductVersion` trong file và từ chối (400 `PRODUCT_VERSION_LICENSE_MISMATCH`) nếu khác sản phẩm hoặc số phiên bản đang tạo.

**App làm gì với file** (trên bản sao, macro tắt):

- Dán gói code VBA mới nhất, trong đó có module bản quyền `modLicenseApiV1`, rồi điền `Cfg_ProductKey`, `Cfg_ProductVersion` và `Cfg_ProtectedSheetNames` (mọi sheet trừ sheet hướng dẫn).
- Thêm sheet hướng dẫn **"Huong dan"** làm sheet đầu, luôn hiện (cách Unblock file tải về, link `https://shop.sunprime.vn/guides/macros`). Mọi sheet khác lưu ở `xlSheetVeryHidden`.
- Ghép lời gọi bản quyền vào thủ tục sự kiện có sẵn của sản phẩm: `BeforeSave` và `AfterSave` ở **đầu** thủ tục, `BeforeClose` ở **cuối** dưới dạng `If Not Cancel And (ThisWorkbook.Saved) Then …` (sản phẩm hỏi "đóng file?" và trả lời Không, hoặc khách bấm Cancel ở hộp "Save changes?" của Excel, thì không mất quyền). Chỗ nào sản phẩm chưa có thủ tục thì thủ tục của gói nằm trong vùng đánh dấu `'=== tool-vba BEGIN code … ===`. Lời gọi chen vào thủ tục có sẵn nằm giữa hai dòng `'=== tool-vba BEGIN/END license hook … ===`.
- Lưu, mở lại bản đã lưu và kiểm: giá trị Cfg đúng, không sheet nào lộ, sheet hướng dẫn hiện.

**App dừng, nêu lý do và không tạo file khi:**

| Trường hợp | Vì sao |
| --- | --- |
| File có thủ tục `Auto_Open` riêng (ngoài module của gói) | Hai `Auto_Open` làm Excel **không chạy cái nào** và không báo lỗi, tức cổng bản quyền im lặng không chạy ([kết quả thử](tests/manual/LicenseGateSmokeTest.md#ca-hai-auto_open)). |
| `Workbook_BeforeClose` có `Exit Sub`, `Exit Function` hoặc `End` | Lời gọi ở cuối thủ tục có thể bị bỏ qua. |
| Workbook khóa cấu trúc (Review › Protect Workbook) | Không thêm hay ẩn sheet được. |
| Có chart sheet (kể cả sheet macro Excel 4 hay hộp thoại, đang bị gọi chung là "chart sheet") | Module chỉ ẩn được worksheet. |
| Đã có sheet tên "Huong dan" không do tool-vba tạo | Tránh ghi đè sheet của sản phẩm. |
| File chỉ có sheet hướng dẫn, không còn sheet nào để bảo vệ | Không có gì để ẩn. |
| Module ngoài gói có `Cfg_ProductKey` hoặc `License_*` (bản cũ của module bản quyền) | Tránh hai bản module. |
| Gói code trên máy chủ chưa có module bản quyền, hoặc có nhiều module định nghĩa `Cfg_ProductKey` | Xuất bản lại gói từ file mẫu (xem bên dưới). |
| Mã đóng gói hoặc cặp mã/phiên bản nhập tay sai dạng | `productKey` theo mẫu của cửa hàng, phiên bản `x.y.z`. |
| Thủ tục trùng tên ngoài vùng đánh dấu (ví dụ `Workbook_Open`) | Xem bảng dưới. |

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

1. Chuẩn bị file Excel mẫu `.xlsm` (module `.bas`, class `.cls`, ThisWorkbook, code của sheet). Với module bản quyền, dựng bằng script:

   ```powershell
   .\scripts\build-license-template.ps1
   ```

   Script dùng Excel ẩn, dán `modLicenseApiV1.bas` và các sự kiện từ [`ThisWorkbook-events.txt`](license-sample/api-v1-one-file/ThisWorkbook-events.txt) (bỏ dòng chú thích), lưu `out\license-template\license-template.xlsm` rồi mở lại để kiểm. **Ba giá trị theo sản phẩm (`Cfg_ProductKey`, `Cfg_ProductVersion`, `Cfg_ProtectedSheetNames`) phải trống** trong mẫu; app điền chúng khi bấm Tạo. Sai thì script báo lỗi và xóa file. Excel treo thì chỉ dừng đúng PID ghi trong `excel.pid`.
   - Máy chủ giữ module `modLicense*` trong gói (DEC-115 bỏ quy tắc bỏ qua cũ).
2. Tải file `.xlsm` lên (hoặc các file VBA lẻ `.bas`, `.cls`, `.frm` + `.frx`) → **Kiểm tra** để xem danh sách (phải thấy `modLicenseApiV1`) → nhập ghi chú → **Xuất bản**.
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

Script build bản win-x64 **không kèm .NET** (cả gói khoảng 16 MB), đóng gói bằng Velopack (có gói delta) và mở thư mục `out\release\upload`. Bộ cài tự tải .NET 10 Desktop Runtime từ Microsoft nếu máy chưa có, nên lần cài đầu cần mạng. Lý do: bản kèm .NET nặng khoảng 139 MB, vượt giới hạn 100 MB mỗi lần gửi của Cloudflare (DEC-117 ở repo API). Nếu báo "There is a release in channel win which is equal or greater", xóa `out\release\releases` (bản build cũ chưa tải lên) rồi chạy lại. Tải toàn bộ file trong thư mục đó lên Admin web: **Tải công cụ → Phát hành tool-vba → Phát hành ứng dụng**. Người dùng thấy banner "Có phiên bản ứng dụng mới" và bấm **Cập nhật ngay**.

**Chuyển đổi:** ai đang dùng bản zip cũ cần chạy `ToolVba-win-Setup.exe` một lần; từ đó app tự cập nhật. Bản zip và thư mục `artifacts/` cũ đã gỡ khỏi repo ngày 09/11/2026 (vẫn còn trong lịch sử Git).

**Chưa ký mã Authenticode:** dùng nội bộ, ít người, nên chưa mua chứng chỉ. Windows SmartScreen sẽ cảnh báo ở lần cài đầu.

Bản build không có `release.props` đã điền (chạy dev) vẫn mở được, nhưng hiển thị “chưa được cấu hình máy chủ cập nhật” và không dán code.

## Thứ tự deploy và kiểm thử production

Danh sách việc cho chủ dự án, làm **đúng thứ tự** (DEC-115 mục 6: API deploy trước khi tải gói mẫu hoặc file sản phẩm có module bản quyền lên). Bước 1 xong ngày 09/11/2026: API, Admin và Store chạy bản `0efd869`. File mẫu của bước 2 đã dựng, chưa tải lên. Các bước sau chưa làm.

1. **Deploy API, Admin và Store** từ nhánh `phase/license-packaging` của `digital-license-api` (đã gồm `phase/device-management`: tên máy, trang Thiết bị, tab Khách hàng). Không có migration mới.
2. **Tải file mẫu lên Admin web.** Chạy `.\scripts\build-license-template.ps1`, rồi ở **Tải công cụ → Phát hành tool-vba → Xuất bản code VBA** tải `out\license-template\license-template.xlsm` lên, bấm **Kiểm tra** (dryRun). Danh sách phải có module bản quyền `modLicenseApiV1`. Đúng thì bấm **Xuất bản**.
3. **Build và phát hành bản app mới:** `.\scripts\build-app.ps1 -Version <x.y.z>`, tải thư mục `out\release\upload` lên **Phát hành ứng dụng**. Nhân viên mở app, thấy banner và bấm **Cập nhật ngay**.
4. **Đóng gói thử trong app:** lấy mã đóng gói của `test-01` trên Admin web (**Sao chép mã đóng gói**), mở app, dán mã vào ô **Mã đóng gói**, thả file Excel thử, bấm **Tạo file có bản quyền**.
5. **Kiểm trên cửa hàng:** tải file kết quả lên sản phẩm `test-01` (đúng phiên bản), rồi dùng tài khoản khách tải về, Unblock file, mở, bấm Enable Content và kích hoạt máy. Kiểm lần lượt:
   - các sheet sản phẩm hiện ra sau khi kích hoạt (trước đó chỉ thấy "Huong dan");
   - **tên máy** hiện ở trang **Thiết bị** của khách;
   - tab **Khách hàng** trong Admin thấy khách, giấy phép và máy.

Nếu bước 2 không thấy `modLicenseApiV1` hoặc bước 5 bị từ chối 400 `PRODUCT_VERSION_LICENSE_MISMATCH`, dừng lại và xem [mục Phát hành](#xuất-bản-gói-code-vba-admin-web) hoặc báo lại.

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

Self-test chạy lần lượt (24 bước, kỳ vọng `PASS: Self-test hoàn tất.`):

1. Đóng gói bản quyền v1 vào `.xlsx`: sheet đầu là "Huong dan" và đang hiện, có đúng một sheet hướng dẫn, mọi sheet khác `xlSheetVeryHidden`, module bản quyền có mã sản phẩm và phiên bản, tên sheet có dấu ghi bằng `ChrW`.
2. Thêm một module của nhân viên, rồi cập nhật lên v2 và kiểm lại như trên.
3. Kiểm tra: module của nhân viên còn nguyên, module đã bỏ bị gỡ, không có tên kiểu `modReport1`, ThisWorkbook tên `ThisWorkbook` và có đúng một vùng code.
4. Xác nhận app từ chối khi file có sẵn `Workbook_Open` trùng tên, và khi sản phẩm có `Auto_Open` riêng.
5. Xác nhận lời gọi bản quyền được ghép vào `Workbook_BeforeSave` có sẵn: nằm ở đầu thủ tục, có đúng một thủ tục, code của nhân viên giữ nguyên.

Integration runner còn giữ hai chế độ cũ là `--bas` (một module + manifest) và chèn `vbaProject.bin`.

Module bản quyền có bộ kiểm riêng chạy trong Excel thật (64 ca mock và 57 ca hồi quy có chữ ký): xem [Kiểm tra trước phát hành](license-sample/api-v1-one-file/README.md#kiểm-tra-trước-phát-hành). Dựng sẵn mẫu `.xlsm` để xuất bản lên Admin: `.\scripts\build-license-template.ps1`.

Kiểm thử tay: [tests/manual/WorkbookSmokeTest.md](tests/manual/WorkbookSmokeTest.md) (app), [tests/manual/LicenseGateSmokeTest.md](tests/manual/LicenseGateSmokeTest.md) (đóng file và hai `Auto_Open`; file thử dựng bằng `tests/manual/build-manual-test-files.ps1`) và [tests/manual/DeviceCodeActivationSmokeTest.md](tests/manual/DeviceCodeActivationSmokeTest.md) (kích hoạt và cache của module bản quyền).

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
| `scripts/` | Tạo khóa ký, build bản app để tải lên Admin web, dựng file `.xlsm` mẫu có module bản quyền, chạy mọi kiểm thử module không cần máy chủ (`run-license-tests.ps1`), dựng file thử module với máy chủ thật (`build-license-dev-workbook.ps1`) |
| `docs/` | Mục lục, hợp đồng API phát hành, lưu trữ |

## Trạng thái và việc tiếp theo

Cập nhật 09/11/2026, nhánh `feat/license-packaging` (đã push; làm cùng nhánh `phase/license-packaging` của `digital-license-api`, DEC-115).

### Đã có

- App một màn hình như mô tả ở trên: thả file, dán mã đóng gói, Tạo, lưu `<productKey>_<phiên bản>_ban-quyen.xlsm`; gói ký ECDSA P-256 tải từ API, có cache và dùng offline; tự cập nhật bằng Velopack; nút **Bật giúp tôi**; xuất bản và **Kích hoạt lại** trên Admin web.
- Đóng gói bản quyền trong app (B1 đến B4): điền `productKey`/phiên bản/sheet bảo vệ, ghép sự kiện vào thủ tục có sẵn (`BeforeClose` đặt cuối), sheet "Huong dan", ẩn sheet khác, kiểm lại bản đã lưu; từ chối các trường hợp nêu ở [Dành cho người đóng gói](#dành-cho-người-đóng-gói).
- Script `scripts/build-license-template.ps1` dựng file mẫu để xuất bản (B5); ba giá trị theo sản phẩm được kiểm là trống.
- Kết quả kiểm với Excel thật (08/11/2026): `dotnet test ToolVba.sln` 247/247, build cả solution không cảnh báo, self-test 24/24 bước, `contract-check.ps1` đạt, **64/64** mock và **57/57** hồi quy có chữ ký; ca đóng file (người dùng trả lời Không, và đối chứng không hủy) và ca hai `Auto_Open` đã ghi trong [module README](license-sample/api-v1-one-file/README.md#ca-đóng-file-và-auto_open). Kích hoạt thật sản phẩm thử `test-01` trên `shop.sunprime.vn` đã đạt trước đó.
- Đã deploy trên `digital-license-api` (09/11/2026, DEC-113): trang **Thiết bị** cho khách (xem máy, file trên từng máy, gỡ máy; tự gỡ tối đa bằng số thiết bị của gói trong 30 ngày, quá thì liên hệ hỗ trợ). Admin tách tab **Nhân viên** và **Khách hàng** (chi tiết khách có giấy phép và máy; admin gỡ máy không giới hạn).

### Quyết định đã chốt

- Không chọn “sheet premium”. Mọi sheet trừ sheet hướng dẫn đều ẩn cho tới khi máy được kích hoạt. Khách đã mua file và chỉ cần kích hoạt máy một lần.
- File tải từ internet (Mark of the Web): khách tự Unblock từng file một lần. Không làm script thêm trusted location (khách để file trong thư mục riêng hoặc của công ty). Sheet hướng dẫn nêu cách Unblock và link `https://shop.sunprime.vn/guides/macros`.
- Ký VBA project: hoãn.
- Không mua chứng chỉ ký mã Authenticode lúc này.
- `AccessVBOM` bật bằng **Bật giúp tôi** thì giữ nguyên.
- Một tài khoản ADMIN do chủ dự án quản lý.
- DEC-115 (xem `docs/DECISIONS.md` ở repo API, kèm hai phụ lục): app không đăng nhập, nhận mã đóng gói dán từ Admin; mỗi lần một file một sản phẩm; kết quả `<productKey>_<phiên bản>_ban-quyen.xlsm`; `BeforeSave`/`AfterSave` đặt đầu thủ tục, `BeforeClose` đặt cuối (`If Not Cancel And (ThisWorkbook.Saved) Then …`, phụ lục 3 sau ca CL-1) và bị từ chối nếu có `Exit Sub`/`Exit Function`/`End`; file có `Auto_Open` riêng bị từ chối; máy chủ kiểm `Cfg_ProductKey`/`Cfg_ProductVersion` khi tải file lên.
- Zip/artifact phân phối cũ: chủ dự án quyết định gỡ khỏi repo ngày 09/11/2026, sớm hơn kế hoạch 1–2 tuần.

### Việc tiếp theo

0. **Module bản quyền bản 11/11/2026 (chưa commit, chưa phát hành):** ghi nhớ lần bị máy chủ từ chối, coi 429/5xx/phản hồi lạ như mất mạng (nới DEC-107 mục 5), các bản sao dùng chung mốc kiểm tra, token được cấp sớm hơn đồng hồ máy tối đa 300 giây (nới DEC-107 mục 4). Đã đạt ngoài và trong Excel với HTTP giả lập (contract-check, static-check, `dotnet test` 250/250, **71/71** mock, **104/104** hồi quy có chữ ký), và đã thử kích hoạt, tạm khóa, mở lại với máy chủ thật bằng file dựng tại chỗ. Trước khi phát hành còn: Compile và hai bộ kiểm thử trên Excel do người chạy, thử tay DC-45 đến DC-47 với máy chủ thật, đặt gói `rvi = 900` và `ogh = 72` (giờ), chép phụ lục sang `docs/DECISIONS.md` của repo API. Xem [phụ lục DEC-107 trong module README](license-sample/api-v1-one-file/README.md#phụ-lục-dec-107-11112026-thu-hồi-lỗi-máy-chủ-nhiều-bản-sao).
1. Làm tiếp từ bước 2 của [Thứ tự deploy và kiểm thử production](#thứ-tự-deploy-và-kiểm-thử-production).
2. Các ca thử tay trong [LicenseGateSmokeTest.md](tests/manual/LicenseGateSmokeTest.md) đã có kết quả (CL-1 đến CL-3 ĐẠT ngày 09/11/2026; AO-3: không `Auto_Open` nào chạy). Còn thiếu nội dung hộp thông báo hiện ra ở AO-3; ghi vào module README khi thử lại.

Để đợt sau (từ review ngày 08/11/2026, chủ dự án đồng ý hoãn):

- Test cho ViewModel của app (`MainViewModel`, `ProductInputViewModel`); app hiện chưa có project test.
- Thông báo đúng loại sheet: sheet macro Excel 4 và sheet hộp thoại đang bị báo là "chart sheet".
- Câu khuyên khi Excel báo lỗi COM lúc thao tác sheet: hiện mọi lỗi đều nhận "đóng hết Excel rồi thử lại".
- Từ chối nhầm theo hướng an toàn: `Exit Sub`, `End`, `GoTo`, `Resume` nằm trong chú thích cuối dòng hoặc trong chuỗi vẫn làm file bị từ chối.
