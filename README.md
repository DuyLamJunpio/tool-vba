# Công cụ đóng gói VBA cho Excel

Ứng dụng WPF chạy trên Windows. Nhân viên **thả file Excel vào và bấm Tạo**: app tự dán bộ code VBA mới nhất của công ty vào một bản sao `.xlsm`, file gốc không bị sửa. Code VBA và bản thân ứng dụng đều được cập nhật tự động từ `digital-license-api`, nhân viên không phải tải file về.

## Dành cho nhân viên

1. Cài một lần bằng `ToolVba-win-Setup.exe` (link `…/api/v1/tool/app/ToolVba-win-Setup.exe`). Không cần quyền admin.
2. Mở app, kéo thả một hoặc nhiều file `.xlsx`/`.xlsm` vào cửa sổ (hoặc bấm **Chọn file…**), rồi bấm **Tạo file có VBA**.
3. Kết quả được lưu cạnh file gốc với tên `<tên>_vba.xlsm`; nếu tên đã có thì tự đánh số `(2)`, `(3)`… Bấm **Đổi nơi lưu** nếu muốn lưu chỗ khác.

Các thông báo trên màn hình:

- **Chip phiên bản** ở góc trên: bản code VBA đang dùng. App tự tải bản mới khi mở và trước mỗi lần bấm Tạo. Khi mất mạng, app dùng bản đã lưu gần nhất.
- **Có phiên bản ứng dụng mới**: bấm **Cập nhật ngay**, app tự tải, cài và mở lại.
- **Excel chưa cho phép công cụ dán code VBA**: bấm **Bật giúp tôi**. App chỉ bật “Trust access to the VBA project object model” cho riêng tài khoản Windows đó (HKCU) sau khi bạn xác nhận. Nếu chính sách công ty (Group Policy) đang chặn, app chỉ báo liên hệ IT, không ghi gì.

### File đã có VBA và file chưa có VBA

| File nguồn | App làm gì |
| --- | --- |
| `.xlsx` (chưa có VBA) | Dán toàn bộ code công ty, lưu thành `.xlsm`. |
| `.xlsm` đã có VBA | Module trùng tên với code công ty được **thay bằng bản mới**. Module do nhân viên tự viết được giữ nguyên. Module do bản cũ của công ty thêm vào nhưng nay đã bỏ thì bị gỡ. |
| ThisWorkbook / sheet đã có code | Code công ty nằm trong vùng đánh dấu `'=== tool-vba BEGIN … ===`; mỗi lần cập nhật chỉ thay vùng này. Nếu ngoài vùng đã có thủ tục trùng tên (ví dụ `Workbook_Open`), app dừng và nêu tên thủ tục trùng. |

Sau mỗi lần dán, app ghi vào thuộc tính tài liệu `ToolVba.PackageVersion` (phiên bản code) và `ToolVba.Managed.N` (danh sách module do công ty quản lý).

## Dành cho quản trị viên: xuất bản code VBA mới

Việc xuất bản nằm trên **Admin web của digital-license-api** (chỉ tài khoản quản trị), mục **Tải công cụ → Phát hành tool-vba → Xuất bản code VBA**. App desktop không có phần quản trị.

1. Sửa code trong file Excel mẫu `.xlsm` (module `.bas`, class `.cls`, ThisWorkbook, code của sheet). Module tên bắt đầu bằng `modLicense` không được đóng gói.
2. Tải file `.xlsm` lên (hoặc các file VBA lẻ `.bas`, `.cls`, `.frm` + `.frx`) → **Kiểm tra** để xem danh sách → nhập ghi chú → **Xuất bản cho nhân viên**.
3. Máy chủ tự tách code, cấp phiên bản (`YYYY.M.D.N`) và ký gói. Nhân viên nhận bản mới ở lần bấm Tạo tiếp theo.
4. Lỡ xuất bản bản lỗi: ở danh sách bên dưới bấm **Kích hoạt lại** cạnh bản trước.

UserForm: máy chủ không có Excel nên không dựng lại được phần giao diện của form từ `.xlsm`. Nếu file mẫu có UserForm, trang sẽ yêu cầu tải thêm `.frm` + `.frx` (Alt+F11 → chuột phải vào form → Export File).

Ô **Bản ứng dụng tối thiểu** chỉ điền khi code mới cần một bản app mới hơn; khi đó app cũ chặn nút Tạo và yêu cầu cập nhật.

## Phát hành và cấu hình

Hợp đồng API giữa app và máy chủ nằm trong [docs/tool-release-api-contract.md](docs/tool-release-api-contract.md).

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

Script build bản self-contained win-x64, đóng gói bằng Velopack (có gói delta) và mở thư mục `out\release\upload`. Tải toàn bộ file trong thư mục đó lên Admin web: **Tải công cụ → Phát hành tool-vba → Phát hành ứng dụng**. Nhân viên thấy banner "Có phiên bản ứng dụng mới" và bấm **Cập nhật ngay**.

**Chuyển đổi:** nhân viên đang dùng bản zip cũ cần chạy `ToolVba-win-Setup.exe` một lần; từ đó app tự cập nhật. Chưa có chứng chỉ ký mã thì Windows SmartScreen sẽ cảnh báo ở lần cài đầu.

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

Self-test chạy lần lượt các bước:

1. Cài gói v1 vào `.xlsx`.
2. Thêm một module của nhân viên, rồi cập nhật lên v2.
3. Kiểm tra: module của nhân viên còn nguyên, module đã bỏ bị gỡ, không có tên kiểu `modReport1`, ThisWorkbook có đúng một vùng code.
4. Xác nhận app từ chối khi file có sẵn `Workbook_Open` trùng tên.

Integration runner còn giữ hai chế độ cũ là `--bas` (một module + manifest) và chèn `vbaProject.bin`.

## An toàn dữ liệu

- File nguồn luôn được giữ nguyên.
  - Excel chỉ làm việc trên bản tạm, với macro, sự kiện và cập nhật liên kết đều tắt.
  - App mở lại bản đã lưu để xác minh, so lại hash nguồn, rồi mới chuyển bản tạm thành file kết quả.
- Gói code được kiểm tra theo nhiều lớp:
  - Chữ ký ECDSA P-256 của máy chủ.
  - SHA-256 của file zip và của từng file bên trong.
  - Chặn zip-slip và zip bomb.
- App copy gói sang thư mục riêng và kiểm tra lại ngay trước khi import, nên cache bị sửa trên đĩa không ảnh hưởng tới code được dán.
- Log cục bộ `%LocalAppData%\ToolVba\logs\operations.jsonl` chỉ ghi tên file, hash, phiên bản gói và kết quả; không ghi nội dung workbook hay code VBA.
- Rủi ro đã chấp nhận: máy chủ ký gói, nên **tài khoản ADMIN bị lộ có thể phát tán macro**. Dùng mật khẩu mạnh và giới hạn số người có quyền ADMIN.
- Chữ ký VBA hoặc chữ ký package cũ có thể mất hiệu lực sau khi dán code.
- Excel tiếng Việt đặt tên module workbook là `SổlàmviệcNày`. Excel làm mất tên này khi lưu `.xlsm` (module mất tên, sự kiện workbook không chạy). Vì vậy app đổi tên module này thành `ThisWorkbook` trước khi dán; từ khóa `ThisWorkbook` trong VBA không bị ảnh hưởng. File nào đã bị Excel làm mất tên module thì app dừng và hướng dẫn sửa.
- Không hỗ trợ `.xls`, `.xlsb`, `.xltx` hoặc định dạng khác.

## Thành phần

| Thư mục | Nội dung |
| --- | --- |
| `src/ToolVba.App` | WPF một màn hình cho nhân viên; tự cập nhật bằng Velopack |
| `src/ToolVba.Core/Packages` | Đọc/tạo gói, gộp ThisWorkbook/sheet, kiểm chữ ký |
| `src/ToolVba.Core/Updates` | Tải, kiểm chữ ký và cache gói từ API |
| `src/ToolVba.Core/Excel` | Thao tác VBA project qua Excel COM |
| `tests/ToolVba.Tests` | Unit test, không cần Excel |
| `tests/ToolVba.IntegrationRunner` | Self-test và runner thủ công với Excel thật |
| `scripts/` | Tạo khóa ký, build bản app để tải lên Admin web |
