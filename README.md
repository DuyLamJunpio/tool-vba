# Công cụ đóng gói VBA cho Excel

Ứng dụng desktop WPF chạy cục bộ trên Windows. Có hai chế độ tách biệt:

- **Thêm module `.bas` vào `.xlsx` hoặc `.xlsm`:** Excel mở bản tạm, import module đã duyệt và lưu output `.xlsm` riêng. Với `.xlsm` có VBA, các component cũ phải còn hiện diện sau khi lưu.
- **Chèn toàn bộ `vbaProject.bin`:** dùng Open XML SDK để thêm/thay trọn project binary; chế độ này không ghép module.

Không chạy macro trong lúc đóng gói. Không gửi workbook hoặc mã VBA ra ngoài máy.

## Yêu cầu

- Windows 10/11 và .NET 10 SDK.
- NuGet được truy cập ở lần restore đầu để tải `DocumentFormat.OpenXml`, xUnit và Microsoft.NET.Test.Sdk.
- Chế độ `.bas` cần Microsoft Excel desktop đã cài. Quản trị viên phải tự bật “Trust access to the VBA project object model” trên máy đóng gói theo chính sách tổ chức. Tool không bật Trust Center, sửa Registry hay thay đổi thiết lập bảo mật.

Chế độ `.bas` dùng Excel COM Automation cục bộ, không cài package Office Interop. Nó chỉ chạy trên máy Windows có Excel; không dùng cho web server, Windows Service hay xử lý file khách hàng tải lên.

## Build, test và chạy

Mở PowerShell tại thư mục project:

```powershell
cd D:\Code\tool-vba
dotnet restore ToolVba.sln
dotnet test ToolVba.sln
dotnet build src\ToolVba.App\ToolVba.App.csproj -c Release
dotnet run --project src\ToolVba.App\ToolVba.App.csproj
```

Sau khi build, đóng mọi cửa sổ `tool-vba` cũ trước khi mở bản mới. Bản mới nằm tại `artifacts\tool-vba-vba-source-fix-5\tool-vba.exe`; tiêu đề cửa sổ có chuỗi `VBA source fix 5` để phân biệt với các bản cũ.

Để phát hành bản tự chứa cho Windows x64:

```powershell
dotnet publish src\ToolVba.App\ToolVba.App.csproj -c Release -r win-x64 --self-contained true
```

Integration runner cho chế độ binary:

```powershell
dotnet run --project tests\ToolVba.IntegrationRunner\ToolVba.IntegrationRunner.csproj -- "D:\path\source.xlsx" "D:\path\vbaProject.bin" "D:\path\result.xlsm"
```

Integration runner cho chế độ `.bas` (cần Excel và quyền VBA Project Object Model đã được quản trị viên cấu hình):

```powershell
dotnet run --project tests\ToolVba.IntegrationRunner\ToolVba.IntegrationRunner.csproj -- --bas "D:\path\source.xlsx" "D:\path\module.bas" "D:\path\module.manifest.json" "D:\path\result.xlsm"
```

Runner không ghi đè output; nó so hash nguồn trước/sau và giữ output để kiểm tra thủ công.

## Cách dùng

### Thêm module `.bas`

1. Chọn `.xlsx` hoặc `.xlsm` hợp lệ, file `.bas`, manifest JSON có `moduleName` và SHA-256, rồi chọn output `.xlsm` mới. Không cần file `vbaProject.bin`.
2. Xác nhận tóm tắt. Nếu output đã tồn tại, xác nhận ghi đè riêng.
3. Đóng workbook nguồn trong các cửa sổ Excel đang mở. Tool xác minh hash module, tên module không trùng, tạo bản tạm và mở bản tạm bằng Excel với macro bị vô hiệu, event tắt, liên kết ngoài không cập nhật và calculation ở chế độ thủ công.
4. Tool import module. Với nguồn `.xlsx`, Excel lưu bản tạm thành `.xlsm`; với nguồn `.xlsm`, Excel lưu bản tạm hiện có. Tool mở lại file `.xlsm` để xác nhận component cũ và module mới còn hiện diện, kiểm tra package Open XML và hash nguồn, rồi mới chuyển thành output.

Nếu Excel chặn truy cập project, tool dừng và báo cần quản trị viên cấu hình quyền phù hợp. Tool không tự bật quyền này. Nếu import thất bại, input và output có sẵn được giữ nguyên; file tạm được dọn best-effort.

Mẫu giấy phép một file ở `license-sample/api-v1-one-file/` đã có cổng khởi động: chưa có quyền thì thông báo rồi đóng **workbook đó** sau OK/X, không thoát Excel hoặc workbook khác. Mẫu còn thiếu API mã thiết bị, cấu hình khóa công khai và bộ xác minh Ed25519, nên chưa dùng để phát hành cho khách. Hợp đồng API cần bổ sung và ca kiểm thử nằm trong `license-sample/api-v1-one-file/DEVICE-CODE-CONTRACT.md` và `tests/manual/DeviceCodeActivationSmokeTest.md`.

Nếu Excel Automation báo COM `0x800A03EC`, công cụ nêu chính xác bước cấu hình Excel gặp lỗi. Công cụ tạo một workbook trống tạm để đặt tính toán thủ công trước khi mở file nguồn, rồi đóng workbook trống mà không lưu. Macro và sự kiện vẫn bị tắt trước khi workbook nguồn được mở. Với nguồn `.xlsx` chưa có VBA project, khi kiểm tra sau lưu công cụ yêu cầu module mới tồn tại; với nguồn `.xlsm` đã có VBA project, công cụ còn kiểm tra tất cả tên component VBA gốc. Số component có thể tăng do Excel tạo thêm component, nên không còn yêu cầu số lượng đúng bằng số cũ cộng một.

### Chèn toàn bộ `vbaProject.bin`

Chọn `.xlsx` hoặc `.xlsm`, binary `.bin` đã được duyệt và output `.xlsm`. Nếu input đã có project, chế độ này sẽ yêu cầu xác nhận thay **toàn bộ** project. Chỉ tiếp tục nếu binary mới đã chứa mọi mã cần giữ.

## Kiểm thử thủ công bằng Excel

Thực hiện trên bản sao hoặc output riêng, theo chính sách bảo mật của tổ chức:

1. Chạy `tests/manual/WorkbookSmokeTest.md` với workbook `.xlsm` có VBA và module test mẫu.
2. Mở output trong Excel trên máy thử nghiệm. Trong lúc đóng gói, tool không bật hay chạy macro.
3. Xác nhận module mới và component VBA cũ còn trong Project Explorer. Chỉ chạy `ToolVbaSmokeTest` nếu được chính sách tổ chức cho phép.
4. Đối chiếu hash input trước/sau. Không xem việc module xuất hiện là bằng chứng macro an toàn, hợp lệ nghiệp vụ hay có chữ ký hợp lệ.

## Giới hạn và bảo vệ dữ liệu

- Cả hai chế độ nhận nguồn `.xlsx` hoặc `.xlsm`. Output luôn là `.xlsm`; chế độ `.bas` cần Excel desktop và quyền truy cập VBA Project Object Model.
- Chế độ thêm module chỉ nhận một module chuẩn `.bas`; không hỗ trợ `.cls`, `.frm`, `.frx`, ActiveX hoặc tự thêm references. Tên module đã tồn tại thì từ chối, không thay module cũ.
- Workbook nguồn luôn được giữ nguyên. Output chỉ thay sau khi kiểm tra đạt; không ghi đè âm thầm.
- Chế độ Excel có thể chuẩn hóa metadata khi lưu. Tool xác nhận component VBA cũ còn tên và không bị xóa, nhưng không tuyên bố project cũ giữ nguyên từng byte; chữ ký cũ có thể mất hiệu lực.
- Không ghi nội dung workbook/VBA hay stack trace vào log hoặc giao diện. Tool chỉ hash file module và đọc tên các component VBA để xác nhận giữ lại.
- Thay đổi project có thể làm chữ ký VBA/package cũ mất hiệu lực. Tool cảnh báo artifact chữ ký nhưng không xác minh hay tự ký.
- Không hỗ trợ `.xls`, `.xlsb`, `.xltx`, Word, Access, PowerPoint hoặc định dạng khác.
- Workbook `Tạo hồ sơ.xlsx` và `HR_Quan_ly_nhan_su_v2.0.xlsm` do người dùng cung cấp được giữ tại chỗ và loại trừ bằng `.gitignore`; không đưa workbook khách hàng hay binary VBA thật vào repository.

## Thành phần

- `src/ToolVba.App`: WPF UI/MVVM với hai chế độ.
- `src/ToolVba.Core`: xử lý Open XML, xác minh manifest, Excel Automation và ghi file tạm an toàn.
- `tests/ToolVba.Tests`: unit test fixture tổng hợp, chạy không cần mở Excel.
- `tests/ToolVba.IntegrationRunner`: runner thủ công cho workbook thật; `.bas` runner cần Excel đã được cấu hình theo chính sách.
# tool-vba
