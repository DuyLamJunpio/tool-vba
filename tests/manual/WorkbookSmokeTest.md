# Kiểm thử tay app đóng gói bản quyền một màn hình

Dùng máy Windows có Excel desktop và bản app đã cấu hình `release.props`. Chỉ dùng file tổng hợp hoặc bản sao. Ghi SHA-256 file nguồn trước mỗi ca; sau ca phải không đổi. Không chạy macro của file nguồn.

Mỗi ca "Tạo" cần một mã đóng gói: dán mã lấy trên Admin web (**Sao chép mã đóng gói** ở sản phẩm), hoặc nhập tay `productKey` và phiên bản `x.y.z`. Ví dụ trong bảng dùng `test-01` và `1.0.0`.

Chạy self-test tự động trước, kỳ vọng `PASS: Self-test hoàn tất.` (thư mục phải trống):

```powershell
dotnet run --project tests\ToolVba.IntegrationRunner\ToolVba.IntegrationRunner.csproj -- --selftest D:\Temp\tool-vba-selftest
```

Các ca của cổng bản quyền khi mở và đóng file nằm ở [LicenseGateSmokeTest.md](LicenseGateSmokeTest.md); kích hoạt máy nằm ở [DeviceCodeActivationSmokeTest.md](DeviceCodeActivationSmokeTest.md).

## Đóng gói

| ID | Thao tác | Kỳ vọng |
| --- | --- | --- |
| TV-01 | Thả một file `.xlsx` có hai sheet, dán mã đóng gói `test-01` `1.0.0`, bấm **Tạo file có bản quyền** | Tạo `test-01_1.0.0_ban-quyen.xlsm` cạnh file gốc. Sheet đầu là "Huong dan" và hiện; sheet kia `xlSheetVeryHidden`. Module `modLicenseApiV1` có `Cfg_ProductKey = "test-01"`, `Cfg_ProductVersion = "1.0.0"`, `Cfg_ProtectedSheetNames` chứa sheet kia. Báo "Xong!" kèm phiên bản gói. |
| TV-02 | Tạo lại cùng file và cùng mã | Kết quả mới là `test-01_1.0.0_ban-quyen (2).xlsm`; file trước không bị ghi đè. |
| TV-03 | Bấm **Đổi nơi lưu**, chọn thư mục khác, Tạo | Kết quả nằm ở thư mục đã chọn. **Lưu cạnh file gốc** đưa về mặc định. |
| TV-04 | Thả hai file cùng lúc; thả một file `.xls` | Hai file: app lấy file đầu và báo "Mỗi lần chỉ đóng gói một file…". `.xls`: báo "Chỉ nhận file Excel .xlsx hoặc .xlsm." |
| TV-05 | `.xlsm` có module nhân viên tự viết và module trùng tên với gói | Module trùng được thay bằng bản mới. Module nhân viên giữ nguyên. Module bản gói cũ nay đã bỏ bị gỡ. |
| TV-06 | Mở kết quả, Alt+F11 xem ThisWorkbook | Đúng một vùng `'=== tool-vba BEGIN … ===` cho từng loại; code ngoài vùng giữ nguyên. Thuộc tính `ToolVba.PackageVersion` bằng phiên bản trên chip. |
| TV-07 | File có sẵn `Workbook_Open` ngoài vùng, gói cũng có `Workbook_Open` | App dừng, nêu tên thủ tục trùng, không tạo file. |
| TV-08 | Excel tiếng Việt (module `SổlàmviệcNày`) | Module được đổi tên `ThisWorkbook`; sự kiện workbook chạy sau khi lưu. |
| TV-09 | Tắt "Trust access to the VBA project object model", mở app | Báo "Excel chưa cho phép…". **Bật giúp tôi** → hộp xác nhận → "Đã bật". Đóng/mở app: vẫn bật. |
| TV-10 | Đặt policy `AccessVBOM` (VM) rồi bấm **Bật giúp tôi** | Chỉ báo liên hệ IT, không ghi Registry. |
| TV-11 | Đã có gói trong cache, ngắt mạng, mở app và Tạo | Báo "Không kết nối được máy chủ. Đang dùng code VBA đã lưu v…"; Tạo vẫn chạy. |
| TV-12 | Ngắt mạng, sửa một byte file zip trong `%LocalAppData%\ToolVba\vba\<phiên bản>\` | App bỏ gói đó (sai SHA-256/chữ ký), báo cần kết nối mạng để tải code VBA, không dán code. Có mạng lại thì tải gói sạch. |
| TV-13 | Phát hành bản app mới hơn trên Admin web, mở app cũ | Banner "Có phiên bản ứng dụng mới". **Cập nhật ngay** → app tự cài, mở lại đúng phiên bản. |
| TV-14 | Xuất bản gói có "Bản ứng dụng tối thiểu" cao hơn app | Nút Tạo bị chặn, yêu cầu cập nhật. |
| TV-15 | Admin web: **Kích hoạt lại** bản gói trước, rồi Tạo | Chip và file kết quả dùng bản gói vừa kích hoạt lại. |

## Mã đóng gói và sản phẩm

| ID | Thao tác | Kỳ vọng |
| --- | --- | --- |
| TV-16 | Dán chuỗi không phải mã đóng gói (văn bản tự do, JSON thiếu `version`, JSON có trường thừa). Làm thêm: dán mã đúng, rồi dán đè một mã sai | Báo "Mã đóng gói không đúng…" kèm cách lấy mã trên Admin web. Nút Tạo bị chặn, kể cả khi hai ô còn giá trị của mã đúng trước đó. Không tạo file. |
| TV-17 | Nhập tay phiên bản `1.0` hoặc `v1.0.0`; nhập tay `productKey` có ký tự lạ (khoảng trắng giữa chừng, `_`, dấu) ; nhập `Test-01` | Hai ca đầu báo lỗi định dạng, không tạo file. `Test-01` được chấp nhận và đổi thành `test-01` (cửa hàng cũng đổi chữ hoa thành chữ thường). |
| TV-18 | Dán mã đúng, rồi sửa tay ô `productKey` hoặc phiên bản | App dùng giá trị đang hiển thị trong hai ô; tên file kết quả và `Cfg_ProductKey`/`Cfg_ProductVersion` khớp giá trị đó. |

## Từ chối an toàn (không tạo file, file nguồn không đổi)

| ID | File nguồn | Kỳ vọng |
| --- | --- | --- |
| TV-19 | Có module riêng với `Public Sub Auto_Open()` | Nêu tên module và lý do (module bản quyền dùng `Auto_Open`). |
| TV-20 | `Workbook_BeforeClose` có `Exit Sub`, `Exit Function` hoặc `End` | Nêu tên thủ tục và lý do (lời gọi cuối thủ tục có thể không chạy). |
| TV-21 | Workbook đang khóa cấu trúc (Review › Protect Workbook) | Hướng dẫn bỏ khóa cấu trúc trong file gốc. |
| TV-22 | Có chart sheet | Nêu tên sheet; module chỉ ẩn được worksheet. |
| TV-23 | Đã có sheet tên "Huong dan" do người dùng tạo | Hướng dẫn đổi tên sheet đó. |
| TV-24 | Chỉ có sheet "Huong dan", không còn sheet nào khác | Báo cần ít nhất một sheet ngoài sheet hướng dẫn. |
| TV-25 | Có module ngoài gói định nghĩa `Cfg_ProductKey` hoặc `License_*` | Từ chối vì có bản cũ của module bản quyền. |

## Ghép sự kiện

| ID | File nguồn | Kỳ vọng |
| --- | --- | --- |
| TV-26 | ThisWorkbook đã có `Workbook_BeforeSave` và `Workbook_AfterSave` | Lời gọi bản quyền nằm **đầu** từng thủ tục giữa hai dòng `'=== tool-vba BEGIN/END license hook … ===`; code của sản phẩm giữ nguyên. |
| TV-27 | ThisWorkbook đã có `Workbook_BeforeClose` không có `Exit` | Lời gọi nằm **cuối** thủ tục, dạng `If Not Cancel And (ThisWorkbook.Saved) Then …`. Mở kết quả, trả lời Không ở hộp của sản phẩm: sheet và quyền không đổi (xem thêm LicenseGateSmokeTest). |
| TV-28 | ThisWorkbook chưa có ba thủ tục trên | Thủ tục của gói nằm trong vùng đánh dấu; mở kết quả không báo lỗi biên dịch. |

Ghi kết quả: ID, phiên bản app, phiên bản gói, PASS/FAIL, lý do. Không ghi nội dung workbook hay code VBA.
