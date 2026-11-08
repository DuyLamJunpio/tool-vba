# Kiểm thử tay app đóng gói VBA một màn hình

Dùng máy Windows có Excel desktop và bản app đã cấu hình `release.props`. Chỉ dùng file tổng hợp hoặc bản sao. Ghi SHA-256 file nguồn trước mỗi ca; sau ca phải không đổi. Không chạy macro của file nguồn.

Chạy self-test tự động trước, kỳ vọng `PASS: Self-test hoàn tất.`:

```powershell
dotnet run --project tests\ToolVba.IntegrationRunner\ToolVba.IntegrationRunner.csproj -- --selftest D:\Temp\tool-vba-selftest
```

| ID | Thao tác | Kỳ vọng |
| --- | --- | --- |
| TV-01 | Thả một file `.xlsx`, bấm **Tạo file có VBA** | Tạo `<tên>_vba.xlsm` cạnh file gốc, có đủ component của gói. Báo "Xong!" kèm phiên bản gói. |
| TV-02 | Tạo lại cùng file | Kết quả mới là `<tên>_vba (2).xlsm`; file trước không bị ghi đè. |
| TV-03 | Bấm **Đổi nơi lưu**, chọn thư mục khác, Tạo | Kết quả nằm ở thư mục đã chọn. |
| TV-04 | Thả nhiều file, có một file `.xls` | Bỏ qua `.xls` và báo số file bỏ qua. Mỗi file hợp lệ có một kết quả; file lỗi nêu lý do trên dòng của nó. |
| TV-05 | `.xlsm` có module nhân viên tự viết và module trùng tên với gói | Module trùng được thay bằng bản mới. Module nhân viên giữ nguyên. Module bản gói cũ nay đã bỏ bị gỡ. |
| TV-06 | Mở kết quả, Alt+F11 xem ThisWorkbook | Đúng một vùng `'=== tool-vba BEGIN … ===`; code ngoài vùng giữ nguyên. Thuộc tính `ToolVba.PackageVersion` bằng phiên bản trên chip. |
| TV-07 | File có sẵn `Workbook_Open` ngoài vùng, gói cũng có `Workbook_Open` | App dừng, nêu tên thủ tục trùng, không tạo file. |
| TV-08 | Excel tiếng Việt (module `SổlàmviệcNày`) | Module được đổi tên `ThisWorkbook`; sự kiện workbook chạy sau khi lưu. |
| TV-09 | Tắt "Trust access to the VBA project object model", mở app | Báo "Excel chưa cho phép…". **Bật giúp tôi** → hộp xác nhận → "Đã bật". Đóng/mở app: vẫn bật. |
| TV-10 | Đặt policy `AccessVBOM` (VM) rồi bấm **Bật giúp tôi** | Chỉ báo liên hệ IT, không ghi Registry. |
| TV-11 | Đã có gói trong cache, ngắt mạng, mở app và Tạo | Báo "Không kết nối được máy chủ. Đang dùng code VBA đã lưu v…"; Tạo vẫn chạy. |
| TV-12 | Ngắt mạng, sửa một byte file zip trong `%LocalAppData%\ToolVba\vba\<phiên bản>\` | App bỏ gói đó (sai SHA-256/chữ ký), báo cần kết nối mạng để tải code VBA, không dán code. Có mạng lại thì tải gói sạch. |
| TV-13 | Phát hành bản app mới hơn trên Admin web, mở app cũ | Banner "Có phiên bản ứng dụng mới". **Cập nhật ngay** → app tự cài, mở lại đúng phiên bản. |
| TV-14 | Xuất bản gói có "Bản ứng dụng tối thiểu" cao hơn app | Nút Tạo bị chặn, yêu cầu cập nhật. |
| TV-15 | Admin web: **Kích hoạt lại** bản gói trước, rồi Tạo | Chip và file kết quả dùng bản gói vừa kích hoạt lại. |

Ghi kết quả: ID, phiên bản app, phiên bản gói, PASS/FAIL, lý do. Không ghi nội dung workbook hay code VBA.
