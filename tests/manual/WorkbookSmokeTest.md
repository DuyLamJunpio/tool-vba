# Kiểm thử thêm module VBA vào workbook Excel

Fixture `tests/fixtures/ToolVbaSmokeTest.bas` chỉ kiểm tra định dạng workbook rồi hiện hộp thoại; nó không ghi vào workbook. Manifest cố định tại `tests/fixtures/ToolVbaSmokeTest.manifest.json` xác nhận SHA-256. Không dùng module này với dữ liệu phát hành.

## Test case TV-REAL-001: thêm module, giữ workbook nguồn và project VBA gốc

**Điều kiện:** máy Windows thử nghiệm có Excel; quản trị viên đã cấu hình quyền Trust access theo chính sách; output là đường dẫn mới; macro không được chạy trong lúc đóng gói.

1. Ghi SHA-256 của `HR_Quan_ly_nhan_su_v2.0.xlsm` trước thử nghiệm.
2. Chạy tool, chọn chế độ “Thêm module .bas vào project hiện có”. Chọn workbook HR, file `.bas`, manifest, và output mới, ví dụ `HR_Quan_ly_nhan_su_v2.0_ToolVbaTest.xlsm`.
3. Xác nhận tóm tắt. Tool phải làm việc trên file tạm, không tự bật Trust Center hoặc chạy macro.
4. Kỳ vọng thành công chỉ khi SHA-256 module khớp manifest, tên module chưa tồn tại, component cũ vẫn còn, module mới đã import, package `.xlsm` hợp lệ và hash nguồn không đổi.
5. So hash nguồn sau chạy. Nó phải trùng giá trị ban đầu. Trong Excel, mở output trên máy thử nghiệm và xác nhận module mới cùng các component cũ hiện diện. Không chạy macro nếu chưa được chính sách tổ chức cho phép.

## Test case TV-REAL-002: tên module trùng

1. Dùng một workbook `.xlsm` test có component mang tên `ToolVbaSmokeTest`.
2. Chạy chế độ thêm module với fixture và chọn output mới.
3. Kỳ vọng tool từ chối trước khi import; không component cũ nào bị thay; input không đổi; output không được tạo; file tạm được dọn.

## Test case TV-REAL-005: thêm module trực tiếp vào `.xlsx`

1. Chọn `Tạo hồ sơ.xlsx`, module `.bas`, manifest và một output mới `.xlsm` trong chế độ “Thêm module .bas”.
2. Kỳ vọng Excel chỉ mở bản tạm với macro bị vô hiệu, import module và lưu bản tạm thành `.xlsm`.
3. Xác nhận output có VBA Project và module mới; dữ liệu trang tính vẫn mở được; SHA-256 của `.xlsx` nguồn không đổi.

## Test case TV-REAL-003: Excel không cho truy cập VBA Project

1. Trên máy test chưa được cấp quyền VBA Project Object Model theo chính sách, chọn workbook `.xlsm` và module fixture.
2. Kỳ vọng tool báo cần quản trị viên cấu hình Excel, không tự sửa Registry/Trust Center, không chạy macro, giữ nguyên input/output.

## Test case TV-REAL-004: project binary mode không ghép module

1. Dùng một workbook `.xlsm` disposable có project VBA và một binary test.
2. Không chọn xác nhận thay project. Kỳ vọng từ chối và input/output giữ nguyên.
3. Chỉ xác nhận thay toàn bộ khi binary đã chứa mọi component cần giữ; không dùng binary chỉ có module test cho workbook có VBA gốc.

## Giới hạn kiểm thử ở môi trường hiện tại

Test tự động chạy được với fixture tổng hợp. Trên phiên máy hiện tại, Excel COM chưa khởi động được nên vẫn cần chạy các test thủ công trên máy đóng gói có phiên Excel hoạt động và quyền VBA Project Object Model đã được quản trị viên cấu hình. Không dùng file nguồn làm output.
