# Kiểm thử tay cổng bản quyền: đóng file và hai `Auto_Open`

Dùng máy Windows có Excel desktop. Chỉ dùng các file thử do script dựng, không dùng file thật của khách. Các file không gọi mạng và không ghi Registry.

Dựng file thử (Excel chạy ẩn, mất vài giây):

```powershell
.\tests\manual\build-manual-test-files.ps1
```

Kết quả nằm ở `out\manual-test\` (thư mục `out\` không commit):

| File | Dùng cho |
| --- | --- |
| `close-test.xlsm` | Ca đóng file CL-1 đến CL-3. Module bản quyền thật, sản phẩm giả `close-test` 1.0.0, sheet "Du lieu" đang **hiện**. `Auto_Open` được đổi tên để mở file không gọi mạng và không tự đóng. `BeforeClose` của sản phẩm ghi "BeforeClose da chay" vào `Huong dan!B1`, rồi đến dòng bản quyền đặt cuối. |
| `auto-open-license-first.xlsm`, `auto-open-other-first.xlsm` | Ca AO-3. Hai module giả đều có `Auto_Open`, chỉ khác thứ tự tạo module. |

Khi mở mỗi file, bấm **Enable Content**. Nếu Excel chặn vì file từ thư mục lạ, chuột phải file → Properties → Unblock.

## Ca đóng file

Cổng bản quyền gọi `License_OnClose` từ `Workbook_BeforeClose` (và từ `Auto_Close` khi đóng bằng giao diện). `License_OnClose` hủy quyền trong phiên và ẩn sheet bảo vệ. Excel chạy `BeforeClose` **trước** hộp "Save changes?", nên bấm Cancel ở hộp đó có thể xảy ra khi `License_OnClose` đã chạy.

| ID | Thao tác | Điều cần ghi lại |
| --- | --- | --- |
| CL-1 | Mở `close-test.xlsm`. Kiểm hai tab "Huong dan" và "Du lieu" đang hiện. Đóng workbook (nút X của cửa sổ workbook, hoặc File → Close). Hộp "Save changes?" hiện ra: bấm **Cancel**. | (1) File còn mở không. (2) Tab "Du lieu" còn hiện hay đã mất. (3) `Huong dan!B1` có chữ "BeforeClose da chay" không. |
| CL-2 | Đối chứng không hủy: từ trạng thái sau CL-1, đóng lại, ở hộp "Save changes?" bấm **Don't Save**. Mở lại `close-test.xlsm`. | Tab "Du lieu" hiện lại không; `Huong dan!B1` trống không. Kỳ vọng: file về đúng bản đã lưu, không lộ gì thêm. |
| CL-3 | Mở lại file thử. Đóng, ở hộp "Save changes?" bấm **Save**. Mở lại file. | Tab "Du lieu" có còn thấy không. Kỳ vọng thiết kế: bản đã lưu chỉ còn sheet "Huong dan" (`BeforeSave` ẩn sheet trước khi lưu). |

Sau CL-3 file thử đã bị ghi đè. Chạy lại script dựng file nếu cần thử lại.

### Kết quả đã có

| Ca | Cách thử | Kết quả (08/11/2026) |
| --- | --- | --- |
| Sản phẩm hỏi "đóng file?" và trả lời **Không** (`BeforeClose` đặt `Cancel = True`), đóng bằng lệnh giao diện `FileClose` trong Excel ẩn, macro bật | Script (Excel ẩn) | **ĐẠT.** File vẫn mở, sheet bảo vệ vẫn hiện, `License_OnClose` không chạy (cả qua dòng bản quyền cuối `BeforeClose` lẫn qua `Auto_Close`). Lý do: dòng bản quyền là `If Not Cancel Then …`. |
| Đối chứng: `BeforeClose` không hủy, đóng bằng lệnh giao diện `FileClose`, workbook không có thay đổi chưa lưu | Script (Excel ẩn) | **ĐẠT.** File đóng; `BeforeClose` chạy với `Cancel = False`; ngay sau dòng bản quyền sheet "Du lieu" ở trạng thái `xlSheetVeryHidden`, tức `License_OnClose` đã chạy. |
| CL-1 (Cancel ở hộp "Save changes?"), dòng bản quyền `If Not Cancel Then …` | Chủ dự án làm tay, 08/11/2026, Excel 16.0.20527 x64 | **KHÔNG ĐẠT.** File còn mở, "Du lieu" đã bị ẩn, B1 có chữ. Đã sửa: dòng thành `If Not Cancel And (ThisWorkbook.Saved) Then …` (và `If ThisWorkbook.Saved Then …` khi sản phẩm chưa có `BeforeClose`). |
| CL-1 lần 2 (dòng hook đã sửa) | Chủ dự án làm tay, 09/11/2026 | **KHÔNG ĐẠT.** B1 có chữ, nhưng "Du lieu" bị ẩn ngay khi bấm đóng, trước khi hộp hiện; Cancel chỉ đóng hộp, file vẫn mở. Nguyên nhân: `Auto_Close` của module chạy trước hộp và gọi `License_OnClose` không điều kiện. Đã sửa: `If ThisWorkbook.Saved Then License_OnClose`. |
| CL-1 lần 3 (module đã sửa, file thử dựng lại) | Script điều khiển một Excel **hiện** (không phải Excel ẩn), 09/11/2026, cùng máy: gửi lệnh đóng tới cửa sổ workbook như bấm nút X, chụp màn hình khi hộp hỏi hiện, bấm nút trong hộp qua UI Automation. Macro bật bằng `AutomationSecurity` thay cho nút Enable Content. | **ĐẠT.** Khi hộp "Bạn có muốn lưu thay đổi…" hiện, tab "Du lieu" vẫn còn, B1 có chữ. Bấm **Hủy bỏ**: file vẫn mở, "Du lieu" hiện (`xlSheetVisible`), B1 có chữ. |
| CL-2 (Không Lưu ở hộp đó, nối tiếp CL-1 lần 3) | Cùng script | **ĐẠT.** File đóng; mở lại thì "Du lieu" hiện như bản đã lưu, B1 trống. |
| CL-3 lần 2 (Lưu ở hộp đó, bản sao mới) | Cùng script | **ĐẠT.** Mở lại thì "Du lieu" ở `xlSheetVeryHidden` (giá trị 2), chỉ "Huong dan" hiện. |
| CL-3 (Save ở hộp đó) | Chủ dự án làm tay | **ĐẠT.** File đã lưu chỉ còn "Huong dan" hiện. |

Ba ca script ở trên dùng dòng `If Not Cancel Then …` cũ. Ca `Cancel = True` không đổi kết quả với dòng mới.

Đóng bằng `Workbook.Close` từ mã (COM/VBA) không chạy `Auto_Close`; `Workbook_BeforeClose` vẫn chạy.

## Ca hai `Auto_Open`

App từ chối khi file của sản phẩm đã có `Auto_Open` riêng (unit test `Stops_when_the_product_has_its_own_Auto_Open` và self-test với Excel thật). Lý do kiểm với Excel thật dưới đây. Dùng hai module giả có `Auto_Open` ghi `license` hoặc `other` vào cột A (module bản quyền thật gọi mạng và hiện hộp thoại nên không dùng được trong Excel ẩn).

| ID | Thao tác | Kết quả |
| --- | --- | --- |
| AO-1 | Một module có `Auto_Open`, chạy `RunAutoMacros(xlAutoOpen)` (Excel ẩn, macro bật) | **Chạy**: cột A có `license`. Gọi lần hai thì chạy lại. |
| AO-2 | Hai module (module bản quyền và một module khác) đều có `Public Sub Auto_Open`, chạy `RunAutoMacros(xlAutoOpen)`, cả hai thứ tự tạo module | **Không module nào chạy** và **không có lỗi**. `Application.Run "'file'!Auto_Open"` báo "Macro không có trong sổ làm việc này", còn gọi kèm tên module (`modLicenseApiV1.Auto_Open`, `modOther.Auto_Open`) thì từng cái chạy bình thường. |
| AO-3 | Mở bằng giao diện (double-click) `auto-open-license-first.xlsm` và `auto-open-other-first.xlsm`, bấm Enable Content, xem cột A. A1 có sẵn tiêu đề "log"; mỗi `Auto_Open` chạy sẽ thêm một dòng `license` hoặc `other` bên dưới. | **Không module nào chạy** (chủ dự án, 08/11/2026): cột A chỉ có "log", cả hai file. Có một hộp thông báo hiện ra, chưa ghi nội dung. |

Hệ quả: nếu file có hai `Auto_Open`, cổng bản quyền **im lặng không chạy** khi mở file, không có dấu hiệu nào. Đó là lý do app từ chối thay vì ghép chung hai thủ tục.

Ghi kết quả: ID, phiên bản Excel, PASS/FAIL, quan sát. Không ghi nội dung workbook thật.
