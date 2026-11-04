# Cách VBA hạn chế chia sẻ file Excel

Tài liệu này giải thích bằng ngôn ngữ đơn giản luồng cấp phép cho file Excel có VBA. Đây là mô tả thiết kế; một số bước phát hành/ký macro còn phụ thuộc hạ tầng và việc cấu hình của chủ hệ thống.

## 1. Ý tưởng cốt lõi

Hãy hình dung file Excel như một căn phòng có khóa:

- File chứa bảng tính, công thức và giao diện làm việc.
- VBA là chốt kiểm tra ở cửa: nó xác nhận máy đang mở file có giấy phép hay không.
- Máy chủ là nơi quyết định giấy phép còn hiệu lực, thiết bị có được phép dùng và gói có còn suất thiết bị hay không.

VBA **không tự quyết định** rằng khách hàng có giấy phép chỉ vì một ô trong file ghi “đã kích hoạt”. Nó kiểm tra token đã ký; khi đến hạn xác minh lại thì hỏi máy chủ. Máy chủ kiểm tra dữ liệu giấy phép trong hệ thống rồi trả kết quả có chữ ký số. Theo yêu cầu mới cho bản mẫu một file, nếu khách không kích hoạt hoặc bị từ chối, Excel hiện hướng dẫn rồi đóng riêng workbook đó sau OK/X. Tính năng cần giấy phép chỉ hoạt động sau khi xác minh thành công.

## 2. Sơ đồ hoạt động

```text
Khách mở file Excel
        │
        ▼
VBA kiểm tra khi workbook mở
        │
        ├─ Tạo dấu vân tay của máy dưới dạng hash
        ├─ Đọc token kích hoạt đã lưu
        └─ Kiểm tra chữ ký, thời hạn và máy được gắn với token
                   │
           ┌───────┴────────┐
           │                │
      Token hợp lệ     Cần xác minh / lỗi
           │                │
           ▼                ▼
  Cho dùng tính năng   Gửi yêu cầu qua HTTPS
  đã được cấp quyền          │
                             ▼
                  Máy chủ kiểm tra giấy phép
                    │                     │
                 Hợp lệ               Bị từ chối
                    │                     │
                    ▼                     ▼
             Cấp token mới       Hiện hướng dẫn kích hoạt;
                                 OK/X đóng workbook này
```

## 3. Các bước chi tiết

### Bước 1: Khách mở workbook

Bản mẫu một file dùng `Auto_Open` để kiểm tra giấy phép khi khách mở workbook trực tiếp và đã bật macro. Nếu chưa có quyền, Excel cho khách kích hoạt bằng mã trên trang web; khi khách hủy hoặc bị từ chối, OK/X trên hướng dẫn đóng riêng workbook này mà không lưu. Excel và các workbook khác vẫn mở. Nếu macro bị tắt, VBA không thể tự đóng file.

### Bước 2: Nhận diện thiết bị

VBA đọc một số định danh của Windows/phần cứng, chuẩn hóa chúng và tạo giá trị SHA-256 gọi là `fp_hash` (dấu vân tay dạng hash). Thiết kế liệt kê các thành phần như MachineGuid, số sê-ri ổ đĩa, CPU ID, Windows Product ID và user SID.

Máy chủ nhận hash, không nhận các định danh thô. Hash được dùng để nhận biết lần kích hoạt sau có phải cùng thiết bị hay không. Cài lại Windows hoặc thay đổi phần cứng có thể làm dấu vân tay đổi; khi đó khách hàng có thể cần vào cổng tài khoản để tắt thiết bị cũ và kích hoạt lại. Chính sách dung sai khi phần cứng thay đổi vẫn cần được xác định.

### Bước 3: Đọc và kiểm tra token

Token kích hoạt được lưu cùng thông tin thiết bị, theo thiết kế ở sheet ẩn rất sâu tên `_lic`; có phương án dự phòng là lưu trong vùng Registry của người dùng Windows. Khi mở file, VBA kiểm tra chữ ký và xác nhận token được cấp cho thiết bị này.

Token được máy chủ ký bằng khóa riêng Ed25519. VBA nhúng **khóa công khai** để kiểm tra token có bị sửa hay không. Khóa riêng không được nhúng vào file VBA. Sửa token sẽ làm chữ ký không hợp lệ; token sao chép sang máy khác sẽ không khớp dấu vân tay thiết bị.

Token có thời hạn cứng. Theo thiết kế kiến trúc, thời hạn là giá trị nhỏ hơn giữa hạn dùng của giấy phép và thời điểm hiện tại cộng tối đa 30 ngày. Khoảng cách cần gọi xác minh lại (`rvi`) và thời gian ân hạn khi mất mạng (`ogh`) lấy từ cấu hình gói/giấy phép.

### Bước 4: Xác minh lại với máy chủ khi cần

Khi token đã đến kỳ xác minh lại, VBA gọi API qua HTTPS và gửi token cùng `fp_hash`. Máy chủ tự tra giấy phép, thiết bị, sản phẩm và phiên bản trong cơ sở dữ liệu. Máy chủ không tin các thông tin quyền hạn do file khách hàng tự khai.

Máy chủ kiểm tra những điểm chính:

- Giấy phép có đang hoạt động, chưa hết hạn, chưa bị tạm ngưng hoặc thu hồi không.
- Thiết bị này có thuộc giấy phép đó không.
- Số thiết bị đang hoạt động có vượt giới hạn gói không.
- Token có đúng sản phẩm/phiên bản và chưa bị thu hồi không.

Giới hạn thiết bị được đọc từ cấu hình gói hoặc quyền ghi đè của giấy phép. Khi khách kích hoạt lại cùng một thiết bị đang hoạt động, việc cấp lại token không tính thành một thiết bị mới.

### Bước 5: Cho dùng tính năng hoặc hiện hướng dẫn

Nếu hợp lệ, máy chủ trả token mới có chữ ký số và VBA cho phép dùng tính năng được cấp phép. Nếu giấy phép hết hạn, thiết bị bị thu hồi, chữ ký sai hoặc máy chủ từ chối, VBA chặn tính năng premium, hiện hướng dẫn rồi đóng workbook sau OK/X khi đang ở bước mở file.

Nếu mất mạng trong thời gian ân hạn đã cấu hình, VBA có thể cho phép dùng tạm theo token hiện có. Hết ân hạn mà chưa xác minh được, tính năng cần giấy phép bị chặn và workbook đóng sau hướng dẫn khi mở lại.

Thiết kế này **không dùng mật khẩu riêng để bảo vệ từng sheet**. Việc cấp quyền được kiểm tra bằng token và logic VBA; nó không mã hóa nội dung sheet. Nếu công thức hoặc dữ liệu nằm trong workbook, người có kỹ thuật vẫn có thể tìm cách xem hoặc sao chép chúng.

## 4. Khi mất kết nối Internet

Token đã ký có thể được kiểm tra cục bộ mà không cần gọi máy chủ mỗi lần mở. Nếu đang trong khoảng ân hạn offline được cấu hình, VBA cho dùng tạm. Hết ân hạn mà vẫn không kết nối để xác minh, VBA chặn tính năng premium và đóng workbook sau OK/X trên hướng dẫn lúc mở file.

## 5. Điều gì xảy ra khi chia sẻ file?

| Tình huống | Kết quả theo thiết kế |
|---|---|
| Gửi bản sao file cho người khác trên máy khác | Token gắn với dấu vân tay máy ban đầu nên không xác thực được trên máy mới. Người nhận thấy hướng dẫn đăng nhập/kích hoạt; tính năng premium chưa dùng được cho đến khi kích hoạt hợp lệ và còn suất thiết bị. |
| Cùng người dùng đổi hoặc cài lại máy | Dấu vân tay có thể thay đổi. Người dùng có thể phải tắt thiết bị cũ trong tài khoản rồi kích hoạt lại. |
| Chép token sang workbook/máy khác | Chữ ký token có thể đúng, nhưng thông tin thiết bị không khớp; VBA yêu cầu xác minh/kích hoạt lại và chặn tính năng premium. |
| Giấy phép hết hạn hoặc bị thu hồi | Lần xác minh tiếp theo bị từ chối; tính năng premium bị chặn và workbook đóng sau hướng dẫn khi mở lại. |
| Người dùng mở file khi đang mất mạng | Có thể dùng trong khoảng ân hạn offline nếu token còn hợp lệ; hết ân hạn thì cần kết nối lại, workbook đóng sau hướng dẫn. |

## 6. Cách này bảo vệ được gì và không bảo vệ được gì?

### Giúp hạn chế

- Chia sẻ file cho người khác để dùng ngay trên máy khác.
- Vượt số thiết bị được cấp trong gói.
- Chặn các tính năng cần giấy phép khi thiết bị không khớp, giấy phép hết hạn/bị thu hồi hoặc đã hết ân hạn offline.
- Tự tạo token hợp lệ bằng cách sửa một cờ hoặc một ô trong workbook.

### Không thể đảm bảo

- VBA không thể chống lại người có kỹ thuật gỡ hoặc sửa macro.
- Kiểm tra cấp phép trong VBA không phải mã hóa nội dung; nó không chống được việc trích xuất có chủ đích.
- Công thức và dữ liệu nằm trong file tải về có thể bị phân tích hoặc sao chép.
- Chữ ký macro xác nhận nguồn phát hành và giúp Excel nhận diện macro đã ký; nó không làm cho file không thể sao chép.
- VBA có thể bị gỡ hoặc sửa bởi người có kỹ thuật; vì vậy nó là lớp răn đe chia sẻ thông thường, không phải ranh giới bảo mật tuyệt đối.

Do đó, mục tiêu thực tế là làm cho việc chia sẻ thông thường không thuận tiện. Nếu logic nghiệp vụ cần được bảo vệ mạnh hơn khỏi việc trích xuất, kiến trúc dự án khuyến nghị đặt logic có giá trị cao trong web app do máy chủ vận hành, thay vì để toàn bộ công thức trong file tải xuống.

## 7. Vai trò của chữ ký macro

Ngoài token cấp phép, kế hoạch phát hành file Excel đề xuất ký dự án macro bằng chứng thư OV. Khóa ký cần được bảo vệ bằng thiết bị phần cứng hoặc cơ chế được phê duyệt, do người phụ trách phát hành/quản lý bảo mật giữ; không xuất khóa riêng và không dùng chứng thư tự ký cho khách hàng.

Chữ ký macro và token kích hoạt giải quyết hai việc khác nhau:

| Cơ chế | Câu hỏi nó trả lời |
|---|---|
| Chữ ký macro | Macro này có đến từ nhà phát hành đã ký và còn nguyên vẹn kể từ lúc ký không? |
| Token kích hoạt | Giấy phép này có cho phép thiết bị hiện tại dùng sản phẩm không? |
| Kiểm tra trong VBA | Có thể từ chối chạy thao tác/tính năng cần giấy phép khi thiết bị không hợp lệ không? |

Chữ ký macro không thay thế kiểm tra giấy phép; token không thay thế việc ký macro; kiểm tra trong VBA không mã hóa các sheet hoặc công thức trong workbook.

## 8. Những nguyên tắc an toàn quan trọng

- Không nhúng khóa riêng Ed25519, shared secret hay quyết định cấp phép vào VBA.
- Chỉ nhúng khóa công khai và địa chỉ API cần thiết.
- Gửi hash dấu vân tay thay vì mã định danh máy thô.
- Máy chủ luôn tự tra quyền từ dữ liệu của mình; không tin `license_id`, `device_id`, `fp_hash` hoặc gói do client tự khai như bằng chứng cấp quyền.
- Không ghi token hoặc dấu vân tay vào thông báo, URL hay log dễ bị lộ.
- VBA chỉ là lớp răn đe phía client. Quyết định cấp quyền đáng tin cậy thuộc về máy chủ.

## 9. Trạng thái trong dự án

Luồng token, xác minh theo chu kỳ và khoảng ân hạn offline dựa trên `docs/ARCHITECTURE-PLAN.md` của API. Chính sách **đóng workbook sau OK/X** là thay đổi theo yêu cầu mới, hiện chỉ có trong bản mẫu một file của `tool-vba`; tài liệu thiết kế ở repo API còn mô tả giữ file mở và cần được cập nhật khi tích hợp. Danh mục test ghi nhận ánh xạ lỗi, hành vi ân hạn và việc không ghi token/hash vào worksheet, log hoặc hộp thoại.

Tuy nhiên, tài liệu tiến độ cho biết phát hành Excel hoàn chỉnh còn phụ thuộc worker Windows để chèn/xử lý VBA và ký file. Vì thế, nội dung ở trên cần được hiểu là **thiết kế và mục tiêu vận hành của dự án**, không phải khẳng định rằng mọi bước phát hành/ký macro đã được bật trong môi trường hiện tại.
