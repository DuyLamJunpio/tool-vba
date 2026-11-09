# Báo cáo: thay đổi hướng xử lý VBA cho sản phẩm Excel số

**Mục đích:** giải thích vì sao hướng cũ dựa trên `vbaProject.bin` không đáp ứng việc bổ sung mã vào workbook đã có VBA, và đề xuất hướng mới: đóng gói hàng loạt bằng Microsoft Excel Automation trên một máy Windows chuyên dụng.

## 1. Bối cảnh và mục tiêu kinh doanh

Sản phẩm được bán dưới dạng file Excel có macro. Mỗi file khách hàng tải về cần tự mang đầy đủ chức năng VBA; không thể phụ thuộc vào Excel Add-in cài sẵn trên máy khách.

Nhiều file nguồn đã có VBA riêng, gồm module thông thường, mã gắn với `ThisWorkbook` và mã gắn với các worksheet. Ví dụ, workbook HR hiện có module `ModInHangLoat` và các mã sự kiện trong những sheet. Nhu cầu là **bổ sung một module VBA sản phẩm vào từng file mà không làm mất mã VBA gốc**.

Các nguyên tắc không thay đổi:

- Không sửa file nguồn hoặc ghi đè âm thầm.
- Chỉ phát hành file kết quả `.xlsm` đã hoàn tất kiểm tra.
- Không chạy macro của file trong lúc đóng gói.
- Không đưa workbook hay mã VBA của khách hàng vào log, dịch vụ web hoặc bên thứ ba.
- Không coi việc khóa VBA bằng mật khẩu là bảo vệ tuyệt đối.

## 2. Hướng cũ: thay `vbaProject.bin`

### 2.1 `vbaProject.bin` thực chất là gì?

Một workbook Excel macro-enabled (`.xlsm`) là một gói Open XML/ZIP. Dự án VBA của workbook thường nằm trong entry nội bộ:

```text
xl/vbaProject.bin
```

`vbaProject.bin` là toàn bộ container nhị phân của dự án VBA, không phải một file module `.bas` riêng lẻ. Nó có thể chứa:

- Mã của `ThisWorkbook`.
- Mã gắn với từng worksheet.
- Các module chuẩn, ví dụ `ModInHangLoat`.
- Class module, UserForm, resource liên quan, reference và metadata của dự án.

Vì vậy, hình cây `VBAProject` trong VBA Editor chính là nội dung được đóng gói bên trong một `vbaProject.bin`.

### 2.2 Hướng cũ hoạt động như thế nào?

Tool ban đầu dùng Open XML SDK để:

1. Sao chép workbook nguồn thành file tạm.
2. Thêm hoặc thay `VbaProjectPart`.
3. Ghi file `vbaProject.bin` đã được chọn vào part đó.
4. Đổi loại workbook thành `.xlsm` và kiểm tra cấu trúc package.

Hướng này đúng cho hai tình huống hẹp:

- File `.xlsx` chưa có VBA và đã có một `vbaProject.bin` hoàn chỉnh, được chuẩn bị sẵn.
- File `.xlsm` có VBA được phép thay **toàn bộ** bằng một dự án VBA khác đã được phê duyệt.

### 2.3 Vấn đề cốt lõi

Open XML SDK chỉ nhìn `vbaProject.bin` như dữ liệu nhị phân gắn vào workbook. Nó không cung cấp API để mở, hiểu, biên dịch hoặc ghép một module VBA vào bên trong container này.

Do đó, nếu workbook HR đang có VBA gốc và tool thay nó bằng một `vbaProject.bin` chỉ chứa macro mới, file kết quả sẽ:

- Mất `ModInHangLoat`.
- Mất code trong `ThisWorkbook`.
- Mất code của worksheet.
- Có thể mất form, reference hoặc cấu hình dự án VBA cũ.

Nói ngắn gọn: hướng cũ thay cả chiếc hộp VBA, chứ không thêm một món đồ vào chiếc hộp đang có.

### 2.4 Tại sao không “merge” trực tiếp bằng Open XML SDK?

Không có thao tác Open XML chuẩn nào mang nghĩa “thêm module VBA”. Làm việc trực tiếp với cấu trúc nhị phân của VBA project sẽ cần parser/ghi lại OLE Compound File và định dạng nội bộ của VBA. Đây là phạm vi rủi ro cao:

- Dễ làm hỏng dự án VBA hoặc workbook.
- Không đáng tin cậy với UserForm, reference và mã sheet/workbook.
- Khó kiểm chứng trên nhiều phiên bản Office.
- Không phù hợp để dùng làm đường phát hành sản phẩm.

Vì vậy không nên dùng cách sửa ZIP hoặc sửa binary thiếu kiểm chứng để cố ghép VBA.

### 2.5 Các giới hạn khác của hướng cũ

- Nó không tạo hoặc biên dịch VBA từ file `.bas`.
- Nó không biết macro gốc có tương thích với macro mới hay không.
- Mọi chữ ký VBA cũ sẽ có thể không còn hợp lệ sau khi thay đổi dự án.
- Kiểm tra `xl/vbaProject.bin` tồn tại không chứng minh macro an toàn, chạy đúng hoặc được ký số.

## 3. Kết luận về hướng cũ

Hướng dùng `vbaProject.bin` không sai, nhưng không đúng bài toán sản phẩm hiện tại.

Nó nên được giữ như một chức năng riêng cho trường hợp **đóng gói một dự án VBA hoàn chỉnh vào workbook không có VBA** hoặc **thay toàn bộ VBA khi người vận hành chủ động xác nhận**. Nó không được dùng để thêm macro vào các workbook HR đã có VBA gốc.

## 4. Hướng mới được đề xuất: Excel Automation để thêm module vào bản sao

### 4.1 Ý tưởng

Thay vì can thiệp vào `vbaProject.bin`, dùng chính Microsoft Excel trên Windows để import một module `.bas` đã được phê duyệt vào VBA Project hiện có.

Excel hiểu cấu trúc VBA nội bộ, nên nó có thể thêm module mới mà giữ các module, mã sheet và mã workbook cũ.

### 4.2 Luồng cho một workbook

```text
Workbook nguồn .xlsm có VBA gốc
        +
Module sản phẩm đã duyệt .bas
        |
        v
Tạo bản sao output riêng
        |
        v
Mở bản sao trong Microsoft Excel với macro bị vô hiệu
        |
        v
Import module .bas vào VBA Project hiện có
        |
        v
Lưu output .xlsm, đóng Excel, kiểm tra và ghi log
```

Ví dụ:

```text
HR_Quan_ly_nhan_su_v2.0.xlsm
  + Tekmonk_Product_v1.bas
  -> HR_Quan_ly_nhan_su_v2.0_Tekmonk_v1.xlsm
```

Trong file kết quả, `ModInHangLoat`, `ThisWorkbook`, code các sheet và mọi module cũ được giữ; `Tekmonk_Product_v1` là module mới được thêm.

### 4.3 Phạm vi phiên bản đầu

Để giảm rủi ro, phiên bản đầu chỉ nên hỗ trợ:

- Input: `.xlsm` hợp lệ, có VBA project sẵn.
- Module bổ sung: một file `.bas` được phê duyệt.
- Output: file `.xlsm` mới, đặt tên khác file nguồn.
- Hành vi khi tên module đã tồn tại: từ chối và báo lỗi; không tự xóa hoặc thay thế module cũ.

Chưa nên hỗ trợ tự động import `.cls`, `.frm`, `.frx`, ActiveX control hoặc tự quản lý VBA references. Các loại này cần quy trình kiểm thử riêng.

## 5. Kiến trúc triển khai phù hợp cho sản phẩm số

### 5.1 Tách máy đóng gói khỏi website bán hàng

Không chạy Excel Automation trong backend web, API, Windows Service hoặc máy chủ tải file của cửa hàng.

Thay vào đó có hai vùng độc lập:

1. **Máy đóng gói Windows riêng:** dùng để tạo các file sản phẩm `.xlsm` trước khi phát hành.
2. **Website bán hàng:** chỉ lưu metadata đơn hàng và phân phối các file `.xlsm` đã đóng gói, đã kiểm tra.

```text
File mẫu + module đã duyệt
        |
        v
Máy đóng gói Windows riêng
        |
        v
Kiểm thử / ký số / phê duyệt phát hành
        |
        v
Kho file phát hành của website
        |
        v
Khách hàng tải file thành phẩm
```

Điều này tránh để website mở workbook hoặc macro từ nguồn không tin cậy. Office Automation không phù hợp để chạy như tiến trình web không giám sát, đặc biệt khi có macro và file từ nhiều nguồn.

### 5.2 Trường hợp sản phẩm không cá nhân hóa

Đây là trường hợp tối ưu:

1. Đóng gói mỗi phiên bản sản phẩm một lần.
2. Kiểm thử workbook kết quả.
3. Ký số dự án VBA của bản cuối nếu tổ chức có quy trình ký.
4. Upload file thành phẩm lên kho phát hành.
5. Mọi khách mua cùng phiên bản đều tải cùng artifact đã kiểm thử.

Không cần chạy Excel theo từng đơn hàng.

### 5.3 Trường hợp cần cá nhân hóa theo đơn hàng

Nếu chỉ cần tên khách, mã đơn hoặc dữ liệu hiển thị, ưu tiên tạo bản cá nhân hóa từ một `.xlsm` thành phẩm bằng xử lý workbook có kiểm soát, đồng thời giữ nguyên VBA đã đóng gói.

Nếu mỗi đơn lại đòi hỏi mã VBA khác nhau, không nên thực hiện Excel Automation trong web server. Cần một hàng đợi đóng gói tách biệt, trên máy Windows chuyên dụng, có giám sát vận hành và chính sách cấp phép phù hợp. Trước khi triển khai mô hình này, cần đánh giá tải, khôi phục khi Excel treo và yêu cầu cấp phép Office/RPA.

## 6. Điều kiện bắt buộc trên máy đóng gói

- Windows và Microsoft Excel desktop hợp lệ, đã kích hoạt.
- Một tài khoản vận hành riêng, không dùng chung với tài khoản quản trị website.
- Chỉ người vận hành được phép đăng nhập và chạy tool.
- Thư mục input/module/output có quyền NTFS giới hạn.
- Quản trị viên tự bật trong Excel: `Trust access to the VBA project object model`.
- Tool không tự bật Trust Center, không sửa Registry và không thay đổi chính sách bảo mật Office.
- Máy không nhận workbook do khách hàng upload để xử lý tự động.
- Mọi macro trong workbook phải bị vô hiệu trong phiên đóng gói.

Quyền “Trust access to the VBA project object model” là điều kiện để Excel cho phép tool import module vào VBA project. Đây là quyền mạnh; chỉ nên bật trên máy đóng gói cô lập, không bật đại trà trên máy nhân viên hoặc máy web.

## 7. Yêu cầu chức năng của tool batch mới

### 7.1 Input

- Một hoặc nhiều `.xlsm` nguồn được nội bộ phê duyệt.
- Một file `.bas` đã duyệt.
- Thư mục output mới.
- Tên phiên bản phát hành.

### 7.2 Kiểm tra trước xử lý

- Input tồn tại, phần mở rộng `.xlsm`, không trùng output.
- Workbook là Open XML package hợp lệ.
- Workbook có `xl/vbaProject.bin`.
- Module `.bas` tồn tại và SHA-256 trùng manifest phát hành.
- Output chưa tồn tại, hoặc người vận hành xác nhận rõ việc thay output.
- Tên module đích không có sẵn trong VBA Project.

### 7.3 Xử lý an toàn

- Sao chép input sang một file tạm trong thư mục output.
- Dùng Microsoft Excel để mở file tạm; tắt hiển thị cảnh báo, tắt event và đặt bảo mật automation để không chạy macro.
- Import duy nhất module `.bas` đã duyệt.
- Lưu workbook dưới dạng `.xlsm`.
- Đóng workbook và giải phóng hoàn toàn tiến trình Excel do tool tạo.
- Chỉ đổi tên file tạm thành output khi mọi bước kiểm tra đã đạt.
- Nếu thất bại: xóa file tạm do tool tạo, giữ nguyên input và output hiện hữu.

### 7.4 Kiểm tra sau xử lý

- Output là Open XML package hợp lệ.
- Output có `xl/vbaProject.bin`.
- Output có đúng phần mở rộng và content type `.xlsm`.
- VBA project chứa module mới với đúng tên kỳ vọng.
- VBA project vẫn chứa danh sách các component VBA gốc đã ghi nhận trước khi import.
- Không có macro nào được thực thi trong quá trình đóng gói.
- File nguồn có hash không đổi trước/sau.

Việc xác minh component cũ vẫn tồn tại không nên được diễn giải là chứng minh tuyệt đối mã gốc chưa đổi từng byte. Excel có thể chuẩn hóa metadata khi lưu. Chính sách an toàn là luôn giữ nguyên file nguồn, tạo output mới và thử nghiệm trên output.

### 7.5 Log phát hành

Log chỉ nên có metadata, không có nội dung workbook hoặc nội dung VBA:

```text
thời gian | file nguồn | SHA-256 nguồn | module | SHA-256 module |
output | trạng thái | lý do thất bại (nếu có) | phiên bản tool
```

Không ghi source code VBA, dữ liệu worksheet, mật khẩu hoặc binary vào log.

## 8. Chữ ký số và bảo vệ mã

Khi thay đổi VBA project, chữ ký VBA hiện có có thể không còn hợp lệ. Vì vậy:

1. Kiểm tra trước khi xử lý xem file có dấu hiệu có chữ ký VBA/package để cảnh báo.
2. Sau khi import module, coi output là một bản phát hành mới.
3. Nếu tổ chức dùng code signing, ký **toàn bộ VBA project ở bước cuối** bằng chứng thư ký mã của doanh nghiệp theo quy trình đã phê duyệt.
4. Không tự tạo chứng thư, không tự ký giả và không báo “đã ký” nếu chưa kiểm chứng bằng cơ chế tin cậy.

Mật khẩu khóa VBA Project chỉ giúp hạn chế việc xem/sửa thông thường; nó không bảo vệ tuyệt đối mã VBA trước người có kiến thức và công cụ phù hợp. Bảo vệ thực tế cần kết hợp kiểm soát bản phát hành, ký số, phân quyền thư mục và quy trình cập nhật phiên bản.

## 9. Kế hoạch triển khai đề xuất

### Giai đoạn 1 — Proof of concept có kiểm soát

- Chọn 3–5 bản sao workbook mẫu đại diện.
- Chọn một module `.bas` nhỏ, không ảnh hưởng nghiệp vụ.
- Chạy tool trên máy Windows chuyên dụng.
- Xác nhận module mới xuất hiện, VBA cũ vẫn tồn tại, workbook mở được và macro không tự chạy khi đóng gói.

### Giai đoạn 2 — Kiểm thử phát hành

- Chạy 20–50 file mẫu đa dạng.
- Kiểm tra tên module trùng, file bị khóa, file hỏng, đường dẫn Unicode/tiếng Việt và dung lượng lớn.
- Đánh giá việc Excel đóng sạch, không còn tiến trình treo.
- Thực hiện kiểm thử thủ công mở các output bằng Excel trên máy thử nghiệm, theo chính sách macro của tổ chức.

### Giai đoạn 3 — Đóng gói hàng loạt

- Khóa version của module và manifest hash.
- Chạy batch trên các workbook đã được phê duyệt.
- Rà soát log, tách các file lỗi để xử lý riêng.
- Chỉ đưa các output PASS sang kho phát hành.

### Giai đoạn 4 — Vận hành phiên bản mới

- Mỗi lần cập nhật macro tạo version module mới.
- Không sửa lại artifact đã phát hành.
- Lưu manifest, log và checksum cho từng đợt phát hành để truy vết.
- Nếu cần ký số, ký ở cuối mỗi đợt đóng gói trước khi upload.

## 10. Những điều không nên làm

- Không dùng `vbaProject.bin` chỉ chứa macro mới để thay thế VBA của workbook đang có macro.
- Không tự ghép/sửa binary VBA hoặc ZIP package bằng cách không có kiểm chứng.
- Không để web server nhận file Excel của khách rồi mở Excel để import VBA.
- Không chạy tool bằng tài khoản có đặc quyền không cần thiết.
- Không bật Trust Center bằng Registry/script mà không có thao tác và phê duyệt của quản trị viên.
- Không chạy macro từ file đầu vào để “test” trong lúc batch đóng gói.
- Không ghi đè file nguồn hoặc phát hành file khi một bước kiểm tra thất bại.

## 11. Quyết định đề xuất

Chọn **hướng 2: tool batch dùng Microsoft Excel Automation để import module `.bas` vào bản sao `.xlsm`**.

Tool phải hoạt động trên máy Windows đóng gói riêng, tách biệt khỏi website bán sản phẩm. Website chỉ phân phối các `.xlsm` đã đóng gói, kiểm thử và phê duyệt. Đây là hướng đáp ứng yêu cầu “mỗi file bán ra tự có macro”, đồng thời giữ được VBA gốc trong từng workbook.

## 12. Bước tiếp theo

1. Chốt một workbook mẫu và module `.bas` thử nghiệm.
2. Thiết kế lại tool thành hai chế độ tách biệt: 
   - **Chèn/thay toàn bộ `vbaProject.bin`** cho use case cũ.
   - **Thêm module `.bas` vào VBA project có sẵn** cho sản phẩm số.
3. Hoàn thiện proof of concept trên 3–5 bản sao trước.
4. Viết unit test cho các phần kiểm tra file/hash/path và integration test riêng trên máy có Excel.
5. Chỉ sau khi kiểm thử đạt mới triển khai batch hàng trăm file và đưa output lên hệ thống bán hàng.

