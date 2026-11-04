# Bàn giao triển khai API kích hoạt giấy phép Excel

Ngày lập: 26/10/2026. Dùng tệp này khi mở một phiên Codex mới với thư mục làm việc và quyền ghi tại `D:\Code\digital-license-api`.

## Prompt để dán vào Codex mới

```text
Hãy tiếp tục triển khai luồng kích hoạt giấy phép bằng mã thiết bị cho Excel trong repo D:\Code\digital-license-api. Tôi đã duyệt triển khai luồng này. Đọc toàn bộ tệp D:\Code\tool-vba\BAN-GIAO-AI-DIGITAL-LICENSE-API.md, sau đó đọc AGENTS.md, docs/AUTONOMY.md, docs/PROGRESS.md, 10 quyết định mới nhất trong docs/DECISIONS.md, docs/api/v1.md và các tài liệu liên quan trong docs trước khi sửa mã.

Hợp đồng đề xuất của client VBA nằm tại D:\Code\tool-vba\license-sample\api-v1-one-file\DEVICE-CODE-CONTRACT.md. Đọc mã VBA hiện tại tại D:\Code\tool-vba\license-sample\api-v1-one-file\modLicenseApiV1.bas để đối chiếu chính xác request/response, tên trường và mã lỗi. Đọc cả IMPLEMENTATION-STATUS.md và tests/manual/DeviceCodeActivationSmokeTest.md cùng thư mục.

Triển khai API và trang Store để người mua đăng nhập/đăng ký, nhập user_code, xem đúng sản phẩm/phiên bản và xác nhận kích hoạt máy đang dùng. Tạo các route cấp mã, tra cứu, duyệt, redeem và xác minh activation token theo hợp đồng đề xuất; chốt contract thực tế trong docs/api/v1.md. Tái sử dụng logic cấp quyền, giới hạn thiết bị, ký token và audit hiện có. Mã có hạn, dùng một lần, lưu hash của bí mật; ràng buộc đúng fingerprint; redeem phải an toàn trước yêu cầu đồng thời. Không gửi account JWT cho VBA, không đặt code/token/fingerprint trong URL hay log, không ghi khóa riêng hoặc shared secret vào VBA. Cấu hình thời hạn/giới hạn tốc độ bằng cấu hình thay vì số hardcode. Mọi route phải có guard hoặc @Public kèm giải thích.

Repo hiện có rất nhiều thay đổi chưa commit và tệp/thư mục mới. Giữ nguyên tất cả thay đổi hiện có; tuyệt đối không reset, clean, checkout đè, xóa hoặc ghi đè hàng loạt. Trước khi sửa, xem git status và diff từng tệp sẽ đụng tới, rồi chỉnh sửa theo ngữ cảnh hiện tại. Không đưa thay đổi có trước của tôi vào commit của bạn. Không đọc/sửa .env, không đụng dữ liệu production hoặc dữ liệu khách hàng thật. Nếu cần thay đổi thiết kế bảo mật ngoài luồng đã duyệt, thực hiện phần độc lập trước rồi nêu rõ điểm cần tôi quyết định theo AGENTS.md.

Viết test case trước theo quy định repo, sau đó triển khai và chạy test thực tế cho phần vừa làm. Kiểm thử tối thiểu: mã hết hạn/sai/đã dùng, chỉ chủ license ACTIVE duyệt được, sai sản phẩm/phiên bản/fingerprint, chưa duyệt trả pending, redeem một lần kể cả khi gọi đồng thời, kích hoạt lại cùng máy không tăng số thiết bị, vượt giới hạn bị chặn, license/device bị thu hồi, validate bằng activation token và từ chối account JWT/sai token. Chạy các lệnh kiểm tra phù hợp và pnpm verify nếu môi trường cho phép. Không reset database để chạy test nếu guard/migration hiện đang lỗi; ghi rõ trở ngại thay vì thao tác phá hủy.

Cập nhật tài liệu contract, test case, tiến độ và quyết định theo AGENTS.md. Sau cùng báo cáo tệp đã sửa, test đã chạy với số lượng/kết quả, test chưa chạy và lý do, những điều còn ngăn phát hành. Không tuyên bố client Excel đã cấp quyền thật khi VerifyEd25519Detached trong .bas còn trả False hoặc khóa ký/public key chưa được cấu hình.
```

## Bối cảnh đã xác nhận

- Repo API: `D:\Code\digital-license-api`. Phiên Codex cũ chỉ có quyền ghi `D:\Code\tool-vba`, nên **chưa hề sửa repo API** trong phần bàn giao này. Lần thử ghi tệp tạm vào repo API bị từ chối; yêu cầu vượt sandbox cũng bị chính sách phê duyệt tự động từ chối.
- Repo API đang có working tree rất bẩn: các tệp `apps/api`, Prisma, `docs`, `apps/admin`, `apps/store` đã được sửa/tạo/xóa bởi công việc có sẵn. Trạng thái cụ thể phải kiểm tra lại bằng `git status --short` tại phiên mới. Không mặc định những thay đổi này do Codex cũ tạo ra.
- API hiện có `POST /activations` và `POST /activations/validate` dưới account JWT. Chưa xác nhận có các route mã thiết bị hay `/activations/validate-device`; hãy kiểm tra lại mã mới nhất trước khi thêm. `ActivationService.activate` đã có logic khóa license row, giới hạn thiết bị, cùng máy không chiếm thêm suất. `ActivationService.validate` kiểm tra quyền và cấp token mới. `TokenService` dùng Ed25519 JWT.
- Client VBA đề xuất `POST /api/v1/auth/device-code`, `POST /api/v1/auth/device-code/redeem`, `POST /api/v1/activations/validate-device`; Store cần route lookup/approve có phiên khách hàng. Đây là **hợp đồng đề xuất**, chưa được triển khai trên API. Xem tệp contract để lấy JSON và quy tắc đầy đủ; nếu phải thay đổi JSON, cập nhật client một cách có chủ đích và ghi rõ lý do.
- Client VBA có `Auto_Open`, hướng dẫn kích hoạt, mở URL Store cố định, hiển thị `user_code`, redeem thủ công, cache DPAPI và đóng **chỉ workbook hiện tại** khi khách không có quyền và bấm OK/X. Macro bị vô hiệu là giới hạn của VBA: workbook vẫn có thể mở. Không dùng mật khẩu từng sheet. Cấu hình URL, product key/version, key ID và public key hiện còn placeholder.
- `VerifyEd25519Detached` trong `.bas` cố ý trả `False`. Vì thế chưa thể xác minh token thật hay phát hành workbook cho khách, dù API và Store hoàn tất. Đừng sửa thành `True` hoặc bỏ qua kiểm tra chữ ký. Đây là việc còn thiếu riêng cần giải pháp đã kiểm thử trên Excel Windows.
- Bộ `tool-vba` đã chạy `dotnet test ToolVba.sln --no-restore`: **34/34 đạt** ở lượt trước, gồm kiểm tra manifest `.bas`. Đây không phải test API hay thực thi VBA. Chưa có test API nào được chạy cho luồng mới.
- `D:\Code\tool-vba\tests\manual\DeviceCodeActivationSmokeTest.md` có 16 ca DC-01 đến DC-16. Phần lớn chờ API, cấu hình khóa và bộ xác minh chữ ký; không ghi nhận là đã đạt nếu chưa chạy trên Excel Windows.
- `docs/PROGRESS.md` của API từng ghi `pnpm db:test:reset` đang lỗi vì Prisma schema/migration engine. Đọc trạng thái mới nhất trước khi chạy. Không chạy reset hoặc migration phá hủy khi chưa xác minh đích là test schema/dev cho phép và điều kiện theo `AGENTS.md`.

## Phạm vi và tiêu chí hoàn tất

1. Backend lưu trạng thái yêu cầu mã thiết bị lâu bền, có hạn và chống dùng lại; mã bí mật chỉ lưu dạng hash. Thiết kế transaction đảm bảo không cấp hai token/chiếm hai suất khi redeem đồng thời.
2. API public cho Excel cấp mã/redeem được giới hạn tốc độ và không tiết lộ thông tin nhạy cảm. API xác thực cho Store tra cứu và duyệt chỉ thao tác trên license ACTIVE thuộc tài khoản. Redeem kiểm tra lại quyền ngay lúc cấp token.
3. Route xác minh bằng activation token hoạt động mà không cần account JWT; không làm hỏng route account JWT cũ. Token mới và mã lỗi khớp client VBA hoặc contract được cập nhật đồng bộ.
4. Store có đường dẫn kích hoạt dễ hiểu cho người mới đăng ký/đăng nhập, nhập mã và xác nhận máy. Không đưa mã hay fingerprint vào query string, analytics hoặc log.
5. Unit/integration test tập trung vào quyền sở hữu, vòng đời mã, thiết bị, cạnh tranh, thu hồi và contract. Ghi rõ test nào dùng DB thật và test nào chỉ mock. Cập nhật `docs/TEST-CASES.md`, `docs/api/v1.md`, `docs/PROGRESS.md` và `docs/DECISIONS.md` theo quy định repo.
6. Chạy lệnh kiểm tra và báo cáo kết quả thực, không suy từ code review. Nếu khóa ký, database test, Store hoặc Excel chưa sẵn sàng, phân biệt rõ phần đã triển khai với phần chưa xác nhận bằng thử nghiệm.

## Tệp tham chiếu chính

| Mục đích | Đường dẫn |
| --- | --- |
| Quy tắc repo API | `D:\Code\digital-license-api\AGENTS.md`, `docs\AUTONOMY.md` |
| Contract API hiện tại | `D:\Code\digital-license-api\docs\api\v1.md` |
| Kiến trúc, tiến độ, quyết định | `D:\Code\digital-license-api\docs\ARCHITECTURE-PLAN.md`, `docs\PROGRESS.md`, `docs\DECISIONS.md` |
| Contract mã thiết bị đề xuất | `D:\Code\tool-vba\license-sample\api-v1-one-file\DEVICE-CODE-CONTRACT.md` |
| Mã VBA hiện tại và manifest | `D:\Code\tool-vba\license-sample\api-v1-one-file\modLicenseApiV1.bas`, `modLicenseApiV1.manifest.json` |
| Tình trạng và smoke test Excel | `D:\Code\tool-vba\license-sample\api-v1-one-file\IMPLEMENTATION-STATUS.md`, `D:\Code\tool-vba\tests\manual\DeviceCodeActivationSmokeTest.md` |

## Lưu ý khi mở phiên mới

Mở thư mục `D:\Code\digital-license-api` làm workspace của AI **có quyền ghi**, sau đó dán nguyên prompt ở đầu tệp. Cho phiên mới quyền **đọc** `D:\Code\tool-vba` để đối chiếu client. Nếu phiên mới không thể ghi ngược vào `tool-vba`, vẫn có thể hoàn thành API/Store và ghi lại thay đổi cần đồng bộ client; đừng âm thầm tuyên bố đã cập nhật `.bas`.
