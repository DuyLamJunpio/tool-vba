# Trạng thái triển khai (29/10/2026)

Bản C1–C7/C9 đã sửa source và thêm kiểm tra, **chưa Compile/chạy VBA mới trong Excel**. Kết quả Excel của bản trước không chứng minh bản này chạy được.

## Hành vi đã có trong mã

- Giữ nguyên DEC-104 (expired token dùng để gia hạn), `ogh = 0`, bỏ đường JWT tài khoản và quy tắc xóa cache/đề nghị mã mới đã làm trước đây. Không thay server contract hay chính sách cấp quyền.
- C1: danh sách sheet cấu hình mặc định rỗng; visible khi có quyền, very hidden khi không có quyền/đóng. Có hàm trước/sau lưu để bản lưu không chứa sheet premium visible. Chủ workbook tự tạo sheet hướng dẫn visible và ghép [sự kiện ThisWorkbook](ThisWorkbook-events.txt); không dùng mật khẩu sheet. Lỗi cấu hình sheet từ chối quyền.
- C2: cache DPAPI có thêm mốc giờ lớn nhất chỉ tăng, giữ nguyên thời điểm validate. Giờ nhỏ hơn mốc hoặc `iat` đã ký là `CLOCK_ROLLBACK`, không xóa cache; hướng dẫn đồng bộ giờ rồi mở lại. Kiểm tra cục bộ không chống xóa hoặc phục hồi toàn bộ cache.
- C3: hàm kiểm tra quyền thử validate giữa phiên tối đa một lần mỗi chu kỳ `rvi`, chống gọi lại, dùng timeout hiện có. Thành công không hiện thông báo; lỗi mạng chỉ được xét cache trong `exp`/`ogh`. Khi hết quyền trong phiên, ẩn sheet, hướng dẫn rồi đóng chỉ ThisWorkbook không lưu. Không có timer: cần gọi hàm kiểm tra quyền ở điểm vào của tính năng.
- C4: DPAPI protect/unprotect dùng cùng entropy công khai cố định theo product key. Cache cũ không giải mã được coi như chưa kích hoạt, không hiện lỗi; nâng cấp có thể cần mã mới một lần.
- C5: TLS 1.2/1.3 trước `http.Open`; vẫn xác minh chứng chỉ và không redirect.
- C6: mặc định dùng proxy Windows người dùng (tĩnh/PAC/tự phát hiện), rồi WinHTTP hệ thống; `Cfg_ProxyServer` tùy chọn trống, không chứa credentials. Không sửa cấu hình proxy máy. Cần smoke test mạng công ty/PAC thực tế.
- C7: serial ổ lấy từ `%SystemDrive%`. Nếu SystemDrive là C:, thứ tự/thành phần/chuỗi hash không đổi. **Đổi cách tính có thể đổi fp_hash của máy đã kích hoạt dùng ổ hệ thống khác C**; phải hướng dẫn gỡ máy cũ nếu hết suất rồi kích hoạt lại.
- C9: hai cặp `kid`/public key; cặp đầu giữ nguyên `k2026-10` và khóa đang có, cặp thứ hai trống. Chọn theo header, bind payload với header, không thử khóa khác; khóa lạ hoặc chữ ký sai là `TOKEN_INVALID`. Cặp thiếu/trùng/sai định dạng làm cấu hình không hợp lệ.

API/Store URL, product key và version vẫn để trống. Không có private key/shared secret/JWT tài khoản trong VBA; `.bas` thuần ASCII, thông báo tiếng Việt không dấu. Token/fingerprint/device_code chỉ ở bộ nhớ hoặc cache mã hóa, không đi vào worksheet/log/URL/hộp thoại.

## Kiểm tra và giới hạn

- `License_RunMockTests`: **64 ca** = 8 ca gốc (có KAT RFC 8032), 8 policy, 16 phục hồi cache và 32 ca clock/cycle/key ring/DPAPI/SystemDrive.
- `tests/build-excel-regression.mjs`: sinh **57 ca** = 31 ca hồi quy cũ và 26 ca mới cho clock, refresh giữa phiên, hai khóa, DPAPI và cổng sheet. API signer dùng khóa test tạm, không đọc `.env` hay dùng private key triển khai. Registry/HTTP/UI/thời gian/fingerprint/sheet là stub; các ca DPAPI gọi Windows thật nhưng ciphertext chỉ lưu trong bộ nhớ.
- `tests/vba-static-check.mjs`: kiểm tra ASCII, cấu trúc khối và số tham số các lời gọi đã đổi; **không phải trình biên dịch VBA**.
- `contract-check.ps1`: kiểm tra cả ràng buộc bảo mật cũ và mới, SHA-256 manifest chữ hoa. Không nới lỏng kiểm tra để qua test.
- Hướng dẫn kiểm tra: [RELEASE-CHECKS.md](RELEASE-CHECKS.md), [44 ca thủ công](tests/manual/DeviceCodeActivationSmokeTest.md).

Các route create/redeem/lookup/approve/deny/validate-device và cơ chế gia hạn DEC-104 đã có trong repo API local. Đây chưa phải xác nhận deploy hoặc smoke test production. Verifier Ed25519 và số học TweetNaCl giữ nguyên; mọi lỗi trả False.

## Còn phải làm trước phát hành

1. Chủ dự án Debug → Compile VBAProject, chạy mock **64/64** và hồi quy có chữ ký **57/57** trên Excel Windows. Kiểm tra cả workbook thật với sự kiện đóng/lưu và sheet hướng dẫn.
2. Điền cấu hình môi trường đã duyệt, xác nhận hai public key/kid theo kế hoạch xoay khóa. Chạy API/Store thật, proxy/TLS và các ca clock trên VM thử nghiệm; không thay giờ máy đang làm việc.
3. Kiểm tra bản phân phối được lưu với sheet premium very hidden, ký macro theo quy trình phát hành. Ẩn sheet chỉ giảm chia sẻ thông thường, không mã hóa nội dung và không chống sửa macro.
