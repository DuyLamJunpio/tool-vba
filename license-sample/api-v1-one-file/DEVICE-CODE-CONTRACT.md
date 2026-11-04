# Hợp đồng mã thiết bị với digital-license-api

Trạng thái đối chiếu ngày 29/10/2026: **mã nguồn API local đã có create/redeem, lookup/approve/deny và validate-device**, tại `apps/api/src/modules/licensing/device-code.controller.ts` và `device-activation.controller.ts`, kèm test. Chưa xác nhận trạng thái triển khai production. Không phát hành cho khách cho đến khi API, Store, khóa phát hành và bộ xác minh Ed25519 đều vượt qua kiểm thử trên Excel Windows. Mọi URL dưới đây có tiền tố `/api/v1` của API; không điền hostname phát hành trong mã mẫu.

## Luồng khách hàng

1. Excel tạo `fp_hash` trên máy hiện tại và gọi `POST /auth/device-code` qua HTTPS. API trả `device_code` ngẫu nhiên (chỉ giữ trong bộ nhớ Excel), `user_code` ngắn để khách gõ và `expires_in` (giây, 30–1800). API lưu **hash** của `device_code`, `fp_hash`, sản phẩm, hạn dùng và trạng thái; không cấp quyền ở bước này.
2. Excel chỉ hiển thị `user_code`, mở `Cfg_ActivationWebUrl()` cố định. URL không mang mã, token hoặc fingerprint. Khách đăng nhập/đăng ký trên Store, nhập mã và xác nhận sản phẩm/máy. Nếu chưa có license ACTIVE, Store hướng khách mua hoặc chờ thanh toán; không tự duyệt mã.
3. Store gọi route xác nhận có xác thực khách hàng. API kiểm tra sơ bộ license ACTIVE cho sản phẩm và đánh dấu mã đã được chủ tài khoản duyệt; bước này **chưa chiếm suất thiết bị**. Excel gọi `POST /auth/device-code/redeem` sau khi khách bấm kiểm tra. Tại lúc redeem, API kiểm tra lại quyền, dùng cùng logic `ActivationService.activate` và giới hạn thiết bị trong transaction, ghi audit rồi trả token Ed25519. Chỉ đúng `device_code` và `fp_hash` đã gắn mới nhận token; mã chỉ đổi được một lần.
4. Khi token đến kỳ `rvi`, đã hết `exp` hoặc không được dùng cache offline (kể cả `ogh = 0`), Excel gọi `POST /activations/validate-device` với token kích hoạt, `fp_hash`, `request_id`, `timestamp`. Route dùng guard riêng xác minh chữ ký/JTI, tra license/device và gọi lại logic validate hiện có. Không cấp access/refresh token tài khoản cho VBA. Lỗi mạng chỉ được dùng token cũ chưa hết `exp` và còn trong `ogh`; lỗi từ chối rõ ràng từ API khóa quyền ngay. Theo DEC-104, riêng validate-device chấp nhận token hết hạn có chữ ký và ràng buộc hợp lệ để server tra lại license/device và cấp token mới. Token hết hạn không cấp quyền offline, không bắt khách tạo mã mới; mất mạng thì yêu cầu kết nối mạng.

## JSON client đã dùng

`POST /auth/device-code` (public, rate-limited):

```json
{"product_key":"<public product key>","version":"<workbook SemVer>","fp_hash":"<SHA-256 hex>","request_id":"<UUID>","timestamp":0}
```

Thành công: `{"success":true,"data":{"device_code":"<unguessable base64url>","user_code":"XXXX-XXXX","expires_in":600}}`. Giá trị `600` chỉ minh họa; API quyết định thời hạn qua cấu hình.

`POST /auth/device-code/redeem` (public, rate-limited):

```json
{"device_code":"<unguessable base64url>","fp_hash":"<SHA-256 hex>","request_id":"<UUID>","timestamp":0}
```

Trả HTTP 202 với `code: "AUTHORIZATION_PENDING"` khi chờ; HTTP 200 với `data.token` chỉ sau khi khách đã xác nhận và API đã kích hoạt. Mã hết hạn, từ chối, đã dùng lại hoặc không khớp thiết bị trả lỗi ổn định và không cấp token.

`POST /activations/validate-device` gửi Bearer activation token và thân request giống `/activations/validate` hiện tại. Route này cần xác thực **activation token**, không nhận JWT phiên khách hàng làm bằng chứng ở luồng này. Nó không thay thế route `/activations/validate` đang dùng bởi client cũ.

Store dùng `POST /auth/device-code/lookup`, `POST /auth/device-code/approve` và `POST /auth/device-code/deny` qua phiên khách hàng đã đăng nhập. Cả ba nhận `user_code` trong JSON body, không đặt mã trong URL. Lookup chỉ trả thông tin sản phẩm/phiên bản cần xác nhận, không trả `device_code` hoặc `fp_hash`; approve đánh dấu yêu cầu thuộc chủ tài khoản; deny từ chối yêu cầu. Route và JSON đã được cài đặt trong controller/DTO và mô tả tại `docs/api/v1.md`; việc triển khai trên URL phát hành cần kiểm tra riêng.

## Ràng buộc an toàn

- Sau validate-device trả `TOKEN_INVALID`, `DEVICE_REVOKED`, `DEVICE_MISMATCH` qua HTTP 4xx (trừ 429), client xóa cache và có thể đề nghị mã mới một lần ngay trong lần mở đó. Token cache sai chữ ký/`kid`/`prd`/`ver` hoặc fingerprint khác máy cũng được xóa trước khi đề nghị lại. Không cấp quyền cũ sau lỗi; token mới vẫn phải qua toàn bộ kiểm tra. Mất mạng, rate limit, timestamp sai, 5xx và license hết hạn/tạm ngưng giữ cache để validate lại khi điều kiện được khôi phục. Xóa cache thất bại thì khóa với `CACHE_UNAVAILABLE`.
- Code có hạn, một lần, dùng nguồn ngẫu nhiên mật mã; lưu hash của bí mật thiết bị, không log mã, token hoặc `fp_hash`. Giới hạn tạo mã, nhập sai, redeem, và polling theo IP/tài khoản.
- Browser chỉ duyệt khi user sở hữu license ACTIVE cho sản phẩm và phiên bản workbook. API tự chọn license, không tin id/plan do Excel gửi. Cùng máy kích hoạt lại không tiêu thêm suất. VBA so khớp `prd` và `ver` trong token đã ký với cấu hình phát hành.
- Sau redeem, VBA chỉ nhận quyền sau khi kiểm tra chữ ký Ed25519, `kid`, `prd`, `ver`, `dev`, `lic`, `jti`, `exp`, `rvi`, `ogh`, `iat`, `feat`. Verifier Ed25519 đã được cài đặt bằng VBA thuần, mọi lỗi trả False; token mới trả về vẫn phải chưa hết hạn. Cần Compile và chạy thử thật trong Excel theo [RELEASE-CHECKS.md](RELEASE-CHECKS.md).
- Excel chỉ tự đóng workbook nếu macro đang được phép chạy. Không tắt bảo mật macro toàn cục, không dùng mật khẩu từng sheet, không đóng các workbook khác.
