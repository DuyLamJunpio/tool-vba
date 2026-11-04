# Kiểm tra trước phát hành

Phạm vi C1–C7/C9 đã được chủ dự án duyệt ngày 29/10/2026. Giữ nguyên DEC-104, `ogh = 0`, bỏ JWT tài khoản và phục hồi cache. Không có quyền thì hướng dẫn, OK/X đóng chỉ ThisWorkbook không lưu. Manual hiện có [44 ca](tests/manual/DeviceCodeActivationSmokeTest.md).

## Trạng thái xác minh

**VBA mới chưa được Compile/chạy trong Excel.** Kết quả Excel bản trước không áp dụng cho bản này. Cần chủ dự án Debug → Compile VBAProject rồi chạy **64/64 mock** và **57/57 signed regression**, sau đó smoke test workbook/API/Store thật. URL sản phẩm và môi trường vẫn để trống, chưa đủ để phát hành.

Kiểm tra tự động ngoài Excel gồm contract/SHA-256, sinh harness bằng API signer và kiểm tra cấu trúc/arity. Chúng không thực thi VBA, WinHTTP, sheet events hay chứng minh tương thích Excel 32/64 bit. Kết quả .NET và contract được ghi ở cuối tài liệu sau lần chạy cuối.

## Chạy trong Excel Windows

1. Tạo `.xlsm` thử nghiệm riêng, import `modLicenseApiV1.bas`, không import cùng module cũ. Debug → Compile VBAProject.
2. Chạy `License_RunMockTests`: kỳ vọng **64/64** (8 gốc/KAT RFC 8032, 8 policy, 16 phục hồi cache, 32 clock/cycle/key ring/DPAPI/SystemDrive). Không gọi API, không ghi/xóa Registry và không cấp quyền; các ca DPAPI dùng API Windows với dữ liệu test trong bộ nhớ.
3. Chạy `node .\tests\build-excel-regression.mjs D:\Code\digital-license-api`. Import `TestResults/modLicenseRegression.bas` vào workbook thử nghiệm riêng, Compile, chạy `License_RunSignedRegressionTests`: kỳ vọng **57/57**. Không nhập harness vào workbook khách hàng.
4. Điền cấu hình được duyệt và thử token do API môi trường thử nghiệm ký với đúng `kid`/public key. Không ghi/chụp token, fingerprint, device_code hoặc private key. Chỉ lưu tên ca, trạng thái và kết quả.

Bộ sinh dùng `ActivationTokenService` từ API với hai cặp khóa test tạm trong bộ nhớ, không đọc `.env` hay khóa triển khai. Token có claim license/device/JTI giả, không có quyền trong DB. HTTP/fingerprint/thời gian/cache/UI/sheet được thay stub; ca DPAPI giữ protect/unprotect/cache parser thật, chỉ thay Registry bằng bộ nhớ. Hồi quy giữ nguyên 31 ca cũ (26 trạng thái/cache, 3 Auto_Open, 2 resume), thêm 26 ca: 7 key ring, 6 clock, 7 refresh giữa phiên, 3 DPAPI, 3 sheet gate. Cổng đóng file được đếm thay vì đóng workbook thử nghiệm. Đây không phải smoke test HTTP/DB/worksheet thật.

## Các kết quả bắt buộc

| Nhóm | Kết quả cần xác nhận |
|---|---|
| Lần đầu và phục hồi cache | Chỉ hiển thị user_code; duyệt Store rồi redeem. Cache invalid/revoked/mismatch chỉ đề nghị mã mới một lần sau xóa thành công. No/X đóng riêng workbook. Xóa lỗi phải khóa, không lặp. |
| DEC-104 và ogh = 0 | Chỉ quá exp thì gia hạn online, không mã mới; offline khóa. Fresh token hết exp bị từ chối. ogh = 0 vẫn validate khi mở trước rvi; mất mạng không được dùng cache. |
| Từ chối và lỗi tạm thời | Server từ chối không được grace. Chỉ ba lỗi TOKEN_INVALID/DEVICE_REVOKED/DEVICE_MISMATCH qua 4xx trừ 429 mới xóa. Mất mạng/429/timestamp/5xx/license expired/suspended giữ credential; gia hạn/mở khóa rồi dùng lại. Mốc giờ có thể tăng, không dời thời điểm validate. |
| C1: sheet | Có sheet hướng dẫn visible riêng; premium lưu very hidden, macro tắt chỉ thấy hướng dẫn. Có quyền mới visible; từ chối/đóng/BeforeSave ẩn, AfterSave khôi phục khi còn quyền. Kiểm tra cả Close bằng code, Save/SaveAs và bản lưu mở lại với macro tắt. Sai tên/all-premium/khóa cấu trúc phải fail-closed. |
| C2: đồng hồ | Giờ thấp hơn max đã lưu hoặc iat bị CLOCK_ROLLBACK, giữ cache, ẩn sheet và đóng theo hướng dẫn. Bằng mốc được chấp nhận. Mốc không giảm sau mở lại; tăng mốc không kéo dài rvi/ogh. Đồng bộ giờ rồi mở lại được. Dùng VM thử nghiệm. |
| C3: giữa phiên | Gọi hàm kiểm tra quyền khi qua rvi: thành công không có thông báo. Mất mạng chỉ thử một lần trong chu kỳ; lần gọi lặp không gửi thêm; chu kỳ tiếp theo được thử lại nếu phiên vẫn trong grace. Chặn gọi lồng; token hết exp được gia hạn, server từ chối không dùng grace. Hết quyền thì hướng dẫn và đóng. Không có timer nền. |
| C4: DPAPI | Cache mới round-trip đủ bốn trường, mốc không giảm. Cache cũ không entropy/khác product không giải mã được: coi chưa kích hoạt, không hiện lỗi giải mã; kích hoạt lại một lần nếu cần. |
| C5/C6: HTTPS và proxy | TLS 1.2/1.3; chứng chỉ sai không được bỏ qua, không redirect. Thử trực tiếp, proxy tĩnh Windows, PAC, override và proxy lỗi với timeout; không nhập credentials vào VBA. |
| C7: ổ hệ thống | SystemDrive C: giữ fingerprint cũ. VM hệ thống khác C: dùng serial ổ hệ thống; có thể cần kích hoạt lại và gỡ máy cũ nếu hết suất. Không in hash. |
| C9: hai khóa | Token mỗi khóa hợp lệ được nhận khi cặp tương ứng cấu hình đủ. Kid lạ, header/payload khác nhau hoặc chữ ký bằng khóa khác đều bị từ chối. Cặp thứ hai trống vẫn dùng khóa đầu; thiếu một nửa/trùng/sai kích thước không cấp quyền. |

Ghép [ThisWorkbook-events.txt](ThisWorkbook-events.txt) vào các sự kiện hiện có. Chủ workbook tạo sheet hướng dẫn; không dùng password sheet. Đảm bảo startup chỉ gọi một lần, kể cả `Workbooks.Open`. Bản phân phối phải lưu sheet premium ở trạng thái very hidden; standard module một mình không tự thêm event handler vào ThisWorkbook.

## Lệnh kiểm tra ngoài Excel

```powershell
powershell -NoProfile -File .\contract-check.ps1
node .\tests\build-excel-regression.mjs D:\Code\digital-license-api
node .\tests\vba-static-check.mjs
```

Tại `D:\Code\tool-vba`: `dotnet test ToolVba.sln --no-restore`. Nếu môi trường giới hạn quyền ghi chỉ cho phép thư mục mẫu, dùng cùng lệnh với `-p:DirectoryBuildPropsPath=D:\Code\tool-vba\license-sample\api-v1-one-file\tests\workspace-output.props --results-directory D:\Code\tool-vba\license-sample\api-v1-one-file\TestResults\dotnet-results`. Tệp props chỉ chuyển build artifacts vào `TestResults`, giữ assets đã restore trong repo; không bỏ test hay giảm ràng buộc bảo mật.

## Tài liệu ngoài phạm vi ghi

Phiên làm việc chỉ được ghi thư mục mẫu. Filesystem đã từ chối áp dụng trực tiếp tài liệu API và manual gốc; đã kiểm tra tài liệu API không thay đổi. [Bản vá API](patches/digital-license-api-docs.patch) chỉ sửa kiến trúc §7, nối DEC-107 và cập nhật `docs/GIOI-HAN-CHIA-SE-FILE-VBA.md`. [Bản vá manual gốc](tests/manual/repo-manual-update.patch) đồng bộ 44 ca cho `D:\Code\tool-vba\tests\manual\DeviceCodeActivationSmokeTest.md`. Cả hai đã qua `git apply --check`, **chưa áp dụng**. [Hướng dẫn áp dụng](patches/README.md) kèm SHA-256 trước/sau. Không commit.

## Kết quả ngoài Excel ngày 29/10/2026

- `contract-check.ps1`: **PASS**, bao gồm ASCII, SHA-256, TLS, entropy, key ring, clock, refresh guard và các ràng buộc DEC-104/ogh/cache cũ.
- Sinh harness: **57 ca**, thành công với API signer và khóa test tạm; kiểm tra cấu trúc/arity cho module phát hành và harness đều **PASS**. Đây chưa phải kết quả chạy 64/64 hoặc 57/57 trong Excel.
- `dotnet test ToolVba.sln --no-restore` bản nguyên lệnh bị chặn quyền ghi build artifacts ở repo cha. Cùng toàn bộ solution với props chuyển artifacts vào workspace: **34/34 PASS, 0 failed, 0 skipped**, không sửa/giảm test .NET. Kết quả [TRX](TestResults/dotnet-results/hardening-c.trx).
- Đối chiếu source trước/sau: verifier Ed25519, các thủ tục số học Ed*, `PolicyCanUse`, `PolicyRequiresOnline`, `NeedsDeviceActivation`, `ShouldClearCacheForServerError` và `CacheClear` giữ nguyên.
- **VBA mới chưa Compile hoặc chạy trên Excel**; chủ dự án vẫn phải thực hiện các bước Excel bên trên và smoke test môi trường thật.
