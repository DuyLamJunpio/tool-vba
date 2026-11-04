# Bản vá tài liệu ngoài phạm vi ghi

Filesystem của phiên chỉ cho ghi `D:\Code\tool-vba\license-sample\api-v1-one-file`. Thử áp dụng trực tiếp đã bị `Permission denied`; không reset/stash/commit hoặc thay đổi các tệp khác của chủ dự án. Các bản vá dưới đây **chưa được áp dụng** và đã qua `git apply --check` với working tree hiện tại ngày 29/10/2026.

- `digital-license-api-docs.patch`: ba tệp trong `D:\Code\digital-license-api\docs`. `ARCHITECTURE-PLAN.md` chỉ đổi §7; `DECISIONS.md` giữ nguyên toàn bộ nội dung và nối DEC-107 sau DEC-106; tài liệu giới hạn chia sẻ cập nhật cache DPAPI, sheet ẩn, clock, hai khóa và giới hạn macro. Không sửa source API hoặc cấu hình bí mật.
- `digital-license-api-docs.sha256.json`: hash trước/sau để kiểm tra. Đọc lại số DEC cuối trước áp dụng; nếu đã có DEC-107 do công việc khác, đổi số mục mới và tham chiếu tương ứng, không ghi đè.
- `architecture-section-7.md`, `DEC-107.md`: nội dung riêng để duyệt dễ hơn; nội dung này đã nằm trong patch.
- `../tests/manual/repo-manual-update.patch`: đồng bộ manual gốc của repo tool-vba lên 44 ca, 64 mock/57 signed. Manual hiện hành trong thư mục mẫu đã được cập nhật trực tiếp.

Chạy ở shell có quyền ghi hai repo, sau khi kiểm tra lại working tree. Không dùng reset/stash để ép patch khớp; nếu check thất bại thì cần ghép theo nội dung mới của chủ dự án.

```powershell
git -C D:\Code\digital-license-api apply --check D:\Code\tool-vba\license-sample\api-v1-one-file\patches\digital-license-api-docs.patch
git -C D:\Code\digital-license-api apply D:\Code\tool-vba\license-sample\api-v1-one-file\patches\digital-license-api-docs.patch
git -C D:\Code\tool-vba apply --check D:\Code\tool-vba\license-sample\api-v1-one-file\tests\manual\repo-manual-update.patch
git -C D:\Code\tool-vba apply D:\Code\tool-vba\license-sample\api-v1-one-file\tests\manual\repo-manual-update.patch
```

Nếu Git báo dubious ownership, chạy từ tài khoản sở hữu repo hoặc dùng `-c safe.directory=D:/Code/digital-license-api` cho riêng lệnh khi đã xác nhận repo. Không cần đổi cấu hình Git toàn cục. Các lệnh trên không commit.
