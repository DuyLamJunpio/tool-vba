# Mục lục tài liệu tool-vba

Đọc file nào khi nào. Cập nhật 09/11/2026.

## Trong repo này

| File | Đọc khi |
| --- | --- |
| [README.md](../README.md) | Bắt đầu: mục đích, cách dùng app, xuất bản gói, thứ tự deploy và kiểm thử production, build, trạng thái và việc tiếp theo. |
| [docs/tool-release-api-contract.md](tool-release-api-contract.md) | Sửa luồng phát hành giữa app và máy chủ (gói VBA ký, feed Velopack, route admin). |
| [license-sample/api-v1-one-file/README.md](../license-sample/api-v1-one-file/README.md) | Sửa hoặc gắn module bản quyền: cấu hình, giá trị theo sản phẩm, kích hoạt, kiểm tra trước phát hành. |
| [tests/manual/WorkbookSmokeTest.md](../tests/manual/WorkbookSmokeTest.md) | Thử tay app một màn hình với Excel thật: mã đóng gói, các trường hợp bị từ chối, ghép sự kiện. |
| [tests/manual/LicenseGateSmokeTest.md](../tests/manual/LicenseGateSmokeTest.md) | Thử tay cổng bản quyền khi đóng file (hộp "Save changes?") và khi file có hai `Auto_Open`; kết quả đã có. File thử dựng bằng `tests/manual/build-manual-test-files.ps1`. |
| [tests/manual/DeviceCodeActivationSmokeTest.md](../tests/manual/DeviceCodeActivationSmokeTest.md) | Thử tay kích hoạt và cache của module bản quyền (44 ca). |
| [docs/archive/README.md](archive/README.md) | Tra lịch sử: báo cáo cũ đã thay thế. |

## Tài liệu chuẩn ở digital-license-api

Repo `D:\Code\digital-license-api`. Khi khác với repo này, các file dưới đây là chuẩn.

| File | Nội dung |
| --- | --- |
| `docs/DECISIONS.md` — DEC-107 | Cổng bản quyền Excel, sheet ẩn, gia cố client VBA. |
| `docs/DECISIONS.md` — DEC-111 | Phát hành tool-vba qua API; phụ lục 08/11/2026: xuất bản chuyển lên Admin web. |
| `docs/DECISIONS.md` — DEC-104 | Token hết hạn được gia hạn qua validate-device. |
| `docs/DECISIONS.md` — DEC-115 (+ hai phụ lục) | Đóng gói module bản quyền: gói VBA chứa `modLicense*`, mã đóng gói từ Admin, kiểm productKey/phiên bản khi tải lên, `BeforeClose` đặt cuối. |
| `docs/api/v1.md` — Excel device-code activation | Route device-code, redeem, validate-device (hợp đồng của module bản quyền). |
| `docs/api/v1.md` — tool-vba releases | Route phát hành gói VBA và app, kể cả `from-source`. |
| `docs/GIOI-HAN-CHIA-SE-FILE-VBA.md` | VBA hạn chế chia sẻ file được gì, không được gì. |
