# Bàn giao: tool-vba đóng gói bản quyền (DEC-115), còn bước B5

Cập nhật 08/11/2026. Chưa push nhánh nào. Mỗi commit cần chủ dự án đồng ý trước.

## Nhánh và commit

**tool-vba** `D:\Code\tool-vba`, nhánh `feat/license-packaging` (tạo từ `feat/license-device-name`):

| Commit | Nội dung |
| --- | --- |
| `df565c0` | B1: `LicenseProduct` (mã đóng gói JSON / nhập tay), `LicenseModuleConfigurator` (điền 3 dòng Cfg), `VbaStringLiteral` (ChrW) |
| `dbdc777` | B2: ghép sự kiện bản quyền vào handler có sẵn (`DocumentModuleMerger`), `ThisWorkbook-events.txt` BeforeSave `Cancel = True: Exit Sub` |
| `7a774d9` | B3+B4: quy trình Tạo mới (module bản quyền, sheet "Huong dan", ẩn VeryHidden, kiểm lại), thao tác sheet qua COM, giao diện một file + ô mã đóng gói, tên `<productKey>_<version>_ban-quyen.xlsm`, BeforeClose đặt CUỐI `If Not Cancel Then …` |
| `18e9d44` | Kiểm sau lưu chặt hơn, từ chối module ngoài danh sách có `Cfg_ProductKey`/`License_*`, xóa trạng thái "Xong…"; README mục "Để đợt sau" |

**digital-license-api** `D:\Code\digital-license-api`, nhánh `phase/license-packaging` (tạo từ `phase/device-management`):

| Commit | Nội dung |
| --- | --- |
| `480aba2` | DEC-115 |
| `847f94e`, `813bfcf` | Gói VBA chứa `modLicense*` (bỏ luật bỏ qua), fixture `fixture-license.xlsm`, `fixture-sheet-code.xlsm`, docs |
| `e8d5886`, `a3dcc1d` | Admin: nút "Sao chép mã đóng gói" cấp sản phẩm; nhãn "người đóng gói" |
| `b60cb3d` | Kiểm `Cfg_ProductKey`/`Cfg_ProductVersion` khi tạo phiên bản Excel và khi PATCH đổi số (400 `PRODUCT_VERSION_LICENSE_MISMATCH`), đọc chỉ `vbaProject.bin`, chống ReDoS |
| `cc9faf3` | DEC-115 phụ lục 2: BeforeClose đặt cuối |

Kết quả lần cuối: tool-vba `dotnet test ToolVba.sln` 247/247; selftest Excel thật 24/24; API `pnpm verify` xanh (1477 API, 311 script/Admin).

## Việc còn lại (B5)

1. `scripts/build-license-template.ps1` trong tool-vba: dựng `.xlsm` mẫu để tải lên Admin web (Phát hành tool-vba → Xuất bản code VBA) từ `license-sample/api-v1-one-file/modLicenseApiV1.bas` + `ThisWorkbook-events.txt` (bỏ dòng chú thích). **Ba dòng Cfg phải để trống** (`vbNullString`, `Array()`); kiểm lại sau khi lưu. Đổi tên module workbook qua `VBComponent.Properties("_CodeName") = "ThisWorkbook"`. Lưu vào `out\license-template\` (`out/` đã gitignore). Mẫu cách làm: `C:\Users\duyla\AppData\Local\Temp\claude\40810030eb4d9c13\3926dbc0-77f8-4ffd-adc4-2053a39a14ad\scratchpad\a2-real\make-real-template.ps1`.
2. Excel thật, ca hai `Auto_Open` (module bản quyền + một module khác): xem Excel làm gì (dùng `RunAutoMacros 1` hoặc mở có giao diện). App đã từ chối trường hợp này; ghi kết quả vào README module.
3. Ca đóng file (chủ dự án yêu cầu ghi vào README module `license-sample/api-v1-one-file/README.md`):
   - **Đã kiểm, ĐẠT:** sản phẩm có `Workbook_BeforeClose` đặt `Cancel = True` (bấm Hủy), đóng bằng lệnh giao diện `CommandBars.ExecuteMso("FileClose")` trong Excel ẩn, macro bật: file vẫn mở, sheet bảo vệ "Du lieu" vẫn hiện, `License_OnClose` không chạy (cả qua hook lẫn qua `Auto_Close` của module). Script: `…\scratchpad\b5-close\close-cancel.ps1`. Chưa chạy ca đối chứng (không hủy).
   - **Chưa kiểm, cần chủ dự án làm tay:** bấm Cancel ở hộp "Save changes?" của Excel. Đã biết trước: `BeforeClose` chạy trước hộp này, nên `License_OnClose` có thể đã chạy (ẩn sheet, xóa quyền trong phiên) trong khi file vẫn mở. Chuẩn bị file thử + các bước cho chủ dự án, ghi kết quả.
   - Lưu ý: module có `Auto_Close` cũng gọi `License_OnClose`; đóng qua `Workbook.Close` (COM/VBA) không chạy `Auto_Close`.
4. Cập nhật tài liệu: README module (trạng thái, mục "Gắn vào sản phẩm" giờ do app làm, kết quả các ca trên), README gốc (cách dùng mới: một file, mã đóng gói, tên file, các trường hợp bị từ chối; mục "Trạng thái và việc tiếp theo"), `docs/README.md` nếu thêm/đổi file .md, `docs/tool-release-api-contract.md` nếu cần, `tests/manual/WorkbookSmokeTest.md` (luồng mới).
5. Bộ kiểm thử đầy đủ: `powershell -NoProfile -File .\contract-check.ps1` (tại `license-sample/api-v1-one-file`), `node .\tests\build-excel-regression.mjs D:\Code\digital-license-api` rồi chạy 57 ca `License_RunSignedRegressionTests` trong Excel, `License_RunMockTests` 64/64 trong Excel (MsgBox cuối phải đổi thành ghi ô như `build.ps1` cũ để Excel ẩn không treo), `dotnet test ToolVba.sln`, build riêng app (`dotnet test` không build app WPF), selftest `dotnet run --project tests\ToolVba.IntegrationRunner -- --selftest <thư mục trống>`.
6. Trước khi thử với `test-01`: hỏi chủ dự án xác nhận `productKey` thật trên Admin web.
7. Review code (csharp-reviewer) cho phần mới, sửa CRITICAL/HIGH, báo danh sách file trước khi commit.

## Quyết định và ràng buộc phải giữ

- DEC-115 (+ phụ lục 1, 2) trong `D:\Code\digital-license-api\docs\DECISIONS.md` là nguồn chuẩn. Memory: `tool-vba-packaging-rework-decisions`.
- BeforeSave/AfterSave: lời gọi ở ĐẦU handler có sẵn; BeforeClose: ở CUỐI, `If Not Cancel Then …`; BeforeClose có `Exit Sub`/`Exit Function`/`End` thì từ chối.
- Không commit `.vscode/`, không đẩy `Mau_code_VBA.xlsm`, không đọc `.env`/khóa riêng, không chạy migration production, không push khi chưa được đồng ý.
- Hook "Fact-Forcing Gate" giữ nguyên (chủ dự án quyết): mỗi lần tạo/sửa file lần đầu phải nêu dữ kiện rồi thử lại.
- Excel ẩn gặp lỗi biên dịch VBA hoặc hộp thoại sẽ treo: chạy ở nền, ghi PID (`GetWindowThreadProcessId(Hwnd)`), nếu treo chỉ dừng đúng PID đó. Không đụng Excel/tiến trình node của chủ dự án.
- Repo API: chạy pnpm từ thư mục repo (`Push-Location D:\Code\digital-license-api; corepack pnpm run verify`), không dùng `--dir` từ thư mục khác (corepack lấy nhầm bản pnpm). Prettier chỉ cho file mình sửa.
- Module `.bas` thuần ASCII; sửa module thì cập nhật SHA-256 trong `modLicenseApiV1.manifest.json`.
- Excel đổi kiểu hoa/thường tên biến khi lưu dự án VBA: so tên không phân biệt hoa thường.
- Repo API theo AGENTS.md: test trước, TEST-CASES, PROGRESS, DECISIONS; test bị quyết định đảo ngược thì đổi cùng commit với code (DEC-115 mục 7).

## Để đợt sau (đã ghi trong README tool-vba)

Test ViewModel; thông báo sheet macro Excel 4/hộp thoại bị gọi là chart sheet; câu khuyên khi lỗi COM; từ chối nhầm `Exit Sub`/`GoTo` trong chú thích hoặc chuỗi.
