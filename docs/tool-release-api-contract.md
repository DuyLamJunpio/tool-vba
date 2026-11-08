# Hợp đồng API phát hành tool-vba (v1)

Tài liệu này là nguồn chuẩn chung cho `tool-vba` (client WPF) và `digital-license-api` (server). Mọi route có tiền tố `/api/v1`. JSON dùng camelCase, UTF-8.

**Envelope:** API bọc mọi response thành công bằng `TransformResponseInterceptor`: `{ "success": true, "statusCode": 200, "data": <payload>, "timestamp": "…" }`. Response lỗi có dạng `{ "success": false, "statusCode": 4xx, "message": "…" | ["…"], "code"?: "…" }`. Các mẫu JSON dưới đây là phần `data`.

Có hai ngoại lệ không dùng envelope:
- `GET /tool/app/releases.win.json` trả feed nguyên văn, vì Velopack đọc trực tiếp.
- Response 302/304 không có body.

## 1. Gói code VBA (`VBA_PACKAGE`)

### 1.1 Định dạng file zip

- Zip phẳng: mọi entry nằm ở gốc, không có thư mục, không có `..`, `/`, `\`, ký tự `:`. Tên entry khớp `^[A-Za-z0-9_.-]{1,64}$`.
- Kích thước zip ≤ 5 MB; tổng dung lượng giải nén ≤ 20 MB; tối đa 200 entry.
- Bắt buộc có `package.json`:

```json
{
  "formatVersion": 1,
  "components": [
    { "name": "modBaoCao", "kind": "standard", "file": "modBaoCao.bas", "sha256": "<64 hex thường>" },
    { "name": "clsDongDuLieu", "kind": "class", "file": "clsDongDuLieu.cls", "sha256": "…" },
    { "name": "frmNhap", "kind": "form", "file": "frmNhap.frm", "sha256": "…", "extraFile": "frmNhap.frx", "extraSha256": "…" },
    { "name": "ThisWorkbook", "kind": "workbook", "file": "ThisWorkbook.txt", "sha256": "…" },
    { "name": "Sheet1", "kind": "sheet", "file": "Sheet1.txt", "sha256": "…" }
  ]
}
```

| `kind` | Đuôi `file` | Ghi chú |
| --- | --- | --- |
| `standard` | `.bas` | Module chuẩn |
| `class` | `.cls` | Class module |
| `form` | `.frm` | Bắt buộc có `extraFile` `.frx` và `extraSha256` |
| `workbook` | `.txt` | Code của module ThisWorkbook, dạng UTF-8; tối đa 1 component |
| `sheet` | `.txt` | Code của module sheet theo CodeName, dạng UTF-8 |

- `name` khớp `^\p{L}[\p{L}\p{Nd}_]{0,30}$` (cho phép chữ có dấu, ví dụ `Trang_tính1`) và duy nhất (không phân biệt hoa thường). Có ít nhất 1 component.
- Component `workbook` luôn có `name` là `ThisWorkbook`, kể cả khi Excel tiếng Việt đặt tên khác (`SổlàmviệcNày`).
- Component `sheet` có thể có thêm `sheetName` (tên tab sheet, ≤ 31 ký tự). Khi CodeName của file đích khác, client dùng trường này để khớp sheet. Component loại khác không được có `sheetName`.
- Tên file trong zip vẫn chỉ dùng ký tự ASCII theo regex ở trên; tên file có thể khác `name` (ví dụ `Sheet1.txt` cho component `Trang_tính1`).
- `sha256` là SHA-256 của đúng byte trong file, viết bằng hex chữ thường.
- Server từ chối gói chứa chuỗi giống private key hoặc secret (`-----BEGIN`, `PRIVATE KEY`, …).
- `package.json` được đọc chặt chẽ ở cả hai phía: tên trường phân biệt hoa thường; ở cấp gốc chỉ có `formatVersion` và `components`; mỗi component chỉ có các trường liệt kê ở trên.
- Tên file không được là tên thiết bị của Windows (`CON`, `PRN`, `AUX`, `NUL`, `COM1`–`COM9`, `LPT1`–`LPT9`, có hoặc không có phần mở rộng) và không được kết thúc bằng dấu chấm.

### 1.2 Chữ ký

- Thuật toán: ECDSA P-256 + SHA-256. Chữ ký ở định dạng IEEE P1363 (`r||s`, 64 byte), mã hóa base64.
- Node: `crypto.sign('sha256', Buffer.from(payload, 'utf8'), { key, dsaEncoding: 'ieee-p1363' })`.
- Payload (UTF-8, không có xuống dòng): `VBA_PACKAGE|{version}|{sha256}|{minAppVersion}`. Nếu không có `minAppVersion` thì để chuỗi rỗng, ví dụ `VBA_PACKAGE|2026.11.7.1|ab12…|`.
- Khóa riêng lấy từ env `TOOL_RELEASE_SIGNING_PRIVATE_KEY` (PEM PKCS#8). Client chỉ nhúng public key (SPKI PEM).

### 1.3 `GET /tool/vba-packages/latest` — public

`@Public()` vì code VBA vốn đọc được trong mọi file `.xlsm` đã phát hành; tính toàn vẹn được bảo vệ bằng chữ ký. Giới hạn tốc độ theo IP.

- 200:

```json
{
  "version": "2026.11.7.1",
  "sha256": "<64 hex thường của file zip>",
  "sizeBytes": 18036,
  "minAppVersion": "1.0.0",
  "releaseNotes": "Sửa định dạng báo cáo",
  "publishedAt": "2026-11-07T03:12:45.000Z",
  "signature": "<base64>",
  "downloadUrl": "https://…signed…"
}
```

  - Có header `ETag: "<sha256>"`. `minAppVersion` và `releaseNotes` có thể là `null`.
  - `downloadUrl` là URL `https` tuyệt đối, có hiệu lực ít nhất 5 phút. Client chỉ chấp nhận `http` khi chính API chạy trên localhost (môi trường dev).
  - Toàn bộ response ≤ 64 KiB.
- Request có `If-None-Match` bằng ETag hiện tại → trả 304, không có body.
- Chưa có bản active nào → 404.

### 1.4 `POST /admin/tool/vba-packages` — JWT, vai trò ADMIN / SUPER_ADMIN

- Body multipart: `file` (zip, bắt buộc), `releaseNotes` (tùy chọn, ≤ 2000 ký tự), `minAppVersion` (tùy chọn, dạng `x.y.z`).
- Server làm các bước sau:
  1. Kiểm tra zip theo §1.1.
  2. Tính SHA-256 của zip.
  3. Cấp `version = YYYY.M.D.N` theo ngày ở múi giờ Asia/Ho_Chi_Minh, với N = 1, 2, … trong ngày. Version mới luôn lớn hơn các version đã có.
  4. Ký theo §1.2, lưu file vào storage.
  5. Đặt bản mới làm active (bản cũ thành inactive), ghi audit log.
- 201: cùng dạng §1.3 nhưng thêm `id` và không có `downloadUrl`.
- 400: zip không hợp lệ, `message` tiếng Việt nêu lý do cụ thể.

## 2. Ứng dụng desktop (`DESKTOP_APP`, feed Velopack)

### 2.1 `GET /tool/app/releases.win.json` — public

Trả nguyên văn feed `releases.win.json` của bản active (`Content-Type: application/json`). Bỏ qua mọi query string. Chưa có bản nào → 404.

### 2.2 `GET /tool/app/:fileName` — public

- `fileName` phải khớp `^[A-Za-z0-9._-]{1,200}$`.
- Server tìm file theo thứ tự: trong `assets` của bản active trước, rồi các bản cũ hơn, mới nhất trước.
- Tìm thấy → 302 tới signed URL (≥ 5 phút). Không có → 404.
- Link tải bộ cài lần đầu: `/api/v1/tool/app/ToolVba-win-Setup.exe`.

### 2.3 `POST /admin/tool/app-releases` — JWT, vai trò ADMIN / SUPER_ADMIN

- Body multipart:
  - `version` (bắt buộc, dạng `x.y.z`)
  - `releaseNotes` (tùy chọn)
  - `feed` (file `releases.win.json`)
  - `files` (một hoặc nhiều file: `*.nupkg`, `*-Setup.exe`, …; mỗi file ≤ 200 MB)
- Kiểm tra:
  - `feed` là JSON có mảng `Assets`.
  - Mọi asset có `Version == version` phải có file tương ứng trong `files`.
  - Tên file khớp regex ở §2.2.
- Server lưu storage key `tool-app/{version}/{releaseId}/{fileName}` (mỗi lần tải lên một thư mục riêng, nên hai lần phát hành cùng version không ghi đè nhau), lưu `feedJson` và danh sách `assets` (`fileName`, `sha256`, `sizeBytes`), đặt bản này làm active, ghi audit log.
- 201: `{ id, version, publishedAt }`.

## 3. Quản trị chung

- `GET /admin/tool/releases?channel=VBA_PACKAGE|DESKTOP_APP`: danh sách phân trang (mặc định 20, tối đa 100), gồm `id, channel, version, sha256, sizeBytes, isActive, releaseNotes, publishedAt`.
- `POST /admin/tool/releases/:id/activate`: đặt một bản cũ thành active để rollback, ghi audit log.

## 4. Xuất bản từ Admin web

App desktop không còn đăng nhập hay xuất bản. Admin web gọi:

- `POST /admin/tool/vba-packages/from-source` (JWT ADMIN): multipart `files` (một `.xlsm` và/hoặc `.bas`, `.cls`, `.frm` + `.frx`), `releaseNotes?`, `minAppVersion?`, `dryRun` (`"true"` chỉ phân tích). Máy chủ tách code từ `xl/vbaProject.bin` theo [MS-OVBA], dựng gói đúng định dạng §1.1 rồi đi qua cùng bước kiểm tra, ký và lưu như §1.4. Code của ThisWorkbook/sheet có thể còn dòng đánh dấu `'=== tool-vba …`; client tự bỏ các dòng này khi dán.
- `POST /admin/tool/app-releases` (§2.3) với các file do `scripts/build-app.ps1` tạo.
