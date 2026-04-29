# USB Barcode Reader → Excel

Tự động bắt dữ liệu từ nhiều máy quét mã vạch USB/Bluetooth và ghi vào file Excel — mỗi scanner một cột, tự tạo sheet mới theo tháng.

> Tương thích: CLABEL T27H và hầu hết máy quét mã vạch dạng HID (keyboard emulation)

---

## Yêu cầu

- Windows 7 trở lên
- Microsoft Excel đã được cài đặt
- Máy quét mã vạch kết nối qua USB hoặc Bluetooth (chế độ HID / keyboard emulation)

---

## Cấu trúc file

```
detech_scanner/
├── Setup_Autostart.bat          # Đăng ký tự khởi động cùng Windows (chạy 1 lần)
├── thoi_gian_dong_hang.xlsx     # File Excel dữ liệu (tự tạo khi chạy)
├── src/
│   ├── USB_Reader_HID.ps1       # Logic chính (không cần chỉnh)
│   ├── USB_Reader_HID.bat       # Launcher — chỉnh cấu hình ở đây
│   ├── scanner_map.txt          # Ánh xạ thiết bị → cột (tự tạo, ẩn)
│   └── USB_Reader.log           # File log (tự tạo khi chạy)
└── test/
    ├── Test_Scanner.ps1         # Giả lập máy quét để test
    └── Test_Scanner.bat         # Launcher cho test
```

---

## Hướng dẫn cài đặt

### Bước 1 — Cấu hình (tuỳ chọn)

Mở file `src\USB_Reader_HID.bat` bằng Notepad, chỉnh nếu cần:

```bat
set SCANNER_SPEED_MS=100    :: Tốc độ phân biệt scanner vs bàn phím (ms)
set MIN_BARCODE_LEN=3       :: Độ dài mã vạch tối thiểu
```

### Bước 2 — Đăng ký tự khởi động

1. Double-click vào `Setup_Autostart.bat` (không cần Run as administrator)
2. Khi hỏi *"Chạy ngay bây giờ không?"* → nhập `y` để chạy luôn

Script sẽ chạy **ẩn hoàn toàn** (không có cửa sổ) mỗi khi đăng nhập Windows.

> Tương thích Windows 10 và Windows 11. Dùng Registry HKCU\Run — không cần quyền Administrator.

### Bước 3 — Kiểm tra hoạt động

Cắm scanner, quét 1 mã vạch bất kỳ, sau đó mở `src\USB_Reader.log`:

- Thấy dòng `Dang lang nghe ma vach` → script đang chạy
- Thấy dòng `Keyboard suppress: bat` → tính năng chặn phím scanner đang hoạt động
- Thấy dòng `Ghi STT ...` → barcode đã được ghi vào Excel

> **Lưu ý:** Scanner mới cần **quét ít nhất 1 mã vạch** thì mới được nhận diện và gán cột. Chỉ cắm vào chưa đủ.

---

## Cách hoạt động

```
Máy quét mã vạch (HID)
      ↓  gửi ký tự như bàn phím + Enter
WH_KEYBOARD_LL hook (KeyboardSuppressor)
      ↓  chặn phím scanner, không cho hiện vào app đang focus
Windows Raw Input API
      ↓  phân biệt từng thiết bị theo device handle
Nhận diện scanner
      ↓  lưu tên thiết bị + gán cột (scanner_map.txt)
Ghi vào Excel: STT | Thời gian | Scanner A | Scanner B | ...
      ↓  sheet mới tạo tự động khi sang tháng mới
```

**Cơ chế chặn phím scanner (Keyboard Suppressor)**

Phím từ máy quét sẽ không hiện vào ứng dụng đang mở (Word, Notepad, trình duyệt…). Script dùng Windows low-level keyboard hook để phát hiện và chặn chuỗi ký tự của scanner dựa vào tốc độ gõ:

- **Scanner đã biết** (đã quét ít nhất 1 lần): ký tự đầu bị giữ lại (không vào app), timer `SCANNER_SPEED_MS × 3` bắt đầu đếm
  - Ký tự thứ 2 đến nhanh → xác nhận scanner → chặn toàn bộ, ký tự đầu không bao giờ vào app
  - Timer hết giờ mà không có ký tự thứ 2 → tái phát ký tự đầu vào app (người dùng gõ tay)
- **Scanner chưa biết** (lần quét đầu tiên): ký tự đầu được thả qua, ký tự thứ 2 đến nhanh → chặn + xoá ký tự đầu bằng Backspace

> Phím người dùng gõ tay (chậm hơn ngưỡng) vẫn hoạt động bình thường.

**Cấu trúc file Excel** (`thoi_gian_dong_hang.xlsx` ở thư mục gốc):

| Cột A | Cột B | Cột C (Scanner 1) | Cột D (Scanner 2) |
|-------|-------|-------------------|-------------------|
| STT | Thời gian | T27H | Honeywell |
| 1 | 2026-04-25 08:30:01 | 8935001234567 | |
| 2 | 2026-04-25 08:30:05 | | 8935009876543 |

- Mỗi tháng tạo một sheet mới (tên sheet: `MM-yyyy`, VD: `04-2026`)
- Sheet sắp xếp tăng dần: tháng cũ bên trái, tháng mới bên phải
- Mỗi scanner được nhận diện tự động theo tên thiết bị và gán vào cột riêng
- Cột mới tự thêm khi có scanner mới kết nối và quét lần đầu
- Ghi được ngay kể cả khi file Excel đang mở

---

## Test / Giả lập

Chạy `test\Test_Scanner.bat` để giả lập máy quét:

```
[1] Gửi barcode vào tháng hiện tại (script chính phải đang chạy)
[2] Giả lập qua tháng (script chính vẫn chạy, tạo 3 sheet: tháng trước | tháng này | tháng sau)
```

Hoặc dùng PowerShell trực tiếp:

```powershell
# Gửi barcode vào tháng hiện tại
.\test\Test_Scanner.ps1

# Gửi vào tháng cụ thể
.\test\Test_Scanner.ps1 -Date "04-2026"

# Giả lập chuyển tháng (tạo 3 sheet: tháng trước | tháng này | tháng sau)
.\test\Test_Scanner.ps1 -TestDateChange
```

---

## Xử lý sự cố

### Kiểm tra script có đang chạy không

Mở **Task Manager** → tab **Details** → tìm tiến trình `powershell.exe`. Hoặc xem `src\USB_Reader.log` — nếu có dòng `Dang lang nghe` là đang chạy.

### Dừng script thủ công

Mở **Task Manager** → tab **Details** → chuột phải `powershell.exe` → **End task**.  
Hoặc chạy lệnh:
```bat
taskkill /f /im powershell.exe
```

### Mã vạch không được ghi vào Excel

- Kiểm tra `src\USB_Reader.log` xem có lỗi không
- Đảm bảo Microsoft Excel đã được cài đặt
- Thử chạy thẳng `src\USB_Reader_HID.bat` để xem lỗi

### Phím bàn phím bị nhận nhầm là mã vạch

Mở `src\USB_Reader_HID.bat`, giảm giá trị `SCANNER_SPEED_MS`:

```bat
set SCANNER_SPEED_MS=50
```

### Máy quét bị bỏ qua (không ghi)

Mở `src\USB_Reader_HID.bat`, tăng giá trị `SCANNER_SPEED_MS`:

```bat
set SCANNER_SPEED_MS=150
```

### Ký tự máy quét vẫn rớt vào ứng dụng

Đây là hành vi bình thường ở **lần quét đầu tiên của scanner mới** — ký tự đầu sẽ xuất hiện rồi bị xoá ngay bằng Backspace. Từ lần quét thứ 2 trở đi, scanner đã được nhận diện và ký tự đầu bị giữ lại hoàn toàn (không bao giờ vào app).

Nếu vẫn rớt ký tự ở lần quét đầu tiên dù đã xoá `scanner_map.txt`, thử tăng `SCANNER_SPEED_MS`:

```bat
set SCANNER_SPEED_MS=150
```

### Cột bị lệch khi thêm/bớt scanner

Xóa file `src\scanner_map.txt` để reset ánh xạ (script sẽ tự tạo lại khi quét lần đầu).

### Gỡ tự khởi động

Double-click vào `Uninstall_Autostart.bat` (không cần Run as administrator)

---

## Gỡ cài đặt hoàn toàn

1. Chạy `Uninstall_Autostart.bat` với quyền Administrator
2. Xóa toàn bộ thư mục `detech_scanner`
