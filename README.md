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

1. Chuột phải vào `Setup_Autostart.bat`
2. Chọn **"Run as administrator"**
3. Khi hỏi *"Chạy ngay bây giờ không?"* → nhập `y` để test luôn

### Bước 3 — Kiểm tra hoạt động

Quét 1 mã vạch bất kỳ, sau đó mở:

```
src\USB_Reader.log
```

Nếu thấy dòng `Ghi STT ...` → đang hoạt động đúng.

---

## Cách hoạt động

```
Máy quét mã vạch (HID)
      ↓  gửi ký tự như bàn phím + Enter
Windows Raw Input API
      ↓  phân biệt từng thiết bị theo device handle
Nhận diện scanner
      ↓  lưu tên thiết bị + gán cột (scanner_map.txt)
Ghi vào Excel: STT | Thời gian | Scanner A | Scanner B | ...
      ↓  sheet mới tạo tự động khi sang tháng mới
```

**Cấu trúc file Excel:**

| Cột A | Cột B | Cột C (Scanner 1) | Cột D (Scanner 2) |
|-------|-------|-------------------|-------------------|
| STT | Thời gian | T27H | Honeywell |
| 1 | 2026-04-25 08:30:01 | 8935001234567 | |
| 2 | 2026-04-25 08:30:05 | | 8935009876543 |

- Mỗi tháng tạo một sheet mới (tên sheet: `MM-yyyy`, VD: `04-2026`)
- Mỗi scanner được nhận diện tự động và gán vào cột riêng
- Cột mới tự thêm khi có scanner mới kết nối

---

## Test / Giả lập

Chạy `test\Test_Scanner.bat` để giả lập máy quét:

```
[1] Gửi barcode vào tháng hiện tại (script chính phải đang chạy)
[2] Giả lập qua tháng mới (tự động restart script)
```

Hoặc dùng PowerShell trực tiếp:

```powershell
# Gửi barcode vào tháng hiện tại
.\test\Test_Scanner.ps1

# Gửi vào tháng cụ thể
.\test\Test_Scanner.ps1 -Date "04-2026"

# Giả lập chuyển tháng (tạo 2 sheet)
.\test\Test_Scanner.ps1 -TestDateChange
```

---

## Xử lý sự cố

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

### Cột bị lệch khi thêm/bớt scanner

Xóa file `src\scanner_map.txt` để reset ánh xạ (script sẽ tự tạo lại).

### Gỡ tự khởi động

Chuột phải vào `Uninstall_Autostart.bat` → **"Run as administrator"**

---

## Gỡ cài đặt hoàn toàn

1. Chạy `Uninstall_Autostart.bat` với quyền Administrator
2. Xóa toàn bộ thư mục `detech_scanner`
