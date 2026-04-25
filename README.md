# USB Barcode Reader → Excel

Tự động bắt dữ liệu từ máy quét mã vạch USB và ghi vào file Excel.

> Tương thích: CLABEL T27H và hầu hết máy quét mã vạch USB dạng HID (keyboard emulation)

---

## Yêu cầu

- Windows 7 trở lên
- Microsoft Excel đã được cài đặt
- Máy quét mã vạch kết nối qua USB (chế độ HID / keyboard emulation)

---

## Cấu trúc file

```
detech_scanner/
├── USB_Reader_HID.ps1      # Logic chính (không cần chỉnh)
├── USB_Reader_HID.bat      # Launcher — chỉnh cấu hình ở đây
├── Setup_Autostart.bat     # Đăng ký tự khởi động cùng Windows
├── Uninstall_Autostart.bat # Gỡ tự khởi động (tự tạo sau khi chạy Setup)
├── ABC.xlsx                # File Excel dữ liệu (tự tạo khi chạy)
└── USB_Reader.log          # File log (tự tạo khi chạy)
```

---

## Hướng dẫn cài đặt

### Bước 1 — Cấu hình

Mở file `USB_Reader_HID.bat` bằng Notepad, chỉnh các dòng sau:

```bat
set EXCEL_FILE=%~dp0ABC.xlsx        :: Lưu cùng thư mục script (mặc định)
set LOG_FILE=%~dp0USB_Reader.log    :: Log cùng thư mục script (mặc định)
set SCANNER_SPEED_MS=50             :: Tốc độ phân biệt scanner vs bàn phím (ms)
set MIN_BARCODE_LEN=3               :: Độ dài mã vạch tối thiểu
```

### Bước 2 — Đăng ký tự khởi động

1. Chuột phải vào `Setup_Autostart.bat`
2. Chọn **"Run as administrator"**
3. Làm theo hướng dẫn trên màn hình
4. Khi hỏi *"Chạy ngay bây giờ không?"* → nhập `y` để test luôn

### Bước 3 — Kiểm tra hoạt động

Quét 1 mã vạch bất kỳ, sau đó mở file log để kiểm tra:

```
detech_scanner\USB_Reader.log
```

Nếu thấy dòng `Barcode: ...` → đang hoạt động đúng.

---

## Cách hoạt động

```
Máy quét mã vạch
      ↓  quét mã → gửi ký tự như bàn phím + Enter
Script bắt phím toàn hệ thống
      ↓  đo thời gian giữa các phím
Nếu mỗi phím < 50ms → là scanner (người gõ chậm hơn nhiều)
      ↓
Ghi vào Excel:  STT | Thời gian | Mã vạch
```

**Cấu trúc file Excel:**

| Cột A | Cột B | Cột C |
|-------|-------|-------|
| STT | Thời gian | Mã vạch |
| 1 | 2026-04-25 08:30:01 | 8935001234567 |
| 2 | 2026-04-25 08:30:05 | 8935009876543 |

---

## Xử lý sự cố

### Mã vạch không được ghi vào Excel

- Kiểm tra file log `USB_Reader.log` trong thư mục script xem có lỗi không
- Đảm bảo Microsoft Excel đã được cài đặt
- Thử chạy thẳng `USB_Reader_HID.bat` (không ẩn) để xem lỗi

### Phím bàn phím bị nhận nhầm là mã vạch

Mở `USB_Reader_HID.bat`, giảm giá trị `SCANNER_SPEED_MS`:

```bat
set SCANNER_SPEED_MS=30   :: giảm xuống 30ms
```

### Máy quét bị bỏ qua (không ghi)

Mở `USB_Reader_HID.bat`, tăng giá trị `SCANNER_SPEED_MS`:

```bat
set SCANNER_SPEED_MS=80   :: tăng lên 80ms
```

### Gỡ tự khởi động

Chuột phải vào `Uninstall_Autostart.bat` → **"Run as administrator"**

---

## Gỡ cài đặt hoàn toàn

1. Chạy `Uninstall_Autostart.bat` với quyền Administrator
2. Xóa toàn bộ thư mục `detech_scanner` (bao gồm luôn Excel và log)
