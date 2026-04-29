# CLAUDE.md — detech_scanner

Tài liệu này giúp Claude Code hiểu nhanh project khi quay lại.

---

## Mục đích project

Script PowerShell chạy ngầm trên Windows, bắt dữ liệu từ nhiều máy quét mã vạch HID và ghi vào Excel. Mỗi scanner một cột, tự tạo sheet theo tháng.

**File chính:** `src/USB_Reader_HID.ps1` — toàn bộ logic C# inline trong Add-Type block.

---

## Kiến trúc quan trọng

### Hai luồng xử lý song song

```
Barcode scanner (HID keyboard emulation)
    │
    ├─► WH_KEYBOARD_LL hook (KeyboardSuppressor class)
    │       Chặn phím scanner khỏi app đang focus
    │       BUILD BARCODE ← sole source khi hook active
    │       Enqueue vào BarcodeRawInput.Queue
    │
    └─► Windows Raw Input API (BarcodeRawInput.ProcessInput)
            KHI HOOK ACTIVE: chỉ cập nhật _lastActiveDevice, KHÔNG build barcode
            KHI HOOK KHÔNG ACTIVE: build barcode (fallback)
```

**Quan trọng:** Khi `KeyboardSuppressor.IsInstalled == true`, `ProcessInput` bỏ qua toàn bộ buffer building và chỉ set `_lastActiveDevice`. Nếu làm cả hai path build barcode → merged entries khi quét nhanh.

### Cơ chế suppress phím (KeyboardSuppressor)

Hai sub-path dựa vào scanner có "đã biết" hay chưa:

**Scanner đã biết** (`GetLastActiveEncoded() != ""`):
- Char 1 bị giữ (suppress + lưu vào buffer), start WinForms Timer (interval = thresholdMs × 3)
- Char 2 đến nhanh (< detectionMs) → cancel timer, vào scan mode, char 1 KHÔNG bao giờ vào app
- Timer hết → `ReleaseHeldChar()` re-inject char 1 qua `keybd_event` (LLKHF_INJECTED, hook bỏ qua)

**Scanner chưa biết** (lần đầu tiên):
- Char 1 pass through (`CallNextHookEx`), `_firstCharPending = true`
- Char 2 đến nhanh → `GetOrAssignLastActive()` đăng ký device, vào scan mode, `SendBackspace()` xoá char 1 khỏi app
- Backspace được inject với LLKHF_INJECTED → hook bỏ qua (không suppress chính mình)

### Device identification

- `_lastActiveDevice` được set trong `ProcessInput` (Raw Input) mỗi keydown
- `GetLastActiveEncoded()` tra cứu non-creating: nếu device không có trong `_pathToName` → trả về ""
- `GetOrAssignLastActive()` tra cứu creating: tạo entry mới nếu chưa có, enqueue vào `NewDevices`
- `scanner_map.txt` lưu mapping HID path → display name + col index (tab-separated, UTF-8, ẩn)

### Timer (WinForms Timer)

- Chạy trên STA thread của message loop (cùng thread với hook callback) → không cần lock
- Interval = `thresholdMs * 3` = `SCANNER_SPEED_MS × 3` (default 90ms cho 30ms scanner)
- `ResetScan()` luôn stop timer trước khi clear state

---

## Các class C# chính (trong Add-Type block)

| Class | Vai trò |
|-------|---------|
| `ExcelFinder` | Ghi vào Excel đang mở qua COM/Accessibility API |
| `BarcodeRawInput` | Raw Input registration, device lookup, device map |
| `ScannerForm` | WinForms Form để nhận WM_INPUT |
| `KeyboardSuppressor` | WH_KEYBOARD_LL hook, suppress + hold char 1 |

---

## Config

`src/USB_Reader_HID.bat` → truyền params vào script:

```bat
set SCANNER_SPEED_MS=30    :: Ngưỡng phân biệt scanner vs bàn phím (ms)
set MIN_BARCODE_LEN=3      :: Độ dài mã vạch tối thiểu
```

Detection window thực tế = `SCANNER_SPEED_MS × 3` (để bù scanner gửi hơi chậm hơn nominal).

---

## Các bug đã fix (để không fix lại)

### 1. Char 1 leak vào app (VS Code / Electron)
- **Root cause:** Backspace correction có race condition trong Electron — char arrive nhưng Chromium chưa xử lý kịp khi Backspace đến
- **Fix:** Timer-based hold — char 1 không bao giờ vào app với known scanner

### 2. Barcode merged khi quét nhanh
- **Root cause:** Raw Input nhận chars độc lập với hook suppress, `ClearLastActiveBuffer` timing không đảm bảo → sbuf tích lũy qua nhiều scan
- **Fix:** Khi `KeyboardSuppressor.IsInstalled`, `ProcessInput` return sớm sau khi set `_lastActiveDevice`

### 3. Missing leading chars trong barcode
- **Root cause:** Char 1 không được lưu vào `_hookBuf` trước khi `CallNextHookEx`
- **Fix:** Lưu char 1 vào `_hookBuf` TRƯỚC khi return pass-through

### 4. Scan đầu tiên của scanner mới bị ghi vào cùng ô với scan trước
- **Root cause:** `GetOrAssignLastActive()` được gọi ở Enter thay vì char 2 → device registered quá muộn
- **Fix:** Gọi `GetOrAssignLastActive()` ngay tại char 2 detection

### 5. Em-dash parse error
- **Root cause:** `—` (U+2014) trong string PowerShell đọc là Windows-1252 byte 0x94 = right curly quote → string terminator
- **Fix:** Thay bằng ASCII `-`

---

## Luật code quan trọng

- **Không dùng Vietnamese có dấu trong C# string literals** — PowerShell 5.1 đọc UTF-8-without-BOM là Windows-1252, byte 0x94 = right curly quote gây parse error
- `Add-Type` block compile cùng lúc → các class có thể cross-reference nhau (`KeyboardSuppressor` ref `BarcodeRawInput` và ngược lại)
- `System.Windows.Forms.Timer` (không phải `System.Timers.Timer`) — phải chạy trên STA message loop thread
- `keybd_event` inject luôn có `LLKHF_INJECTED` flag → hook check flag này để không suppress chính mình
- `_lastActiveDevice` được set từ Raw Input thread (WM_INPUT) → hook đọc nó trong callback. Cùng 1 thread (STA) nên không cần lock.

---

## Test

```bat
test\Test_Scanner.bat
  [1] Gửi barcode vào tháng hiện tại
  [2] Giả lập qua tháng (tạo 3 sheet)
```

File inject: `src/test_inject.queue` — mỗi dòng là một entry `"name|colIdx\tbarcode"`.
Simulate date: `src/simulate_date.txt` — override tháng hiện tại (format `MM-yyyy`).

---

## Lưu ý khi chạy lại

1. Khi test, xoá `src/scanner_map.txt` để reset device mapping
2. Log: `src/USB_Reader.log` — ghi đè mỗi lần khởi động
3. Script chạy qua `src/USB_Reader_HID.bat` (không chạy .ps1 trực tiếp vì thiếu params)
4. Dừng: `taskkill /f /im powershell.exe` hoặc tạo file `src/stop_signal`
