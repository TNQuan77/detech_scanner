#Requires AutoHotkey v2.0
#SingleInstance Force

; ════════════════════════════════════════════════════════════════════
;  suppress_scanner.ahk
;  Chặn barcode scanner HID gõ vào app đang focus
;  Phân biệt scanner vs bàn phím bằng tốc độ gõ:
;    scanner  → gõ liên tục < SCANNER_MS ms/phím → suppress
;    bàn phím → gõ chậm hơn                      → pass through
;
;  Script USB_Reader_HID.ps1 vẫn nhận barcode qua Raw Input riêng,
;  không bị ảnh hưởng bởi suppress ở đây.
; ════════════════════════════════════════════════════════════════════

SCANNER_MS := 50     ; ms giữa 2 phím liên tiếp — nhanh hơn = scanner
IDLE_MS    := 300    ; ms không có phím → reset về chế độ bàn phím

global _t    := 0
global _scan := false

; ── Handler chung ────────────────────────────────────────────────
_k(ch, passthru, *) {
    global _t, _scan, SCANNER_MS, IDLE_MS
    now := A_TickCount
    gap := _t ? now - _t : 99999
    _t  := now

    if (gap < SCANNER_MS || (_scan && gap < IDLE_MS)) {
        _scan := true
        return                  ; suppress — không pass qua app
    }
    if (_scan)
        _scan := false
    SendInput passthru          ; bàn phím thường → pass through
}

_e(*) {     ; Enter — kết thúc barcode hoặc Enter bình thường
    global _scan, _t
    wasScan := _scan
    _scan := false
    _t    := 0
    if (!wasScan)
        SendInput "{Enter}"     ; Enter bàn phím → pass through
                                ; Enter scanner  → suppress (PS đã nhận qua Raw Input)
}

; Điều kiện: chỉ bắt phím khi Ctrl / Alt / Win KHÔNG được giữ
_noMod(*) {
    return !GetKeyState("Ctrl") && !GetKeyState("Alt")
        && !GetKeyState("LWin") && !GetKeyState("RWin")
}

; ── Đăng ký hotkeys ──────────────────────────────────────────────
; → Ctrl+C, Alt+F4, Win+D, v.v. hoàn toàn không bị ảnh hưởng

HotIf(_noMod)

; Chữ số 0-9
Loop 10 {
    c := String(A_Index - 1)
    HotKey "*" c, _k.Bind(c, c)
}

; Chữ cái thường a-z và hoa A-Z (scanner thường gửi hoa qua Shift+letter)
Loop 26 {
    lo := Chr(96 + A_Index)        ; a … z
    up := Chr(64 + A_Index)        ; A … Z
    HotKey "*"  lo, _k.Bind(lo, lo)
    HotKey "*+" lo, _k.Bind(up, "+" lo)
}

; Ký tự đặc biệt phổ biến trong barcode
HotKey "*-",     _k.Bind("-",  "-")
HotKey "*.",     _k.Bind(".",  ".")
HotKey "*/",     _k.Bind("/",  "/")
HotKey "*Space", _k.Bind(" ",  "{Space}")
HotKey "*+=",    _k.Bind("+",  "+{=}")    ; Shift+= → +
HotKey "*+-",    _k.Bind("_",  "+{-}")    ; Shift+- → _
HotKey "*+;",    _k.Bind(":",  "+{;}")    ; Shift+; → :

; Enter
HotKey "*Enter", _e

HotIf()  ; reset — các hotkey sau (nếu có) không bị điều kiện trên
