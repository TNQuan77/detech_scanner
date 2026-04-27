@echo off
setlocal

:: ============================================================
::  USB_Reader_HID.bat
::  Danh cho may quet ma vach HID (keyboard emulation)
::  Vi du: CLABEL T27H, Honeywell, Zebra, Datalogic, v.v.
:: ============================================================

:: ---------- CAU HINH ----------
set EXCEL_FILE=%~dp0..\thoi_gian_dong_hang.xlsx
set LOG_FILE=%~dp0USB_Reader.log

:: Nguong toc do (ms): phim den trong vong bao nhieu ms = may quet (khong phai nguoi gõ)
:: Giam xuong neu bi bat nham phim ban phim, tang len neu bo qua ma vach
set SCANNER_SPEED_MS=30

:: Do dai ma vach toi thieu (bo qua cac phim le / ngan hon)
set MIN_BARCODE_LEN=3

:: AutoHotkey — suppress keystroke tu scanner vao cac app khac
set AHK_EXE=%~dp0..\bin\AutoHotkey64.exe
set AHK_SCRIPT=%~dp0..\suppress_scanner.ahk
:: --------------------------------

set SCRIPT=%~dp0USB_Reader_HID.ps1

if not exist "%SCRIPT%" (
    echo [LOI] Khong tim thay: %SCRIPT%
    pause
    exit /b 1
)

:: Khoi dong AHK suppress (neu co san, bo qua neu chua cai)
if exist "%AHK_EXE%" if exist "%AHK_SCRIPT%" (
    start "" /b "%AHK_EXE%" "%AHK_SCRIPT%"
)

powershell.exe -NoProfile -STA -ExecutionPolicy Bypass -WindowStyle Hidden ^
    -File "%SCRIPT%" ^
    -ExcelFile "%EXCEL_FILE%" ^
    -LogFile "%LOG_FILE%" ^
    -ScannerSpeedMs %SCANNER_SPEED_MS% ^
    -MinBarcodeLength %MIN_BARCODE_LEN%

endlocal
