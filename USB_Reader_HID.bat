@echo off
setlocal

:: ============================================================
::  USB_Reader_HID.bat
::  Danh cho may quet ma vach HID (keyboard emulation)
::  Vi du: CLABEL T27H, Honeywell, Zebra, Datalogic, v.v.
:: ============================================================

:: ---------- CAU HINH ----------
set EXCEL_FILE=%~dp0ABC.xlsx
set LOG_FILE=%~dp0USB_Reader.log

:: Nguong toc do (ms): phim den trong vong bao nhieu ms = may quet (khong phai nguoi gõ)
:: Giam xuong neu bi bat nham phim ban phim, tang len neu bo qua ma vach
@REM set SCANNER_SPEED_MS=50
set SCANNER_SPEED_MS=100

:: Do dai ma vach toi thieu (bo qua cac phim le / ngan hon)
set MIN_BARCODE_LEN=3
:: --------------------------------

set SCRIPT=%~dp0USB_Reader_HID.ps1

if not exist "%SCRIPT%" (
    echo [LOI] Khong tim thay: %SCRIPT%
    pause
    exit /b 1
)

powershell.exe -NoProfile -STA -ExecutionPolicy Bypass -WindowStyle Hidden ^
    -File "%SCRIPT%" ^
    -ExcelFile "%EXCEL_FILE%" ^
    -LogFile "%LOG_FILE%" ^
    -ScannerSpeedMs %SCANNER_SPEED_MS% ^
    -MinBarcodeLength %MIN_BARCODE_LEN%

endlocal
