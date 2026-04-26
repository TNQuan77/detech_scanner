@echo off
setlocal
:: ============================================================
::  Setup_Autostart.bat
::  Dang ky USB_Reader_HID chay tu dong khi dang nhap Windows
::  Dung Registry HKCU\Run — khong can quyen Administrator
::  Tuong thich: Windows 7 / 8 / 10 / 11
:: ============================================================

set TASK_NAME=USB_Barcode_Reader
set BAT_FILE=%~dp0src\USB_Reader_HID.bat
set REG_KEY=HKCU\SOFTWARE\Microsoft\Windows\CurrentVersion\Run

if not exist "%BAT_FILE%" (
    echo [LOI] Khong tim thay: %BAT_FILE%
    pause
    exit /b 1
)

echo ============================================
echo  Dang ky tu dong khoi dong...
echo ============================================

:: Xoa entry cu neu co
reg delete "%REG_KEY%" /v "%TASK_NAME%" /f >nul 2>&1

:: Them vao Registry — chay khi user dang nhap, khong can Admin
reg add "%REG_KEY%" /v "%TASK_NAME%" /t REG_SZ /d "cmd.exe /c \"%BAT_FILE%\"" /f

if %errorLevel% equ 0 (
    echo.
    echo [OK] Dang ky thanh cong!
    echo      Ten     : %TASK_NAME%
    echo      Bat dau : ngay khi dang nhap Windows
    echo      File    : %BAT_FILE%
    echo.
) else (
    echo [LOI] Dang ky that bai.
)

:: Tao file go cai
(
echo @echo off
echo reg delete "%REG_KEY%" /v "%TASK_NAME%" /f
echo echo Da xoa: %TASK_NAME%
echo pause
) > "%~dp0Uninstall_Autostart.bat"

echo [INFO] Da tao Uninstall_Autostart.bat
echo.

set /p RUN_NOW="Chay USB Reader ngay bay gio khong? (y/n): "
if /i "%RUN_NOW%"=="y" (
    start "" /b cmd.exe /c "%BAT_FILE%"
    echo [OK] Da chay ngam. Kiem tra log tai: %~dp0src\USB_Reader.log
)

echo.
pause
endlocal
