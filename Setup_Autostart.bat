@echo off
setlocal
:: ============================================================
::  Setup_Autostart.bat
::  Chay MOT LAN voi quyen Administrator
::  Dang ky USB_Reader_HID chay tu dong khi dang nhap Windows
:: ============================================================

net session >nul 2>&1
if %errorLevel% neq 0 (
    echo [!] Can quyen Administrator.
    echo Chuot phai vao file nay ^> "Run as administrator"
    pause
    exit /b 1
)

set TASK_NAME=USB_Barcode_Reader
set BAT_FILE=%~dp0USB_Reader_HID.bat

if not exist "%BAT_FILE%" (
    echo [LOI] Khong tim thay: %BAT_FILE%
    pause
    exit /b 1
)

echo ============================================
echo  Dang ky Task Scheduler...
echo ============================================

schtasks /delete /tn "%TASK_NAME%" /f >nul 2>&1

:: Tao task: chay khi dang nhap, delay 30 giay cho Windows on dinh
schtasks /create ^
    /tn "%TASK_NAME%" ^
    /tr "cmd.exe /c \"%BAT_FILE%\"" ^
    /sc ONLOGON ^
    /rl HIGHEST ^
    /f ^
    /delay 0000:30

if %errorLevel% equ 0 (
    echo.
    echo [OK] Dang ky thanh cong!
    echo      Task name : %TASK_NAME%
    echo      Bat dau   : 30 giay sau khi dang nhap Windows
    echo      File      : %BAT_FILE%
    echo.
) else (
    echo [LOI] Dang ky that bai. Kiem tra quyen Administrator.
)

:: Tao file go cai
(
echo @echo off
echo net session ^>nul 2^>^&1
echo if %%errorLevel%% neq 0 ^( echo Can quyen Admin ^& pause ^& exit /b 1 ^)
echo schtasks /delete /tn "%TASK_NAME%" /f
echo echo Da xoa task: %TASK_NAME%
echo pause
) > "%~dp0Uninstall_Autostart.bat"

echo [INFO] Da tao Uninstall_Autostart.bat
echo.

set /p RUN_NOW="Chay USB Reader ngay bay gio khong? (y/n): "
if /i "%RUN_NOW%"=="y" (
    start "" /b cmd.exe /c "%BAT_FILE%"
    echo [OK] Da chay ngam. Kiem tra log tai: %~dp0USB_Reader.log
)

echo.
pause
endlocal
