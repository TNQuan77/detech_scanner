@echo off
setlocal
echo === Gia lap may quet ma vach ===
echo Script chinh (USB_Reader_HID.bat) phai dang chay truoc!
echo.
powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%~dp0Test_Scanner.ps1"
pause
endlocal
