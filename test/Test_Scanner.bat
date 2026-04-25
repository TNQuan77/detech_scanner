@echo off
setlocal
echo === Gia lap may quet ma vach ===
echo.
echo [1] Gui barcode vao thang hien tai (script chinh phai dang chay)
echo [2] Gia lap qua thang moi (tu dong restart script)
echo.
set /p CHOICE="Chon (1/2): "

if "%CHOICE%"=="2" (
    powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%~dp0Test_Scanner.ps1" -TestDateChange
) else (
    powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%~dp0Test_Scanner.ps1"
)
pause
endlocal
