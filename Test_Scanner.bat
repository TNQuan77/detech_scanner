@echo off
setlocal
echo === Gia lap may quet ma vach ===
echo.
echo [1] Gui barcode vao ngay hom nay (script chinh phai dang chay)
echo [2] Gia lap qua ngay moi (tu dong restart script)
echo.
set /p CHOICE="Chon (1/2): "

if "%CHOICE%"=="2" (
    powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%~dp0Test_Scanner.ps1" -TestDateChange
) else (
    powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%~dp0Test_Scanner.ps1"
)
pause
endlocal
