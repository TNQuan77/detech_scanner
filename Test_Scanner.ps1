# Test_Scanner.ps1 — Gia lap may quet ma vach
# Script chinh (USB_Reader_HID.bat) phai dang chay truoc
param(
    [string[]]$Barcodes = @("BARCODE001", "TEST123456", "9876543210987"),
    [int]$DelayMs       = 500   # delay giua cac lan quet (ms)
)

Add-Type -AssemblyName System.Windows.Forms

Write-Host "=== Gia lap scanner ==="
Write-Host "Barcode se gui: $($Barcodes -join ', ')"
Write-Host ""

foreach ($b in $Barcodes) {
    # Escape ky tu dac biet cua SendKeys: + ^ % ~ ( ) { } [ ]
    $escaped = $b -replace '([+^%~(){}\[\]])', '{$1}'
    [System.Windows.Forms.SendKeys]::SendWait($escaped + "{ENTER}")
    Write-Host "Sent: $b"
    Start-Sleep -Milliseconds $DelayMs
}

Write-Host ""
Write-Host "Xong! Kiem tra file Excel va log."
