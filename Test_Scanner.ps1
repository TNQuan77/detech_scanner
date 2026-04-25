# Test_Scanner.ps1 — Gia lap may quet ma vach
# Co the gia lap thay doi ngay de test tao sheet moi
#
# Cach dung:
#   .\Test_Scanner.ps1                          -> gui barcode vao ngay hom nay
#   .\Test_Scanner.ps1 -Date "23-04-2026"       -> gui barcode vao ngay cu the
#   .\Test_Scanner.ps1 -TestDateChange          -> gia lap qua ngay (hom qua -> hom nay)

param(
    [string[]]$Barcodes       = @("BARCODE001", "TEST123456", "9876543210987"),
    [string]$Date             = "",      # Override ngay, VD: "23-04-2026"
    [switch]$TestDateChange,             # Gia lap qua ngay moi
    [int]$DelayMs             = 500
)

Add-Type -AssemblyName System.Windows.Forms

$scriptDir  = Split-Path -Parent $MyInvocation.MyCommand.Path
$mainScript = Join-Path $scriptDir "USB_Reader_HID.ps1"
$mainBat    = Join-Path $scriptDir "USB_Reader_HID.bat"

function Send-Barcodes {
    param([string[]]$Codes)
    foreach ($b in $Codes) {
        $escaped = $b -replace '([+^%~(){}\[\]])', '{$1}'
        [System.Windows.Forms.SendKeys]::SendWait($escaped + "{ENTER}")
        Write-Host "  Sent: $b"
        Start-Sleep -Milliseconds $DelayMs
    }
}

function Stop-MainScript {
    Get-Process powershell -ErrorAction SilentlyContinue |
        Where-Object { $_.MainWindowTitle -eq "" } |
        ForEach-Object {
            $cmdline = (Get-WmiObject Win32_Process -Filter "ProcessId=$($_.Id)").CommandLine
            if ($cmdline -like "*USB_Reader_HID*") { $_ | Stop-Process -Force }
        }
    Start-Sleep -Milliseconds 1500
}

function Start-MainScript {
    param([string]$SimDate = "")
    $args = @("-NoProfile", "-STA", "-ExecutionPolicy", "Bypass", "-WindowStyle", "Hidden",
              "-File", $mainScript)
    if ($SimDate) { $args += @("-SimulateDate", $SimDate) }
    Start-Process powershell.exe -ArgumentList $args
    Start-Sleep -Milliseconds 3000   # cho script khoi dong
}

# ----------------------------------------------------------------
if ($TestDateChange) {
    $yesterday = (Get-Date).AddDays(-1).ToString("dd-MM-yyyy")
    $today     = Get-Date -Format "dd-MM-yyyy"

    Write-Host "=== Test thay doi ngay ==="
    Write-Host "Buoc 1: Gia lap ngay hom qua ($yesterday)"
    Stop-MainScript
    Start-MainScript -SimDate $yesterday
    Send-Barcodes -Codes $Barcodes

    Write-Host ""
    Write-Host "Buoc 2: Gia lap ngay hom nay ($today)"
    Stop-MainScript
    Start-MainScript -SimDate $today
    Send-Barcodes -Codes $Barcodes

    Write-Host ""
    Write-Host "Xong! Mo file Excel kiem tra co 2 sheet: '$yesterday' va '$today'"

} elseif ($Date) {
    Write-Host "=== Gia lap scanner (ngay: $Date) ==="
    Stop-MainScript
    Start-MainScript -SimDate $Date
    Send-Barcodes -Codes $Barcodes
    Write-Host "Xong!"

} else {
    Write-Host "=== Gia lap scanner (ngay hom nay) ==="
    Write-Host "(Script chinh phai dang chay truoc)"
    Send-Barcodes -Codes $Barcodes
    Write-Host "Xong! Kiem tra file Excel va log."
}
