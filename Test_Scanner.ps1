# Test_Scanner.ps1 — Gia lap may quet ma vach
# Co the gia lap thay doi ngay de test tao sheet moi
#
# Cach dung:
#   .\Test_Scanner.ps1                          -> gui barcode vao thang hien tai
#   .\Test_Scanner.ps1 -Date "04-2026"          -> gui barcode vao thang cu the
#   .\Test_Scanner.ps1 -TestDateChange          -> gia lap qua thang (thang truoc -> thang nay)

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

$logFile = Join-Path $scriptDir "USB_Reader.log"

function Stop-MainScript {
    Get-Process powershell -ErrorAction SilentlyContinue |
        Where-Object { $_.MainWindowTitle -eq "" } |
        ForEach-Object {
            $cmdline = (Get-WmiObject Win32_Process -Filter "ProcessId=$($_.Id)").CommandLine
            if ($cmdline -like "*USB_Reader_HID*") { $_ | Stop-Process -Force }
        }
    Start-Sleep -Milliseconds 1000
}

function Start-MainScript {
    param([string]$SimDate = "")
    # Xoa log truoc khi start de Wait-ForReady khong bi lua boi log cu
    try { [System.IO.File]::WriteAllText($logFile, "", [System.Text.Encoding]::UTF8) } catch {}
    $psArgs = @("-NoProfile", "-STA", "-ExecutionPolicy", "Bypass", "-WindowStyle", "Hidden",
                "-File", $mainScript)
    if ($SimDate) { $psArgs += @("-SimulateDate", $SimDate) }
    Start-Process powershell.exe -ArgumentList $psArgs
}

function Wait-ForReady {
    param([int]$TimeoutSec = 30)
    Write-Host "  Dang cho script san sang..." -NoNewline
    $deadline = [DateTime]::Now.AddSeconds($TimeoutSec)
    while ([DateTime]::Now -lt $deadline) {
        if (Test-Path $logFile) {
            $lines = Get-Content $logFile -ErrorAction SilentlyContinue
            if ($lines -match "lang nghe") { Write-Host " OK"; return }
        }
        Write-Host "." -NoNewline
        Start-Sleep -Milliseconds 500
    }
    Write-Host " TIMEOUT"
}

# ----------------------------------------------------------------
if ($TestDateChange) {
    $lastMonth = (Get-Date).AddMonths(-1).ToString("MM-yyyy")
    $thisMonth = Get-Date -Format "MM-yyyy"

    Write-Host "=== Test thay doi thang ==="
    Write-Host "Buoc 1: Gia lap thang truoc ($lastMonth)"
    Stop-MainScript
    Start-MainScript -SimDate $lastMonth
    Wait-ForReady
    Send-Barcodes -Codes $Barcodes
    Start-Sleep -Milliseconds 3000   # cho timer flush Excel (2s) kip chay truoc khi kill

    Write-Host ""
    Write-Host "Buoc 2: Gia lap thang nay ($thisMonth)"
    Stop-MainScript
    Start-MainScript -SimDate $thisMonth
    Wait-ForReady
    Send-Barcodes -Codes $Barcodes

    Write-Host ""
    Write-Host "Xong! Mo file Excel kiem tra co 2 sheet: '$lastMonth' va '$thisMonth'"

} elseif ($Date) {
    Write-Host "=== Gia lap scanner (thang: $Date) ==="
    Stop-MainScript
    Start-MainScript -SimDate $Date
    Wait-ForReady
    Send-Barcodes -Codes $Barcodes
    Write-Host "Xong!"

} else {
    Write-Host "=== Gia lap scanner (thang hien tai) ==="
    Write-Host "(Script chinh phai dang chay truoc)"
    Send-Barcodes -Codes $Barcodes
    Write-Host "Xong! Kiem tra file Excel va log."
}
