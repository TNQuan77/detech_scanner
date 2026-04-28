# Test_Scanner.ps1 — Gia lap may quet ma vach
# Co the gia lap thay doi thang de test tao sheet moi
#
# Cach dung:
#   .\Test_Scanner.ps1                          -> gui barcode vao thang hien tai
#   .\Test_Scanner.ps1 -Date "04-2026"          -> gui barcode vao thang cu the
#   .\Test_Scanner.ps1 -TestDateChange          -> gia lap qua thang (thang truoc -> thang nay)

param(
    [string[]]$Barcodes       = @("BARCODE001", "TEST123456", "9876543210987"),
    [string]$Date             = "",      # Override thang, VD: "04-2026"
    [switch]$TestDateChange,             # Gia lap qua thang moi
    [int]$DelayMs             = 500
)

$scriptDir  = Split-Path -Parent $MyInvocation.MyCommand.Path
$srcDir     = Join-Path (Split-Path -Parent $scriptDir) "src"
$mainScript = Join-Path $srcDir "USB_Reader_HID.ps1"
$mainBat    = Join-Path $srcDir "USB_Reader_HID.bat"
$injectFile = Join-Path $srcDir "test_inject.queue"

function Send-Barcodes {
    param([string[]]$Codes)
    foreach ($b in $Codes) {
        $line = "Test Scanner|1`t$b"
        [System.IO.File]::AppendAllText($injectFile, $line + [Environment]::NewLine, [System.Text.Encoding]::UTF8)
        Write-Host "  Sent: $b"
        Start-Sleep -Milliseconds $DelayMs
    }
}

function Get-LogLineCount {
    if (-not (Test-Path $logFile)) { return 0 }
    return @(Get-Content $logFile -ErrorAction SilentlyContinue).Count
}

function Wait-ForFlush {
    param(
        [string[]]$Codes,
        [int]$StartLine = 0,
        [int]$TimeoutSec = 15
    )
    Write-Host "  Chap Excel ghi xong..." -NoNewline
    $deadline = [DateTime]::Now.AddSeconds($TimeoutSec)
    
    # Wait for the newly queued batch to be consumed and written.
    while ([DateTime]::Now -lt $deadline) {
        $queueDrained = -not (Test-Path $injectFile)
        if (Test-Path $logFile) {
            $content = @(Get-Content $logFile -ErrorAction SilentlyContinue | Select-Object -Skip $StartLine)
            $saveOk  = @($content | Select-String -Pattern "OK: Luu file thanh cong").Count -gt 0
            $saveErr = @($content | Select-String -Pattern "LOI flush|LOI khoi dong Hidden Excel").Count -gt 0
            if ($queueDrained -and ($saveOk -or $saveErr)) {
                if ($saveErr) {
                    Write-Host " CO LOI (xem log)"
                } else {
                    Write-Host " OK"
                }
                return
            }
        }
        Write-Host "." -NoNewline
        Start-Sleep -Milliseconds 500
    }
    Write-Host " TIMEOUT"
}

$logFile     = Join-Path $srcDir "USB_Reader.log"
$simDateFile = Join-Path $srcDir "simulate_date.txt"

function Set-SimDate {
    param([string]$Date)
    if ($Date) {
        [System.IO.File]::WriteAllText($simDateFile, $Date, [System.Text.Encoding]::UTF8)
    } else {
        Remove-Item $simDateFile -Force -ErrorAction SilentlyContinue
    }
}

function Assert-MainScriptRunning {
    $running = Get-Process powershell -ErrorAction SilentlyContinue |
        Where-Object { $_.MainWindowTitle -eq "" } |
        Where-Object {
            try { (Get-WmiObject Win32_Process -Filter "ProcessId=$($_.Id)").CommandLine -like "*USB_Reader_HID*" } catch { $false }
        }
    if (-not $running) {
        Write-Host "  [!] Script chinh chua chay. Khoi dong..." -ForegroundColor Yellow
        Start-MainScript
        Wait-ForReady
    }
}

function Stop-MainScript {
    $signalFile = Join-Path $srcDir "stop_signal"

    # Graceful stop: signal file -> main script timer se detect va goi form.Close()
    try { [System.IO.File]::WriteAllText($signalFile, "") } catch {}

    # Wait up to 5s for graceful exit
    $deadline = [DateTime]::Now.AddSeconds(5)
    $exited   = $false
    while ([DateTime]::Now -lt $deadline) {
        $running = Get-Process powershell -ErrorAction SilentlyContinue |
            Where-Object { $_.MainWindowTitle -eq "" } |
            Where-Object {
                try { (Get-WmiObject Win32_Process -Filter "ProcessId=$($_.Id)").CommandLine -like "*USB_Reader_HID*" } catch { $false }
            }
        if (-not $running) { $exited = $true; break }
        Start-Sleep -Milliseconds 200
    }

    if (-not $exited) {
        # Force kill neu graceful shutdown that bai
        Get-Process powershell -ErrorAction SilentlyContinue |
            Where-Object { $_.MainWindowTitle -eq "" } |
            ForEach-Object {
                try {
                    if ((Get-WmiObject Win32_Process -Filter "ProcessId=$($_.Id)").CommandLine -like "*USB_Reader_HID*") {
                        $_ | Stop-Process -Force -ErrorAction SilentlyContinue
                    }
                } catch {}
            }
        Start-Sleep -Milliseconds 800

        # Kill orphaned hidden Excel (Visible=false, khong co main window)
        Get-Process excel -ErrorAction SilentlyContinue |
            Where-Object { $_.MainWindowHandle -eq 0 } |
            Stop-Process -Force -ErrorAction SilentlyContinue
        Start-Sleep -Milliseconds 500
    }

    Remove-Item $signalFile -Force -ErrorAction SilentlyContinue
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
    $nextMonth = (Get-Date).AddMonths(1).ToString("MM-yyyy")

    Write-Host "=== Test thay doi thang (script chinh van chay) ==="
    Assert-MainScriptRunning

    Write-Host "Buoc 1: Gia lap thang truoc ($lastMonth)"
    Set-SimDate $lastMonth
    $logCursor = Get-LogLineCount
    Send-Barcodes -Codes $Barcodes
    Wait-ForFlush -Codes $Barcodes -StartLine $logCursor

    Write-Host ""
    Write-Host "Buoc 2: Gia lap thang nay ($thisMonth)"
    Set-SimDate $thisMonth
    $logCursor = Get-LogLineCount
    Send-Barcodes -Codes $Barcodes
    Wait-ForFlush -Codes $Barcodes -StartLine $logCursor

    Write-Host ""
    Write-Host "Buoc 3: Gia lap thang sau ($nextMonth)"
    Set-SimDate $nextMonth
    $logCursor = Get-LogLineCount
    Send-Barcodes -Codes $Barcodes
    Wait-ForFlush -Codes $Barcodes -StartLine $logCursor

    Set-SimDate ""  # Xoa override, tra ve thang thuc

    Write-Host ""
    Write-Host "Xong! Mo file Excel kiem tra thu tu sheet (trai->phai): '$lastMonth' | '$thisMonth' | '$nextMonth'"

} elseif ($Date) {
    Write-Host "=== Gia lap scanner (thang: $Date) ==="
    Stop-MainScript
    Start-MainScript -SimDate $Date
    Wait-ForReady
    $logCursor = Get-LogLineCount
    Send-Barcodes -Codes $Barcodes
    Wait-ForFlush -Codes $Barcodes -StartLine $logCursor
    Stop-MainScript
    Write-Host "Xong!"

} else {
    Write-Host "=== Gia lap scanner (thang hien tai) ==="
    Write-Host "(Script chinh phai dang chay truoc)"
    Send-Barcodes -Codes $Barcodes
    Write-Host "Xong! Kiem tra file Excel va log."
}
