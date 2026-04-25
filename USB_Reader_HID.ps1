# USB_Reader_HID.ps1
param(
    [string]$ExcelFile     = "$PSScriptRoot\ABC.xlsx",
    [string]$LogFile       = "$PSScriptRoot\USB_Reader.log",
    [int]$ScannerSpeedMs   = 100,
    [int]$MinBarcodeLength = 3
)

Add-Type -AssemblyName System.Windows.Forms

# ----------------------------------------------------------------
# C#: Ghi barcode truc tiep vao Excel cua user qua Window Handle
# Lam moi thu trong C# - khong tra COM object ve PowerShell (tranh OLE variant error)
# ----------------------------------------------------------------
Add-Type -TypeDefinition @"
using System;
using System.Reflection;
using System.Runtime.InteropServices;

public class ExcelFinder {
    [DllImport("user32.dll")]
    private static extern IntPtr FindWindowEx(IntPtr parent, IntPtr after, string cls, string title);

    [DllImport("oleacc.dll")]
    private static extern int AccessibleObjectFromWindow(IntPtr hwnd, uint dwId, ref Guid riid, out IntPtr ppvObject);

    private const uint OBJID_NATIVEOM = 0xFFFFFFF0;
    private static readonly Guid IID_IDispatch = new Guid("{00020400-0000-0000-C000-000000000046}");

    // Lay property qua late-binding (tranh dung dynamic de khong can Microsoft.CSharp.dll)
    private static object Get(object obj, string prop, object[] args = null) {
        return obj.GetType().InvokeMember(prop, BindingFlags.GetProperty, null, obj, args);
    }
    private static void Set(object obj, string prop, object[] args) {
        obj.GetType().InvokeMember(prop, BindingFlags.SetProperty, null, obj, args);
    }
    private static void Call(object obj, string method, object[] args = null) {
        obj.GetType().InvokeMember(method, BindingFlags.InvokeMethod, null, obj, args);
    }

    // Ghi tat ca barcodes vao Excel dang mo cua user.
    // Tra ve so STT bat dau neu thanh cong, -1 neu khong tim thay Excel.
    public static int AppendBarcodes(string filePath, string[] timestamps, string[] barcodes) {
        string normPath = System.IO.Path.GetFullPath(filePath).ToLower();
        IntPtr hMain = IntPtr.Zero;

        while (true) {
            hMain = FindWindowEx(IntPtr.Zero, hMain, "XLMAIN", null);
            if (hMain == IntPtr.Zero) break;

            IntPtr hDesk = FindWindowEx(hMain, IntPtr.Zero, "XLDESK", null);
            if (hDesk == IntPtr.Zero) continue;

            IntPtr h7 = IntPtr.Zero;
            while (true) {
                h7 = FindWindowEx(hDesk, h7, "EXCEL7", null);
                if (h7 == IntPtr.Zero) break;

                IntPtr ptr = IntPtr.Zero;
                Guid g = IID_IDispatch;
                int hr = AccessibleObjectFromWindow(h7, OBJID_NATIVEOM, ref g, out ptr);
                if (hr != 0 || ptr == IntPtr.Zero) continue;

                try {
                    object win = Marshal.GetObjectForIUnknown(ptr);

                    // Kiem tra Excel co visible khong
                    object app = Get(win, "Application");
                    bool visible = (bool)Get(app, "Visible");
                    if (!visible) continue;

                    // Kiem tra duong dan workbook
                    object wb  = Get(win, "Parent");
                    string wbFull = (string)Get(wb, "FullName");
                    if (System.IO.Path.GetFullPath(wbFull).ToLower() != normPath) continue;

                    // Tim dong cuoi cung
                    object ws        = Get(Get(wb, "Sheets"), "Item", new object[] { 1 });
                    object usedRange = Get(ws, "UsedRange");
                    int lastRow      = (int)Get(Get(usedRange, "Rows"), "Count");
                    int nextRow      = Math.Max(2, lastRow + 1);
                    int firstStt     = nextRow - 1;
                    object cells     = Get(ws, "Cells");

                    for (int i = 0; i < barcodes.Length; i++) {
                        Set(Get(cells, "Item", new object[] { nextRow, 1 }), "Value", new object[] { nextRow - 1 });
                        Set(Get(cells, "Item", new object[] { nextRow, 2 }), "Value", new object[] { timestamps[i] });
                        Set(Get(cells, "Item", new object[] { nextRow, 3 }), "Value", new object[] { barcodes[i] });
                        nextRow++;
                    }

                    Call(wb, "Save");
                    return firstStt;  // thanh cong
                } catch { }
                finally {
                    if (ptr != IntPtr.Zero) Marshal.Release(ptr);
                }
            }
        }
        return -1;  // khong tim thay Excel
    }
}
"@ -Language CSharp

# ----------------------------------------------------------------
# C#: Global keyboard hook
# ----------------------------------------------------------------
Add-Type -TypeDefinition @"
using System;
using System.Runtime.InteropServices;
using System.Text;
using System.Diagnostics;
using System.Collections.Concurrent;

public class BarcodeHook {
    private const int WH_KEYBOARD_LL = 13;
    private const int WM_KEYDOWN     = 0x0100;
    private const int WM_SYSKEYDOWN  = 0x0104;

    public static ConcurrentQueue<string> Queue = new ConcurrentQueue<string>();

    private static LowLevelKeyboardProc _proc;
    private static IntPtr               _hook      = IntPtr.Zero;
    private static StringBuilder        _buf       = new StringBuilder(256);
    private static DateTime             _lastKey   = DateTime.MinValue;
    private static int                  _threshold = 100;
    private static int                  _minLen    = 3;

    public delegate IntPtr LowLevelKeyboardProc(int nCode, IntPtr wParam, IntPtr lParam);

    [StructLayout(LayoutKind.Sequential)]
    private struct KBDLLHOOKSTRUCT {
        public uint vkCode, scanCode, flags, time;
        public UIntPtr dwExtraInfo;
    }

    [DllImport("user32.dll")] private static extern IntPtr SetWindowsHookEx(int id, LowLevelKeyboardProc fn, IntPtr hMod, uint tid);
    [DllImport("user32.dll")] private static extern bool   UnhookWindowsHookEx(IntPtr h);
    [DllImport("user32.dll")] private static extern IntPtr CallNextHookEx(IntPtr h, int n, IntPtr w, IntPtr l);
    [DllImport("kernel32.dll")] private static extern IntPtr GetModuleHandle(string name);
    [DllImport("user32.dll")] private static extern short   GetKeyState(int vk);
    [DllImport("user32.dll")] private static extern IntPtr  GetKeyboardLayout(uint tid);
    [DllImport("user32.dll")] private static extern int     ToUnicodeEx(
        uint vk, uint scan, byte[] state,
        [Out, MarshalAs(UnmanagedType.LPWStr)] StringBuilder sb,
        int cap, uint flags, IntPtr hkl);

    public static void Install(int thresholdMs, int minLen) {
        _threshold = thresholdMs; _minLen = minLen; _proc = HookProc;
        using (var p = Process.GetCurrentProcess())
        using (var m = p.MainModule)
            _hook = SetWindowsHookEx(WH_KEYBOARD_LL, _proc, GetModuleHandle(m.ModuleName), 0);
    }

    public static void Uninstall() {
        if (_hook != IntPtr.Zero) { UnhookWindowsHookEx(_hook); _hook = IntPtr.Zero; }
    }

    private static IntPtr HookProc(int nCode, IntPtr wParam, IntPtr lParam) {
        if (nCode >= 0 && (wParam == (IntPtr)WM_KEYDOWN || wParam == (IntPtr)WM_SYSKEYDOWN)) {
            var    ks      = (KBDLLHOOKSTRUCT)Marshal.PtrToStructure(lParam, typeof(KBDLLHOOKSTRUCT));
            uint   vk      = ks.vkCode;
            double elapsed = (_lastKey == DateTime.MinValue) ? 0 : (DateTime.Now - _lastKey).TotalMilliseconds;

            if (_buf.Length > 0 && elapsed > _threshold * 4) _buf.Clear();
            _lastKey = DateTime.Now;

            if (vk == 13) {
                string code = _buf.ToString().Trim();
                _buf.Clear();
                if (code.Length >= _minLen) Queue.Enqueue(code);
                return CallNextHookEx(_hook, nCode, wParam, lParam);
            }
            if (vk == 8) {
                if (_buf.Length > 0) _buf.Remove(_buf.Length - 1, 1);
                return CallNextHookEx(_hook, nCode, wParam, lParam);
            }
            if (vk < 32 || (vk >= 91 && vk <= 93) || vk == 20 || vk == 16 || vk == 17 || vk == 18)
                return CallNextHookEx(_hook, nCode, wParam, lParam);

            byte[] state = new byte[256];
            for (int i = 0; i < 256; i++) state[i] = (byte)(GetKeyState(i) & 0xFF);
            var sb  = new StringBuilder(4);
            int res = ToUnicodeEx(vk, ks.scanCode, state, sb, sb.Capacity, 0, GetKeyboardLayout(0));
            if (res >= 1) _buf.Append(sb[0]);
        }
        return CallNextHookEx(_hook, nCode, wParam, lParam);
    }
}
"@ -Language CSharp

# ----------------------------------------------------------------
# Helper
# ----------------------------------------------------------------
function Write-Log {
    param([string]$msg)
    $line = "[$(Get-Date -f 'yyyy-MM-dd HH:mm:ss')] $msg"
    Add-Content -Path $LogFile -Value $line -Encoding UTF8
    Write-Host $line
}

# ----------------------------------------------------------------
# Hidden Excel: khoi dong 1 lan, tai su dung cho moi flush
# -> loai bo delay 3-5 giay moi lan ghi khi Excel user khong mo
# ----------------------------------------------------------------
$script:xl = $null

function Get-HiddenExcel {
    # Kiem tra instance cu con song khong
    if ($null -ne $script:xl) {
        try { $null = $script:xl.Version } catch { $script:xl = $null }
    }
    if ($null -eq $script:xl) {
        $script:xl               = New-Object -ComObject Excel.Application
        $script:xl.Visible       = $false
        $script:xl.DisplayAlerts = $false
        Write-Log "Hidden Excel khoi dong"
    }
    return $script:xl
}

# ----------------------------------------------------------------
# Flush batch vao Excel
# - Neu user dang mo file: ghi qua COM cua user (hien thi ngay)
# - Neu khong: dung hidden Excel da warm-up (chi mo/dong workbook)
# ----------------------------------------------------------------
function Flush-ToExcel {
    param([string]$Path, [string[]]$Barcodes)

    # Uu tien: ghi truc tiep vao Excel visible cua user (neu dang mo file nay)
    $timestamps = $Barcodes | ForEach-Object { Get-Date -Format "yyyy-MM-dd HH:mm:ss" }
    $firstStt   = [ExcelFinder]::AppendBarcodes($Path, $timestamps, $Barcodes)
    if ($firstStt -ge 0) {
        for ($i = 0; $i -lt $Barcodes.Length; $i++) {
            Write-Log "Ghi STT $($firstStt + $i): $($Barcodes[$i])"
        }
        return
    }

    # Khong co Excel cua user -> dung hidden Excel (Excel.Application da chay san)
    $wb = $null
    try {
        $xl = Get-HiddenExcel

        if (Test-Path $Path) {
            $wb = $xl.Workbooks.Open($Path)
        } else {
            $wb  = $xl.Workbooks.Add()
            $ws0 = $wb.Sheets.Item(1)
            $ws0.Cells.Item(1,1) = "STT"
            $ws0.Cells.Item(1,2) = "Thoi gian"
            $ws0.Cells.Item(1,3) = "Ma vach"
            $ws0.Rows.Item(1).Font.Bold      = $true
            $ws0.Columns.Item(1).ColumnWidth = 6
            $ws0.Columns.Item(2).ColumnWidth = 22
            $ws0.Columns.Item(3).ColumnWidth = 40
            $wb.SaveAs($Path, 51)
        }

        $ws      = $wb.Sheets.Item(1)
        $nextRow = [Math]::Max(2, $ws.UsedRange.Rows.Count + 1)

        foreach ($barcode in $Barcodes) {
            $ts  = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
            $stt = $nextRow - 1
            $ws.Cells.Item($nextRow, 1) = $stt
            $ws.Cells.Item($nextRow, 2) = $ts
            $ws.Cells.Item($nextRow, 3) = $barcode
            $nextRow++
            Write-Log "Ghi STT ${stt}: $barcode"
        }

        $wb.Save()

    } finally {
        # Chi dong workbook, giu Excel.Application song de lan sau dung lai (nhanh hon)
        if ($null -ne $wb) {
            try { $wb.Close($false) } catch {}
            [System.Runtime.InteropServices.Marshal]::ReleaseComObject($wb) | Out-Null
        }
    }
}

# ----------------------------------------------------------------
# Buffer: barcode duoc ghi vao day truoc, flush dinh ky
# -> khong bao gio mat data du file co bi lock tam thoi
# ----------------------------------------------------------------
$script:pending = [System.Collections.Generic.List[string]]::new()
$script:lastFlush = [DateTime]::Now
$FLUSH_INTERVAL_MS = 2000   # flush moi 2 giay

# ----------------------------------------------------------------
# Khoi dong
# ----------------------------------------------------------------
Write-Log "=== USB Reader khoi dong | ScannerSpeed: ${ScannerSpeedMs}ms | MinLen: $MinBarcodeLength ==="
Write-Log "File: $ExcelFile | Flush interval: ${FLUSH_INTERVAL_MS}ms"

# Tao file Excel neu chua co
if (-not (Test-Path $ExcelFile)) {
    Flush-ToExcel -Path $ExcelFile -Barcodes @()
    Write-Log "Tao file moi: $ExcelFile"
}

Write-Log "San sang."

# Pre-warm hidden Excel ngay khi khoi dong -> lan dau qua het ghi se nhanh
try { Get-HiddenExcel | Out-Null } catch { Write-Log "Pre-warm Excel that bai: $_" }

# ----------------------------------------------------------------
# Hidden WinForms + Timer
# ----------------------------------------------------------------
$form               = New-Object System.Windows.Forms.Form
$form.Opacity       = 0
$form.ShowInTaskbar = $false
$form.WindowState   = 'Minimized'

$timer          = New-Object System.Windows.Forms.Timer
$timer.Interval = 100

$timer.Add_Tick({
    # Buoc 1: Thu thap tat ca barcode vao pending (nhanh, khong IO)
    [string]$barcode = $null
    while ([BarcodeHook]::Queue.TryDequeue([ref]$barcode)) {
        if (-not [string]::IsNullOrWhiteSpace($barcode)) {
            $script:pending.Add($barcode)
        }
    }

    # Buoc 2: Flush neu du thoi gian hoac pending nhieu
    $elapsed = ([DateTime]::Now - $script:lastFlush).TotalMilliseconds
    if ($script:pending.Count -eq 0) { return }
    if ($elapsed -lt $FLUSH_INTERVAL_MS -and $script:pending.Count -lt 10) { return }

    try {
        $batch = $script:pending.ToArray()
        Flush-ToExcel -Path $script:ExcelFile -Barcodes $batch
        $script:pending.Clear()
        $script:lastFlush = [DateTime]::Now
    } catch {
        Write-Log "LOI flush (se thu lai): $_"
        # Pending giu nguyen, thu lai lan sau
    }
})

$form.Add_FormClosed({
    $timer.Stop()
    [BarcodeHook]::Uninstall()
    # Flush lan cuoi truoc khi thoat
    if ($script:pending.Count -gt 0) {
        try { Flush-ToExcel -Path $script:ExcelFile -Barcodes $script:pending.ToArray() } catch {}
    }
    # Dong hidden Excel neu dang chay
    if ($null -ne $script:xl) {
        try { $script:xl.Quit() } catch {}
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($script:xl) | Out-Null
        [GC]::Collect()
    }
    Write-Log "Da thoat."
})

[BarcodeHook]::Install($ScannerSpeedMs, $MinBarcodeLength)
$timer.Start()
Write-Log "Dang lang nghe ma vach..."

[System.Windows.Forms.Application]::Run($form)
