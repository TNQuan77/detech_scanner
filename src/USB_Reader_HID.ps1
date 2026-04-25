# USB_Reader_HID.ps1
param(
    [string]$ExcelFile     = "$(Split-Path $PSScriptRoot -Parent)\thoi_gian_dong_hang.xlsx",
    [string]$LogFile       = "$PSScriptRoot\USB_Reader.log",
    [int]$ScannerSpeedMs   = 100,
    [int]$MinBarcodeLength = 3,
    [string]$SimulateDate  = ""   # Override thang de test, VD: "04-2026"
)

Add-Type -AssemblyName System.Windows.Forms

# ----------------------------------------------------------------
# C#: Ghi barcode truc tiep vao Excel cua user qua Window Handle
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

    private static object Get(object obj, string prop, object[] args = null) {
        return obj.GetType().InvokeMember(prop, BindingFlags.GetProperty, null, obj, args);
    }
    private static void Set(object obj, string prop, object[] args) {
        obj.GetType().InvokeMember(prop, BindingFlags.SetProperty, null, obj, args);
    }
    private static void Call(object obj, string method, object[] args = null) {
        obj.GetType().InvokeMember(method, BindingFlags.InvokeMethod, null, obj, args);
    }

    private static object FindOrCreateSheet(object wb, string sheetDate) {
        object sheets   = Get(wb, "Sheets");
        int    cnt      = (int)Get(sheets, "Count");
        // Tim sheet co ten la ngay hom nay
        for (int s = 1; s <= cnt; s++) {
            object sh = Get(sheets, "Item", new object[] { s });
            if (string.Equals((string)Get(sh, "Name"), sheetDate, StringComparison.OrdinalIgnoreCase))
                return sh;
        }
        // Chua co -> tao sheet moi sau sheet cuoi cung
        object last = Get(sheets, "Item", new object[] { cnt });
        object mv   = System.Reflection.Missing.Value;
        object ws   = sheets.GetType().InvokeMember("Add", System.Reflection.BindingFlags.InvokeMethod,
                          null, sheets, new object[] { mv, last, mv, mv });
        Set(ws, "Name", new object[] { sheetDate });
        // Header
        object c = Get(ws, "Cells");
        Set(Get(c, "Item", new object[] { 1, 1 }), "Value", new object[] { "STT" });
        Set(Get(c, "Item", new object[] { 1, 2 }), "Value", new object[] { "Thoi gian" });
        Set(Get(Get(Get(ws, "Rows"), "Item", new object[] { 1 }), "Font"), "Bold", new object[] { true });
        Set(Get(Get(ws, "Columns"), "Item", new object[] { 1 }), "ColumnWidth", new object[] { 6.0 });
        Set(Get(Get(ws, "Columns"), "Item", new object[] { 2 }), "ColumnWidth", new object[] { 22.0 });
        return ws;
    }

    // scannerNames: ten hien thi (VD "T27H"), colIndices: vi tri cot (1-based -> Excel col 3,4,5,...)
    // sheetDate: ten sheet theo thang (VD "04-2026")
    public static int AppendBarcodes(string filePath, string[] timestamps, string[] barcodes,
                                     string[] scannerNames, int[] colIndices, string sheetDate) {
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

                    object app = Get(win, "Application");
                    bool visible = (bool)Get(app, "Visible");
                    if (!visible) continue;

                    object wb     = Get(win, "Parent");
                    string wbFull = (string)Get(wb, "FullName");
                    if (System.IO.Path.GetFullPath(wbFull).ToLower() != normPath) continue;

                    object ws        = FindOrCreateSheet(wb, sheetDate);
                    object usedRange = Get(ws, "UsedRange");
                    int lastRow      = (int)Get(Get(usedRange, "Rows"), "Count");
                    int nextRow      = Math.Max(2, lastRow + 1);
                    int firstStt     = nextRow - 1;
                    object cells     = Get(ws, "Cells");

                    for (int i = 0; i < barcodes.Length; i++) {
                        int scanCol = 2 + colIndices[i];

                        // Them header neu chua co
                        object hdrCell = Get(cells, "Item", new object[] { 1, scanCol });
                        object hdrVal  = null;
                        try { hdrVal = Get(hdrCell, "Value"); } catch {}
                        if (hdrVal == null || string.IsNullOrWhiteSpace(hdrVal.ToString()))
                            Set(hdrCell, "Value", new object[] { scannerNames[i] });

                        Set(Get(cells, "Item", new object[] { nextRow, 1 }), "Value", new object[] { nextRow - 1 });
                        Set(Get(cells, "Item", new object[] { nextRow, 2 }), "Value", new object[] { timestamps[i] });
                        Set(Get(cells, "Item", new object[] { nextRow, scanCol }), "Value", new object[] { barcodes[i] });
                        nextRow++;
                    }

                    Call(wb, "Save");
                    return firstStt;
                } catch { }
                finally {
                    if (ptr != IntPtr.Zero) Marshal.Release(ptr);
                }
            }
        }
        return -1;
    }
}
"@ -Language CSharp

# ----------------------------------------------------------------
# C#: Raw Input API — phan biet tung may quet theo HID device handle
# ----------------------------------------------------------------
Add-Type -TypeDefinition @"
using System;
using System.Collections.Generic;
using System.Collections.Concurrent;
using System.Runtime.InteropServices;
using System.Text;
using System.Windows.Forms;

public class BarcodeRawInput {
    private const int    WM_INPUT          = 0x00FF;
    private const uint   RIDEV_INPUTSINK   = 0x00000100;
    private const uint   RIDEV_REMOVE      = 0x00000001;
    private const ushort USAGE_PAGE_HID    = 0x01;
    private const ushort USAGE_KEYBOARD    = 0x06;
    private const uint   RID_INPUT         = 0x10000003;
    private const ushort RI_KEY_BREAK      = 0x01;
    private const uint   RIDI_DEVICENAME   = 0x20000007;

    [StructLayout(LayoutKind.Sequential)]
    private struct RAWINPUTDEVICE {
        public ushort usUsagePage;
        public ushort usUsage;
        public uint   dwFlags;
        public IntPtr hwndTarget;
    }

    [StructLayout(LayoutKind.Sequential)]
    private struct RAWINPUTHEADER {
        public uint   dwType;
        public uint   dwSize;
        public IntPtr hDevice;
        public IntPtr wParam;
    }

    [StructLayout(LayoutKind.Sequential)]
    private struct RAWKEYBOARD {
        public ushort MakeCode;
        public ushort Flags;
        public ushort Reserved;
        public ushort VKey;
        public uint   Message;
        public uint   ExtraInformation;
    }

    [DllImport("user32.dll", SetLastError = true)]
    private static extern bool RegisterRawInputDevices(
        [MarshalAs(UnmanagedType.LPArray, SizeParamIndex = 1)] RAWINPUTDEVICE[] rid,
        int count, int cbSize);

    [DllImport("user32.dll")]
    private static extern int GetRawInputData(IntPtr hRawInput, uint cmd,
        IntPtr pData, ref int pcbSize, int cbSizeHeader);

    [DllImport("user32.dll")]
    private static extern int GetRawInputDeviceInfoW(IntPtr hDevice, uint cmd,
        IntPtr pData, ref int pcbSize);

    [DllImport("user32.dll")]
    private static extern short  GetKeyState(int vk);
    [DllImport("user32.dll")]
    private static extern IntPtr GetKeyboardLayout(uint tid);
    [DllImport("user32.dll")]
    private static extern int ToUnicodeEx(uint vk, uint scan, byte[] state,
        [Out, MarshalAs(UnmanagedType.LPWStr)] StringBuilder sb,
        int cap, uint flags, IntPtr hkl);

    [DllImport("kernel32.dll", CharSet = CharSet.Unicode, SetLastError = true)]
    private static extern IntPtr CreateFile(string lpFileName, uint dwDesiredAccess,
        uint dwShareMode, IntPtr lpSec, uint dwCreationDisp, uint dwFlags, IntPtr hTemplate);
    [DllImport("kernel32.dll")] private static extern bool CloseHandle(IntPtr h);
    [DllImport("hid.dll", CharSet = CharSet.Unicode)]
    private static extern bool HidD_GetProductString(IntPtr hDev, [Out] char[] buf, uint len);

    private static readonly IntPtr INVALID_HANDLE = new IntPtr(-1);
    private const uint FILE_SHARE_RW = 0x00000001 | 0x00000002;
    private const uint OPEN_EXISTING = 3;

    // Queue entries: "displayName|colIdx\tbarcode"
    public static ConcurrentQueue<string> Queue      = new ConcurrentQueue<string>();
    // New device events: "displayName\tdevicePath"
    public static ConcurrentQueue<string> NewDevices = new ConcurrentQueue<string>();

    private static Dictionary<IntPtr, string>        _ids         = new Dictionary<IntPtr, string>();
    private static Dictionary<IntPtr, StringBuilder> _bufs        = new Dictionary<IntPtr, StringBuilder>();
    private static Dictionary<IntPtr, DateTime>      _times       = new Dictionary<IntPtr, DateTime>();
    private static Dictionary<string, string>        _pathToName  = new Dictionary<string, string>(); // HID path -> display name
    private static Dictionary<string, int>           _pathToCol   = new Dictionary<string, int>();    // HID path -> col index
    private static string _mapFile    = "";
    private static int    _nextColIdx = 1;
    private static int    _threshold  = 100;
    private static int    _minLen     = 3;

    private static string GetFriendlyName(string devicePath) {
        try {
            IntPtr hFile = CreateFile(devicePath, 0, FILE_SHARE_RW, IntPtr.Zero, OPEN_EXISTING, 0, IntPtr.Zero);
            if (hFile == INVALID_HANDLE) return null;
            try {
                char[] buf = new char[256];
                if (HidD_GetProductString(hFile, buf, (uint)(buf.Length * 2))) {
                    string name = new string(buf).TrimEnd('\0').Trim();
                    if (!string.IsNullOrEmpty(name)) return name;
                }
            } finally { CloseHandle(hFile); }
        } catch { }
        return null;
    }

    private static bool IsNameTaken(string name) {
        foreach (var v in _pathToName.Values)
            if (v == name) return true;
        return false;
    }

    private static string EnsureUnique(string baseName) {
        if (!IsNameTaken(baseName)) return baseName;
        int n = 2;
        while (IsNameTaken(baseName + " (" + n + ")")) n++;
        return baseName + " (" + n + ")";
    }

    public static void LoadMap(string mapFilePath) {
        _mapFile = mapFilePath;
        if (!System.IO.File.Exists(mapFilePath)) return;
        foreach (string line in System.IO.File.ReadAllLines(mapFilePath, System.Text.Encoding.UTF8)) {
            string[] p = line.Split('\t');
            if (p.Length < 2) continue;
            string path = p[0], name = p[1];
            int colIdx = 1;
            if (p.Length >= 3) {
                int.TryParse(p[2], out colIdx);
            } else {
                // Backward compat: old format "path\tScanner N" -> extract N as colIdx
                string[] np = name.Split(' ');
                if (np.Length >= 2) int.TryParse(np[np.Length - 1], out colIdx);
            }
            _pathToName[path] = name;
            _pathToCol[path]  = colIdx;
            if (colIdx >= _nextColIdx) _nextColIdx = colIdx + 1;
        }
    }

    public static string LastSaveError = "";

    private static void SaveMap() {
        if (string.IsNullOrEmpty(_mapFile)) return;
        try {
            // Bo attributes truoc de tranh loi ghi de file Hidden/System
            if (System.IO.File.Exists(_mapFile))
                try { System.IO.File.SetAttributes(_mapFile, System.IO.FileAttributes.Normal); } catch {}

            var lines = new List<string>();
            foreach (var kv in _pathToName)
                lines.Add(kv.Key + "\t" + kv.Value + "\t" + _pathToCol[kv.Key]);
            System.IO.File.WriteAllLines(_mapFile, lines.ToArray(), System.Text.Encoding.UTF8);

            try {
                System.IO.File.SetAttributes(_mapFile,
                    System.IO.FileAttributes.Hidden | System.IO.FileAttributes.System);
            } catch {}
        } catch (Exception ex) {
            LastSaveError = ex.Message;
        }
    }

    // Tra ve "displayName|colIdx" de ghi vao queue
    private static string GetOrAssign(IntPtr hDevice, out bool isNew) {
        isNew = false;
        if (_ids.ContainsKey(hDevice)) return _ids[hDevice];

        string path = GetDevicePath(hDevice);
        string name;
        int    colIdx;

        if (_pathToName.ContainsKey(path)) {
            name   = _pathToName[path];
            colIdx = _pathToCol[path];
        } else {
            isNew = true;
            string friendly = GetFriendlyName(path);
            name   = EnsureUnique(string.IsNullOrEmpty(friendly) ? "Scanner " + _nextColIdx : friendly);
            colIdx = _nextColIdx++;
            _pathToName[path] = name;
            _pathToCol[path]  = colIdx;
            SaveMap();
        }

        string encoded = name + "|" + colIdx;
        _ids[hDevice] = encoded;
        return encoded;
    }

    private static string GetDevicePath(IntPtr hDevice) {
        int sz = 0;
        GetRawInputDeviceInfoW(hDevice, RIDI_DEVICENAME, IntPtr.Zero, ref sz);
        if (sz <= 0) return "(unknown)";
        IntPtr buf = Marshal.AllocHGlobal(sz * 2);
        try {
            GetRawInputDeviceInfoW(hDevice, RIDI_DEVICENAME, buf, ref sz);
            return Marshal.PtrToStringUni(buf) ?? "(unknown)";
        } finally { Marshal.FreeHGlobal(buf); }
    }

    public static void Register(IntPtr hwnd, int thresholdMs, int minLen) {
        _threshold = thresholdMs;
        _minLen    = minLen;
        var rid = new RAWINPUTDEVICE[] {
            new RAWINPUTDEVICE {
                usUsagePage = USAGE_PAGE_HID,
                usUsage     = USAGE_KEYBOARD,
                dwFlags     = RIDEV_INPUTSINK,
                hwndTarget  = hwnd
            }
        };
        RegisterRawInputDevices(rid, 1, Marshal.SizeOf(typeof(RAWINPUTDEVICE)));
    }

    public static void Unregister() {
        var rid = new RAWINPUTDEVICE[] {
            new RAWINPUTDEVICE {
                usUsagePage = USAGE_PAGE_HID,
                usUsage     = USAGE_KEYBOARD,
                dwFlags     = RIDEV_REMOVE,
                hwndTarget  = IntPtr.Zero
            }
        };
        RegisterRawInputDevices(rid, 1, Marshal.SizeOf(typeof(RAWINPUTDEVICE)));
    }

    public static void ProcessInput(IntPtr lParam) {
        int headerSz = Marshal.SizeOf(typeof(RAWINPUTHEADER));
        int size = 0;
        GetRawInputData(lParam, RID_INPUT, IntPtr.Zero, ref size, headerSz);
        if (size <= 0) return;

        IntPtr buf = Marshal.AllocHGlobal(size);
        try {
            if (GetRawInputData(lParam, RID_INPUT, buf, ref size, headerSz) < 0) return;

            var header = (RAWINPUTHEADER)Marshal.PtrToStructure(buf, typeof(RAWINPUTHEADER));
            if (header.dwType != 1) return; // 0=mouse 1=keyboard 2=hid

            var kb = (RAWKEYBOARD)Marshal.PtrToStructure(
                new IntPtr(buf.ToInt64() + headerSz), typeof(RAWKEYBOARD));

            if ((kb.Flags & RI_KEY_BREAK) != 0) return; // key-up

            IntPtr hDevice = header.hDevice;
            uint   vk      = kb.VKey;

            if (!_bufs.ContainsKey(hDevice))  _bufs[hDevice]  = new StringBuilder(256);
            if (!_times.ContainsKey(hDevice)) _times[hDevice] = DateTime.MinValue;

            var    sbuf    = _bufs[hDevice];
            double elapsed = (_times[hDevice] == DateTime.MinValue)
                ? 0 : (DateTime.Now - _times[hDevice]).TotalMilliseconds;

            if (sbuf.Length > 0 && elapsed > _threshold * 4) sbuf.Clear();
            _times[hDevice] = DateTime.Now;

            if (vk == 13) { // Enter
                string code = sbuf.ToString().Trim();
                sbuf.Clear();
                // Chi gan Scanner ID khi nhan duoc barcode hop le
                // -> keyboard thuong khong bao gio duoc dang ky
                if (code.Length >= _minLen) {
                    bool   isNew;
                    string encoded = GetOrAssign(hDevice, out isNew); // "name|colIdx"
                    if (isNew) {
                        string displayName = encoded.Split('|')[0];
                        NewDevices.Enqueue(displayName + "\t" + GetDevicePath(hDevice));
                    }
                    Queue.Enqueue(encoded + "\t" + code); // "name|colIdx\tbarcode"
                }
                return;
            }
            if (vk == 8) { // Backspace
                if (sbuf.Length > 0) sbuf.Remove(sbuf.Length - 1, 1);
                return;
            }
            if (vk < 32 || (vk >= 91 && vk <= 93) || vk == 20 || vk == 16 || vk == 17 || vk == 18)
                return;

            byte[] state = new byte[256];
            for (int i = 0; i < 256; i++) state[i] = (byte)(GetKeyState(i) & 0xFF);
            var sb2 = new StringBuilder(4);
            int res = ToUnicodeEx(vk, kb.MakeCode, state, sb2, sb2.Capacity, 0, GetKeyboardLayout(0));
            if (res >= 1) sbuf.Append(sb2[0]);
        } finally {
            Marshal.FreeHGlobal(buf);
        }
    }
}

// Form override de nhan WM_INPUT
public class ScannerForm : System.Windows.Forms.Form {
    private const int WM_INPUT = 0x00FF;
    protected override void WndProc(ref System.Windows.Forms.Message m) {
        if (m.Msg == WM_INPUT) BarcodeRawInput.ProcessInput(m.LParam);
        base.WndProc(ref m);
    }
}
"@ -Language CSharp -ReferencedAssemblies "System.Windows.Forms"

# ----------------------------------------------------------------
# Helper
# ----------------------------------------------------------------
function Get-SheetDate {
    if ($SimulateDate) { return $SimulateDate }
    return (Get-Date -Format "MM-yyyy")
}

function Write-Log {
    param([string]$msg)
    $line = "[$(Get-Date -f 'yyyy-MM-dd HH:mm:ss')] $msg"
    try {
        $fs = [System.IO.FileStream]::new($LogFile,
            [System.IO.FileMode]::Append,
            [System.IO.FileAccess]::Write,
            [System.IO.FileShare]::ReadWrite)
        $sw = [System.IO.StreamWriter]::new($fs, [System.Text.Encoding]::UTF8)
        $sw.WriteLine($line)
        $sw.Dispose()
    } catch {}
    Write-Host $line
}

# ----------------------------------------------------------------
# Hidden Excel: khoi dong 1 lan, tai su dung cho moi flush
# ----------------------------------------------------------------
$script:xl = $null

function Get-HiddenExcel {
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
# ----------------------------------------------------------------
function Find-OrCreateSheet {
    param($Workbook, [string]$SheetName)
    for ($s = 1; $s -le $Workbook.Sheets.Count; $s++) {
        if ($Workbook.Sheets.Item($s).Name -eq $SheetName) { return $Workbook.Sheets.Item($s) }
    }
    # Tao sheet moi sau sheet cuoi
    $ws = $Workbook.Sheets.Add([System.Reflection.Missing]::Value, $Workbook.Sheets.Item($Workbook.Sheets.Count))
    $ws.Name                         = $SheetName
    $ws.Cells.Item(1,1)              = "STT"
    $ws.Cells.Item(1,2)              = "Thoi gian"
    $ws.Rows.Item(1).Font.Bold       = $true
    $ws.Columns.Item(1).ColumnWidth  = 6
    $ws.Columns.Item(2).ColumnWidth  = 22
    return $ws
}

function Flush-ToExcel {
    param([string]$Path, [string[]]$Barcodes, [string[]]$Scanners, [int[]]$Cols, [string]$SheetDate)

    $timestamps = $Barcodes | ForEach-Object { Get-Date -Format "yyyy-MM-dd HH:mm:ss" }
    $firstStt   = [ExcelFinder]::AppendBarcodes($Path, $timestamps, $Barcodes, $Scanners, $Cols, $SheetDate)
    if ($firstStt -ge 0) {
        for ($i = 0; $i -lt $Barcodes.Length; $i++) {
            Write-Log "[$($Scanners[$i])] Ghi STT $($firstStt + $i): $($Barcodes[$i])"
        }
        return
    }

    $wb = $null
    try {
        $xl = Get-HiddenExcel

        if (Test-Path $Path) {
            $wb = $xl.Workbooks.Open($Path)
        } else {
            $wb  = $xl.Workbooks.Add()
            # Dat ten sheet dau tien la ngay hom nay
            $wb.Sheets.Item(1).Name          = $SheetDate
            $wb.Sheets.Item(1).Cells.Item(1,1) = "STT"
            $wb.Sheets.Item(1).Cells.Item(1,2) = "Thoi gian"
            $wb.Sheets.Item(1).Rows.Item(1).Font.Bold      = $true
            $wb.Sheets.Item(1).Columns.Item(1).ColumnWidth = 6
            $wb.Sheets.Item(1).Columns.Item(2).ColumnWidth = 22
            $wb.SaveAs($Path, 51)
        }

        $ws      = Find-OrCreateSheet -Workbook $wb -SheetName $SheetDate
        $nextRow = [Math]::Max(2, $ws.UsedRange.Rows.Count + 1)

        for ($i = 0; $i -lt $Barcodes.Length; $i++) {
            $scanCol = 2 + $Cols[$i]

            if ([string]::IsNullOrWhiteSpace($ws.Cells.Item(1, $scanCol).Value2)) {
                $ws.Cells.Item(1, $scanCol)            = $Scanners[$i]
                $ws.Columns.Item($scanCol).ColumnWidth = 40
                $ws.Cells.Item(1, $scanCol).Font.Bold  = $true
            }

            $ts  = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
            $stt = $nextRow - 1
            $ws.Cells.Item($nextRow, 1)        = $stt
            $ws.Cells.Item($nextRow, 2)        = $ts
            $ws.Cells.Item($nextRow, $scanCol) = $Barcodes[$i]
            $nextRow++
            Write-Log "[$($Scanners[$i])] Ghi STT ${stt}: $($Barcodes[$i])"
        }

        $wb.Save()

    } finally {
        if ($null -ne $wb) {
            try { $wb.Close($false) } catch {}
            [System.Runtime.InteropServices.Marshal]::ReleaseComObject($wb) | Out-Null
        }
    }
}

# ----------------------------------------------------------------
# Buffer
# ----------------------------------------------------------------
$script:pendingBarcodes = [System.Collections.Generic.List[string]]::new()
$script:pendingScanners = [System.Collections.Generic.List[string]]::new()
$script:pendingCols     = [System.Collections.Generic.List[int]]::new()
$script:lastFlush       = [DateTime]::Now
$FLUSH_INTERVAL_MS      = 2000

# ----------------------------------------------------------------
# Khoi dong
# ----------------------------------------------------------------
# Xoa log cu moi lan khoi dong lai
try { [System.IO.File]::WriteAllText($LogFile, "", [System.Text.Encoding]::UTF8) } catch {}

Write-Log "=== USB Reader khoi dong | ScannerSpeed: ${ScannerSpeedMs}ms | MinLen: $MinBarcodeLength ==="
Write-Log "File: $ExcelFile | Flush interval: ${FLUSH_INTERVAL_MS}ms"

if (-not (Test-Path $ExcelFile)) {
    Flush-ToExcel -Path $ExcelFile -Barcodes @() -Scanners @() -Cols @() -SheetDate (Get-SheetDate)
    Write-Log "Tao file moi: $ExcelFile"
}

Write-Log "San sang."

try { Get-HiddenExcel | Out-Null } catch { Write-Log "Pre-warm Excel that bai: $_" }

# ----------------------------------------------------------------
# WinForms (ScannerForm de nhan WM_INPUT) + Timer
# ----------------------------------------------------------------
$form               = New-Object ScannerForm
$form.Opacity       = 0
$form.ShowInTaskbar = $false
$form.WindowState   = 'Minimized'

$timer          = New-Object System.Windows.Forms.Timer
$timer.Interval = 100

$timer.Add_Tick({
    # Thu thap barcode tu queue (format: "name|colIdx\tbarcode")
    [string]$entry = $null
    while ([BarcodeRawInput]::Queue.TryDequeue([ref]$entry)) {
        $tabIdx = $entry.IndexOf("`t")
        if ($tabIdx -lt 0) { continue }
        $meta    = $entry.Substring(0, $tabIdx)   # "name|colIdx"
        $barcode = $entry.Substring($tabIdx + 1)
        $pipeIdx = $meta.LastIndexOf('|')
        if ($pipeIdx -lt 0 -or [string]::IsNullOrWhiteSpace($barcode)) { continue }
        $displayName = $meta.Substring(0, $pipeIdx)
        $colIdx      = [int]$meta.Substring($pipeIdx + 1)
        $script:pendingBarcodes.Add($barcode)
        $script:pendingScanners.Add($displayName)
        $script:pendingCols.Add($colIdx)
    }

    # Log thiet bi moi phat hien
    [string]$newDev = $null
    while ([BarcodeRawInput]::NewDevices.TryDequeue([ref]$newDev)) {
        $devParts = $newDev -split "`t", 2
        Write-Log "Phat hien thiet bi: $($devParts[0]) => $($devParts[1])"
    }

    # Log loi ghi scanner_map neu co
    if (-not [string]::IsNullOrEmpty([BarcodeRawInput]::LastSaveError)) {
        Write-Log "LOI ghi scanner_map: $([BarcodeRawInput]::LastSaveError)"
        [BarcodeRawInput]::LastSaveError = ""
    }

    $elapsed = ([DateTime]::Now - $script:lastFlush).TotalMilliseconds
    if ($script:pendingBarcodes.Count -eq 0) { return }
    if ($elapsed -lt $FLUSH_INTERVAL_MS -and $script:pendingBarcodes.Count -lt 10) { return }

    try {
        $batchBarcodes = $script:pendingBarcodes.ToArray()
        $batchScanners = $script:pendingScanners.ToArray()
        $batchCols     = $script:pendingCols.ToArray()
        Flush-ToExcel -Path $ExcelFile -Barcodes $batchBarcodes -Scanners $batchScanners -Cols $batchCols -SheetDate (Get-SheetDate)
        $script:pendingBarcodes.Clear()
        $script:pendingScanners.Clear()
        $script:pendingCols.Clear()
        $script:lastFlush = [DateTime]::Now
    } catch {
        Write-Log "LOI flush (se thu lai): $_"
    }
})

$form.Add_FormClosed({
    $timer.Stop()
    [BarcodeRawInput]::Unregister()
    if ($script:pendingBarcodes.Count -gt 0) {
        try {
            Flush-ToExcel -Path $ExcelFile `
                -Barcodes   $script:pendingBarcodes.ToArray() `
                -Scanners   $script:pendingScanners.ToArray() `
                -Cols       $script:pendingCols.ToArray() `
                -SheetDate  (Get-SheetDate)
        } catch {}
    }
    if ($null -ne $script:xl) {
        try { $script:xl.Quit() } catch {}
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($script:xl) | Out-Null
        [GC]::Collect()
    }
    Write-Log "Da thoat."
})

[BarcodeRawInput]::LoadMap("$PSScriptRoot\scanner_map.txt")
[BarcodeRawInput]::Register($form.Handle, $ScannerSpeedMs, $MinBarcodeLength)
$timer.Start()
Write-Log "Dang lang nghe ma vach (Raw Input)..."

[System.Windows.Forms.Application]::Run($form)
