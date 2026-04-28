# USB_Reader_HID.ps1
param(
    [string]$ExcelFile     = "$(Split-Path $PSScriptRoot -Parent)\thoi_gian_dong_hang.xlsx",
    [string]$LogFile       = "$PSScriptRoot\USB_Reader.log",
    [int]$ScannerSpeedMs   = 100,
    [int]$MinBarcodeLength = 3,
    [string]$SimulateDate  = ""   # Override thang de test, VD: "04-2026"
)

# Normalize paths to avoid issues with ".." in paths
$ExcelFile = [System.IO.Path]::GetFullPath($ExcelFile)
$LogFile = [System.IO.Path]::GetFullPath($LogFile)

Add-Type -AssemblyName System.Windows.Forms -ErrorAction Stop

# ----------------------------------------------------------------
# Helper: Ensure file is writable
# ----------------------------------------------------------------
function Ensure-FileWritable {
    param([string]$FilePath)
    if (-not (Test-Path $FilePath)) { return }
    try {
        $file = Get-Item $FilePath -Force
        $attrs = $file.Attributes
        if ($attrs -band [System.IO.FileAttributes]::ReadOnly) {
            $oldAttrs = $file.Attributes
            $file.Attributes = $attrs -bxor [System.IO.FileAttributes]::ReadOnly
            Write-Log "Loai bo read-only: $FilePath (oldAttrs=$oldAttrs, newAttrs=$($file.Attributes))"
        }
    } catch { 
        Write-Log "Canh bao: Khong the sua doi attribute: $FilePath - $_" 
    }
}

# ================================================================
# C#: Ghi barcode truc tiep vao Excel cua user qua Window Handle
# ================================================================
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

    private static int ParseMonthKey(string name) {
        string[] p = name.Split('-');
        if (p.Length == 2) {
            int month, year;
            if (int.TryParse(p[0], out month) && int.TryParse(p[1], out year))
                return year * 12 + month;
        }
        return -1;
    }

    private static object FindOrCreateSheet(object wb, string sheetDate) {
        object sheets = Get(wb, "Sheets");
        int    cnt    = (int)Get(sheets, "Count");
        for (int s = 1; s <= cnt; s++) {
            object sh = Get(sheets, "Item", new object[] { s });
            if (string.Equals((string)Get(sh, "Name"), sheetDate, StringComparison.OrdinalIgnoreCase))
                return sh;
        }
        // Chen dung vi tri: cu nhat ben trai (index nho), moi nhat ben phai (index lon)
        int    newKey      = ParseMonthKey(sheetDate);
        object insertBefore = null;
        for (int s = 1; s <= cnt; s++) {
            object sh  = Get(sheets, "Item", new object[] { s });
            int    key = ParseMonthKey((string)Get(sh, "Name"));
            if (key >= 0 && key > newKey) { insertBefore = sh; break; }
        }
        object mv = System.Reflection.Missing.Value;
        object ws;
        if (insertBefore != null)
            ws = sheets.GetType().InvokeMember("Add", System.Reflection.BindingFlags.InvokeMethod,
                     null, sheets, new object[] { insertBefore, mv, mv, mv });
        else {
            object last = Get(sheets, "Item", new object[] { cnt });
            ws = sheets.GetType().InvokeMember("Add", System.Reflection.BindingFlags.InvokeMethod,
                     null, sheets, new object[] { mv, last, mv, mv });
        }
        Set(ws, "Name", new object[] { sheetDate });
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
                        if (hdrVal == null || string.IsNullOrWhiteSpace(hdrVal.ToString())) {
                            Set(hdrCell, "Value", new object[] { scannerNames[i] });
                            // Format toan bo cot scanner la Text de tranh scientific notation
                            object scannerCol = Get(Get(ws, "Columns"), "Item", new object[] { scanCol });
                            Set(scannerCol, "NumberFormat", new object[] { "@" });
                        }

                        Set(Get(cells, "Item", new object[] { nextRow, 1 }), "Value", new object[] { nextRow - 1 });
                        Set(Get(cells, "Item", new object[] { nextRow, 2 }), "Value", new object[] { timestamps[i] });
                        object barCell = Get(cells, "Item", new object[] { nextRow, scanCol });
                        Set(barCell, "NumberFormat", new object[] { "@" });
                        Set(barCell, "Value", new object[] { barcodes[i] });
                        nextRow++;
                    }

                    // AutoFit tat ca cot theo noi dung thuc te
                    Call(Get(Get(ws, "UsedRange"), "Columns"), "AutoFit");

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
"@ -Language CSharp -ErrorAction Stop

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

    [DllImport("kernel32.dll", CharSet = CharSet.Unicode, SetLastError = true)]
    private static extern IntPtr CreateFile(string lpFileName, uint dwDesiredAccess,
        uint dwShareMode, IntPtr lpSec, uint dwCreationDisp, uint dwFlags, IntPtr hTemplate);
    [DllImport("kernel32.dll")] private static extern bool CloseHandle(IntPtr h);
    [DllImport("hid.dll", CharSet = CharSet.Unicode)]
    private static extern bool HidD_GetProductString(IntPtr hDev, [Out] char[] buf, uint len);
    [DllImport("user32.dll")]
    private static extern uint GetMessageTime();
    [DllImport("user32.dll")]
    private static extern IntPtr GetKeyboardLayout(uint tid);
    [DllImport("user32.dll", CharSet = CharSet.Unicode)]
    private static extern IntPtr LoadKeyboardLayout(string pwszKLID, uint Flags);
    [DllImport("user32.dll")]
    private static extern int ToUnicodeEx(uint vk, uint scan, byte[] state,
        [Out, MarshalAs(UnmanagedType.LPWStr)] StringBuilder sb,
        int cap, uint flags, IntPtr hkl);

    private static readonly IntPtr INVALID_HANDLE = new IntPtr(-1);
    private const uint FILE_SHARE_RW = 0x00000001 | 0x00000002;
    private const uint OPEN_EXISTING = 3;

    // Queue entries: "displayName|colIdx\tbarcode"
    public static ConcurrentQueue<string> Queue      = new ConcurrentQueue<string>();
    // New device events: "displayName\tdevicePath"
    public static ConcurrentQueue<string> NewDevices = new ConcurrentQueue<string>();

    private static Dictionary<IntPtr, string>        _ids         = new Dictionary<IntPtr, string>();
    private static Dictionary<IntPtr, StringBuilder> _bufs        = new Dictionary<IntPtr, StringBuilder>();
    private static Dictionary<IntPtr, uint>          _times       = new Dictionary<IntPtr, uint>();
    private static Dictionary<IntPtr, List<double>>  _charTimes   = new Dictionary<IntPtr, List<double>>();
    private static bool _shiftDown = false;
    private static bool _capsOn    = false;
    private static IntPtr _decodeLayout = IntPtr.Zero;
    private static bool   _decodeLayoutLoaded = false;
    private static Dictionary<string, string>        _pathToName  = new Dictionary<string, string>(); // HID path -> display name
    private static Dictionary<string, int>           _pathToCol   = new Dictionary<string, int>();    // HID path -> col index
    private static string _mapFile    = "";
    private static int    _nextColIdx = 1;
    private static int    _threshold  = 100;
    private static int    _minLen     = 3;

    private static IntPtr GetDecodeLayout() {
        if (!_decodeLayoutLoaded) {
            _decodeLayoutLoaded = true;
            try { _decodeLayout = LoadKeyboardLayout("00000409", 0); } catch {}
        }
        return _decodeLayout != IntPtr.Zero ? _decodeLayout : GetKeyboardLayout(0);
    }

    private static char? VkToCharUs(uint vk, bool shift, bool caps) {
        if (vk >= 0x41 && vk <= 0x5A) {
            bool upper = shift ^ caps;
            return upper ? (char)vk : (char)(vk + 32);
        }
        if (vk >= 0x30 && vk <= 0x39) {
            if (!shift) return (char)vk;
            return ")!@#$%^&*("[(int)(vk - 0x30)];
        }
        if (vk >= 0x60 && vk <= 0x69) return (char)('0' + (vk - 0x60));
        if (vk == 0x6A) return '*';
        if (vk == 0x6B) return '+';
        if (vk == 0x6D) return '-';
        if (vk == 0x6E) return '.';
        if (vk == 0x6F) return '/';
        if (vk == 0x20) return ' ';
        if (vk == 0xBA) return shift ? ':' : ';';
        if (vk == 0xBB) return shift ? '+' : '=';
        if (vk == 0xBC) return shift ? '<' : ',';
        if (vk == 0xBD) return shift ? '_' : '-';
        if (vk == 0xBE) return shift ? '>' : '.';
        if (vk == 0xBF) return shift ? '?' : '/';
        if (vk == 0xC0) return shift ? '~' : (char)0x60;
        if (vk == 0xDB) return shift ? '{' : '[';
        if (vk == 0xDC) return shift ? '|' : '\\';
        if (vk == 0xDD) return shift ? '}' : ']';
        if (vk == 0xDE) return shift ? '"' : '\'';
        return null;
    }

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

            bool   isKeyUp = (kb.Flags & RI_KEY_BREAK) != 0;
            uint   vk      = kb.VKey;

            // Track Shift key-up immediately so next character state is accurate.
            if (isKeyUp) {
                if (vk == 0x10 || vk == 0xA0 || vk == 0xA1) _shiftDown = false;
                return;
            }

            IntPtr hDevice = header.hDevice;

            if (!_bufs.ContainsKey(hDevice))  _bufs[hDevice]  = new StringBuilder(256);

            var  sbuf   = _bufs[hDevice];
            uint msgTime = GetMessageTime();
            double elapsed = _times.ContainsKey(hDevice) && _times[hDevice] != 0
                ? (double)(msgTime - _times[hDevice]) : 0;

            if (sbuf.Length > 0 && elapsed > _threshold * 4) {
                sbuf.Clear();
                if (_charTimes.ContainsKey(hDevice)) _charTimes[hDevice].Clear();
            }
            _times[hDevice] = msgTime;

            // Modifier keys still need to update timing above; otherwise uppercase
            // scans can be split because Shift gaps get counted into the next char.
            if (vk == 0x10 || vk == 0xA0 || vk == 0xA1) { _shiftDown = true; return; }
            if (vk == 0x14) { _capsOn = !_capsOn; return; }

            if (vk == 13) { // Enter
                string code = sbuf.ToString().Trim();
                List<double> ct = _charTimes.ContainsKey(hDevice)
                    ? new List<double>(_charTimes[hDevice]) : new List<double>();
                sbuf.Clear();
                if (_charTimes.ContainsKey(hDevice)) _charTimes[hDevice].Clear();

                // Tach cac ma bi ghep: tim gap > threshold*2 giua cac ky tu lien tiep
                var segments = new List<string>();
                int segStart = 0;
                for (int si = 1; si < ct.Count && si < code.Length; si++) {
                    if (ct[si] > _threshold * 2) {
                        string seg = code.Substring(segStart, si - segStart).Trim();
                        if (seg.Length >= _minLen) segments.Add(seg);
                        segStart = si;
                    }
                }
                if (segStart < code.Length) {
                    string last = code.Substring(segStart).Trim();
                    if (last.Length >= _minLen) segments.Add(last);
                }

                // Chi gan Scanner ID khi nhan duoc barcode hop le
                // -> keyboard thuong khong bao gio duoc dang ky
                if (segments.Count > 0) {
                    bool   isNew;
                    string encoded = GetOrAssign(hDevice, out isNew); // "name|colIdx"
                    if (isNew) {
                        string displayName = encoded.Split('|')[0];
                        NewDevices.Enqueue(displayName + "\t" + GetDevicePath(hDevice));
                    }
                    foreach (string seg in segments)
                        Queue.Enqueue(encoded + "\t" + seg);
                }
                return;
            }
            if (vk == 8) { // Backspace
                if (sbuf.Length > 0) {
                    sbuf.Remove(sbuf.Length - 1, 1);
                    if (_charTimes.ContainsKey(hDevice) && _charTimes[hDevice].Count > 0)
                        _charTimes[hDevice].RemoveAt(_charTimes[hDevice].Count - 1);
                }
                return;
            }
            if (vk < 32 || (vk >= 91 && vk <= 93) || vk == 20 || vk == 16 || vk == 17 || vk == 18)
                return;

            byte[] state = new byte[256];
            if (_shiftDown) {
                state[0x10] = 0x80;
                state[0xA0] = 0x80;
                state[0xA1] = 0x80;
            }
            if (_capsOn) state[0x14] = 0x01;
            if (vk < state.Length) state[(int)vk] = 0x80;

            // Prefer a fixed US ASCII mapping for common barcode characters so
            // the result does not depend on UniKey / active input method.
            char? ch = VkToCharUs(vk, _shiftDown, _capsOn);
            if (!ch.HasValue) {
                var sb2 = new StringBuilder(4);
                int res = ToUnicodeEx(vk, kb.MakeCode, state, sb2, sb2.Capacity, 0, GetDecodeLayout());
                if (res >= 1) ch = sb2[0];
            }
            if (ch.HasValue) {
                if (!_charTimes.ContainsKey(hDevice)) _charTimes[hDevice] = new List<double>();
                _charTimes[hDevice].Add(elapsed);
                sbuf.Append(ch.Value);
            }
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

// Chan keystroke cua scanner khoi cac app khac / UniKey bang WH_KEYBOARD_LL
// Raw Input van tiep tuc nhan du lieu nen script tu decode barcode.
public class KeyboardSuppressor {
    private const int WH_KEYBOARD_LL = 13;
    private const int WM_KEYDOWN     = 0x0100;
    private const int WM_SYSKEYDOWN  = 0x0104;

    [DllImport("user32.dll", SetLastError = true)]
    private static extern IntPtr SetWindowsHookEx(int idHook, LowLevelKeyboardProc fn, IntPtr hMod, uint tid);
    [DllImport("user32.dll")]
    private static extern bool UnhookWindowsHookEx(IntPtr hhk);
    [DllImport("user32.dll")]
    private static extern IntPtr CallNextHookEx(IntPtr hhk, int nCode, IntPtr wParam, IntPtr lParam);
    [DllImport("user32.dll")]
    private static extern short GetKeyState(int vk);
    [DllImport("kernel32.dll", CharSet = CharSet.Auto)]
    private static extern IntPtr GetModuleHandle(string name);

    [StructLayout(LayoutKind.Sequential)]
    private struct KBDLLHOOKSTRUCT {
        public uint   vkCode;
        public uint   scanCode;
        public uint   flags;
        public uint   time;
        public IntPtr dwExtraInfo;
    }

    private const uint LLKHF_INJECTED = 0x10;

    public delegate IntPtr LowLevelKeyboardProc(int nCode, IntPtr wParam, IntPtr lParam);

    private static IntPtr               _hook      = IntPtr.Zero;
    private static LowLevelKeyboardProc _proc;
    private static DateTime             _lastKey   = DateTime.MinValue;
    private static bool                 _scanMode  = false;
    private static int                  _threshold = 50;
    private static int                  _idleMs    = 300;
    public static string                LastError  = "";

    public static void Install(int thresholdMs) {
        _threshold = thresholdMs;
        _idleMs    = Math.Max(300, thresholdMs * 6);
        _proc      = Callback;
        _hook      = SetWindowsHookEx(WH_KEYBOARD_LL, _proc, GetModuleHandle(null), 0);
        if (_hook == IntPtr.Zero)
            LastError = "SetWindowsHookEx failed: " + Marshal.GetLastWin32Error();
    }

    public static void Uninstall() {
        if (_hook != IntPtr.Zero) {
            UnhookWindowsHookEx(_hook);
            _hook = IntPtr.Zero;
        }
    }

    private static IntPtr Callback(int nCode, IntPtr wParam, IntPtr lParam) {
        if (nCode >= 0 && ((int)wParam == WM_KEYDOWN || (int)wParam == WM_SYSKEYDOWN)) {
            var ks = (KBDLLHOOKSTRUCT)Marshal.PtrToStructure(lParam, typeof(KBDLLHOOKSTRUCT));

            // Phim gia lap (SendInput / SendKeys) -> pass through de test script van chay.
            if ((ks.flags & LLKHF_INJECTED) != 0)
                return CallNextHookEx(_hook, nCode, wParam, lParam);

            // Neu dang giu Ctrl / Alt / Win thi khong chan.
            bool modified = (GetKeyState(0x11) & 0x8000) != 0
                         || (GetKeyState(0x12) & 0x8000) != 0
                         || (GetKeyState(0x5B) & 0x8000) != 0
                         || (GetKeyState(0x5C) & 0x8000) != 0;
            if (!modified) {
                double gap = _lastKey == DateTime.MinValue ? 99999
                           : (DateTime.Now - _lastKey).TotalMilliseconds;
                _lastKey = DateTime.Now;

                if (gap < _threshold || (_scanMode && gap < _idleMs)) {
                    _scanMode = true;
                    return (IntPtr)1;
                }
                _scanMode = false;
            }
        }
        return CallNextHookEx(_hook, nCode, wParam, lParam);
    }
}

"@ -Language CSharp -ReferencedAssemblies "System.Windows.Forms" -ErrorAction Stop

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

function Invoke-ComMethod {
    param(
        [Parameter(Mandatory)]$ComObject,
        [Parameter(Mandatory)][string]$MethodName,
        [object[]]$Arguments = @()
    )

    return $ComObject.GetType().InvokeMember(
        $MethodName,
        [System.Reflection.BindingFlags]::InvokeMethod,
        $null,
        $ComObject,
        $Arguments
    )
}

# ----------------------------------------------------------------
# Hidden Excel worker: thread STA rieng, pre-warm luc start, khong giu lock workbook
# ----------------------------------------------------------------
$script:excelQueue            = $null
$script:excelLogQueue         = $null
$script:excelWorkerState      = $null
$script:excelWorkerRunspace   = $null
$script:excelWorkerPowerShell = $null
$script:excelWorkerHandle     = $null

function Flush-ExcelWorkerLogs {
    if ($null -eq $script:excelLogQueue) { return }
    [string]$msg = $null
    while ($script:excelLogQueue.TryDequeue([ref]$msg)) {
        Write-Log $msg
    }
}

function Start-ExcelWorker {
    if ($null -ne $script:excelWorkerPowerShell) { return }

    $script:excelQueue       = [System.Collections.Concurrent.BlockingCollection[object]]::new()
    $script:excelLogQueue    = [System.Collections.Concurrent.ConcurrentQueue[string]]::new()
    $script:excelWorkerState = [hashtable]::Synchronized(@{
        Ready   = $false
        Busy    = $false
        Stopped = $false
    })

    $iss = [System.Management.Automation.Runspaces.InitialSessionState]::CreateDefault()
    $iss.Variables.Add((New-Object System.Management.Automation.Runspaces.SessionStateVariableEntry -ArgumentList 'ExcelQueue', $script:excelQueue, 'Excel write queue'))
    $iss.Variables.Add((New-Object System.Management.Automation.Runspaces.SessionStateVariableEntry -ArgumentList 'ExcelLogQueue', $script:excelLogQueue, 'Excel worker logs'))
    $iss.Variables.Add((New-Object System.Management.Automation.Runspaces.SessionStateVariableEntry -ArgumentList 'ExcelWorkerState', $script:excelWorkerState, 'Excel worker state'))

    $script:excelWorkerRunspace = [runspacefactory]::CreateRunspace($iss)
    $script:excelWorkerRunspace.ApartmentState = 'STA'
    $script:excelWorkerRunspace.ThreadOptions  = 'ReuseThread'
    $script:excelWorkerRunspace.Open()

    $workerScript = @'
function Queue-WorkerLog {
    param([string]$Message)
    $ExcelLogQueue.Enqueue($Message)
}

function Invoke-ComMethod {
    param(
        [Parameter(Mandatory)]$ComObject,
        [Parameter(Mandatory)][string]$MethodName,
        [object[]]$Arguments = @()
    )

    return $ComObject.GetType().InvokeMember(
        $MethodName,
        [System.Reflection.BindingFlags]::InvokeMethod,
        $null,
        $ComObject,
        $Arguments
    )
}

function Ensure-FileWritable {
    param([string]$FilePath)
    if (-not (Test-Path $FilePath)) { return }
    try {
        $file = Get-Item $FilePath -Force
        $attrs = $file.Attributes
        if ($attrs -band [System.IO.FileAttributes]::ReadOnly) {
            $file.Attributes = $attrs -bxor [System.IO.FileAttributes]::ReadOnly
            Queue-WorkerLog "Loai bo read-only: $FilePath"
        }
    } catch {
        Queue-WorkerLog "Canh bao: Khong the sua doi attribute: $FilePath - $_"
    }
}

function Find-OrCreateSheet {
    param($Workbook, [string]$SheetName)
    $cnt = $Workbook.Sheets.Count
    for ($s = 1; $s -le $cnt; $s++) {
        if ($Workbook.Sheets.Item($s).Name -eq $SheetName) { return $Workbook.Sheets.Item($s) }
    }

    $keyOf = {
        param([string]$Name)
        $parts = $Name -split '-'
        if ($parts.Length -eq 2) {
            $month = 0
            $year  = 0
            if ([int]::TryParse($parts[0], [ref]$month) -and [int]::TryParse($parts[1], [ref]$year)) {
                return $year * 12 + $month
            }
        }
        return -1
    }

    $newKey = & $keyOf $SheetName
    $insertBefore = $null
    for ($s = 1; $s -le $cnt; $s++) {
        $sheet = $Workbook.Sheets.Item($s)
        $key   = & $keyOf $sheet.Name
        if ($key -ge 0 -and $key -gt $newKey) {
            $insertBefore = $sheet
            break
        }
    }

    $mv = [System.Reflection.Missing]::Value
    if ($null -ne $insertBefore) {
        $ws = Invoke-ComMethod -ComObject $Workbook.Sheets -MethodName Add -Arguments @($insertBefore, $mv, $mv, $mv)
    } else {
        $ws = Invoke-ComMethod -ComObject $Workbook.Sheets -MethodName Add -Arguments @($mv, $Workbook.Sheets.Item($cnt), $mv, $mv)
    }
    $ws.Name                         = $SheetName
    $ws.Cells.Item(1,1)              = "STT"
    $ws.Cells.Item(1,2)              = "Thoi gian"
    $ws.Rows.Item(1).Font.Bold       = $true
    $ws.Columns.Item(1).ColumnWidth  = 6
    $ws.Columns.Item(2).ColumnWidth  = 22
    return $ws
}

function Close-Workbook {
    param($Workbook)
    if ($null -eq $Workbook) { return }
    try { Invoke-ComMethod -ComObject $Workbook -MethodName Close -Arguments @($false) | Out-Null } catch {}
    try { [System.Runtime.InteropServices.Marshal]::ReleaseComObject($Workbook) | Out-Null } catch {}
}

function Open-HiddenWorkbook {
    param(
        $ExcelApp,
        [string]$Path,
        [string]$SheetDate
    )

    Ensure-FileWritable -FilePath $Path
    $workbooks = $ExcelApp.Workbooks

    if (Test-Path $Path) {
        $wb = Invoke-ComMethod -ComObject $workbooks -MethodName Open -Arguments @($Path, 0, $false)
    } else {
        $wb = Invoke-ComMethod -ComObject $workbooks -MethodName Add
        $wb.Sheets.Item(1).Name                        = $SheetDate
        $wb.Sheets.Item(1).Cells.Item(1,1)             = "STT"
        $wb.Sheets.Item(1).Cells.Item(1,2)             = "Thoi gian"
        $wb.Sheets.Item(1).Rows.Item(1).Font.Bold      = $true
        $wb.Sheets.Item(1).Columns.Item(1).ColumnWidth = 6
        $wb.Sheets.Item(1).Columns.Item(2).ColumnWidth = 22
        Invoke-ComMethod -ComObject $wb -MethodName SaveAs -Arguments @($Path, 51) | Out-Null
    }

    if ($wb.ReadOnly) {
        Close-Workbook -Workbook $wb
        throw "Workbook mo ra o che do read-only: $Path"
    }

    return $wb
}

function Flush-Batch {
    param(
        $Batch,
        $ExcelApp
    )

    $path       = [System.IO.Path]::GetFullPath([string]$Batch.Path)
    $barcodes   = @($Batch.Barcodes)
    $scanners   = @($Batch.Scanners)
    $cols       = @($Batch.Cols)
    $sheetDate  = [string]$Batch.SheetDate
    $createdNew = -not (Test-Path $path)

    if ($barcodes.Count -eq 0 -and -not $createdNew) { return }

    $timestamps = $barcodes | ForEach-Object { Get-Date -Format "yyyy-MM-dd HH:mm:ss" }
    Ensure-FileWritable -FilePath $path

    if ($barcodes.Count -gt 0) {
        $firstStt = [ExcelFinder]::AppendBarcodes($path, $timestamps, $barcodes, $scanners, $cols, $sheetDate)
        if ($firstStt -ge 0) {
            for ($i = 0; $i -lt $barcodes.Count; $i++) {
                Queue-WorkerLog "[$($scanners[$i])] Ghi STT $($firstStt + $i): $($barcodes[$i])"
            }
            return
        }
    }

    $wb = $null
    try {
        $wb = Open-HiddenWorkbook -ExcelApp $ExcelApp -Path $path -SheetDate $sheetDate
        $ws = Find-OrCreateSheet -Workbook $wb -SheetName $sheetDate
        $nextRow = [Math]::Max(2, $ws.UsedRange.Rows.Count + 1)

        for ($i = 0; $i -lt $barcodes.Count; $i++) {
            $scanCol = 2 + [int]$cols[$i]

            if ([string]::IsNullOrWhiteSpace($ws.Cells.Item(1, $scanCol).Value2)) {
                $ws.Cells.Item(1, $scanCol)             = $scanners[$i]
                $ws.Columns.Item($scanCol).NumberFormat = "@"
                $ws.Cells.Item(1, $scanCol).Font.Bold   = $true
            }

            $ts  = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
            $stt = $nextRow - 1
            $ws.Cells.Item($nextRow, 1)                     = $stt
            $ws.Cells.Item($nextRow, 2)                     = $ts
            $ws.Cells.Item($nextRow, $scanCol).NumberFormat = "@"
            $ws.Cells.Item($nextRow, $scanCol).Value2       = $barcodes[$i]
            $nextRow++
            Queue-WorkerLog "[$($scanners[$i])] Ghi STT ${stt}: $($barcodes[$i])"
        }

        Invoke-ComMethod -ComObject $wb -MethodName Save | Out-Null
        if ($createdNew) {
            Queue-WorkerLog "Tao file moi: $path"
        } elseif ($barcodes.Count -gt 0) {
            Queue-WorkerLog "OK: Luu file thanh cong"
        }
    } finally {
        Close-Workbook -Workbook $wb
    }
}

$xl = $null
try {
    $xl = New-Object -ComObject Excel.Application
    $xl.Visible = $false
    $xl.DisplayAlerts = $false
    $ExcelWorkerState['Ready'] = $true
    Queue-WorkerLog "Hidden Excel khoi dong"

    foreach ($batch in $ExcelQueue.GetConsumingEnumerable()) {
        $ExcelWorkerState['Busy'] = $true
        try {
            Flush-Batch -Batch $batch -ExcelApp $xl
        } catch {
            Queue-WorkerLog "LOI flush (se thu lai): $_"
        } finally {
            $ExcelWorkerState['Busy'] = $false
        }
    }
} catch {
    $ExcelWorkerState['Ready'] = $true
    Queue-WorkerLog "LOI khoi dong Hidden Excel: $_"
} finally {
    if ($null -ne $xl) {
        try { Invoke-ComMethod -ComObject $xl -MethodName Quit | Out-Null } catch {}
        try { [System.Runtime.InteropServices.Marshal]::ReleaseComObject($xl) | Out-Null } catch {}
        [GC]::Collect()
    }
    $ExcelWorkerState['Stopped'] = $true
}
'@

    $script:excelWorkerPowerShell = [powershell]::Create()
    $script:excelWorkerPowerShell.Runspace = $script:excelWorkerRunspace
    [void]$script:excelWorkerPowerShell.AddScript($workerScript)
    $script:excelWorkerHandle = $script:excelWorkerPowerShell.BeginInvoke()
}

function Wait-ExcelWorkerReady {
    param([int]$TimeoutMs = 15000)
    $deadline = [DateTime]::Now.AddMilliseconds($TimeoutMs)
    while ([DateTime]::Now -lt $deadline) {
        Flush-ExcelWorkerLogs
        if ($script:excelWorkerState['Ready']) { return $true }
        Start-Sleep -Milliseconds 100
    }
    Flush-ExcelWorkerLogs
    return [bool]$script:excelWorkerState['Ready']
}

function Queue-ExcelFlush {
    param(
        [string]$Path,
        [string[]]$Barcodes,
        [string[]]$Scanners,
        [int[]]$Cols,
        [string]$SheetDate
    )

    if ($null -eq $script:excelQueue -or $script:excelQueue.IsAddingCompleted) {
        throw "Excel worker khong san sang"
    }

    $batch = [pscustomobject]@{
        Path      = [System.IO.Path]::GetFullPath($Path)
        Barcodes  = @($Barcodes)
        Scanners  = @($Scanners)
        Cols      = @($Cols)
        SheetDate = $SheetDate
    }
    $script:excelQueue.Add($batch)
}

function Wait-ExcelWorkerIdle {
    param([int]$TimeoutMs = 15000)
    $deadline = [DateTime]::Now.AddMilliseconds($TimeoutMs)
    while ([DateTime]::Now -lt $deadline) {
        Flush-ExcelWorkerLogs
        if ($script:excelQueue.Count -eq 0 -and -not [bool]$script:excelWorkerState['Busy']) {
            return $true
        }
        Start-Sleep -Milliseconds 100
    }
    Flush-ExcelWorkerLogs
    return ($script:excelQueue.Count -eq 0 -and -not [bool]$script:excelWorkerState['Busy'])
}

function Stop-ExcelWorker {
    if ($null -eq $script:excelWorkerPowerShell) { return }

    try {
        if (-not $script:excelQueue.IsAddingCompleted) {
            $script:excelQueue.CompleteAdding()
        }
    } catch {}

    if ($null -ne $script:excelWorkerHandle) {
        $null = $script:excelWorkerHandle.AsyncWaitHandle.WaitOne(15000)
        if ($script:excelWorkerHandle.IsCompleted) {
            try { $script:excelWorkerPowerShell.EndInvoke($script:excelWorkerHandle) | Out-Null } catch { Write-Log "Loi worker khi dung: $_" }
        } else {
            Write-Log "Canh bao: Excel worker khong dung kip trong 15s"
        }
    }

    Flush-ExcelWorkerLogs

    try { $script:excelWorkerPowerShell.Dispose() } catch {}
    try { $script:excelWorkerRunspace.Dispose() } catch {}
    try { $script:excelQueue.Dispose() } catch {}

    $script:excelQueue            = $null
    $script:excelLogQueue         = $null
    $script:excelWorkerState      = $null
    $script:excelWorkerRunspace   = $null
    $script:excelWorkerPowerShell = $null
    $script:excelWorkerHandle     = $null
}

# ----------------------------------------------------------------
# Hidden Excel: giu app + workbook mo giua cac lan flush
# ----------------------------------------------------------------
$script:xl = $null
$script:wb = $null   # workbook mo thuong truc, tranh open/close moi lan

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

function Close-HiddenWorkbook {
    if ($null -ne $script:wb) {
        try { 
            # Check if workbook is in read-only mode and try to disable it
            if ($script:wb.ReadOnly) {
                Write-Log "Canh bao: Workbook dang o che do read-only, dong va thap lai..."
                Invoke-ComMethod -ComObject $script:wb -MethodName Close -Arguments @($false) | Out-Null
                $script:wb = $null
                Start-Sleep -Milliseconds 500
                return
            }
            Invoke-ComMethod -ComObject $script:wb -MethodName Save | Out-Null
        } catch { Write-Log "Canh bao khi save trong Close-HiddenWorkbook: $_" }
        try { Invoke-ComMethod -ComObject $script:wb -MethodName Close -Arguments @($false) | Out-Null } catch {}
        try { [System.Runtime.InteropServices.Marshal]::ReleaseComObject($script:wb) | Out-Null } catch {}
        $script:wb = $null
    }
}

function Get-HiddenWorkbook {
    param([string]$Path, [string]$SheetDate)
    $xl = Get-HiddenExcel
    if ($null -ne $script:wb) {
        try { $null = $script:wb.Name } catch { $script:wb = $null }
    }
    if ($null -eq $script:wb) {
        Ensure-FileWritable -FilePath $Path
        $workbooks = $xl.Workbooks
        if (Test-Path $Path) {
            # Open with UpdateLinks=0 to avoid prompts, ReadOnly=$false to ensure writable
            $script:wb = Invoke-ComMethod -ComObject $workbooks -MethodName Open -Arguments @($Path, 0, $false)
            # Double-check: disable ReadOnly on workbook if Excel set it
            if ($script:wb.ReadOnly) {
                Invoke-ComMethod -ComObject $script:wb -MethodName Close -Arguments @($false) | Out-Null
                [System.Runtime.InteropServices.Marshal]::ReleaseComObject($script:wb) | Out-Null
                $script:wb = $null
                # Try opening again with different parameters
                Ensure-FileWritable -FilePath $Path
                $script:wb = Invoke-ComMethod -ComObject $workbooks -MethodName Open -Arguments @($Path, 0, $false)
            }
        } else {
            $script:wb = Invoke-ComMethod -ComObject $workbooks -MethodName Add
            $script:wb.Sheets.Item(1).Name                          = $SheetDate
            $script:wb.Sheets.Item(1).Cells.Item(1,1)              = "STT"
            $script:wb.Sheets.Item(1).Cells.Item(1,2)              = "Thoi gian"
            $script:wb.Sheets.Item(1).Rows.Item(1).Font.Bold       = $true
            $script:wb.Sheets.Item(1).Columns.Item(1).ColumnWidth  = 6
            $script:wb.Sheets.Item(1).Columns.Item(2).ColumnWidth  = 22
            Invoke-ComMethod -ComObject $script:wb -MethodName SaveAs -Arguments @($Path, 51) | Out-Null
        }
    }
    return $script:wb
}

# ----------------------------------------------------------------
# Flush batch vao Excel
# ----------------------------------------------------------------
function Find-OrCreateSheet {
    param($Workbook, [string]$SheetName)
    $cnt = $Workbook.Sheets.Count
    for ($s = 1; $s -le $cnt; $s++) {
        if ($Workbook.Sheets.Item($s).Name -eq $SheetName) { return $Workbook.Sheets.Item($s) }
    }
    # Parse "MM-yyyy" -> sort key
    $keyOf = { param([string]$n)
        $p = $n -split '-'
        if ($p.Length -eq 2) {
            $m = 0; $y = 0
            if ([int]::TryParse($p[0],[ref]$m) -and [int]::TryParse($p[1],[ref]$y)) { return $y*12+$m }
        }
        return -1
    }
    $newKey = & $keyOf $SheetName
    $insertBefore = $null
    for ($s = 1; $s -le $cnt; $s++) {
        $sh = $Workbook.Sheets.Item($s)
        $k  = & $keyOf $sh.Name
        if ($k -ge 0 -and $k -gt $newKey) { $insertBefore = $sh; break }
    }
    $mv = [System.Reflection.Missing]::Value
    if ($null -ne $insertBefore) {
        $ws = Invoke-ComMethod -ComObject $Workbook.Sheets -MethodName Add -Arguments @($insertBefore, $mv, $mv, $mv)
    } else {
        $ws = Invoke-ComMethod -ComObject $Workbook.Sheets -MethodName Add -Arguments @($mv, $Workbook.Sheets.Item($cnt), $mv, $mv)
    }
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
    Ensure-FileWritable -FilePath $Path
    $firstStt   = [ExcelFinder]::AppendBarcodes($Path, $timestamps, $Barcodes, $Scanners, $Cols, $SheetDate)
    if ($firstStt -ge 0) {
        # Live Excel dang mo → dong hidden wb tranh file lock conflict
        Close-HiddenWorkbook
        for ($i = 0; $i -lt $Barcodes.Length; $i++) {
            Write-Log "[$($Scanners[$i])] Ghi STT $($firstStt + $i): $($Barcodes[$i])"
        }
        return
    }

    # Khong co live Excel → dung hidden workbook (giu mo giua cac lan flush)
    $wb = Get-HiddenWorkbook -Path $Path -SheetDate $SheetDate
    $ws      = Find-OrCreateSheet -Workbook $wb -SheetName $SheetDate
    $nextRow = [Math]::Max(2, $ws.UsedRange.Rows.Count + 1)

    for ($i = 0; $i -lt $Barcodes.Length; $i++) {
        $scanCol = 2 + $Cols[$i]

        if ([string]::IsNullOrWhiteSpace($ws.Cells.Item(1, $scanCol).Value2)) {
            $ws.Cells.Item(1, $scanCol)             = $Scanners[$i]
            $ws.Columns.Item($scanCol).NumberFormat = "@"
            $ws.Cells.Item(1, $scanCol).Font.Bold   = $true
        }

        $ts  = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
        $stt = $nextRow - 1
        $ws.Cells.Item($nextRow, 1)                     = $stt
        $ws.Cells.Item($nextRow, 2)                     = $ts
        $ws.Cells.Item($nextRow, $scanCol).NumberFormat = "@"
        $ws.Cells.Item($nextRow, $scanCol).Value2       = $Barcodes[$i]
        $nextRow++
        Write-Log "[$($Scanners[$i])] Ghi STT ${stt}: $($Barcodes[$i])"
    }

    # Try save with retry logic - use SaveAs as fallback
    $saveOk = $false
    for ($retry = 0; $retry -lt 3; $retry++) {
        try {
            Ensure-FileWritable -FilePath $Path
            
            # Try regular Save first
            Invoke-ComMethod -ComObject $wb -MethodName Save | Out-Null
            $saveOk = $true
            Write-Log "OK: Luu file thanh cong"
            break
        } catch {
            Write-Log "Canh bao: Loi Save (lan $($retry+1)/3): $_"
            
            # If regular Save fails, try SaveAs with explicit parameters
            if ($retry -eq 1) {
                try {
                    Write-Log "Thu SaveAs voi xlOpenXMLWorkbook..."
                    Ensure-FileWritable -FilePath $Path
                    # SaveAs: FileName, FileFormat, Password, WriteResPassword, ReadOnlyRecommended, CreateBackup, AccessMode
                    # AccessMode: 1=xlShared, 2=xlExclusive
                    Invoke-ComMethod -ComObject $wb -MethodName SaveAs -Arguments @($Path, 51, "", "", $false, $false, 1) | Out-Null  # 51 = xlOpenXMLWorkbook, 1 = xlShared
                    $saveOk = $true
                    Write-Log "OK: SaveAs thanh cong"
                    break
                } catch {
                    Write-Log "SaveAs cung that bai: $_"
                }
            }
            
            # If both fail, close and retry with fresh workbook
            if ($retry -lt 2) {
                Write-Log "Dong workbook va thu lai..."
                Close-HiddenWorkbook
                $script:wb = $null
                Start-Sleep -Milliseconds 500
                $wb = Get-HiddenWorkbook -Path $Path -SheetDate $SheetDate
                $ws = Find-OrCreateSheet -Workbook $wb -SheetName $SheetDate
            } else {
                # Last resort: try SaveAs with xlWorkbookDefault format
                try {
                    Write-Log "Thu SaveAs voi xlWorkbookDefault (format 56)..."
                    Ensure-FileWritable -FilePath $Path
                    Invoke-ComMethod -ComObject $wb -MethodName SaveAs -Arguments @($Path, 56) | Out-Null  # 56 = xlWorkbookDefault
                    $saveOk = $true
                    Write-Log "OK: SaveAs (format 56) thanh cong"
                } catch {
                    Write-Log "SaveAs format 56 that bai: $_"
                }
            }
        }
    }
    
    if (-not $saveOk) {
        Write-Log "LOI: Khong the save file sau $($retry+1) lan thu"
    }
}

# ----------------------------------------------------------------
# Buffer
# ----------------------------------------------------------------
$script:pendingBarcodes = [System.Collections.Generic.List[string]]::new()
$script:pendingScanners = [System.Collections.Generic.List[string]]::new()
$script:pendingCols     = [System.Collections.Generic.List[int]]::new()
$script:lastInputAt     = [DateTime]::MinValue
$script:cleanupStarted  = $false
$script:testInjectQueueFile = Join-Path $PSScriptRoot "test_inject.queue"
$FLUSH_INTERVAL_MS      = 200

function Add-PendingEntry {
    param([string]$Entry)
    $tabIdx = $Entry.IndexOf("`t")
    if ($tabIdx -lt 0) { return }
    $meta    = $Entry.Substring(0, $tabIdx)   # "name|colIdx"
    $barcode = $Entry.Substring($tabIdx + 1)
    $pipeIdx = $meta.LastIndexOf('|')
    if ($pipeIdx -lt 0 -or [string]::IsNullOrWhiteSpace($barcode)) { return }
    $displayName = $meta.Substring(0, $pipeIdx)
    $colIdx      = [int]$meta.Substring($pipeIdx + 1)
    $script:pendingBarcodes.Add($barcode)
    $script:pendingScanners.Add($displayName)
    $script:pendingCols.Add($colIdx)
    $script:lastInputAt = [DateTime]::Now
}

function Drain-TestInjectQueue {
    if (-not (Test-Path $script:testInjectQueueFile)) { return @() }

    $processingFile = "$($script:testInjectQueueFile).$([guid]::NewGuid().ToString('N')).processing"
    try {
        Move-Item -LiteralPath $script:testInjectQueueFile -Destination $processingFile -ErrorAction Stop
    } catch {
        return @()
    }

    try {
        return @(Get-Content -LiteralPath $processingFile -ErrorAction SilentlyContinue)
    } finally {
        Remove-Item -LiteralPath $processingFile -Force -ErrorAction SilentlyContinue
    }
}

# ----------------------------------------------------------------
# Khoi dong
# ----------------------------------------------------------------
# Xoa log cu moi lan khoi dong lai
try { [System.IO.File]::WriteAllText($LogFile, "", [System.Text.Encoding]::UTF8) } catch {}

Write-Log "=== USB Reader khoi dong | ScannerSpeed: ${ScannerSpeedMs}ms | MinLen: $MinBarcodeLength ==="
Write-Log "File: $ExcelFile | Flush interval: ${FLUSH_INTERVAL_MS}ms"

Ensure-FileWritable -FilePath $ExcelFile
Remove-Item -LiteralPath $script:testInjectQueueFile -Force -ErrorAction SilentlyContinue
Start-ExcelWorker
if (-not (Wait-ExcelWorkerReady -TimeoutMs 15000)) {
    Write-Log "Canh bao: Hidden Excel khoi dong cham hoac that bai"
}

if (-not (Test-Path $ExcelFile)) {
    Queue-ExcelFlush -Path $ExcelFile -Barcodes @() -Scanners @() -Cols @() -SheetDate (Get-SheetDate)
    [void](Wait-ExcelWorkerIdle -TimeoutMs 15000)
}

Flush-ExcelWorkerLogs
Write-Log "San sang."

# ----------------------------------------------------------------
# Graceful Shutdown Handler
# ================================================================
function Invoke-Cleanup {
    if ($script:cleanupStarted) { return }
    $script:cleanupStarted = $true
    Write-Log "Dang dong Excel va lam sach..."
    
    # Stop timer
    if ($null -ne $timer) { $timer.Stop() }
    
    # Unregister Raw Input
    [BarcodeRawInput]::Unregister()
    [KeyboardSuppressor]::Uninstall()
    Flush-ExcelWorkerLogs
    
    # Flush remaining barcodes
    if ($script:pendingBarcodes.Count -gt 0) {
        try {
            Write-Log "Flush du lieu con lai ($($script:pendingBarcodes.Count) records)..."
            Queue-ExcelFlush -Path $ExcelFile `
                -Barcodes  $script:pendingBarcodes.ToArray() `
                -Scanners  $script:pendingScanners.ToArray() `
                -Cols      $script:pendingCols.ToArray() `
                -SheetDate (Get-SheetDate)
            $script:pendingBarcodes.Clear()
            $script:pendingScanners.Clear()
            $script:pendingCols.Clear()
            [void](Wait-ExcelWorkerIdle -TimeoutMs 15000)
        } catch { Write-Log "Loi khi flush cuoi: $_" }
    }
    
    Stop-ExcelWorker
    Flush-ExcelWorkerLogs
    Write-Log "Da thoat."
}

# Trap Ctrl+C va System shutdown events
$null = Register-EngineEvent -SourceIdentifier PowerShell.Exiting -Action {
    Invoke-Cleanup
} -ErrorAction SilentlyContinue

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
    Flush-ExcelWorkerLogs

    # Thu thap barcode tu queue (format: "name|colIdx\tbarcode")
    [string]$entry = $null
    while ([BarcodeRawInput]::Queue.TryDequeue([ref]$entry)) {
        Add-PendingEntry -Entry $entry
    }

    foreach ($injectEntry in (Drain-TestInjectQueue)) {
        Add-PendingEntry -Entry $injectEntry
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

    if ($script:pendingBarcodes.Count -eq 0) { return }
    $idleMs = ([DateTime]::Now - $script:lastInputAt).TotalMilliseconds
    if ($idleMs -lt $FLUSH_INTERVAL_MS -and $script:pendingBarcodes.Count -lt 10) { return }

    try {
        $batchBarcodes = $script:pendingBarcodes.ToArray()
        $batchScanners = $script:pendingScanners.ToArray()
        $batchCols     = $script:pendingCols.ToArray()
        Queue-ExcelFlush -Path $ExcelFile -Barcodes $batchBarcodes -Scanners $batchScanners -Cols $batchCols -SheetDate (Get-SheetDate)
        $script:pendingBarcodes.Clear()
        $script:pendingScanners.Clear()
        $script:pendingCols.Clear()
    } catch {
        Write-Log "LOI flush (se thu lai): $_"
    }
})

$form.Add_FormClosed({
    Invoke-Cleanup
})

[BarcodeRawInput]::LoadMap("$PSScriptRoot\scanner_map.txt")
[BarcodeRawInput]::Register($form.Handle, $ScannerSpeedMs, $MinBarcodeLength)
[KeyboardSuppressor]::Install($ScannerSpeedMs)
if ([KeyboardSuppressor]::LastError) {
    Write-Log "CANH BAO: Keyboard suppress that bai: $([KeyboardSuppressor]::LastError)"
} else {
    Write-Log "Keyboard suppress: bat (threshold=${ScannerSpeedMs}ms, Ctrl/Alt/Win mien tru)"
}

$timer.Start()
Write-Log "Dang lang nghe ma vach (Raw Input)..."

[System.Windows.Forms.Application]::Run($form)
