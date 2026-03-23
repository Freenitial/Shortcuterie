<# :
    @echo off & Title Shortcuterie

    :: Windows version check (requires Windows 7 or later)
    for /f "tokens=2 delims=[]" %%v in ('ver') do for /f "tokens=2,3 delims=. " %%m in ("%%v") do (set "WINMAJOR=%%m" & set "WINMINOR=%%n")
    if not defined WINMAJOR set "WINMAJOR=0"
    if not defined WINMINOR set "WINMINOR=0"
    if %WINMAJOR% GTR 6 goto :winVersionOk
    if %WINMAJOR% EQU 6 if %WINMINOR% GEQ 1 goto :winVersionOk
    echo. & echo  [ERROR] This tool requires Windows 7 or later. & echo.
    pause
    exit /b 1
    :winVersionOk
    
    powershell -NoLogo -NoProfile -STA -Window Hidden -Command ^
        ^
        %= Create loading popup =% ^
        "$M=[Runtime.InteropServices.Marshal];" ^
        "$d=[AppDomain]::CurrentDomain.DefineDynamicAssembly(" ^
        "(New-Object Reflection.AssemblyName('W')),'Run').DefineDynamicModule('W');" ^
        "$t=$d.DefineType('A','Public,Class');" ^
        "$z=$t.DefinePInvokeMethod('CreateWindowExW','user32.dll'," ^
        "'Public,Static,PinvokeImpl','Standard',([IntPtr])," ^
        "@([Int32],[String],[String],[Int32],[Int32],[Int32],[Int32],[Int32]," ^
        "[IntPtr],[IntPtr],[IntPtr],[IntPtr]),'Winapi','Unicode');" ^
        "$z.SetImplementationFlags($z.GetMethodImplementationFlags()-bor128);" ^
        "$z=$t.DefinePInvokeMethod('ShowWindow','user32.dll'," ^
        "'Public,Static,PinvokeImpl','Standard',([Bool])," ^
        "@([IntPtr],[Int32]),'Winapi','Unicode');" ^
        "$z.SetImplementationFlags($z.GetMethodImplementationFlags()-bor128);" ^
        "$z=$t.DefinePInvokeMethod('GetSystemMetrics','user32.dll'," ^
        "'Public,Static,PinvokeImpl','Standard',([Int32])," ^
        "@([Int32]),'Winapi','Unicode');" ^
        "$z.SetImplementationFlags($z.GetMethodImplementationFlags()-bor128);" ^
        "$z=$t.DefinePInvokeMethod('SendMessageW','user32.dll'," ^
        "'Public,Static,PinvokeImpl','Standard',([IntPtr])," ^
        "@([IntPtr],[UInt32],[IntPtr],[IntPtr]),'Winapi','Unicode');" ^
        "$z.SetImplementationFlags($z.GetMethodImplementationFlags()-bor128);" ^
        "$z=$t.DefinePInvokeMethod('GetStockObject','gdi32.dll'," ^
        "'Public,Static,PinvokeImpl','Standard',([IntPtr])," ^
        "@([Int32]),'Winapi','Unicode');" ^
        "$z.SetImplementationFlags($z.GetMethodImplementationFlags()-bor128);" ^
        "$z=$t.DefinePInvokeMethod('InitCommonControlsEx','comctl32.dll'," ^
        "'Public,Static,PinvokeImpl','Standard',([Bool])," ^
        "@([IntPtr]),'Winapi','Unicode');" ^
        "$z.SetImplementationFlags($z.GetMethodImplementationFlags()-bor128);" ^
        "$A=$t.CreateType();" ^
        "$sw=$A::GetSystemMetrics(0);$sh=$A::GetSystemMetrics(1);" ^
        "$hw=$A::CreateWindowExW(9,'#32770','Shortcuterie',0x10C00000," ^
        "[int](($sw-440)/2),[int](($sh-130)/2),440,130," ^
        "[IntPtr]::Zero,[IntPtr]::Zero,[IntPtr]::Zero,[IntPtr]::Zero);" ^
        "$null=$A::ShowWindow($hw,5);" ^
        "$pc=$M::AllocHGlobal(8);$M::WriteInt32($pc,0,8);$M::WriteInt32($pc,4,0x20);" ^
        "$null=$A::InitCommonControlsEx($pc);$M::FreeHGlobal($pc);" ^
        "$ft=$A::GetStockObject(17);" ^
        "$hl=$A::CreateWindowExW(0,'Static','Initializing...',0x50000000," ^
        "20,15,390,20,$hw,[IntPtr]::Zero,[IntPtr]::Zero,[IntPtr]::Zero);" ^
        "$null=$A::SendMessageW($hl,0x30,$ft,[IntPtr]::Zero);" ^
        "$hb=$A::CreateWindowExW(0,'msctls_progress32','',0x50000000," ^
        "20,42,390,24,$hw,[IntPtr]::Zero,[IntPtr]::Zero,[IntPtr]::Zero);" ^
        ^
        %= PowerShell self-read, skipping batch part =% ^
        "$batFile='%~f0';& ([ScriptBlock]::Create([IO.File]::ReadAllText('%~f0')))"
    exit /b
#>

#region ── VERSION & PATHS ─

$script:AppName       = "Shortcuterie"
$script:Version       = [version]"1.0"

# ---- Remaining functions for Invoke-LoadingPump + updates ----
$t=$d.DefineType('E','Public,Class')
foreach($x in @(
    ,@('SetWindowTextW','user32.dll',([Bool]),@([IntPtr],[String]))
    ,@('DestroyWindow','user32.dll',([Bool]),@([IntPtr]))
    ,@('PeekMessageW','user32.dll',([Bool]),@([IntPtr],[IntPtr],[UInt32],[UInt32],[UInt32]))
    ,@('TranslateMessage','user32.dll',([Bool]),@([IntPtr]))
    ,@('DispatchMessageW','user32.dll',([IntPtr]),@([IntPtr]))
)){$z=$t.DefinePInvokeMethod($x[0],$x[1],'Public,Static,PinvokeImpl','Standard',$x[2],$x[3],'Winapi','Unicode');$z.SetImplementationFlags($z.GetMethodImplementationFlags()-bor128)}
$E=$t.CreateType()

$mg=$M::AllocHGlobal(48)
function Invoke-LoadingPump{try{while($E::PeekMessageW($mg,[IntPtr]::Zero,0,0,1)){$null=$E::TranslateMessage($mg);$null=$E::DispatchMessageW($mg)}}catch{}}
function Update-LoadingPopup([int]$pct,[string]$s){$null=$A::SendMessageW($hb,0x402,[IntPtr]$pct,[IntPtr]::Zero);if($s){$null=$E::SetWindowTextW($hl,$s)};try{Invoke-LoadingPump}catch{}}
function Close-LoadingPopup{$null=$E::DestroyWindow($hw);try{Invoke-LoadingPump}catch{};$M::FreeHGlobal($mg)}
Update-LoadingPopup 5  "Loading..."

$script:AppId         = "$($script:AppName).Freenitial.1" -replace " ",""
$script:LnkName       = "$($script:AppName).lnk"
$script:TaskbarPinDir = [IO.Path]::Combine($env:APPDATA, "Microsoft\Internet Explorer\Quick Launch\User Pinned\TaskBar")
$script:StartMenuDir  = [Environment]::GetFolderPath("Programs")

#region ── ASSEMBLIES & DPI ─

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# DPI awareness must be set before creating any window
if (-not ('DPIAware' -as [type])) {
    Add-Type -TypeDefinition @'
using System;
using System.Runtime.InteropServices;
public class DPIAware
{
    // Modern DPI awareness contexts (Windows 10 1703+)
    public static readonly IntPtr UNAWARE              = (IntPtr) (-1);
    public static readonly IntPtr SYSTEM_AWARE         = (IntPtr) (-2);
    public static readonly IntPtr PER_MONITOR_AWARE    = (IntPtr) (-3);
    public static readonly IntPtr PER_MONITOR_AWARE_V2 = (IntPtr) (-4);
    public static readonly IntPtr UNAWARE_GDISCALED    = (IntPtr) (-5);
    [DllImport("user32.dll", EntryPoint = "SetProcessDpiAwarenessContext", SetLastError = true)]
    private static extern bool NativeSetProcessDpiAwarenessContext(IntPtr Value);
    // Legacy API fallback (Windows Vista+)
    [DllImport("user32.dll", SetLastError = true)]
    private static extern bool SetProcessDPIAware();
    // Try modern API first, then legacy
    public static void SetDpiAwareness(IntPtr context)
    {
        if (!NativeSetProcessDpiAwarenessContext(context))
        {
            SetProcessDPIAware();
        }
    }
}
'@
}
try {[System.Windows.Forms.Application]::EnableVisualStyles()}      catch {}
try {[DPIAware]::SetDpiAwareness([DPIAware]::PER_MONITOR_AWARE_V2)} catch {}

function Get-DisplayPrimaryScaling {
    $VistaAndMore = [Environment]::OSVersion.Version.Major -ge 6
    if (-not $VistaAndMore) {
        try {
            $val = Get-ItemProperty -Path 'HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\FontDPI' -ErrorAction Stop
            if ($val -and $val.LogPixels -is [int] -and $val.LogPixels -gt 0) {return [math]::Round($val.LogPixels/96.0,2)}
        } catch { return 1.0 }
    }
    else {
        if (-not ('DPIHelper' -as [type])) {
        Add-Type @'
using System;
using System.Runtime.InteropServices;
using System.Drawing;
public static class DPIHelper {
    [DllImport("gdi32.dll")] static extern int GetDeviceCaps(IntPtr hdc, int nIndex);
    public enum DeviceCap { VERTRES = 10, DESKTOPVERTRES = 117, LOGPIXELSX = 88 }
    public static float GetScaling() {
        using (Graphics g = Graphics.FromHwnd(IntPtr.Zero)) {
            IntPtr hdc = g.GetHdc();
            try {
                int dpi = GetDeviceCaps(hdc, (int)DeviceCap.LOGPIXELSX);
                if (dpi > 0) { return (float)dpi / 96.0f; }
                int logical  = GetDeviceCaps(hdc, (int)DeviceCap.VERTRES);
                int physical = GetDeviceCaps(hdc, (int)DeviceCap.DESKTOPVERTRES);
                if (logical > 0 && physical > 0) { return (float)physical / (float)logical; }
                return 1.0f;
            } finally { g.ReleaseHdc(hdc); }
        }
    }
}
'@ -ReferencedAssemblies System.Drawing.dll
        }
    return [DPIHelper]::GetScaling()
    }
}
$script:DPI_Factor = Get-DisplayPrimaryScaling
write-host "DPI = $script:DPI_Factor"

$isAdmin = ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)

#region ── C# TYPES : TASKBAR / ICO / SHORTCUT ─

Add-Type -Language CSharp -ReferencedAssemblies System.Drawing -TypeDefinition @'
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Drawing.Imaging;
using System.IO;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using System.Text;

// ── Taskbar AppUserModelID ──
public static class TaskBarHelper
{
    [DllImport("shell32.dll", SetLastError = true)]
    private static extern void SetCurrentProcessExplicitAppUserModelID([MarshalAs(UnmanagedType.LPWStr)] string AppID);
    public static void SetAppId(string id) { SetCurrentProcessExplicitAppUserModelID(id); }
}

// ── Multi-size ICO builder from Bitmap, Base64, or executable ──
public static class IcoBuilder
{
    [DllImport("user32.dll", CharSet = CharSet.Unicode)]
    private static extern uint PrivateExtractIcons(
        string szFileName, int nIconIndex, int cxIcon, int cyIcon,
        IntPtr[] phicon, uint[] piconid, uint nIcons, uint flags);
    [DllImport("user32.dll")]
    private static extern bool DestroyIcon(IntPtr hIcon);
    [DllImport("kernel32.dll", SetLastError = true, CharSet = CharSet.Unicode)]
    private static extern IntPtr LoadLibraryEx(string lpFileName, IntPtr hFile, uint dwFlags);
    [DllImport("kernel32.dll")]
    private static extern bool FreeLibrary(IntPtr hModule);
    [DllImport("kernel32.dll", CharSet = CharSet.Unicode)]
    private static extern bool EnumResourceNames(IntPtr hModule, IntPtr lpType, EnumResNameProc lpEnumFunc, IntPtr lParam);
    private delegate bool EnumResNameProc(IntPtr hModule, IntPtr lpType, IntPtr lpName, IntPtr lParam);
    [DllImport("kernel32.dll", CharSet = CharSet.Unicode)]
    private static extern IntPtr FindResource(IntPtr hModule, IntPtr lpName, IntPtr lpType);
    [DllImport("kernel32.dll")]
    private static extern IntPtr LoadResource(IntPtr hModule, IntPtr hResInfo);
    [DllImport("kernel32.dll")]
    private static extern IntPtr LockResource(IntPtr hResData);
    [DllImport("kernel32.dll")]
    private static extern uint SizeofResource(IntPtr hModule, IntPtr hResInfo);
    private const uint LOAD_LIBRARY_AS_DATAFILE = 0x00000002;
    private static readonly int[] StandardSizes = new int[] { 16, 24, 32, 48, 64, 128, 256 };
    // Diagnostic message from the last operation (readable from PowerShell)
    private static string _lastDiagnostic;
    public static string LastDiagnostic { get { return _lastDiagnostic; } }
    // Helper to join int array as comma-separated string (no lambda)
    private static string JoinIntArray(int[] arr)
    {
        string[] parts = new string[arr.Length];
        for (int i = 0; i < arr.Length; i++) parts[i] = arr[i].ToString();
        return string.Join(",", parts);
    }
    // Build a multi-size ICO byte array from a Bitmap (upscale to all standard sizes)
    public static byte[] BuildFromBitmap(Bitmap source)
    {
        _lastDiagnostic = "BuildFromBitmap : source " + source.Width + "x" + source.Height;
        List<byte[]> pngEntries = new List<byte[]>();
        List<int> entrySizes = new List<int>();
        foreach (int sz in StandardSizes)
        {
            Bitmap bmp = new Bitmap(sz, sz, PixelFormat.Format32bppArgb);
            try
            {
                Graphics g = Graphics.FromImage(bmp);
                try
                {
                    g.Clear(Color.Transparent);
                    g.InterpolationMode  = InterpolationMode.HighQualityBicubic;
                    g.SmoothingMode      = SmoothingMode.HighQuality;
                    g.PixelOffsetMode    = PixelOffsetMode.HighQuality;
                    g.CompositingQuality = CompositingQuality.HighQuality;
                    g.DrawImage(source, 0, 0, sz, sz);
                }
                finally { g.Dispose(); }
                MemoryStream ms = new MemoryStream();
                try { bmp.Save(ms, ImageFormat.Png); pngEntries.Add(ms.ToArray()); entrySizes.Add(sz); }
                finally { ms.Dispose(); }
            }
            finally { bmp.Dispose(); }
        }
        return AssembleIco(pngEntries, entrySizes);
    }
    // Build a multi-size ICO from a base64-encoded image
    public static byte[] BuildFromBase64(string base64)
    {
        byte[] raw = Convert.FromBase64String(base64);
        MemoryStream ms = new MemoryStream(raw);
        try
        {
            Bitmap bmp = new Bitmap(ms);
            try { return BuildFromBitmap(bmp); }
            finally { bmp.Dispose(); }
        }
        finally { ms.Dispose(); }
    }
    // Build a multi-size ICO by extracting native sizes from an EXE/DLL via PE resource table
    public static byte[] BuildFromExecutable(string exePath, int iconIndex)
    {
        int maxNative = GetMaxNativeSize(exePath, iconIndex);
        List<byte[]> pngEntries = new List<byte[]>();
        List<int> entrySizes = new List<int>();
        foreach (int sz in StandardSizes)
        {
            if (sz > maxNative) break;
            IntPtr[] hIcons = new IntPtr[1];
            uint[] ids = new uint[1];
            uint count = PrivateExtractIcons(exePath, iconIndex, sz, sz, hIcons, ids, 1, 0);
            if (count > 0 && hIcons[0] != IntPtr.Zero)
            {
                Icon ico = Icon.FromHandle(hIcons[0]);
                try
                {
                    Bitmap trueColor = ico.ToBitmap();
                    try
                    {
                        Bitmap bmp = new Bitmap(trueColor.Width, trueColor.Height, PixelFormat.Format32bppArgb);
                        try
                        {
                            Graphics g = Graphics.FromImage(bmp);
                            try
                            {
                                g.Clear(Color.Transparent);
                                g.DrawImage(trueColor, 0, 0, trueColor.Width, trueColor.Height);
                            }
                            finally { g.Dispose(); }
                            MemoryStream ms = new MemoryStream();
                            try { bmp.Save(ms, ImageFormat.Png); pngEntries.Add(ms.ToArray()); entrySizes.Add(sz); }
                            finally { ms.Dispose(); }
                        }
                        finally { bmp.Dispose(); }
                    }
                    finally { trueColor.Dispose(); }
                }
                finally { ico.Dispose(); }
                DestroyIcon(hIcons[0]);
            }
        }
        if (pngEntries.Count == 0) return null;
        _lastDiagnostic += " | BuildFromExecutable : " + pngEntries.Count + " sizes extracted (max " + maxNative + "px)";
        return AssembleIco(pngEntries, entrySizes);
    }
    // Build ICO from executable, with fallback probing for resource IDs (negative indices)
    public static byte[] BuildFromExecutableEx(string exePath, int iconIndex)
    {
        // Try standard PE-based extraction first
        byte[] result = BuildFromExecutable(exePath, iconIndex);
        if (result != null) return result;
        // Fallback : probe standard sizes via PrivateExtractIcons (handles resource IDs)
        List<byte[]> pngEntries = new List<byte[]>();
        List<int> entrySizes = new List<int>();
        foreach (int sz in StandardSizes)
        {
            IntPtr[] hIcons = new IntPtr[1];
            uint[] ids = new uint[1];
            uint count = PrivateExtractIcons(exePath, iconIndex, sz, sz, hIcons, ids, 1, 0);
            if (count > 0 && hIcons[0] != IntPtr.Zero)
            {
                Icon ico = Icon.FromHandle(hIcons[0]);
                try
                {
                    Bitmap trueColor = ico.ToBitmap();
                    try
                    {
                        Bitmap bmp = new Bitmap(trueColor.Width, trueColor.Height, PixelFormat.Format32bppArgb);
                        try
                        {
                            Graphics g = Graphics.FromImage(bmp);
                            try
                            {
                                g.Clear(Color.Transparent);
                                g.DrawImage(trueColor, 0, 0, trueColor.Width, trueColor.Height);
                            }
                            finally { g.Dispose(); }
                            // Stop if we get upscaled duplicates (bitmap smaller than requested)
                            if (bmp.Width < sz && entrySizes.Count > 0) break;
                            MemoryStream ms = new MemoryStream();
                            try { bmp.Save(ms, ImageFormat.Png); pngEntries.Add(ms.ToArray()); entrySizes.Add(sz); }
                            finally { ms.Dispose(); }
                        }
                        finally { bmp.Dispose(); }
                    }
                    finally { trueColor.Dispose(); }
                }
                finally { ico.Dispose(); }
                DestroyIcon(hIcons[0]);
            }
        }
        if (pngEntries.Count == 0) return null;
        _lastDiagnostic = "BuildFromExecutableEx : " + pngEntries.Count + " sizes via PrivateExtractIcons probe";
        return AssembleIco(pngEntries, entrySizes);
    }
    // Validate that an icon can actually be extracted at the given index (supports negative resource IDs)
    public static bool CanExtractIcon(string exePath, int iconIndex)
    {
        IntPtr[] hIcons = new IntPtr[1];
        uint[] ids = new uint[1];
        uint count = PrivateExtractIcons(exePath, iconIndex, 16, 16, hIcons, ids, 1, 0);
        if (count > 0 && hIcons[0] != IntPtr.Zero)
        {
            DestroyIcon(hIcons[0]);
            return true;
        }
        return false;
    }
    // Extract a single icon at a specific size as a raw Bitmap (no ICO round-trip)
    public static Bitmap ExtractBitmapAtSize(string exePath, int iconIndex, int size)
    {
        IntPtr[] hIcons = new IntPtr[1];
        uint[] ids = new uint[1];
        uint count = PrivateExtractIcons(exePath, iconIndex, size, size, hIcons, ids, 1, 0);
        if (count > 0 && hIcons[0] != IntPtr.Zero)
        {
            Icon ico = Icon.FromHandle(hIcons[0]);
            Bitmap bmp = new Bitmap(ico.ToBitmap());
            ico.Dispose();
            DestroyIcon(hIcons[0]);
            return bmp;
        }
        return null;
    }
    // Read native icon sizes from the PE resource table (RT_GROUP_ICON)
    public static int[] GetNativeSizes(string exePath, int iconIndex)
    {
        List<int> sizes = new List<int>();
        IntPtr hModule = LoadLibraryEx(exePath, IntPtr.Zero, LOAD_LIBRARY_AS_DATAFILE);
        if (hModule == IntPtr.Zero)
        {
            int err = Marshal.GetLastWin32Error();
            _lastDiagnostic = "GetNativeSizes[PE] : LoadLibraryEx failed (error " + err + ")";
            return sizes.ToArray();
        }
        try
        {
            int currentGroupIndex = 0;
            string callbackError = null;
            // Capture hModule in a local for the delegate closure
            IntPtr hMod = hModule;
            EnumResNameProc callback = delegate(IntPtr hModule2, IntPtr lpType2, IntPtr lpName, IntPtr lParam) {
                if (currentGroupIndex != iconIndex) { currentGroupIndex++; return true; }
                IntPtr hRes = FindResource(hMod, lpName, (IntPtr)14);
                if (hRes == IntPtr.Zero) { callbackError = "FindResource failed"; return false; }
                IntPtr hResData = LoadResource(hMod, hRes);
                if (hResData == IntPtr.Zero) { callbackError = "LoadResource failed"; return false; }
                IntPtr pData = LockResource(hResData);
                uint resSize = SizeofResource(hMod, hRes);
                if (pData == IntPtr.Zero || resSize < 6)
                {
                    callbackError = "LockResource failed or resource too small (" + resSize + " bytes)";
                    return false;
                }
                short count = Marshal.ReadInt16(pData, 4);
                for (int i = 0; i < count; i++)
                {
                    int entryOffset = 6 + (i * 14);
                    if (entryOffset + 14 > resSize) break;
                    byte w = Marshal.ReadByte(pData, entryOffset);
                    int size = (w == 0) ? 256 : (int)w;
                    if (!sizes.Contains(size)) sizes.Add(size);
                }
                return false;
            };
            EnumResourceNames(hModule, (IntPtr)14, callback, IntPtr.Zero);
            GC.KeepAlive(callback);
            if (callbackError != null)
            {
                _lastDiagnostic = "GetNativeSizes[PE] : " + callbackError + " for group " + iconIndex;
                return sizes.ToArray();
            }
            if (sizes.Count == 0 && currentGroupIndex < iconIndex)
            {
                _lastDiagnostic = "GetNativeSizes[PE] : iconIndex " + iconIndex
                    + " out of range (" + (currentGroupIndex + 1) + " groups found)";
                return sizes.ToArray();
            }
        }
        finally { FreeLibrary(hModule); }
        sizes.Sort();
        _lastDiagnostic = "GetNativeSizes[PE] : " + sizes.Count + " sizes found ["
            + JoinIntArray(sizes.ToArray()) + "]";
        return sizes.ToArray();
    }
    // Detect the largest native icon size from the PE resource table
    public static int GetMaxNativeSize(string exePath, int iconIndex)
    {
        int[] sizes = GetNativeSizes(exePath, iconIndex);
        if (sizes.Length > 0)
        {
            int maxSize = sizes[sizes.Length - 1];
            _lastDiagnostic = "GetMaxNativeSize : " + maxSize + "px via PE resource table | " + _lastDiagnostic;
            return maxSize;
        }
        _lastDiagnostic = "GetMaxNativeSize : no sizes found, defaulting to 48px | " + _lastDiagnostic;
        return 48;
    }
    // Return the number of icon resources inside an executable
    public static int GetIconCount(string exePath)
    {
        return (int)PrivateExtractIcons(exePath, 0, 32, 32, null, null, 0, 0);
    }
    // Assemble PNG entries into a valid ICO file
    private static byte[] AssembleIco(List<byte[]> pngEntries, List<int> entrySizes)
    {
        MemoryStream ms = new MemoryStream();
        try
        {
            BinaryWriter bw = new BinaryWriter(ms);
            bw.Write((short)0); bw.Write((short)1); bw.Write((short)pngEntries.Count);
            int offset = 6 + (16 * pngEntries.Count);
            for (int i = 0; i < pngEntries.Count; i++)
            {
                int sz = entrySizes[i];
                bw.Write((byte)(sz >= 256 ? 0 : sz)); bw.Write((byte)(sz >= 256 ? 0 : sz));
                bw.Write((byte)0); bw.Write((byte)0); bw.Write((short)1); bw.Write((short)32);
                bw.Write(pngEntries[i].Length); bw.Write(offset); offset += pngEntries[i].Length;
            }
            for (int i = 0; i < pngEntries.Count; i++) bw.Write(pngEntries[i]);
            return ms.ToArray();
        }
        finally { ms.Dispose(); }
    }
}

// ── Shortcut helper with ADS-embedded icon support ──
public static class ShortcutHelper
{
    [ComImport, Guid("000214F9-0000-0000-C000-000000000046"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    private interface IShellLink {
        void GetPath([Out, MarshalAs(UnmanagedType.LPWStr)] StringBuilder pszFile, int cch, IntPtr pfd, int fFlags);
        void GetIDList(out IntPtr ppidl); void SetIDList(IntPtr pidl);
        void GetDescription([Out, MarshalAs(UnmanagedType.LPWStr)] StringBuilder pszName, int cch);
        void SetDescription([MarshalAs(UnmanagedType.LPWStr)] string pszName);
        void GetWorkingDirectory([Out, MarshalAs(UnmanagedType.LPWStr)] StringBuilder pszDir, int cch);
        void SetWorkingDirectory([MarshalAs(UnmanagedType.LPWStr)] string pszDir);
        void GetArguments([Out, MarshalAs(UnmanagedType.LPWStr)] StringBuilder pszArgs, int cch);
        void SetArguments([MarshalAs(UnmanagedType.LPWStr)] string pszArgs);
        void GetHotkey(out short pwHotkey); void SetHotkey(short wHotkey);
        void GetShowCmd(out int piShowCmd); void SetShowCmd(int iShowCmd);
        void GetIconLocation([Out, MarshalAs(UnmanagedType.LPWStr)] StringBuilder pszIconPath, int cch, out int piIcon);
        void SetIconLocation([MarshalAs(UnmanagedType.LPWStr)] string pszIconPath, int iIcon);
        void SetRelativePath([MarshalAs(UnmanagedType.LPWStr)] string pszPathRel, int dwReserved);
        void Resolve(IntPtr hwnd, int fFlags);
        void SetPath([MarshalAs(UnmanagedType.LPWStr)] string pszFile);
    }
    [ComImport, Guid("886D8EEB-8CF2-4446-8D02-CDBA1DBDCF99"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    private interface IPropertyStore {
        int GetCount(out uint cProps); int GetAt(uint iProp, out PROPERTYKEY pkey);
        int GetValue(ref PROPERTYKEY key, out PROPVARIANT pv);
        int SetValue(ref PROPERTYKEY key, ref PROPVARIANT pv); int Commit();
    }
    [StructLayout(LayoutKind.Sequential, Pack = 4)]
    private struct PROPERTYKEY { public Guid fmtid; public uint pid; }
    [StructLayout(LayoutKind.Sequential)]
    private struct PROPVARIANT {
        public ushort vt; public ushort w1, w2, w3; public IntPtr pwszVal; private IntPtr _padding;
        public static PROPVARIANT FromString(string v) { PROPVARIANT p = new PROPVARIANT(); p.vt = 31; p.pwszVal = Marshal.StringToCoTaskMemUni(v); return p; }
    }
    [ComImport, Guid("00021401-0000-0000-C000-000000000046")] private class ShellLink { }
    private static PROPERTYKEY _pkeyAppUserModelId;
    static ShortcutHelper()
    {
        _pkeyAppUserModelId.fmtid = new Guid("9F4C2855-9F79-4B39-A8D0-E1D42DE1D5F3");
        _pkeyAppUserModelId.pid = 5;
    }
    [DllImport("shell32.dll")]
    private static extern int SHGetKnownFolderPath(ref Guid rfid, uint dwFlags, IntPtr hToken, out IntPtr ppszPath);
    [DllImport("shell32.dll", CharSet = CharSet.Unicode)]
    private static extern IntPtr ILCreateFromPathW(string pszPath);

    private static IntPtr ParseShellPath(string shellPath)
    {
        // Step 1 : direct parse attempts
        string[] candidates = new string[] {
            shellPath,
            "::" + shellPath,
            "shell:::" + shellPath
        };
        for (int i = 0; i < candidates.Length; i++)
        {
            uint sfgao;
            IntPtr pidl;
            int hr = SHParseDisplayName(candidates[i], IntPtr.Zero, out pidl, 0, out sfgao);
            if (hr == 0 && pidl != IntPtr.Zero) return pidl;
        }
        // Step 2 : resolve Known Folder GUID paths like {GUID}\relative.exe
        if (shellPath.Length > 38 && shellPath[0] == '{')
        {
            int closeBrace = shellPath.IndexOf('}');
            if (closeBrace == 37)
            {
                string guidStr = shellPath.Substring(0, 38);
                Guid folderId;
                try { folderId = new Guid(guidStr); } catch { return IntPtr.Zero; }
                IntPtr ppszPath;
                int hr = SHGetKnownFolderPath(ref folderId, 0, IntPtr.Zero, out ppszPath);
                if (hr == 0 && ppszPath != IntPtr.Zero)
                {
                    string folderPath = Marshal.PtrToStringUni(ppszPath);
                    CoTaskMemFree(ppszPath);
                    string remainder = shellPath.Substring(38).TrimStart('\\');
                    string fullPath = string.IsNullOrEmpty(remainder)
                        ? folderPath
                        : System.IO.Path.Combine(folderPath, remainder);
                    // Filesystem path : ILCreateFromPathW
                    IntPtr pidl2 = ILCreateFromPathW(fullPath);
                    if (pidl2 != IntPtr.Zero) return pidl2;
                    // Fallback : parse resolved full path
                    uint sfgao2;
                    hr = SHParseDisplayName(fullPath, IntPtr.Zero, out pidl2, 0, out sfgao2);
                    if (hr == 0 && pidl2 != IntPtr.Zero) return pidl2;
                }
            }
        }
        return IntPtr.Zero;
    }
    private static void ApplyAppId(IShellLink link, string appId)
    {
        if (string.IsNullOrEmpty(appId)) return;
        IPropertyStore store = (IPropertyStore)link;
        PROPVARIANT pv = PROPVARIANT.FromString(appId);
        PROPERTYKEY key = _pkeyAppUserModelId;
        store.SetValue(ref key, ref pv);
        store.Commit();
        Marshal.FreeCoTaskMem(pv.pwszVal);
    }
    public static void CreateWithEmbeddedIcon(string lnkPath, string targetPath, string arguments, byte[] icoBytes, string appId, string description)
    {
        CreateWithEmbeddedIcon(lnkPath, targetPath, arguments, icoBytes, appId, description, null);
    }
    public static void CreateWithEmbeddedIcon(string lnkPath, string targetPath, string arguments, byte[] icoBytes, string appId, string description, string workDir)
    {
        IShellLink link = (IShellLink)new ShellLink();
        link.SetPath(targetPath);
        link.SetArguments(arguments == null ? "" : arguments);
        link.SetIconLocation(lnkPath + ":icon.ico", 0);
        link.SetDescription(description == null ? "" : description);
        if (!string.IsNullOrEmpty(workDir)) link.SetWorkingDirectory(workDir);
        ApplyAppId(link, appId);
        ((IPersistFile)link).Save(lnkPath, true);
    }
    public static void CreateWithStandardIcon(string lnkPath, string targetPath, string arguments, string iconFilePath, int iconIndex, string appId, string description)
    {
        CreateWithStandardIcon(lnkPath, targetPath, arguments, iconFilePath, iconIndex, appId, description, null);
    }
    public static void CreateWithStandardIcon(string lnkPath, string targetPath, string arguments, string iconFilePath, int iconIndex, string appId, string description, string workDir)
    {
        IShellLink link = (IShellLink)new ShellLink();
        link.SetPath(targetPath);
        link.SetArguments(arguments == null ? "" : arguments);
        link.SetIconLocation(iconFilePath, iconIndex);
        link.SetDescription(description == null ? "" : description);
        if (!string.IsNullOrEmpty(workDir)) link.SetWorkingDirectory(workDir);
        ApplyAppId(link, appId);
        ((IPersistFile)link).Save(lnkPath, true);
    }
    public static void UpdateIconOnly(string lnkPath)
    {
        IShellLink link = (IShellLink)new ShellLink();
        ((IPersistFile)link).Load(lnkPath, 0);
        link.SetIconLocation(lnkPath + ":icon.ico", 0);
        ((IPersistFile)link).Save(lnkPath, true);
    }
    public static string GetDescription(string lnkPath)
    {
        IShellLink link = (IShellLink)new ShellLink();
        ((IPersistFile)link).Load(lnkPath, 0);
        StringBuilder sb = new StringBuilder(1024);
        link.GetDescription(sb, sb.Capacity);
        return sb.ToString();
    }
    public static string GetTargetPath(string lnkPath)
    {
        IShellLink link = (IShellLink)new ShellLink();
        ((IPersistFile)link).Load(lnkPath, 0);
        StringBuilder sb = new StringBuilder(260);
        link.GetPath(sb, sb.Capacity, IntPtr.Zero, 0);
        return sb.ToString();
    }
    public static string GetArguments(string lnkPath)
    {
        IShellLink link = (IShellLink)new ShellLink();
        ((IPersistFile)link).Load(lnkPath, 0);
        StringBuilder sb = new StringBuilder(32768);
        link.GetArguments(sb, sb.Capacity);
        return sb.ToString();
    }
    public static string GetWorkingDirectory(string lnkPath)
    {
        IShellLink link = (IShellLink)new ShellLink();
        ((IPersistFile)link).Load(lnkPath, 0);
        StringBuilder sb = new StringBuilder(260);
        link.GetWorkingDirectory(sb, sb.Capacity);
        return sb.ToString();
    }
    public static string GetIconPath(string lnkPath)
    {
        IShellLink link = (IShellLink)new ShellLink();
        ((IPersistFile)link).Load(lnkPath, 0);
        StringBuilder sb = new StringBuilder(260);
        int index;
        link.GetIconLocation(sb, sb.Capacity, out index);
        return sb.ToString();
    }
    public static int GetIconIndex(string lnkPath)
    {
        IShellLink link = (IShellLink)new ShellLink();
        ((IPersistFile)link).Load(lnkPath, 0);
        StringBuilder sb = new StringBuilder(260);
        int index;
        link.GetIconLocation(sb, sb.Capacity, out index);
        return index;
    }
    public static string GetAppUserModelId(string lnkPath)
    {
        IShellLink link = (IShellLink)new ShellLink();
        ((IPersistFile)link).Load(lnkPath, 0);
        IPropertyStore store = (IPropertyStore)link;
        PROPERTYKEY key = _pkeyAppUserModelId;
        PROPVARIANT pv;
        int hr = store.GetValue(ref key, out pv);
        if (hr == 0 && pv.vt == 31 && pv.pwszVal != IntPtr.Zero)
        {
            string result = Marshal.PtrToStringUni(pv.pwszVal);
            Marshal.FreeCoTaskMem(pv.pwszVal);
            return result;
        }
        return null;
    }
    // ── PIDL resolution (shell CLSID shortcuts) ──
    [DllImport("shell32.dll", CharSet = CharSet.Unicode)]
    private static extern int SHGetNameFromIDList(IntPtr pidl, uint sigdnName, out IntPtr ppszName);
    [DllImport("shell32.dll", CharSet = CharSet.Unicode)]
    private static extern int SHParseDisplayName(string pszName, IntPtr pbc, out IntPtr ppidl, uint sfgaoIn, out uint psfgaoOut);
    [DllImport("ole32.dll")]
    private static extern void CoTaskMemFree(IntPtr pv);
    private const uint SIGDN_DESKTOPABSOLUTEPARSING = 0x80028000;
    // Resolve PIDL from an existing .lnk to a parseable shell path
    public static string GetParsedDisplayName(string lnkPath)
    {
        IShellLink link = (IShellLink)new ShellLink();
        ((IPersistFile)link).Load(lnkPath, 0);
        IntPtr pidl;
        link.GetIDList(out pidl);
        if (pidl == IntPtr.Zero) return null;
        try
        {
            IntPtr pszName;
            int hr = SHGetNameFromIDList(pidl, SIGDN_DESKTOPABSOLUTEPARSING, out pszName);
            if (hr != 0 || pszName == IntPtr.Zero) return null;
            string name = Marshal.PtrToStringUni(pszName);
            CoTaskMemFree(pszName);
            return name;
        }
        finally { CoTaskMemFree(pidl); }
    }
    // Create a shortcut from a shell path (CLSID) with embedded ADS icon
    public static void CreateShellWithEmbeddedIcon(string lnkPath, string shellPath, byte[] icoBytes, string appId, string description)
    {
        IntPtr pidl = ParseShellPath(shellPath);
        if (pidl == IntPtr.Zero)
            throw new System.ComponentModel.Win32Exception(0, "SHParseDisplayName failed for: " + shellPath);
        IShellLink link = (IShellLink)new ShellLink();
        link.SetIDList(pidl);
        CoTaskMemFree(pidl);
        link.SetIconLocation(lnkPath + ":icon.ico", 0);
        link.SetDescription(description == null ? "" : description);
        ApplyAppId(link, appId);
        ((IPersistFile)link).Save(lnkPath, true);
    }
    // Create a shortcut from a shell path (CLSID) with standard icon reference
    public static void CreateShellWithStandardIcon(string lnkPath, string shellPath, string iconFilePath, int iconIndex, string appId, string description)
    {
        IntPtr pidl = ParseShellPath(shellPath);
        if (pidl == IntPtr.Zero)
            throw new System.ComponentModel.Win32Exception(0, "SHParseDisplayName failed for: " + shellPath);
        IShellLink link = (IShellLink)new ShellLink();
        link.SetIDList(pidl);
        CoTaskMemFree(pidl);
        link.SetIconLocation(iconFilePath, iconIndex);
        link.SetDescription(description == null ? "" : description);
        ApplyAppId(link, appId);
        ((IPersistFile)link).Save(lnkPath, true);
    }
}
'@
[TaskBarHelper]::SetAppId($script:AppId)

#region ── ICON ─

$iconBase64 = ""
if ([string]::IsNullOrEmpty($iconBase64)) {
    $bmp = New-Object System.Drawing.Bitmap(96, 96)
    $bmp.SetResolution(96, 96)
    $g   = [System.Drawing.Graphics]::FromImage($bmp)
    $g.Clear([System.Drawing.Color]::FromArgb(0, 120, 212))
    $f  = New-Object System.Drawing.Font("Segoe UI", 72, [System.Drawing.FontStyle]::Bold)
    $sf = New-Object System.Drawing.StringFormat
    $sf.Alignment = $sf.LineAlignment = [System.Drawing.StringAlignment]::Center
    $g.DrawString("S", $f, [System.Drawing.Brushes]::White, (New-Object System.Drawing.RectangleF(0, 0, 96, 96)), $sf)
    $f.Dispose(); $sf.Dispose(); $g.Dispose()
    $iconImage = $bmp
    $script:AppIcoBytes = [IcoBuilder]::BuildFromBitmap($bmp)
} else {
    $iconBytes  = [Convert]::FromBase64String($iconBase64)
    $iconStream = New-Object IO.MemoryStream(,$iconBytes)
    $iconImage  = [System.Drawing.Image]::FromStream($iconStream)
    $script:AppIcoBytes = [IcoBuilder]::BuildFromBase64($iconBase64)
}

#region ── LOGGING ─

$script:LogDir  = [System.IO.Path]::Combine($env:TEMP, $script:AppName)
if (!([System.IO.Directory]::Exists($script:LogDir))) { [System.IO.Directory]::CreateDirectory($script:LogDir) | Out-Null }
$script:LogFile = [System.IO.Path]::Combine($script:LogDir, "$($script:AppName)_$(Get-Date -Format 'yyyyMMdd').log")
if ([System.IO.File]::Exists($script:LogFile)) {
    [System.IO.File]::AppendAllText($script:LogFile, "`n`n`n------------------------------`n`n`n")
}
$logFiles = [System.IO.Directory]::GetFiles($script:LogDir, "*.log")
if ($logFiles.Count -gt 10) {
    $sorted = [System.Array]::CreateInstance([System.IO.FileInfo], $logFiles.Count)
    for ($i = 0; $i -lt $logFiles.Count; $i++) { $sorted[$i] = New-Object System.IO.FileInfo($logFiles[$i]) }
    [System.Array]::Sort($sorted, [System.Comparison[System.IO.FileInfo]]{ param($a, $b) $b.LastWriteTimeUtc.CompareTo($a.LastWriteTimeUtc) })
    for ($i = 10; $i -lt $sorted.Count; $i++) { [System.IO.File]::Delete($sorted[$i].FullName) }
}
function Write-Log {
    param([string]$Message, [ValidateSet('Info','Warning','Error','Debug')][string]$Level = 'Info')
    if ([string]::IsNullOrEmpty($script:LogFile)) { return }
    $ts  = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $msg = "[$ts] [$Level] $Message"
    switch ($Level) {
        'Error'   { Write-Host $msg -ForegroundColor Red }
        'Warning' { Write-Host $msg -ForegroundColor Yellow }
        'Debug'   { Write-Host $msg -ForegroundColor Gray }
        default   { Write-Host $msg -ForegroundColor White }
    }
    try { [IO.File]::AppendAllText($script:LogFile, "$msg`r`n") } catch {}
}
Write-Log "═══ $($script:AppName) v$($script:Version) started ═══"
Write-Log "PowerShell version : $($PSVersionTable.PSVersion)"
Write-Log "CLR version : $([Environment]::Version)"
Write-Log "OS : $([Environment]::OSVersion.VersionString)"
if ($isAdmin) { Write-Log "Running with administrator privileges" }
else          { Write-Log "Running without administrator privileges"}

Update-LoadingPopup 20  "Loading..."

#region ── PS2.0 HELPERS ─

# .NET 3.5 does not have [string]::IsNullOrWhiteSpace
function Test-StringEmpty {
    param([string]$Value)
    if ($null -eq $Value) { return $true }
    return ($Value.Trim().Length -eq 0)
}

# Check whether the icon source panel has no valid icon selected
function Test-IconSourceEmpty {
    if ($radioIcon_TargetDefault.Checked) {
        return $false
    }
    if ($radioIcon_Base64.Checked) {
        return (Test-StringEmpty $TextboxIcon_Base64.Text)
    }
    return (Test-StringEmpty (Get-CleanInput $iconPathTextbox.Text))
}

#region ── SCRIPT VARIABLES ─

$script:HitTestPassThruControls = New-Object System.Collections.Generic.List[System.Windows.Forms.Control]
$script:HitTestNativeWindows    = New-Object System.Collections.ArrayList
$script:CleanupDone             = $false

$script:UserPinnedStartMenu     = $false
$script:GroupPadding            = 10

$script:FormBorderPenLight = New-Object System.Drawing.Pen([System.Drawing.Color]::FromArgb(200,200,200), 1)
$script:FormBorderPenDark  = New-Object System.Drawing.Pen([System.Drawing.Color]::FromArgb(60,60,60), 1)
$script:DropZonePen        = New-Object System.Drawing.Pen([System.Drawing.Color]::FromArgb(255, 200, 0), 2)
$script:GroupBoxBorderPen  = New-Object System.Drawing.Pen([System.Drawing.SystemColors]::ControlDark)

$script:CurrentIconIndex    = 0
$script:CurrentExeIconCount = 0
$script:SuppressPreviewUpdate   = $false

$script:AdsProbeCache = @{}

$script:AumidCleanRegex = New-Object System.Text.RegularExpressions.Regex('[^a-zA-Z0-9.\-]', [System.Text.RegularExpressions.RegexOptions]::Compiled)
$script:AumidLeadDotRegex = New-Object System.Text.RegularExpressions.Regex('^\.+', [System.Text.RegularExpressions.RegexOptions]::Compiled)

$script:LastIconMenuPath  = ""
$script:ShellTargetIconCache = $null
$script:LastPinnedTaskbarFile = ""

# Shortcut limits (MS-SHLLINK spec)
$script:MaxTargetPath        = 260
$script:MaxArgsCreateProcess = 32767
$script:MaxArgsCmdExe        = 8191

# Supported file extensions
$script:ImageExtensions      = @('.png', '.jpg', '.jpeg', '.bmp', '.gif', '.ico', '.tiff')
$script:ExeExtensions        = @('.exe', '.dll', '.ocx', '.cpl', '.scr')
$script:StandardIconExt      = @('.ico', '.exe', '.dll', '.ocx', '.cpl', '.scr')
$script:CurrentPreviewBitmap = $null
$script:ActiveDropZone       = $null
$script:SuppressAutoFill     = $false
$script:PreviousTargetText   = ""
$script:SuppressTargetSplit  = $false
$script:ArgsFromAutoSplit    = $false
$script:ShellTargetRegexOpts = [System.Text.RegularExpressions.RegexOptions]::Compiled -bor [System.Text.RegularExpressions.RegexOptions]::IgnoreCase
$script:ShellTargetRegex     = New-Object System.Text.RegularExpressions.Regex('^(::\{|\{[0-9a-fA-F]{8}-|shell:::\{|shell:[a-zA-Z])', $script:ShellTargetRegexOpts)
$script:IsUrlTarget          = $false
$script:IsShellTarget        = $false
$script:UrlTargetRegex       = New-Object System.Text.RegularExpressions.Regex('^[a-zA-Z][a-zA-Z0-9+.\-]*://', $script:ShellTargetRegexOpts)
$script:ExtSplitRegex        = New-Object System.Text.RegularExpressions.Regex('(?i)\.(exe|bat|cmd|ps1|vbs|com|msi|wsf|scr|cpl)\b', [System.Text.RegularExpressions.RegexOptions]::Compiled)
$script:ProbeExts = if (-not [string]::IsNullOrEmpty($env:PATHEXT)) {
    $env:PATHEXT.Split(';') | Where-Object { $_.Length -gt 0 }
} else {
    [string[]]@('.exe', '.bat', '.cmd', '.com')
}
$script:PathDirs = if (-not [string]::IsNullOrEmpty($env:PATH)) {
    $env:PATH.Split(';') | Where-Object { $_.Length -gt 0 -and [IO.Directory]::Exists($_) }
} else {
    [string[]]@()
}

# Theme color
$script:IsDarkMode       = $false
$script:AboutNormalColor = [System.Drawing.Color]::FromArgb(228, 228, 228)
$script:AboutHoverColor  = [System.Drawing.Color]::FromArgb(210, 210, 210)

#region ── C# TYPES : FORM & NATIVE ─

Update-LoadingPopup 30  "Loading..."

Add-Type -ReferencedAssemblies System.Windows.Forms.dll, System.Drawing.dll -TypeDefinition @"
using System;
using System.Windows.Forms;
using System.Drawing;
using System.Runtime.InteropServices;

// ── Borderless resizable form with WndProc event ──
public delegate void WndProcEventHandler(object sender, Message m);
public class CustomForm : Form
{
    public event WndProcEventHandler OnWindowMessage;
    public CustomForm()
    {
        SetStyle(
            ControlStyles.AllPaintingInWmPaint |
            ControlStyles.OptimizedDoubleBuffer,
            true);
        UpdateStyles();
    }
    protected override void WndProc(ref Message m)
    {
        base.WndProc(ref m);
        if (OnWindowMessage != null) OnWindowMessage(this, m);
    }
}

// ── DWM ROUNDED CORNERS ──
public static class DwmHelper {
    [DllImport("dwmapi.dll", CharSet=CharSet.Unicode, SetLastError=true)]
    private static extern long DwmSetWindowAttribute(IntPtr hwnd, uint attr, ref int val, uint sz);
    public static void SetRoundedCorners(Form f) { int v=2; DwmSetWindowAttribute(f.Handle,33,ref v,sizeof(int)); }
}

// ── Drag-and-drop fix for elevated (admin) processes ──
public class DragDropFix {
    [DllImport("shell32.dll")] public static extern void DragAcceptFiles(IntPtr hwnd, bool accept);
    [DllImport("shell32.dll")] public static extern uint DragQueryFile(IntPtr hDrop, uint iFile, [Out] System.Text.StringBuilder lpszFile, uint cch);
    [DllImport("shell32.dll")] public static extern void DragFinish(IntPtr hDrop);
    [DllImport("user32.dll")]  public static extern bool ChangeWindowMessageFilterEx(IntPtr hwnd, uint msg, uint action, IntPtr p);
    public static void Enable(IntPtr hwnd) {
        ChangeWindowMessageFilterEx(hwnd, 0x0233, 1, IntPtr.Zero);
        ChangeWindowMessageFilterEx(hwnd, 0x004A, 1, IntPtr.Zero);
        ChangeWindowMessageFilterEx(hwnd, 0x0049, 1, IntPtr.Zero);
        DragAcceptFiles(hwnd, true);
    }
    public static string[] GetDroppedFiles(IntPtr hDrop) {
        uint count = DragQueryFile(hDrop, 0xFFFFFFFF, null, 0);
        string[] files = new string[count];
        for (uint i = 0; i < count; i++) {
            uint size = DragQueryFile(hDrop, i, null, 0) + 1;
            System.Text.StringBuilder sb = new System.Text.StringBuilder((int)size);
            DragQueryFile(hDrop, i, sb, size);
            files[i] = sb.ToString();
        }
        DragFinish(hDrop);
        return files;
    }
}

// ── Native helpers ──
public class NativeMethods
{
    [DllImport("user32.dll", CharSet = CharSet.Auto)] public static extern int SendMessage(IntPtr hWnd, int msg, int wParam, int lParam);
    [DllImport("uxtheme.dll", CharSet = CharSet.Unicode)] public static extern int SetWindowTheme(IntPtr hwnd, string appName, string idList);
    [DllImport("kernel32.dll")] public static extern IntPtr GetConsoleWindow();
    [DllImport("user32.dll")]   public static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);
    [DllImport("user32.dll", SetLastError = true)] public static extern int GetWindowLong(IntPtr hWnd, int nIndex);
    [DllImport("user32.dll", SetLastError = true)] public static extern int SetWindowLong(IntPtr hWnd, int nIndex, int dwNewLong);
    public const int SW_HIDE = 0, SW_SHOW = 5;
    public const int GWL_EXSTYLE = -20;
    public const int WS_EX_COMPOSITED = 0x02000000;
    [DllImport("user32.dll", CharSet = CharSet.Unicode)]
    public static extern int SendMessageW(IntPtr hWnd, int msg, int wParam, string lParam);
}

// ── Dark mode : scrollbars + window frame ──
public static class DarkMode
{
    [DllImport("uxtheme.dll", EntryPoint = "#135", SetLastError = true)]
    private static extern int SetPreferredAppMode(int mode);
    [DllImport("dwmapi.dll")]
    private static extern int DwmSetWindowAttribute(IntPtr hwnd, int attr, ref int val, int size);
    public static void Init() { try { SetPreferredAppMode(1); } catch { } }
    public static void ApplyControl(IntPtr hwnd, bool dark) {
        NativeMethods.SetWindowTheme(hwnd, dark ? "DarkMode_Explorer" : "Explorer", null);
    }
    public static void ApplyWindowFrame(IntPtr hwnd, bool dark) {
        int v = dark ? 1 : 0;
        if (DwmSetWindowAttribute(hwnd, 20, ref v, sizeof(int)) != 0)
            DwmSetWindowAttribute(hwnd, 19, ref v, sizeof(int));
    }
}
"@ -Language CSharp

[DarkMode]::Init()

#region ── C# TYPES : TITLE BAR ─

Update-LoadingPopup 40 "Loading..."

Add-Type -ReferencedAssemblies System.Windows.Forms, System.Drawing -Language CSharp -TypeDefinition @'
using System; using System.Drawing; using System.Windows.Forms; using System.Runtime.InteropServices;
public class DoubleBufferedPanel : Panel {
    public DoubleBufferedPanel() { DoubleBuffered = true;
        SetStyle(ControlStyles.AllPaintingInWmPaint|ControlStyles.UserPaint|ControlStyles.OptimizedDoubleBuffer,true); }
}
public class HitTestPassThrough : NativeWindow {
    const int WM_NCHITTEST=0x84, HTTRANSPARENT=-1;
    protected override void WndProc(ref Message m) {
        if(m.Msg==WM_NCHITTEST){ m.Result=(IntPtr)HTTRANSPARENT; return; }
        base.WndProc(ref m); }
}
public class Win11TitleBar : Panel {
    [DllImport("user32.dll")] static extern int SendMessage(IntPtr h,int m,int w,int l);
    [DllImport("user32.dll")] static extern bool ReleaseCapture();
    const int WM_NCLBUTTONDOWN=0xA1, HT_CAPTION=2;
    public Win11TitleBar() { DoubleBuffered=true;
        SetStyle(ControlStyles.AllPaintingInWmPaint|ControlStyles.UserPaint|ControlStyles.OptimizedDoubleBuffer,true); }
    public static void DragForm(Form f){ ReleaseCapture(); SendMessage(f.Handle,WM_NCLBUTTONDOWN,HT_CAPTION,0); }
    protected override void OnMouseDown(MouseEventArgs e){ base.OnMouseDown(e); if(e.Button==MouseButtons.Left) DragForm(FindForm()); }
}
public class TitleBarButton : Label {
    private Color _hoverBack;
    private Color _normalBack;
    private Color _normalFore;
    private bool  _isCloseButton;
    public Color HoverBack  { get { return _hoverBack; }  set { _hoverBack = value; } }
    public Color NormalBack { get { return _normalBack; } set { _normalBack = value; } }
    public Color NormalFore { get { return _normalFore; } set { _normalFore = value; } }
    public bool  IsCloseButton { get { return _isCloseButton; } set { _isCloseButton = value; } }
    public TitleBarButton() { TextAlign=ContentAlignment.MiddleCenter;
        _normalBack=Color.FromArgb(240,240,240); _hoverBack=Color.FromArgb(218,218,218);
        _normalFore=Color.FromArgb(50,50,50);
        BackColor=_normalBack; ForeColor=_normalFore;
        Font=new Font("Segoe MDL2 Assets",10f); Cursor=Cursors.Default; Size=new Size(46,34); }
    protected override void OnMouseEnter(EventArgs e){ base.OnMouseEnter(e);
        BackColor=_isCloseButton?Color.FromArgb(196,43,28):_hoverBack;
        if(_isCloseButton) ForeColor=Color.White; }
    protected override void OnMouseLeave(EventArgs e){ base.OnMouseLeave(e);
        BackColor=_normalBack; ForeColor=_normalFore; }
}
'@

#region ── C# TYPES : ADS HELPER & ICON RESOLVER ─

Add-Type -ReferencedAssemblies System.Drawing.dll -Language CSharp -TypeDefinition @'
using System;
using System.IO;
using System.Drawing;
using System.Drawing.Imaging;
using System.Text;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using Microsoft.Win32.SafeHandles;

// ── NTFS Alternate Data Stream reader/writer via CreateFileW ──
public static class AdsHelper
{
    [DllImport("kernel32.dll", CharSet = CharSet.Unicode, SetLastError = true)]
    private static extern SafeFileHandle CreateFileW(
        string lpFileName, uint dwDesiredAccess, uint dwShareMode,
        IntPtr lpSecurityAttributes, uint dwCreationDisposition,
        uint dwFlagsAndAttributes, IntPtr hTemplateFile);
    private const uint GENERIC_READ    = 0x80000000;
    private const uint GENERIC_WRITE   = 0x40000000;
    private const uint CREATE_ALWAYS   = 2;
    private const uint OPEN_EXISTING   = 3;
    private const uint FILE_ATTRIBUTE_NORMAL = 0x80;
    public static void WriteStream(string filePath, string streamName, byte[] data)
    {
        string adsPath = filePath + ":" + streamName;
        using (SafeFileHandle h = CreateFileW(adsPath, GENERIC_WRITE, 0,
            IntPtr.Zero, CREATE_ALWAYS, FILE_ATTRIBUTE_NORMAL, IntPtr.Zero))
        {
            if (h.IsInvalid)
                throw new IOException("CreateFileW failed for : " + adsPath +
                    " (error " + Marshal.GetLastWin32Error() + ")");
            using (FileStream fs = new FileStream(h, FileAccess.Write))
                fs.Write(data, 0, data.Length);
        }
    }
    public static byte[] ReadStream(string filePath, string streamName)
    {
        string adsPath = filePath + ":" + streamName;
        using (SafeFileHandle h = CreateFileW(adsPath, GENERIC_READ, 1,
            IntPtr.Zero, OPEN_EXISTING, FILE_ATTRIBUTE_NORMAL, IntPtr.Zero))
        {
            if (h.IsInvalid) return null;
            using (FileStream fs = new FileStream(h, FileAccess.Read))
            {
                byte[] data = new byte[fs.Length];
                int offset = 0;
                while (offset < data.Length)
                {
                    int read = fs.Read(data, offset, data.Length - offset);
                    if (read <= 0) break;
                    offset += read;
                }
                return data;
            }
        }
    }
    public static long GetStreamLength(string filePath, string streamName)
    {
        string adsPath = filePath + ":" + streamName;
        using (SafeFileHandle h = CreateFileW(adsPath, GENERIC_READ, 1,
            IntPtr.Zero, OPEN_EXISTING, FILE_ATTRIBUTE_NORMAL, IntPtr.Zero))
        {
            if (h.IsInvalid) return -1;
            using (FileStream fs = new FileStream(h, FileAccess.Read))
                return fs.Length;
        }
    }
}

// ── Shell icon resolver for arbitrary file types ──
public static class IconResolver
{
    [DllImport("shell32.dll", CharSet = CharSet.Unicode)]
    private static extern IntPtr SHGetFileInfo(string pszPath, uint dwFileAttributes,
        ref SHFILEINFO psfi, uint cbFileInfo, uint uFlags);
    [DllImport("shell32.dll")]
    private static extern int SHGetImageList(int iImageList, ref Guid riid, out IntPtr ppvObj);
    [DllImport("comctl32.dll")]
    private static extern IntPtr ImageList_GetIcon(IntPtr himl, int i, uint flags);
    [DllImport("user32.dll")]
    private static extern bool DestroyIcon(IntPtr hIcon);
    [DllImport("shlwapi.dll", CharSet = CharSet.Unicode, SetLastError = true)]
    private static extern int AssocQueryString(uint flags, uint str,
        string pszAssoc, string pszExtra, StringBuilder pszOut, ref uint pcchOut);
    [StructLayout(LayoutKind.Sequential, CharSet = CharSet.Unicode)]
    private struct SHFILEINFO
    {
        public IntPtr hIcon;
        public int iIcon;
        public uint dwAttributes;
        [MarshalAs(UnmanagedType.ByValTStr, SizeConst = 260)]
        public string szDisplayName;
        [MarshalAs(UnmanagedType.ByValTStr, SizeConst = 80)]
        public string szTypeName;
    }
    private const uint SHGFI_ICON           = 0x000100;
    private const uint SHGFI_ICONLOCATION   = 0x001000;
    private const uint SHGFI_LARGEICON      = 0x000000;
    private const uint SHGFI_SYSICONINDEX   = 0x004000;
    private const uint ASSOCF_NONE          = 0x00000000;
    private const uint ASSOCSTR_DEFAULTICON = 5;
    private const int SHIL_EXTRALARGE = 2;
    private const int SHIL_JUMBO     = 4;
    private static readonly Guid IID_IImageList = new Guid("46EB5926-582E-4017-9FDF-E8998DAA0950");
    private static string _lastDiagnostic;
    public static string LastDiagnostic { get { return _lastDiagnostic; } }
    // Parse "path,index" format into path + integer index
    private static bool ParseIconReference(string raw, out string path, out int index)
    {
        path = null; index = 0;
        if (string.IsNullOrEmpty(raw)) return false;
        string expanded = Environment.ExpandEnvironmentVariables(raw.Trim().Trim('"'));
        int lastComma = expanded.LastIndexOf(',');
        if (lastComma > 0)
        {
            int idx;
            if (int.TryParse(expanded.Substring(lastComma + 1).Trim(), out idx))
            {
                index = idx;
                expanded = expanded.Substring(0, lastComma).Trim();
            }
        }
        if (File.Exists(expanded)) { path = expanded; return true; }
        return false;
    }
    // Step 1 : SHGetFileInfo SHGFI_ICONLOCATION
    public static string GetIconLocation(string filePath, out int iconIndex)
    {
        iconIndex = 0;
        SHFILEINFO shfi = new SHFILEINFO();
        IntPtr result = SHGetFileInfo(filePath, 0, ref shfi,
            (uint)Marshal.SizeOf(typeof(SHFILEINFO)), SHGFI_ICONLOCATION);
        string raw = shfi.szDisplayName;
        iconIndex = shfi.iIcon;
        _lastDiagnostic = "GetIconLocation : raw='" + (raw == null ? "null" : raw) + "' iIcon=" + shfi.iIcon + " result=" + result;
        if (result != IntPtr.Zero && !string.IsNullOrEmpty(raw))
        {
            string path = Environment.ExpandEnvironmentVariables(raw);
            if (File.Exists(path)) return path;
            _lastDiagnostic += " | expanded='" + path + "' (NOT FOUND)";
        }
        return null;
    }
    // Step 2 : AssocQueryString ASSOCSTR_DEFAULTICON
    public static string GetAssocDefaultIcon(string extension, out int iconIndex)
    {
        iconIndex = 0;
        uint size = 512;
        StringBuilder sb = new StringBuilder((int)size);
        int hr = AssocQueryString(ASSOCF_NONE, ASSOCSTR_DEFAULTICON,
            extension, null, sb, ref size);
        string raw = sb.ToString();
        _lastDiagnostic = "GetAssocDefaultIcon('" + extension + "') : hr=0x" + hr.ToString("X") + " raw='" + raw + "'";
        if (hr != 0 || sb.Length == 0) return null;
        string path;
        if (ParseIconReference(raw, out path, out iconIndex))
        {
            _lastDiagnostic += " | parsed='" + path + "' index=" + iconIndex;
            return path;
        }
        _lastDiagnostic += " | ParseIconReference failed";
        return null;
    }
    // Step 3 : Registry HKCR\.ext -> all ProgIDs -> DefaultIcon
    public static string[] GetRegistryDefaultIcons(string extension, out int primaryCount)
    {
        primaryCount = 0;
        List<string> results = new List<string>();
        List<string> diagnosticParts = new List<string>();
        diagnosticParts.Add("GetRegistryDefaultIcons('" + extension + "')");
        try
        {
            // Primary : direct DefaultIcon on extension key
            Microsoft.Win32.RegistryKey directKey = Microsoft.Win32.Registry.ClassesRoot.OpenSubKey(extension + "\\DefaultIcon");
            if (directKey != null)
            {
                try
                {
                    string val = directKey.GetValue(null) as string;
                    if (!string.IsNullOrEmpty(val) && !results.Contains(val))
                    {
                        results.Add(val);
                        diagnosticParts.Add("direct : '" + val + "'");
                    }
                }
                finally { directKey.Close(); }
            }
            // Collect all candidate ProgIDs
            List<string> primaryProgIds = new List<string>();
            List<string> secondaryProgIds = new List<string>();
            Microsoft.Win32.RegistryKey extKey = Microsoft.Win32.Registry.ClassesRoot.OpenSubKey(extension);
            if (extKey != null)
            {
                try
                {
                    string defaultProgId = extKey.GetValue(null) as string;
                    if (!string.IsNullOrEmpty(defaultProgId)) primaryProgIds.Add(defaultProgId);
                    // UserChoice ProgID is primary
                    try
                    {
                        Microsoft.Win32.RegistryKey ucKey = Microsoft.Win32.Registry.CurrentUser.OpenSubKey(
                            @"Software\Microsoft\Windows\CurrentVersion\Explorer\FileExts\" + extension + @"\UserChoice");
                        if (ucKey != null)
                        {
                            try
                            {
                                string ucProgId = ucKey.GetValue("ProgId") as string;
                                if (!string.IsNullOrEmpty(ucProgId) && !primaryProgIds.Contains(ucProgId))
                                    primaryProgIds.Add(ucProgId);
                            }
                            finally { ucKey.Close(); }
                        }
                    }
                    catch { }
                    // OpenWithProgids are secondary
                    Microsoft.Win32.RegistryKey owpKey = extKey.OpenSubKey("OpenWithProgids");
                    if (owpKey != null)
                    {
                        try
                        {
                            foreach (string name in owpKey.GetValueNames())
                            {
                                if (!string.IsNullOrEmpty(name) && !primaryProgIds.Contains(name) && !secondaryProgIds.Contains(name))
                                    secondaryProgIds.Add(name);
                            }
                        }
                        finally { owpKey.Close(); }
                    }
                }
                finally { extKey.Close(); }
            }
            // Conventional ProgID is secondary
            string conventional = extension.TrimStart('.') + "file";
            if (!primaryProgIds.Contains(conventional) && !secondaryProgIds.Contains(conventional))
                secondaryProgIds.Add(conventional);
            diagnosticParts.Add("primary=[" + string.Join(",", primaryProgIds.ToArray()) + "]");
            diagnosticParts.Add("secondary=[" + string.Join(",", secondaryProgIds.ToArray()) + "]");
            // Add primary ProgID DefaultIcons
            foreach (string progId in primaryProgIds)
            {
                Microsoft.Win32.RegistryKey iconKey = Microsoft.Win32.Registry.ClassesRoot.OpenSubKey(progId + "\\DefaultIcon");
                if (iconKey == null) { diagnosticParts.Add(progId + " : no DefaultIcon"); continue; }
                try
                {
                    string val = iconKey.GetValue(null) as string;
                    if (!string.IsNullOrEmpty(val) && !results.Contains(val))
                    {
                        results.Add(val);
                        diagnosticParts.Add(progId + " : '" + val + "'");
                    }
                    else { diagnosticParts.Add(progId + " : empty or duplicate"); }
                }
                finally { iconKey.Close(); }
            }
            primaryCount = results.Count;
            // Add secondary ProgID DefaultIcons
            foreach (string progId in secondaryProgIds)
            {
                Microsoft.Win32.RegistryKey iconKey = Microsoft.Win32.Registry.ClassesRoot.OpenSubKey(progId + "\\DefaultIcon");
                if (iconKey == null) { diagnosticParts.Add(progId + " : no DefaultIcon"); continue; }
                try
                {
                    string val = iconKey.GetValue(null) as string;
                    if (!string.IsNullOrEmpty(val) && !results.Contains(val))
                    {
                        results.Add(val);
                        diagnosticParts.Add(progId + " : '" + val + "'");
                    }
                    else { diagnosticParts.Add(progId + " : empty or duplicate"); }
                }
                finally { iconKey.Close(); }
            }
        }
        catch (Exception ex) { diagnosticParts.Add("exception : " + ex.Message); }
        _lastDiagnostic = string.Join(" | ", diagnosticParts.ToArray());
        return results.ToArray();
    }
    // Step 4 : High-quality shell bitmap via system image list (jumbo 256x256 or 48x48)
    public static Bitmap ExtractShellIcon(string filePath)
    {
        SHFILEINFO shfi = new SHFILEINFO();
        IntPtr result = SHGetFileInfo(filePath, 0, ref shfi,
            (uint)Marshal.SizeOf(typeof(SHFILEINFO)), SHGFI_SYSICONINDEX);
        if (result == IntPtr.Zero)
        {
            _lastDiagnostic = "ExtractShellIcon : SHGetFileInfo SYSICONINDEX failed";
            return null;
        }
        int sysIndex = shfi.iIcon;
        int[] listIds = new int[] { SHIL_JUMBO, SHIL_EXTRALARGE };
        for (int li = 0; li < listIds.Length; li++)
        {
            int listId = listIds[li];
            IntPtr imageList;
            Guid iid = IID_IImageList;
            if (SHGetImageList(listId, ref iid, out imageList) == 0 && imageList != IntPtr.Zero)
            {
                IntPtr hIcon = ImageList_GetIcon(imageList, sysIndex, 0);
                if (hIcon != IntPtr.Zero)
                {
                    Icon ico = Icon.FromHandle(hIcon);
                    Bitmap bmp = new Bitmap(ico.ToBitmap());
                    ico.Dispose();
                    DestroyIcon(hIcon);
                    if (bmp.Width >= 32)
                    {
                        _lastDiagnostic = "ExtractShellIcon : " + bmp.Width + "x" + bmp.Height
                            + " from SHIL_" + (listId == SHIL_JUMBO ? "JUMBO" : "EXTRALARGE")
                            + " sysIndex=" + sysIndex;
                        return bmp;
                    }
                    bmp.Dispose();
                }
            }
        }
        // Final fallback : basic 32x32
        result = SHGetFileInfo(filePath, 0, ref shfi,
            (uint)Marshal.SizeOf(typeof(SHFILEINFO)), SHGFI_ICON | SHGFI_LARGEICON);
        if (result != IntPtr.Zero && shfi.hIcon != IntPtr.Zero)
        {
            Bitmap bmp = Icon.FromHandle(shfi.hIcon).ToBitmap();
            DestroyIcon(shfi.hIcon);
            _lastDiagnostic = "ExtractShellIcon : " + bmp.Width + "x" + bmp.Height + " from SHGFI_ICON fallback";
            return bmp;
        }
        _lastDiagnostic = "ExtractShellIcon : all methods failed";
        return null;
    }
}

// ── Accept drop including UWP apps shortcuts ──
public static class ShellDropHelper
{
    [DllImport("shell32.dll")]
    private static extern IntPtr ILCombine(IntPtr pidl1, IntPtr pidl2);
    [DllImport("shell32.dll")]
    private static extern void ILFree(IntPtr pidl);
    [DllImport("shell32.dll", CharSet = CharSet.Unicode)]
    private static extern int SHGetNameFromIDList(IntPtr pidl, uint sigdnName, out IntPtr ppszName);
    [DllImport("shell32.dll", CharSet = CharSet.Unicode)]
    private static extern int SHParseDisplayName(string pszName, IntPtr pbc, out IntPtr ppidl, uint sfgaoIn, out uint psfgaoOut);
    [DllImport("shell32.dll", CharSet = CharSet.Unicode)]
    private static extern IntPtr SHGetFileInfo(IntPtr pidl, uint dwFileAttributes,
        ref SHFILEINFO psfi, uint cbFileInfo, uint uFlags);
    [DllImport("shell32.dll")]
    private static extern int SHGetImageList(int iImageList, ref Guid riid, out IntPtr ppvObj);
    [DllImport("comctl32.dll")]
    private static extern IntPtr ImageList_GetIcon(IntPtr himl, int i, uint flags);
    [DllImport("user32.dll")]
    private static extern bool DestroyIcon(IntPtr hIcon);
    [DllImport("ole32.dll")]
    private static extern void CoTaskMemFree(IntPtr pv);

    [StructLayout(LayoutKind.Sequential, CharSet = CharSet.Unicode)]
    private struct SHFILEINFO
    {
        public IntPtr hIcon;
        public int iIcon;
        public uint dwAttributes;
        [MarshalAs(UnmanagedType.ByValTStr, SizeConst = 260)] public string szDisplayName;
        [MarshalAs(UnmanagedType.ByValTStr, SizeConst = 80)] public string szTypeName;
    }

    private const uint SIGDN_DESKTOPABSOLUTEPARSING = 0x80028000;
    private const uint SIGDN_NORMALDISPLAY          = 0x00000000;
    private const uint SHGFI_PIDL         = 0x000000008;
    private const uint SHGFI_SYSICONINDEX = 0x000004000;
    private const int SHIL_JUMBO      = 4;
    private const int SHIL_EXTRALARGE = 2;
    private static readonly Guid IID_IImageList = new Guid("46EB5926-582E-4017-9FDF-E8998DAA0950");

    // Walk ITEMIDLIST chain to compute total byte size including terminator
    private static int GetPidlByteSize(byte[] data, int offset)
    {
        int pos = offset;
        while (pos + 2 <= data.Length)
        {
            ushort cb = BitConverter.ToUInt16(data, pos);
            if (cb == 0) return (pos - offset) + 2;
            pos += cb;
        }
        return pos - offset;
    }

    // Parse CIDA structure into an array of absolute PIDLs (caller must free each with ILFree)
    private static IntPtr[] ParseCida(byte[] cidaBytes, out int count)
    {
        count = 0;
        if (cidaBytes == null || cidaBytes.Length < 8) return null;
        count = (int)BitConverter.ToUInt32(cidaBytes, 0);
        if (count == 0) return null;
        int parentOffset = (int)BitConverter.ToUInt32(cidaBytes, 4);
        int parentSize = GetPidlByteSize(cidaBytes, parentOffset);
        IntPtr[] results = new IntPtr[count];
        for (int i = 0; i < count; i++)
        {
            int childOffset = (int)BitConverter.ToUInt32(cidaBytes, 8 + i * 4);
            int childSize = GetPidlByteSize(cidaBytes, childOffset);
            IntPtr parentPidl = Marshal.AllocHGlobal(parentSize);
            Marshal.Copy(cidaBytes, parentOffset, parentPidl, parentSize);
            IntPtr childPidl = Marshal.AllocHGlobal(childSize);
            Marshal.Copy(cidaBytes, childOffset, childPidl, childSize);
            results[i] = ILCombine(parentPidl, childPidl);
            Marshal.FreeHGlobal(parentPidl);
            Marshal.FreeHGlobal(childPidl);
        }
        return results;
    }

    // Extract highest-quality icon from a PIDL via system image list
    private static Bitmap ExtractIconFromPidl(IntPtr pidl)
    {
        if (pidl == IntPtr.Zero) return null;
        SHFILEINFO shfi = new SHFILEINFO();
        IntPtr res = SHGetFileInfo(pidl, 0, ref shfi,
            (uint)Marshal.SizeOf(typeof(SHFILEINFO)), SHGFI_PIDL | SHGFI_SYSICONINDEX);
        if (res == IntPtr.Zero) return null;
        int[] listIds = { SHIL_JUMBO, SHIL_EXTRALARGE };
        for (int li = 0; li < listIds.Length; li++)
        {
            IntPtr imageList;
            Guid iid = IID_IImageList;
            if (SHGetImageList(listIds[li], ref iid, out imageList) == 0 && imageList != IntPtr.Zero)
            {
                IntPtr hIcon = ImageList_GetIcon(imageList, shfi.iIcon, 0);
                if (hIcon != IntPtr.Zero)
                {
                    Icon ico = Icon.FromHandle(hIcon);
                    Bitmap raw = ico.ToBitmap();
                    ico.Dispose();
                    DestroyIcon(hIcon);
                    Bitmap result = new Bitmap(raw.Width, raw.Height, PixelFormat.Format32bppArgb);
                    using (Graphics g = Graphics.FromImage(result))
                    {
                        g.Clear(Color.Transparent);
                        g.DrawImage(raw, 0, 0, raw.Width, raw.Height);
                    }
                    raw.Dispose();
                    if (result.Width >= 32) return result;
                    result.Dispose();
                }
            }
        }
        return null;
    }

    // Resolve SIGDN name from the first shell item in the CIDA
    private static string ResolveFirstName(byte[] cidaBytes, uint sigdn)
    {
        int count;
        IntPtr[] pidls = ParseCida(cidaBytes, out count);
        if (pidls == null) return null;
        string result = null;
        for (int i = 0; i < count; i++)
        {
            if (pidls[i] == IntPtr.Zero) continue;
            if (result == null)
            {
                IntPtr pszName;
                if (SHGetNameFromIDList(pidls[i], sigdn, out pszName) == 0 && pszName != IntPtr.Zero)
                {
                    result = Marshal.PtrToStringUni(pszName);
                    CoTaskMemFree(pszName);
                }
            }
            ILFree(pidls[i]);
        }
        return result;
    }

    // Resolve absolute parsing name (e.g. "shell:AppsFolder\Microsoft.WindowsCalculator_8wekyb3d8bbwe!App")
    public static string ResolvePath(byte[] cidaBytes)
    {
        return ResolveFirstName(cidaBytes, SIGDN_DESKTOPABSOLUTEPARSING);
    }

    // Resolve user-friendly display name (e.g. "Calculator")
    public static string ResolveDisplayName(byte[] cidaBytes)
    {
        return ResolveFirstName(cidaBytes, SIGDN_NORMALDISPLAY);
    }

    // Extract highest-quality icon bitmap from the first shell item in a CIDA
    public static Bitmap ExtractIconFromCida(byte[] cidaBytes)
    {
        int count;
        IntPtr[] pidls = ParseCida(cidaBytes, out count);
        if (pidls == null) return null;
        Bitmap result = null;
        for (int i = 0; i < count; i++)
        {
            if (pidls[i] == IntPtr.Zero) continue;
            if (result == null) result = ExtractIconFromPidl(pidls[i]);
            ILFree(pidls[i]);
        }
        return result;
    }

    // Extract highest-quality icon from a shell path string (shell:AppsFolder\..., ::{CLSID}, etc.)
    [DllImport("shell32.dll")]
    private static extern int SHGetKnownFolderPath(ref Guid rfid, uint dwFlags, IntPtr hToken, out IntPtr ppszPath);
    [DllImport("shell32.dll", CharSet = CharSet.Unicode)]
    private static extern IntPtr ILCreateFromPathW(string pszPath);

    private static IntPtr ResolveShellPath(string shellPath)
    {
        string[] candidates = new string[] {
            shellPath,
            "::" + shellPath,
            "shell:::" + shellPath
        };
        for (int i = 0; i < candidates.Length; i++)
        {
            uint sfgao;
            IntPtr pidl;
            int hr = SHParseDisplayName(candidates[i], IntPtr.Zero, out pidl, 0, out sfgao);
            if (hr == 0 && pidl != IntPtr.Zero) return pidl;
        }
        if (shellPath.Length > 38 && shellPath[0] == '{')
        {
            int closeBrace = shellPath.IndexOf('}');
            if (closeBrace == 37)
            {
                string guidStr = shellPath.Substring(0, 38);
                Guid folderId;
                try { folderId = new Guid(guidStr); } catch { return IntPtr.Zero; }
                IntPtr ppszPath;
                int hr = SHGetKnownFolderPath(ref folderId, 0, IntPtr.Zero, out ppszPath);
                if (hr == 0 && ppszPath != IntPtr.Zero)
                {
                    string folderPath = Marshal.PtrToStringUni(ppszPath);
                    CoTaskMemFree(ppszPath);
                    string remainder = shellPath.Substring(38).TrimStart('\\');
                    string fullPath = string.IsNullOrEmpty(remainder)
                        ? folderPath
                        : System.IO.Path.Combine(folderPath, remainder);
                    IntPtr pidl2 = ILCreateFromPathW(fullPath);
                    if (pidl2 != IntPtr.Zero) return pidl2;
                    uint sfgao2;
                    hr = SHParseDisplayName(fullPath, IntPtr.Zero, out pidl2, 0, out sfgao2);
                    if (hr == 0 && pidl2 != IntPtr.Zero) return pidl2;
                }
            }
        }
        return IntPtr.Zero;
    }

    public static Bitmap ExtractIconFromShellPath(string shellPath)
    {
        IntPtr pidl = ResolveShellPath(shellPath);
        if (pidl == IntPtr.Zero) return null;
        try
        {
            return ExtractIconFromPidl(pidl);
        }
        finally { ILFree(pidl); }
    }
}
'@

#region ── C# TYPE : TASKBAR PIN ─

Add-Type -Language CSharp -TypeDefinition @'
using System;
using System.Runtime.InteropServices;

public static class TaskbarPinHelper
{
    [DllImport("shell32.dll", CharSet = CharSet.Unicode)]
    private static extern IntPtr ILCreateFromPathW(string pszPath);
    [DllImport("shell32.dll")]
    private static extern void ILFree(IntPtr pidl);
    [DllImport("shell32.dll")]
    private static extern IntPtr ILFindLastID(IntPtr pidl);
    [DllImport("shell32.dll", CharSet = CharSet.Unicode)]
    private static extern int SHParseDisplayName(string pszName, IntPtr pbc, out IntPtr ppidl, uint sfgaoIn, out uint psfgaoOut);
    [DllImport("shell32.dll")]
    private static extern void SHChangeNotify(int wEventId, uint uFlags, IntPtr dwItem1, IntPtr dwItem2);
    [DllImport("kernel32.dll", CharSet = CharSet.Unicode, SetLastError = true)]
    private static extern IntPtr CreateMutexExW(IntPtr lpMutexAttributes, string lpName, uint dwFlags, uint dwDesiredAccess);
    [DllImport("kernel32.dll", SetLastError = true)]
    private static extern uint WaitForSingleObject(IntPtr hHandle, uint dwMilliseconds);
    [DllImport("kernel32.dll", SetLastError = true)]
    private static extern bool ReleaseMutex(IntPtr hMutex);
    [DllImport("kernel32.dll", SetLastError = true)]
    private static extern bool CloseHandle(IntPtr hObject);

    private static byte[] InjectBeef001D(byte[] item, string displayName)
    {
        ushort cb = BitConverter.ToUInt16(item, 0);
        if (cb < 4) return null;
        byte[] nameBytes = System.Text.Encoding.Unicode.GetBytes(displayName + "\0");
        int blockCb = 2 + 2 + 4 + 2 + nameBytes.Length;
        byte[] block = new byte[blockCb];
        Array.Copy(BitConverter.GetBytes((ushort)blockCb), 0, block, 0, 2);
        block[4] = 0x1D; block[5] = 0x00; block[6] = 0xEF; block[7] = 0xBE; block[8] = 0x02; block[9] = 0x00;
        Array.Copy(nameBytes, 0, block, 10, nameBytes.Length);
        ushort extOffset = BitConverter.ToUInt16(item, cb - 2);
        int insertPos;
        if (extOffset > 4 && extOffset < cb - 4)
        {
            int epos = extOffset;
            while (epos + 8 <= cb)
            {
                ushort ecb = BitConverter.ToUInt16(item, epos);
                if (ecb < 8 || epos + ecb > cb) break;
                uint esig = BitConverter.ToUInt32(item, epos + 4);
                if ((esig & 0xFFFF0000) != 0xBEEF0000) break;
                epos += ecb;
            }
            insertPos = epos;
        }
        else { insertPos = cb - 2; extOffset = (ushort)insertPos; }
        int newCb = insertPos + blockCb + 2;
        byte[] result = new byte[newCb];
        Array.Copy(item, 0, result, 0, insertPos);
        Array.Copy(block, 0, result, insertPos, blockCb);
        Array.Copy(BitConverter.GetBytes(extOffset), 0, result, newCb - 2, 2);
        Array.Copy(BitConverter.GetBytes((ushort)newCb), 0, result, 0, 2);
        return result;
    }

    private static byte[] BuildBlobEntry(IntPtr pidl, string beef001dContent)
    {
        IntPtr lastPtr = ILFindLastID(pidl);
        if (lastPtr == IntPtr.Zero) return null;
        int prefixLen = (int)((long)lastPtr - (long)pidl);
        ushort lastCb = (ushort)Marshal.ReadInt16(lastPtr);
        if (lastCb < 4) return null;
        byte[] lastItem = new byte[lastCb];
        Marshal.Copy(lastPtr, lastItem, 0, lastCb);
        byte[] patched = InjectBeef001D(lastItem, beef001dContent);
        if (patched == null) return null;
        int newPidlLen = prefixLen + patched.Length + 2;
        byte[] result = new byte[1 + 4 + newPidlLen];
        result[0] = 0x00;
        Array.Copy(BitConverter.GetBytes((uint)newPidlLen), 0, result, 1, 4);
        Marshal.Copy(pidl, result, 5, prefixLen);
        Array.Copy(patched, 0, result, 5 + prefixLen, patched.Length);
        return result;
    }

    private static byte[] GetBlobEntryInternal(string path, string beef001dContent, bool useFilesystem)
    {
        IntPtr pidl;
        if (useFilesystem) { pidl = ILCreateFromPathW(path); }
        else { uint sfgao; if (SHParseDisplayName(path, IntPtr.Zero, out pidl, 0, out sfgao) != 0) pidl = IntPtr.Zero; }
        if (pidl == IntPtr.Zero) return null;
        try { return BuildBlobEntry(pidl, beef001dContent); }
        finally { ILFree(pidl); }
    }

    public static byte[] GetBlobEntryEx(string lnkFullPath, string beef001dContent)
    {
        return GetBlobEntryInternal(lnkFullPath, beef001dContent, false);
    }

    public static byte[] GetBlobEntryFs(string lnkFullPath, string beef001dContent)
    {
        return GetBlobEntryInternal(lnkFullPath, beef001dContent, true);
    }

    public static int FindBlobEntry(byte[] blob, string filename)
    {
        byte[] needle = System.Text.Encoding.Unicode.GetBytes(filename);
        int pos = 0; int idx = 0;
        while (pos < blob.Length && blob[pos] != 0xFF)
        {
            if (pos + 5 > blob.Length) break;
            uint pidlSize = BitConverter.ToUInt32(blob, pos + 1);
            int pidlStart = pos + 5;
            int pidlEnd = pidlStart + (int)pidlSize;
            if (pidlEnd > blob.Length) break;
            for (int b = pidlStart; b + needle.Length <= pidlEnd; b++)
            {
                bool match = true;
                for (int c = 0; c < needle.Length; c++) { if (blob[b + c] != needle[c]) { match = false; break; } }
                if (match) return idx;
            }
            pos = pidlEnd; idx++;
        }
        return -1;
    }

    public static void SendPinNotify()
    {
        byte[] payload = new byte[12];
        payload[0] = 0x0A; payload[1] = 0x00; payload[2] = 0x0D; payload[3] = 0x00;
        IntPtr ptr = Marshal.AllocHGlobal(12);
        try { Marshal.Copy(payload, 0, ptr, 12); SHChangeNotify(0x04000000, 0x3000, ptr, IntPtr.Zero); }
        finally { Marshal.FreeHGlobal(ptr); }
    }

    private static IntPtr _mutexHandle = IntPtr.Zero;

    public static bool AcquirePinMutex(int timeoutMs)
    {
        IntPtr h = CreateMutexExW(IntPtr.Zero, "TaskbarPinListMutex", 0, 0x001F0001);
        if (h == IntPtr.Zero) return false;
        uint r = WaitForSingleObject(h, (uint)timeoutMs);
        if (r == 0 || r == 0x80) { _mutexHandle = h; return true; }
        CloseHandle(h); return false;
    }

    public static void ReleasePinMutex()
    {
        if (_mutexHandle != IntPtr.Zero) { ReleaseMutex(_mutexHandle); CloseHandle(_mutexHandle); _mutexHandle = IntPtr.Zero; }
    }
}
'@

#region ── HELPER FUNCTIONS ─

Update-LoadingPopup 50 "Loading..."

# Type cache for resolved control types (avoids repeated reflection lookups)
$script:ControlTypeCache = @{}
# Enum type map with pre-resolved [type] objects (avoids hashtable recreation + [type] cast per call)
$script:EnumTypeMap = @{
    Dock          = [System.Windows.Forms.DockStyle]
    Orientation   = [System.Windows.Forms.Orientation]
    FlowDirection = [System.Windows.Forms.FlowDirection]
    BorderStyle   = [System.Windows.Forms.BorderStyle]
    SizeMode      = [System.Windows.Forms.PictureBoxSizeMode]
    TextAlign     = [System.Drawing.ContentAlignment]
    ScrollBars    = [System.Windows.Forms.ScrollBars]
    View          = [System.Windows.Forms.View]
    FlatStyle     = [System.Windows.Forms.FlatStyle]
    StartPosition = [System.Windows.Forms.FormStartPosition]
    AutoSizeMode  = [System.Windows.Forms.AutoSizeMode]
}
# Reusable ref for integer parsing (avoids allocating [ref] per call)
$script:IntParseBuffer = 0
# Function that quick-create majority of controls
function New-Control {
    param(
        [Parameter(Mandatory=$false)][AllowNull()]     $container,
        [Parameter(Mandatory=$true)][string]           $type,
        [Parameter(ValueFromRemainingArguments=$true)] $args
    )
    $text = $x = $y = $width = $height = $null
    $additionalProps = New-Object 'System.Collections.Generic.List[string]'(8)
    $deferredAutoSize = $null
    if ($args -and $args.Count -gt 0) {
        $i = 0
        if ($args[$i] -is [string] -and $args[$i] -notmatch '^\w+=') { $text = $args[$i]; $i++ }
        if (($args.Count - $i) -ge 4) {
            $nums = $args[$i..($i+3)]; $allInt = $true
            foreach ($n in $nums) { if ($n -isnot [int] -and -not ($n -is [string] -and $n -match '^\d+$')) { $allInt = $false; break } }
            if ($allInt) { $x=[int]$nums[0]; $y=[int]$nums[1]; $width=[int]$nums[2]; $height=[int]$nums[3]; $i+=4 }
        }
        while ($i -lt $args.Count) { $additionalProps.Add($args[$i]); $i++ }
    }
    # Resolve and cache control type (avoids repeated -as [type] reflection)
    if (-not $script:ControlTypeCache.ContainsKey($type)) {
        if ($type.Contains('.')) {
            $script:ControlTypeCache[$type] = [type]$type
        } else {
            $wfType = "System.Windows.Forms.$type" -as [type]
            $script:ControlTypeCache[$type] = if ($null -ne $wfType) { $wfType } else { [type]$type }
        }
    }
    $control = [System.Activator]::CreateInstance($script:ControlTypeCache[$type])
    if ($null -ne $text) { try { $control.Text = $text } catch {} }
    if ($null -ne $x) {
        try { $control.Location = New-Object System.Drawing.Point($x, $y) } catch {}
        try { $control.Size     = New-Object System.Drawing.Size($width, $height) } catch {}
    }
    # TextBox auto-behavior : border style, Ctrl+A, suppress Enter beep
    if ($control -is [System.Windows.Forms.TextBox]) {
        $control.BorderStyle = [System.Windows.Forms.BorderStyle]::FixedSingle
        # Explicit height DPI scaling fix for single-line TextBoxes
        $isMultilineArgument = $additionalProps | Where-Object { $_ -match '^Multiline\s*=\s*(\$true|True)$' }
        if ($null -ne $height -and $height -gt 0 -and -not $isMultilineArgument) {
            $dpiAdjustedHeight = [int][Math]::Round($height / $script:DPI_Factor)
            if ($dpiAdjustedHeight % 2 -ne 0) { $dpiAdjustedHeight++ }
            $control.Size = New-Object System.Drawing.Size($width, $dpiAdjustedHeight)
        }
        # Disable AutoSize when an explicit height is provided and no AutoSize argument exists
        $hasExplicitHeight   = ($null -ne $height -and $height -gt 0) -or ($additionalProps | Where-Object { $_ -match '^AutoSize\s*=' })
        $hasAutoSizeArgument = $additionalProps | Where-Object { $_ -match '^AutoSize\s*=' }
        if ($hasExplicitHeight -and -not $hasAutoSizeArgument) { $deferredAutoSize = $false }
        $control.Add_KeyPress({ if ($_.KeyChar -eq [char]13) { $_.Handled = $true } })
        $control.Add_KeyDown({
            if     ($_.Control -and $_.KeyCode -eq [System.Windows.Forms.Keys]::A) {$this.SelectAll();$_.Handled=$true;$_.SuppressKeyPress=$true}
            elseif ($_.KeyCode -eq [System.Windows.Forms.Keys]::Enter)             {$_.Handled=$true;$_.SuppressKeyPress=$true}
        })
    }
    # ListView auto-behavior : Ctrl+A to select all items
    if ($control -is [System.Windows.Forms.ListView]) {
        $control.Add_KeyDown({
            if ($_.Control -and $_.KeyCode -eq [System.Windows.Forms.Keys]::A) {
                $this.BeginUpdate()
                foreach ($item in $this.Items) { $item.Selected = $true }
                $this.EndUpdate()
                $_.Handled          = $true
                $_.SuppressKeyPress = $true
            }
        })
    }
    foreach ($prop in $additionalProps) {
        $eqIdx = $prop.IndexOf('=')
        if ($eqIdx -lt 1) { continue }
        $pn = $prop.Substring(0, $eqIdx).Trim()
        $pv = $prop.Substring($eqIdx + 1).Trim()
        # Nested property access (e.g. HorizontalScroll.Enabled)
        if ($pn.Contains('.')) {
            $parts  = $pn.Split('.')
            $target = $control
            for ($p = 0; $p -lt $parts.Count - 1; $p++) { $target = $target.($parts[$p]) }
            $leafPn = $parts[-1]
            if     ($pv -eq '$true'  -or $pv -eq 'True')  { $target.$leafPn = $true }
            elseif ($pv -eq '$false' -or $pv -eq 'False') { $target.$leafPn = $false }
            elseif ([int]::TryParse($pv, [ref]$script:IntParseBuffer)) { $target.$leafPn = $script:IntParseBuffer }
            else   { $target.$leafPn = $pv }
            continue
        }
        # Enum properties (most frequent in hot loops : Dock, FlowDirection, FlatStyle, AutoSizeMode...)
        if ($script:EnumTypeMap.ContainsKey($pn)) {
            $control.$pn = [Enum]::Parse($script:EnumTypeMap[$pn], $pv)
            continue
        }
        switch ($pn) {
            "AutoSize" {
                if     ($pv -eq '$true'  -or $pv -eq 'True')  { $deferredAutoSize = $true }
                elseif ($pv -eq '$false' -or $pv -eq 'False') { $deferredAutoSize = $false }
            }
            "Font" {
                $fontParts = $pv -split ',\s*'
                $fontName  = $fontParts[0].Trim('"')
                $fontSize  = [float]$fontParts[1].Trim()
                $fontStyle = if ($fontParts.Count -gt 2) { [System.Drawing.FontStyle]($fontParts[2].Trim() -replace '\s','') }
                             else                        { [System.Drawing.FontStyle]::Regular }
                $control.Font = New-Object System.Drawing.Font($fontName, $fontSize, $fontStyle)
            }
            { $_ -eq 'Padding' -or $_ -eq 'Margin' } {
                $values      = $pv -split '\s+'
                $control.$pn = if ($values.Count -eq 1) { New-Object System.Windows.Forms.Padding([int]$values[0]) }
                               else { New-Object System.Windows.Forms.Padding([int]$values[0], [int]$values[1], [int]$values[2], [int]$values[3]) }
            }
            { $_ -eq 'ForeColor' -or $_ -eq 'BackColor' -or $_ -eq 'BorderColor' -or $_ -eq 'HoverColor' } {
                $colorValues = $pv -split '\s+'
                if    ($pv -match 'Color \[A=(\d+), R=(\d+), G=(\d+), B=(\d+)\]'){$parsedColor=[System.Drawing.Color]::FromArgb([int]$Matches[1], [int]$Matches[2], [int]$Matches[3], [int]$Matches[4])}
                elseif($pv -match 'Color \[(\w+)\]')                             {$parsedColor=[System.Drawing.Color]::FromName($Matches[1])}
                elseif($colorValues.Count -eq 1)                                 {$parsedColor=[System.Drawing.ColorTranslator]::FromHtml($colorValues[0])}
                else                                                             {$parsedColor=[System.Drawing.Color]::FromArgb([int]$colorValues[0], [int]$colorValues[1], [int]$colorValues[2])}
                switch($pn) {
                    "ForeColor"   { $control.ForeColor                         = $parsedColor }
                    "BackColor"   { $control.BackColor                         = $parsedColor }
                    "BorderColor" { $control.FlatAppearance.BorderColor        = $parsedColor }
                    "HoverColor"  { $control.FlatAppearance.MouseOverBackColor = $parsedColor }
                }
            }
            "Anchor" {
                $combined = [System.Windows.Forms.AnchorStyles]::None
                foreach ($anchorValue in ($pv -split ',\s*')) { $combined = $combined -bor [System.Windows.Forms.AnchorStyles]::$anchorValue }
                $control.Anchor = $combined
            }
            default {
                if     ($pv -eq '$true'  -or $pv -eq 'True')  { $control.$pn = $true }
                elseif ($pv -eq '$false' -or $pv -eq 'False') { $control.$pn = $false }
                elseif ($pv.Length -gt 0 -and $pv[0] -eq '$') { $control.$pn = Invoke-Expression $pv }
                elseif ($pv -match '^New-Object|^\@\{|^\[.*\]::') { $control.$pn = Invoke-Expression $pv }
                elseif ([int]::TryParse($pv, [ref]$script:IntParseBuffer)) { $control.$pn = $script:IntParseBuffer }
                else   { $control.$pn = $pv }
            }
        }
    }
    if ($null -ne $container) {
        if     ($container -is [System.Windows.Forms.ToolStrip] -or $container -is [System.Windows.Forms.StatusStrip] -or
                $container -is [System.Windows.Forms.ContextMenuStrip]) { [void]$container.Items.Add($control) }
        elseif ($container -is [System.Windows.Forms.SplitContainer])   { [void]$container.Panel1.Controls.Add($control) }
        else   { [void]$container.Controls.Add($control) }
    }
    # Apply AutoSize after control is added to its container
    if ($null -ne $deferredAutoSize) { $control.AutoSize = $deferredAutoSize }
    return $control
}
Set-Alias -Name gen -Value New-Control

$script:GroupBoxPaintHandler = {
    param($s, $e)
    $g   = $e.Graphics
    $box = $s
    $bgBrush = New-Object System.Drawing.SolidBrush($box.BackColor)
    $g.FillRectangle($bgBrush, 0, 0, $box.Width, $box.Height)
    $bgBrush.Dispose()
    $textSize = $g.MeasureString($box.Text, $box.Font)
    $textLeft = 8
    $halfText = [int]($textSize.Height / 2)
    $textBrush = New-Object System.Drawing.SolidBrush($box.ForeColor)
    $g.DrawString($box.Text, $box.Font, $textBrush, $textLeft, 0)
    $textBrush.Dispose()
    $pen  = $script:GroupBoxBorderPen
    $rect = New-Object System.Drawing.Rectangle(0, $halfText, ($box.Width - 1), ($box.Height - $halfText - 1))
    $g.DrawLine($pen, $rect.X, $rect.Y, ($textLeft - 2), $rect.Y)
    $g.DrawLine($pen, ($textLeft + [int]$textSize.Width + 1), $rect.Y, $rect.Right, $rect.Y)
    $g.DrawLine($pen, $rect.X,     $rect.Y,      $rect.X,     $rect.Bottom)
    $g.DrawLine($pen, $rect.X,     $rect.Bottom,  $rect.Right, $rect.Bottom)
    $g.DrawLine($pen, $rect.Right, $rect.Y,       $rect.Right, $rect.Bottom)
}

# Parse a registry DefaultIcon value "path,index" into separate path and index
function Split-IconReference {
    param([string]$RawValue)
    $expanded = [System.Environment]::ExpandEnvironmentVariables($RawValue.Trim().Trim('"', "'"))
    $iconIndex = 0
    $lastComma = $expanded.LastIndexOf(',')
    if ($lastComma -gt 0 -and [int]::TryParse($expanded.Substring($lastComma + 1).Trim(), [ref]$iconIndex)) {
        $iconPath = $expanded.Substring(0, $lastComma).Trim().Trim('"', "'")
    }
    else {
        $iconPath  = $expanded
        $iconIndex = 0
    }
    return @{ Path = $iconPath; Index = $iconIndex }
}

# Process a shell item drop (UWP apps, virtual shell objects) based on target zone
function Invoke-ShellItemDrop {
    param(
        [System.Windows.Forms.IDataObject]$DataObject,
        [string]$Zone
    )
    $form.UseWaitCursor = $true
    $form.Cursor = [System.Windows.Forms.Cursors]::WaitCursor
    [System.Windows.Forms.Application]::DoEvents()
    try {
        $stream = $DataObject.GetData("Shell IDList Array")
        if ($null -eq $stream) { return }
        $cidaBytes = $stream.ToArray()
        $shellPath   = [ShellDropHelper]::ResolvePath($cidaBytes)
        $displayName = [ShellDropHelper]::ResolveDisplayName($cidaBytes)
        if (Test-StringEmpty $shellPath) {
            Write-Log "Shell item drop : CIDA resolved to empty path" -Level Warning
            return
        }
        Write-Log "Shell item drop : '$shellPath' ($displayName) on zone '$Zone'"
        switch ($Zone) {
            'icon' {
                $script:SuppressPreviewUpdate = $true
                try {
                    $bitmap = [ShellDropHelper]::ExtractIconFromCida($cidaBytes)
                    if ($null -ne $bitmap) {
                        $icoBytes = [IcoBuilder]::BuildFromBitmap($bitmap)
                        $bitmap.Dispose()
                        $radioIcon_Base64.Checked = $true
                        $TextboxIcon_Base64.Text = [System.Convert]::ToBase64String($icoBytes)
                        Write-Log "Icon resolved from shell item : $displayName ($($icoBytes.Length) bytes)"
                    }
                    else {
                        Write-Log "Shell item drop : icon extraction failed" -Level Warning
                    }
                }
                finally {
                    $script:SuppressPreviewUpdate = $false
                    Update-IconPreview
                }
            }
            'target' {
                # Pre-extract icon from CIDA before setting text (which triggers Update-IconPreview)
                if ($null -ne $script:ShellTargetIconCache) {
                    try { $script:ShellTargetIconCache.Dispose() } catch {}
                }
                $script:ShellTargetIconCache = [ShellDropHelper]::ExtractIconFromCida($cidaBytes)
                # Bare AUMID detection : UWP apps resolve as "PackageFamily!AppId" without shell: prefix
                if ($shellPath.Contains('!') -and -not $script:ShellTargetRegex.IsMatch($shellPath.TrimStart())) {
                    $shellPath = "shell:AppsFolder\$shellPath"
                    Write-Log "Shell item drop : bare AUMID detected, prefixed to '$shellPath'" -Level Debug
                }
                $script:SuppressTargetSplit = $true
                $textTarget.Text = $shellPath
                $script:SuppressTargetSplit = $false
                $appsPrefix = "shell:AppsFolder\"
                if ($shellPath.StartsWith($appsPrefix, [System.StringComparison]::OrdinalIgnoreCase)) {
                    $textAumid.Text = $shellPath.Substring($appsPrefix.Length)
                }
                if (-not (Test-StringEmpty $displayName)) {
                    $textDescription.Text = $displayName
                }
                if (Test-IconSourceEmpty) {
                    $radioIcon_TargetDefault.Checked = $true
                }
            }
            'lnk' {
                $name = if (-not (Test-StringEmpty $displayName)) { $displayName } else { "Shortcut" }
                $currentLnk = Get-CleanInput $textLnkPath.Text
                if (Test-StringEmpty $currentLnk) {
                    $hasExistingFields = -not (Test-StringEmpty (Get-CleanInput $textTarget.Text))
                    if ($hasExistingFields) { $script:SuppressAutoFill = $true }
                    try {
                        $textLnkPath.Text = [IO.Path]::Combine([Environment]::GetFolderPath("Desktop"), "$name.lnk")
                    }
                    finally {
                        if ($hasExistingFields) { $script:SuppressAutoFill = $false }
                    }
                }
                if (Test-IconSourceEmpty) {
                    $radioIcon_TargetDefault.Checked = $true
                }
            }
            'workdir' {
                Write-Log "Shell item drop on workdir zone ignored (virtual shell item has no working directory)" -Level Debug
            }
        }
    }
    catch {
        Write-Log "Shell item drop error : $($_.Exception.Message)" -Level Error
    }
    finally {
        $form.UseWaitCursor = $false
        $form.Cursor = [System.Windows.Forms.Cursors]::Default
    }
}

# ── Make a control transparent to hit-testing (clicks pass through to parent) ──
function Enable-HitTestPassThrough {
    param([System.Windows.Forms.Control]$Control)
    $script:HitTestPassThruControls.Add($Control)
}

function Set-ControlRedraw {
    param([System.Windows.Forms.Control]$Control, [bool]$Enable)
    [NativeMethods]::SendMessage($Control.Handle, 0x000B, [int]$Enable, 0) | Out-Null
    if ($Enable) { $Control.Invalidate($true); $Control.Update() }
}

# Debounce timer for Update-CreateButtonState (avoids File.Exists on every keystroke)
$script:CreateBtnDebounceTimer = New-Object System.Windows.Forms.Timer
$script:CreateBtnDebounceTimer.Interval = 300
$script:CreateBtnDebounceTimer.Add_Tick({
    $this.Stop()
    Update-CreateButtonState
})
$script:CreateBtnFlashTimer = $null
function Reset-CreateButtonFlash {
    if ($null -ne $script:CreateBtnFlashTimer) {
        $script:CreateBtnFlashTimer.Stop()
        $script:CreateBtnFlashTimer.Dispose()
        $script:CreateBtnFlashTimer = $null
        $btnCreate.Text = "Create Shortcut"
    }
}
function Invoke-CreateButtonFlash {
    param([string]$Text)
    Reset-CreateButtonFlash
    $btnCreate.Text = $Text
    $script:CreateBtnFlashTimer = New-Object System.Windows.Forms.Timer
    $script:CreateBtnFlashTimer.Interval = 2500
    $script:CreateBtnFlashTimer.Add_Tick({
        $this.Stop()
        $this.Dispose()
        $script:CreateBtnFlashTimer = $null
        $btnCreate.Text = "Create Shortcut"
    })
    $script:CreateBtnFlashTimer.Start()
}
function Request-CreateButtonUpdate {
    Reset-CreateButtonFlash
    $script:CreateBtnDebounceTimer.Stop()
    $script:CreateBtnDebounceTimer.Start()
}

# ── Trim whitespace and surrounding quotes from a string ──
function Get-CleanInput {
    param([string]$Text)
    if (Test-StringEmpty $Text) { return "" }
    return $Text.Trim().Trim('"').Trim("'").Trim()
}

# Check if path is on NTFS volume
function Test-NtfsVolume {
    param([string]$Path)
    try {
        $root = [System.IO.Path]::GetPathRoot($Path)
        if (Test-StringEmpty $root) { return $true }
        # UNC paths : assume ADS-capable (most SMB shares are NTFS-backed)
        if ($root.StartsWith('\\')) { return $true }
        $drive = New-Object System.IO.DriveInfo($root)
        if (-not $drive.IsReady) { return $true }
        if ($drive.DriveFormat -ne 'NTFS') { return $false }
        # Find the deepest existing directory in the path
        $testDir = [System.IO.Path]::GetDirectoryName($Path)
        while (-not (Test-StringEmpty $testDir) -and -not [System.IO.Directory]::Exists($testDir)) {
            $testDir = [System.IO.Path]::GetDirectoryName($testDir)
        }
        if (Test-StringEmpty $testDir) { return $true }
        # Root-level directory : pure NTFS, no probe needed
        if ($testDir.TrimEnd('\') -eq $root.TrimEnd('\')) { return $true }
        # Check cache (avoids re-probing while user types)
        $cacheKey = $testDir.ToLower().TrimEnd('\')
        if ($script:AdsProbeCache.ContainsKey($cacheKey)) { return $script:AdsProbeCache[$cacheKey] }
        # Probe actual ADS support (handles CrossDevice, OneDrive, cloud filter drivers...)
        $probeFile = [IO.Path]::Combine($testDir, ".ads_probe_" + [Guid]::NewGuid().ToString('N').Substring(0,8))
        try {
            [IO.File]::WriteAllBytes($probeFile, [byte[]]@(0))
            [AdsHelper]::WriteStream($probeFile, "probe", [byte[]]@(0))
            $supported = ([AdsHelper]::GetStreamLength($probeFile, "probe") -gt 0)
            $script:AdsProbeCache[$cacheKey] = $supported
            Write-Log "ADS probe '$testDir' : $(if ($supported) {'supported'} else {'NOT supported'})" -Level Debug
            return $supported
        }
        catch {
            Write-Log "ADS probe failed '$testDir' : $($_.Exception.Message)" -Level Debug
            $script:AdsProbeCache[$cacheKey] = $false
            return $false
        }
        finally {
            try { if ([IO.File]::Exists($probeFile)) { [IO.File]::Delete($probeFile) } } catch {}
        }
    }
    catch { return $true }
}

# Update NTFS warning in labelLnkInfo
function Update-NtfsWarning {
    $lnkPath = Get-CleanInput $textLnkPath.Text
    if (Test-StringEmpty $lnkPath) {
        if ($labelLnkInfo.Text -like "*NTFS*") { $labelLnkInfo.Text = "" }
        return
    }
    if ($radioIcon_TargetDefault.Checked) {
        if ($labelLnkInfo.Text -like "*NTFS*") { $labelLnkInfo.Text = "" }
        return
    }
    $embedMode = $radioIcon_Base64.Checked -or ($radioIcon_AnyFile.Checked -and $radioEmbedIcon.Checked)
    $hasIcon = if ($radioIcon_Base64.Checked) { -not (Test-StringEmpty $TextboxIcon_Base64.Text) }
               else { -not (Test-StringEmpty (Get-CleanInput $iconPathTextbox.Text)) }
    if ($embedMode -and $hasIcon -and -not (Test-NtfsVolume $lnkPath)) {
        $labelLnkInfo.Text = "Destination not on NTFS - ADS icon embedding will fail."
        $labelLnkInfo.ForeColor = [System.Drawing.Color]::Red
    }
    else {
        if ($labelLnkInfo.Text -like "*NTFS*") { $labelLnkInfo.Text = "" }
    }
}

# ── Validate base64 string and return decoded bitmap or null ──
function Test-Base64Image {
    param([string]$Base64)
    $ms = $null; $srcBmp = $null
    try {
        $clean = Get-CleanInput $Base64
        if (Test-StringEmpty $clean) { return $null }
        $bytes = [System.Convert]::FromBase64String($clean)
        # Detect ICO format (magic bytes : 00 00 01 00)
        $isIco = ($bytes.Length -ge 6 -and $bytes[0] -eq 0 -and $bytes[1] -eq 0 -and $bytes[2] -eq 1 -and $bytes[3] -eq 0)
        if ($isIco) {
            # Parse ICO directory to find the largest entry
            $entryCount = [BitConverter]::ToInt16($bytes, 4)
            $bestSize = 0; $bestOffset = -1; $bestLength = 0
            for ($i = 0; $i -lt $entryCount; $i++) {
                $entryBase = 6 + ($i * 16)
                if ($entryBase + 16 -gt $bytes.Length) { break }
                $w = $bytes[$entryBase]; $h = $bytes[$entryBase + 1]
                # 0 means 256 in ICO spec
                $realW = if ($w -eq 0) { 256 } else { [int]$w }
                $realH = if ($h -eq 0) { 256 } else { [int]$h }
                $dataSize   = [BitConverter]::ToInt32($bytes, $entryBase + 8)
                $dataOffset = [BitConverter]::ToInt32($bytes, $entryBase + 12)
                $dimension  = [Math]::Max($realW, $realH)
                if ($dimension -gt $bestSize -and ($dataOffset + $dataSize) -le $bytes.Length) {
                    $bestSize = $dimension; $bestOffset = $dataOffset; $bestLength = $dataSize
                }
            }
            if ($bestOffset -lt 0) { return $null }
            # PNG entries start with 89 50 4E 47, BMP entries start with 28 00 00 00
            $isPng = ($bytes[$bestOffset] -eq 0x89 -and $bytes[$bestOffset + 1] -eq 0x50)
            if ($isPng) {
                $ms = New-Object System.IO.MemoryStream($bytes, $bestOffset, $bestLength)
                $srcBmp = New-Object System.Drawing.Bitmap($ms)
            }
            else {
                $ms = New-Object System.IO.MemoryStream(,$bytes)
                $ico = New-Object System.Drawing.Icon($ms, $bestSize, $bestSize)
                $srcBmp = $ico.ToBitmap()
                $ico.Dispose()
            }
        }
        else {
            $ms = New-Object System.IO.MemoryStream(,$bytes)
            $srcBmp = New-Object System.Drawing.Bitmap($ms)
        }
        $clone = New-Object System.Drawing.Bitmap($srcBmp.Width, $srcBmp.Height, [System.Drawing.Imaging.PixelFormat]::Format32bppArgb)
        $g = [System.Drawing.Graphics]::FromImage($clone)
        $g.Clear([System.Drawing.Color]::Transparent)
        $g.DrawImage($srcBmp, 0, 0, $srcBmp.Width, $srcBmp.Height)
        $g.Dispose()
        return $clone
    }
    catch {
        Write-Log "Test-Base64Image failed : $($_.Exception.Message)" -Level Debug
        return $null
    }
    finally {
        if ($null -ne $srcBmp) { try { $srcBmp.Dispose() } catch {} }
        if ($null -ne $ms)     { try { $ms.Dispose() }     catch {} }
    }
}

# ── Load icon preview from file (image or executable) ──
function Get-PreviewFromFile {
    param([string]$FilePath)
    try {
        if (-not [System.IO.File]::Exists($FilePath)) { return $null }
        $ext = [System.IO.Path]::GetExtension($FilePath).ToLower()
        if ($script:ExeExtensions -contains $ext) {
            $idx = $script:CurrentIconIndex
            # For negative resource IDs or when PE table is unreadable, probe standard sizes directly
            $nativeSizes = [IcoBuilder]::GetNativeSizes($FilePath, $idx)
            if ($nativeSizes.Length -gt 0) {
                $maxNative = $nativeSizes[$nativeSizes.Length - 1]
            }
            else {
                # Probe descending : PrivateExtractIcons handles resource IDs natively
                $maxNative = 0
                foreach ($probe in @(256, 128, 64, 48, 32, 16)) {
                    if ([IcoBuilder]::CanExtractIcon($FilePath, $idx)) {
                        $testBmp = [IcoBuilder]::ExtractBitmapAtSize($FilePath, $idx, $probe)
                        if ($null -ne $testBmp) {
                            $maxNative = $probe
                            $testBmp.Dispose()
                            break
                        }
                    }
                }
            }
            Write-Log "IcoBuilder diagnostic : $([IcoBuilder]::LastDiagnostic)" -Level Debug
            if ($maxNative -le 0) {
                # PE extraction failed, fallback to shell icon (handles .cpl and similar)
                return [IconResolver]::ExtractShellIcon($FilePath)
            }
            return [IcoBuilder]::ExtractBitmapAtSize($FilePath, $idx, $maxNative)
        }
        elseif ($script:ImageExtensions -contains $ext) {
            $ms = $null; $srcBmp = $null
            try {
                $bytes = [System.IO.File]::ReadAllBytes($FilePath)
                $ms = New-Object System.IO.MemoryStream(,$bytes)
                if ($ext -eq '.ico') {
                    $ico = New-Object System.Drawing.Icon($ms, 256, 256)
                    $srcBmp = $ico.ToBitmap()
                    $ico.Dispose()
                }
                else {
                    $srcBmp = New-Object System.Drawing.Bitmap($ms)
                }
                $clone = New-Object System.Drawing.Bitmap($srcBmp.Width, $srcBmp.Height, [System.Drawing.Imaging.PixelFormat]::Format32bppArgb)
                $g = [System.Drawing.Graphics]::FromImage($clone)
                $g.Clear([System.Drawing.Color]::Transparent)
                $g.DrawImage($srcBmp, 0, 0, $srcBmp.Width, $srcBmp.Height)
                $g.Dispose()
                return $clone
            }
            finally {
                if ($null -ne $srcBmp) { try { $srcBmp.Dispose() } catch {} }
                if ($null -ne $ms)     { try { $ms.Dispose() }     catch {} }
            }
        }
        return $null
    }
    catch {
        Write-Log "Get-PreviewFromFile failed for '$FilePath' : $($_.Exception.Message)" -Level Debug
        return $null
    }
}

# ── Build ICO bytes from the currently selected icon source ──
function Get-IcoBytesFromCurrentSource {
    try {
        if ($radioIcon_Base64.Checked) {
            $clean = Get-CleanInput $TextboxIcon_Base64.Text
            if (Test-StringEmpty $clean) { return $null }
            return [IcoBuilder]::BuildFromBase64($clean)
        }
        else {
            $filePath = Get-CleanInput $iconPathTextbox.Text
            if (-not [System.IO.File]::Exists($filePath)) { return $null }
            $ext = [System.IO.Path]::GetExtension($filePath).ToLower()
            if ($script:ExeExtensions -contains $ext) {
                $result = [IcoBuilder]::BuildFromExecutableEx($filePath, $script:CurrentIconIndex)
                Write-Log "IcoBuilder diagnostic : $([IcoBuilder]::LastDiagnostic)" -Level Debug
                return $result
            }
            elseif ($script:ImageExtensions -contains $ext) {
                $ms = $null; $bmp = $null; $ico = $null
                try {
                    $bytes = [System.IO.File]::ReadAllBytes($filePath)
                    $ms = New-Object System.IO.MemoryStream(,$bytes)
                    if ($ext -eq '.ico') {
                        $ico = New-Object System.Drawing.Icon($ms, 256, 256)
                        $bmp = $ico.ToBitmap()
                    }
                    else {
                        $bmp = New-Object System.Drawing.Bitmap($ms)
                    }
                    $result = [IcoBuilder]::BuildFromBitmap($bmp)
                    return $result
                }
                finally {
                    if ($null -ne $ico) { try { $ico.Dispose() } catch {} }
                    if ($null -ne $bmp) { try { $bmp.Dispose() } catch {} }
                    if ($null -ne $ms)  { try { $ms.Dispose() }  catch {} }
                }
            }
        }
        return $null
    }
    catch {
        Write-Log "Get-IcoBytesFromCurrentSource failed : $($_.Exception.Message)" -Level Debug
        return $null
    }
}

# ── Create a shortcut (.lnk or .url) from the current form fields ──
function New-ShortcutFromFields {
    param([string]$OutputPath)
    # Returns @{ IconMode; IcoBytes; IconFilePath; IconIndex } or $null on icon build failure
    $targetPath   = Get-CleanInput $textTarget.Text
    $userDesc     = $textDescription.Text.Trim()
    if (Test-StringEmpty $userDesc) { $userDesc = [IO.Path]::GetFileNameWithoutExtension($OutputPath) }
    $userAumid    = Get-CleanInput $textAumid.Text
    $useEmbed     = $radioIcon_Base64.Checked -or ($radioIcon_AnyFile.Checked -and $radioEmbedIcon.Checked)
    $hasIcon      = if ($radioIcon_TargetDefault.Checked) { $false }
                    elseif ($radioIcon_Base64.Checked) { -not (Test-StringEmpty $TextboxIcon_Base64.Text) }
                    else { -not (Test-StringEmpty (Get-CleanInput $iconPathTextbox.Text)) }
    $iconFilePath = Get-CleanInput $iconPathTextbox.Text
    $iconIndex    = $script:CurrentIconIndex
    # Resolve icon mode and data
    $iconMode = 'default'
    $icoBytes = $null
    if ($hasIcon -and $useEmbed) {
        $icoBytes = Get-IcoBytesFromCurrentSource
        if ($null -eq $icoBytes -or $icoBytes.Length -eq 0) {
            Write-Log "New-ShortcutFromFields : failed to build ICO bytes" -Level Error
            return $null
        }
        $iconMode = 'embed'
    }
    elseif ($hasIcon) {
        if (-not [IO.File]::Exists($iconFilePath)) {
            Write-Log "New-ShortcutFromFields : icon file not found : $iconFilePath" -Level Error
            return $null
        }
        $iconMode = 'standard'
    }
    Write-Log "New-ShortcutFromFields : $OutputPath | mode=$iconMode" -Level Debug
    if ($script:IsUrlTarget) {
        $url = Get-CleanInput $targetPath
        $urlIconFile = ''; $urlIconIndex = 0
        if ($iconMode -eq 'standard') {
            $urlIconFile  = $iconFilePath
            $urlIconIndex = $iconIndex
        }
        New-UrlShortcut -FilePath $OutputPath -Url $url -IconFile $urlIconFile -IconIndex $urlIconIndex
        Write-Log "URL shortcut written : $OutputPath -> $url"
    }
    elseif ($script:IsShellTarget) {
        $shellPath = Get-CleanInput $targetPath
        switch ($iconMode) {
            'embed' {
                [ShortcutHelper]::CreateShellWithEmbeddedIcon($OutputPath, $shellPath, ([byte[]]$icoBytes), $userAumid, $userDesc)
                Write-AdsIcon -FilePath $OutputPath -IcoBytes ([byte[]]$icoBytes)
            }
            'standard' {
                [ShortcutHelper]::CreateShellWithStandardIcon($OutputPath, $shellPath, $iconFilePath, $iconIndex, $userAumid, $userDesc)
            }
            default {
                [ShortcutHelper]::CreateShellWithStandardIcon($OutputPath, $shellPath, "", 0, $userAumid, $userDesc)
            }
        }
        Write-Log "Shell shortcut written : $OutputPath -> $shellPath ($iconMode)"
    }
    else {
        $arguments = $textArgs.Text.Trim()
        $workDir   = Get-CleanInput $textWorkDir.Text
        if (Test-StringEmpty $workDir -and (-not (Test-StringEmpty $targetPath)) -and [IO.File]::Exists($targetPath)) {
            $workDir = [IO.Path]::GetDirectoryName($targetPath)
        }
        switch ($iconMode) {
            'embed' {
                [ShortcutHelper]::CreateWithEmbeddedIcon($OutputPath, $targetPath, $arguments, ([byte[]]$icoBytes), $userAumid, $userDesc, $workDir)
                Write-AdsIcon -FilePath $OutputPath -IcoBytes ([byte[]]$icoBytes)
            }
            'standard' {
                [ShortcutHelper]::CreateWithStandardIcon($OutputPath, $targetPath, $arguments, $iconFilePath, $iconIndex, $userAumid, $userDesc, $workDir)
            }
            default {
                [ShortcutHelper]::CreateWithStandardIcon($OutputPath, $targetPath, $arguments, "", 0, $userAumid, $userDesc, $workDir)
            }
        }
        Write-Log "Shortcut written : $OutputPath -> $targetPath ($iconMode)"
    }
    return @{ IconMode = $iconMode; IcoBytes = $icoBytes; IconFilePath = $iconFilePath; IconIndex = $iconIndex }
}

# ── Pin a single .lnk to the taskbar via blob injection ──
function Invoke-TaskbarPin {
    param([string]$LnkPath)
    $taskBarDir = [IO.Path]::Combine($env:APPDATA, 'Microsoft\Internet Explorer\Quick Launch\User Pinned\TaskBar')
    $regSubKey  = 'Software\Microsoft\Windows\CurrentVersion\Explorer\Taskband'
    if (-not [IO.Directory]::Exists($taskBarDir)) {
        Write-Log "Taskbar pin directory not found : $taskBarDir" -Level Error
        return $false
    }
    $regProbe = [Microsoft.Win32.Registry]::CurrentUser.OpenSubKey($regSubKey, $false)
    if (-not $regProbe) {
        Write-Log "Taskband registry key not found" -Level Error
        return $false
    }
    $regProbe.Close()
    # Resolve beef001d content (display name for the blob entry)
    $beef001d = $null
    foreach ($beef001dResolver in @(
        { [ShortcutHelper]::GetTargetPath($LnkPath) },
        { [ShortcutHelper]::GetParsedDisplayName($LnkPath) },
        { [ShortcutHelper]::GetAppUserModelId($LnkPath) },
        { [IO.Path]::GetFileNameWithoutExtension($LnkPath) }
    )) {
        try { $beef001d = & $beef001dResolver } catch {}
        if (-not (Test-StringEmpty $beef001d)) { break }
    }
    if ((-not (Test-StringEmpty $beef001d)) -and $beef001d.EndsWith('.cpl', [System.StringComparison]::OrdinalIgnoreCase)) {
        $resolvedCplBeef = $beef001d
        if (-not [IO.File]::Exists($resolvedCplBeef)) {
            $system32Candidate = [IO.Path]::Combine($env:SystemRoot, 'System32', [IO.Path]::GetFileName($beef001d))
            if ([IO.File]::Exists($system32Candidate)) { $resolvedCplBeef = $system32Candidate }
        }
        $cplItem = Resolve-CplControlPanelItem $resolvedCplBeef
        if ($null -ne $cplItem) {
            $beef001d = $cplItem.Path
            Write-Log "Taskbar pin : .cpl beef001d resolved to '$beef001d'" -Level Debug
        }
    }
    Write-Log "Taskbar pin : beef001d = '$beef001d'" -Level Debug
    # Copy .lnk to taskbar pinned directory
    $destLnk = [IO.Path]::Combine($taskBarDir, [IO.Path]::GetFileName($LnkPath))
    try { [IO.File]::Copy($LnkPath, $destLnk, $true) }
    catch {
        Write-Log "Failed to copy .lnk to taskbar directory : $($_.Exception.Message)" -Level Error
        return $false
    }
    $script:LastPinnedTaskbarFile = [IO.Path]::GetFileName($destLnk)
    # Repoint ADS IconLocation to the copied file path, then re-embed ADS (Save wipes it)
    try {
        $currentIconPath = [ShortcutHelper]::GetIconPath($destLnk)
        if ((-not [string]::IsNullOrEmpty($currentIconPath)) -and $currentIconPath.EndsWith(':icon.ico', [System.StringComparison]::OrdinalIgnoreCase)) {
            $adsBytes = [AdsHelper]::ReadStream($destLnk, "icon.ico")
            [ShortcutHelper]::UpdateIconOnly($destLnk)
            if ($null -ne $adsBytes -and $adsBytes.Length -gt 0) {
                [AdsHelper]::WriteStream($destLnk, "icon.ico", $adsBytes)
                Write-Log "Taskbar pin : IconLocation repointed + ADS re-embedded ($($adsBytes.Length) bytes)" -Level Debug
            }
        }
    }
    catch { Write-Log "Taskbar pin : failed to update IconLocation : $($_.Exception.Message)" -Level Warning }
    # Build serialized PIDL blob entry
    $blobEntry = [TaskbarPinHelper]::GetBlobEntryEx($destLnk, $beef001d)
    if (-not $blobEntry) {
        $blobEntry = [TaskbarPinHelper]::GetBlobEntryFs($destLnk, $beef001d)
    }
    if (-not $blobEntry) {
        Write-Log "Failed to build blob entry for : $destLnk" -Level Error
        try { [IO.File]::Delete($destLnk) } catch {}
        return $false
    }
    # Inject into registry Favorites blob
    $mutexAcquired = [TaskbarPinHelper]::AcquirePinMutex(5000)
    try {
        $key = [Microsoft.Win32.Registry]::CurrentUser.OpenSubKey($regSubKey, $true)
        if (-not $key) {
            Write-Log "Cannot open Taskband registry key for writing" -Level Error
            return $false
        }
        try {
            $doNotExpand = [Microsoft.Win32.RegistryValueOptions]::DoNotExpandEnvironmentNames
            $favBlob = $key.GetValue('Favorites', $null, $doNotExpand)
            if (-not $favBlob -or $favBlob.Length -lt 2) { $favBlob = [byte[]]@(0xFF) }
            # Check if already pinned
            $fileName = [IO.Path]::GetFileName($destLnk)
            if ([TaskbarPinHelper]::FindBlobEntry($favBlob, $fileName) -ge 0) {
                Write-Log "Already pinned : $fileName" -Level Info
                return $true
            }
            # Find insertion point (before 0xFF terminator)
            $insertPos = 0
            while ($insertPos -lt $favBlob.Length -and $favBlob[$insertPos] -ne 0xFF) {
                if ($insertPos + 5 -gt $favBlob.Length) { break }
                $pidlSize = [BitConverter]::ToUInt32($favBlob, $insertPos + 1)
                $insertPos += 1 + 4 + $pidlSize
            }
            # Assemble new blob
            $ms = New-Object System.IO.MemoryStream
            if ($insertPos -gt 0) { $ms.Write($favBlob, 0, $insertPos) }
            $ms.Write($blobEntry, 0, $blobEntry.Length)
            $ms.WriteByte(0xFF)
            $newBlob = $ms.ToArray()
            $ms.Dispose()
            $changes = [int]$key.GetValue('FavoritesChanges', 0, $doNotExpand)
            $key.SetValue('Favorites',        $newBlob,          [Microsoft.Win32.RegistryValueKind]::Binary)
            $key.SetValue('FavoritesVersion',  3,                [Microsoft.Win32.RegistryValueKind]::DWord)
            $key.SetValue('FavoritesChanges', ($changes + 1),   [Microsoft.Win32.RegistryValueKind]::DWord)
            Write-Log "Taskbar pin injected : $fileName ($($blobEntry.Length) bytes)" -Level Info
        }
        finally { $key.Close() }
    }
    finally {
        if ($mutexAcquired) { [TaskbarPinHelper]::ReleasePinMutex() }
    }
    [TaskbarPinHelper]::SendPinNotify()
    return $true
}

# ── Update the icon preview display ──
function Update-IconPreview {
    if ($script:SuppressPreviewUpdate) { return }
    Set-ControlRedraw $btnIconPreview $false
    $newBmp = $null
    if ($radioIcon_TargetDefault.Checked) {
        $targetPath = Get-CleanInput $textTarget.Text
        if (-not (Test-StringEmpty $targetPath) -and -not $script:IsUrlTarget) {
            if ($script:ShellTargetRegex.IsMatch($targetPath.TrimStart())) {
                # Try cached icon first (set during shell item drop), then live resolution
                if ($null -ne $script:ShellTargetIconCache) {
                    $newBmp = New-Object System.Drawing.Bitmap($script:ShellTargetIconCache)
                }
                else {
                    $newBmp = [ShellDropHelper]::ExtractIconFromShellPath($targetPath)
                }
            }
            else {
                $ext = [IO.Path]::GetExtension($targetPath).ToLower()
                if ($ext -eq '.cpl') {
                    # .cpl in Target Default : show the shell icon Windows would actually assign to the shortcut
                    if ([IO.File]::Exists($targetPath)) {
                        $newBmp = [IconResolver]::ExtractShellIcon($targetPath)
                    }
                }
                elseif ($script:ExeExtensions -contains $ext -and [IO.File]::Exists($targetPath)) {
                    $maxN = [IcoBuilder]::GetMaxNativeSize($targetPath, 0)
                    $newBmp = [IcoBuilder]::ExtractBitmapAtSize($targetPath, 0, $maxN)
                }
                elseif ([IO.File]::Exists($targetPath) -or [IO.Directory]::Exists($targetPath)) {
                    $newBmp = [IconResolver]::ExtractShellIcon($targetPath)
                }
            }
        }
        if ($null -ne $newBmp) {
            $labelPreviewIconInfo.Text = "$($newBmp.Width) x $($newBmp.Height) px (default)"
            Write-Log "Icon preview updated from target default : $targetPath ($($newBmp.Width)x$($newBmp.Height))" -Level Debug
        }
        else {
            $labelPreviewIconInfo.Text = if (Test-StringEmpty $targetPath) { "No target defined" } else { "Default icon" }
        }
    }
    elseif ($radioIcon_Base64.Checked) {
        $newBmp = Test-Base64Image $TextboxIcon_Base64.Text
        if ($newBmp) {
            $labelPreviewIconInfo.Text = "$($newBmp.Width) x $($newBmp.Height) px"
            Write-Log "Icon preview updated from Base64 : $($newBmp.Width)x$($newBmp.Height)" -Level Debug
        }
        else {
            $labelPreviewIconInfo.Text = if (Test-StringEmpty $TextboxIcon_Base64.Text) { "No icon" } else { "Invalid Base64" }
        }
    }
    else {
        $filePath = Get-CleanInput $iconPathTextbox.Text
        if (Test-StringEmpty $filePath) {
            # Empty icon path : show target default preview as a hint
            $targetPath = Get-CleanInput $textTarget.Text
            if (-not (Test-StringEmpty $targetPath) -and -not $script:IsUrlTarget) {
                if ($script:ShellTargetRegex.IsMatch($targetPath.TrimStart())) {
                    if ($null -ne $script:ShellTargetIconCache) {
                        $newBmp = New-Object System.Drawing.Bitmap($script:ShellTargetIconCache)
                    }
                    else {
                        $newBmp = [ShellDropHelper]::ExtractIconFromShellPath($targetPath)
                    }
                }
                else {
                    $ext = [IO.Path]::GetExtension($targetPath).ToLower()
                    if ($ext -eq '.cpl') {
                        if ([IO.File]::Exists($targetPath)) { $newBmp = [IconResolver]::ExtractShellIcon($targetPath) }
                    }
                    elseif ($script:ExeExtensions -contains $ext -and [IO.File]::Exists($targetPath)) {
                        $maxN = [IcoBuilder]::GetMaxNativeSize($targetPath, 0)
                        $newBmp = [IcoBuilder]::ExtractBitmapAtSize($targetPath, 0, $maxN)
                    }
                    elseif ([IO.File]::Exists($targetPath) -or [IO.Directory]::Exists($targetPath)) {
                        $newBmp = [IconResolver]::ExtractShellIcon($targetPath)
                    }
                }
            }
            if ($null -ne $newBmp) {
                $labelPreviewIconInfo.Text = "$($newBmp.Width) x $($newBmp.Height) px (default)"
            }
            else {
                $labelPreviewIconInfo.Text = "No icon"
            }
        }
        else {
            $newBmp = Get-PreviewFromFile $filePath
            if ($newBmp) {
                $ext = [System.IO.Path]::GetExtension($filePath).ToLower()
                $showIndex = ($script:ExeExtensions -contains $ext) -and ($script:CurrentExeIconCount -gt 1 -or $script:CurrentIconIndex -ne 0)
                $labelPreviewIconInfo.Text = if ($showIndex) { "$($newBmp.Width) x $($newBmp.Height) px (index $($script:CurrentIconIndex))" }
                                         else            { "$($newBmp.Width) x $($newBmp.Height) px" }
                Write-Log "Icon preview updated from file : $filePath index $($script:CurrentIconIndex) ($($newBmp.Width)x$($newBmp.Height))" -Level Debug
            }
            else {
                $labelPreviewIconInfo.Text = "Invalid file"
            }
        }
        Update-IconIndexMenu
    }
    $oldBmp = $script:CurrentPreviewBitmap
    $script:CurrentPreviewBitmap = $newBmp
    $btnIconPreview.Invalidate()
    $btnIconPreview.Update()
    if ($oldBmp -and $oldBmp -ne $newBmp) {
        try { $oldBmp.Dispose() } catch { }
    }
    $clickable = $radioIcon_AnyFile.Checked -and ($script:CurrentExeIconCount -gt 1) -and ($null -ne $newBmp)
    if ($clickable) {
        $btnIconPreview.Cursor = [System.Windows.Forms.Cursors]::Hand
        $btnIconPreview.FlatAppearance.MouseOverBackColor = if ($script:IsDarkMode) { [System.Drawing.Color]::FromArgb(40,40,40) } else { [System.Drawing.Color]::FromArgb(220,220,230) }
        $btnIconPreview.FlatAppearance.MouseDownBackColor = if ($script:IsDarkMode) { [System.Drawing.Color]::FromArgb(50,50,50) } else { [System.Drawing.Color]::FromArgb(210,210,220) }
    }
    else {
        $btnIconPreview.Cursor = [System.Windows.Forms.Cursors]::Default
        $btnIconPreview.FlatAppearance.MouseOverBackColor = $btnIconPreview.BackColor
        $btnIconPreview.FlatAppearance.MouseDownBackColor = $btnIconPreview.BackColor
    }
    Set-ControlRedraw $btnIconPreview $true
    Update-CreateButtonState
}

# ── Populate icon index dropdown menu from an executable file ──
function Update-IconIndexMenu {
    # Dispose old menu item images
    $script:IconIndexMenu.SuspendLayout()
    foreach ($item in $script:IconIndexMenu.Items) {
        if ($item.Image) { $item.Image.Dispose(); $item.Image = $null }
    }
    $script:IconIndexMenu.Items.Clear()
    $script:IconIndexMenu.ResumeLayout()
    $script:CurrentExeIconCount = 0
    if ($radioIcon_TargetDefault.Checked -or $radioIcon_Base64.Checked) { $btnIconPreview.Invalidate(); return }
    # Negative resource IDs are a specific icon reference, not a browsable index
    if ($script:CurrentIconIndex -lt 0) { $btnIconPreview.Invalidate(); return }
    $filePath = Get-CleanInput $iconPathTextbox.Text
    if (Test-StringEmpty $filePath) { $script:LastIconMenuPath = ""; $btnIconPreview.Invalidate(); return }
    # Skip rebuild if same file (avoid repeated PE extraction while user types)
    if ($filePath -eq $script:LastIconMenuPath -and $script:CurrentExeIconCount -gt 0) {
        $btnIconPreview.Invalidate(); return
    }
    $script:LastIconMenuPath = $filePath
    if (-not [IO.File]::Exists($filePath)) { $btnIconPreview.Invalidate(); return }
    $ext = [IO.Path]::GetExtension($filePath).ToLower()
    if ($script:ExeExtensions -notcontains $ext) { $btnIconPreview.Invalidate(); return }
    $iconCount = [IcoBuilder]::GetIconCount($filePath)
    $script:CurrentExeIconCount = $iconCount
    Write-Log "Icon index menu : $iconCount icons in $filePath" -Level Debug
    if ($iconCount -le 1) { $btnIconPreview.Invalidate(); return }
    for ($idx = 0; $idx -lt $iconCount; $idx++) {
        $bmp = [IcoBuilder]::ExtractBitmapAtSize($filePath, $idx, 32)
        if ($null -eq $bmp) { continue }
        $menuItem = New-Object System.Windows.Forms.ToolStripMenuItem
        $menuItem.Text  = "Index $idx"
        $menuItem.Image = $bmp
        $menuItem.Tag   = $idx
        if ($idx -eq $script:CurrentIconIndex) {
            $menuItem.Checked = $true
        }
        $menuItem.Add_Click({
            param($s, $e)
            $selectedIndex = $s.Tag
            $script:CurrentIconIndex = $selectedIndex
            Write-Log "Icon index selected : $selectedIndex" -Level Debug
            foreach ($mi in $script:IconIndexMenu.Items) { $mi.Checked = ($mi.Tag -eq $selectedIndex) }
            Update-IconPreview
        })
        [void]$script:IconIndexMenu.Items.Add($menuItem)
    }
    $btnIconPreview.Invalidate()
}

# ── Update arguments length indicator with color coding and Explorer warning ──
function Update-ArgsLength {
    $len = $textArgs.Text.Length
    $labelArgsLen.Text = "Max length : $len / $($script:MaxArgsCreateProcess)"
    if ($len -gt $script:MaxArgsCreateProcess) {
        $labelArgsLen.ForeColor = [System.Drawing.Color]::Red
    }
    elseif ($len -gt $script:MaxArgsCmdExe) {
        $labelArgsLen.ForeColor = [System.Drawing.Color]::FromArgb(200, 150, 0)
    }
    else {
        $labelArgsLen.ForeColor = if ($script:IsDarkMode) { [System.Drawing.Color]::FromArgb(140,140,140) } else { [System.Drawing.Color]::Gray }
    }
    # Explorer combined length warning (Target + space + Arguments > 260)
    $targetLen = (Get-CleanInput $textTarget.Text).Length
    $combinedLen = $targetLen + 1 + $len
    if ($targetLen -eq 0) { $combinedLen = $len }
    if ($len -gt $script:MaxArgsCmdExe) {
        $labelExplorerWarn.Text = "Args > $script:MaxArgsCmdExe chars. May fail if launched via CMD."
        $labelExplorerWarn.ForeColor = [System.Drawing.Color]::FromArgb(220, 160, 0)
    }
    elseif ($combinedLen -gt 260) {
        $labelExplorerWarn.Text = "Full Target = $combinedLen/260. Not editable from Properties"
        $labelExplorerWarn.ForeColor = [System.Drawing.Color]::FromArgb(220, 160, 0)
    }
    else {
        $labelExplorerWarn.Text = ""
    }
    Request-CreateButtonUpdate
}

# ── Validate all fields and toggle the Create button ──
function Update-CreateButtonState {
    $iconValid = $false
    if ($radioIcon_TargetDefault.Checked) {
        $iconValid = $true
    }
    elseif ($radioIcon_Base64.Checked) {
        $iconValid = ($null -ne $script:CurrentPreviewBitmap)
    }
    else {
        $filePath = Get-CleanInput $iconPathTextbox.Text
        if ([System.IO.File]::Exists($filePath)) {
            $ext = [System.IO.Path]::GetExtension($filePath).ToLower()
            $iconValid = ($script:ExeExtensions -contains $ext) -or ($script:ImageExtensions -contains $ext)
        }
    }
    $lnkFilled = -not (Test-StringEmpty $textLnkPath.Text)
    $argsOk    = $textArgs.Text.Length -le $script:MaxArgsCreateProcess
    $lnkPath   = Get-CleanInput $textLnkPath.Text
    $lnkExists = (-not (Test-StringEmpty $lnkPath)) -and [System.IO.File]::Exists($lnkPath)
    # Target validation : must have content (or existing lnk) + must respect 260 char limit
    $targetLen    = (Get-CleanInput $textTarget.Text).Length
    $targetOk     = ($targetLen -gt 0)
    $targetLenOk  = $script:IsShellTarget -or $script:IsUrlTarget -or ($targetLen -le $script:MaxTargetPath) -or ($targetLen -eq 0)
    # AUMID validation
    $aumidOk   = $true
    $aumidText = $textAumid.Text
    if ($aumidText.Length -gt 0) {
        if ($aumidText[0] -eq '.' -or $aumidText[$aumidText.Length - 1] -eq '.' -or
            $script:AumidCleanRegex.IsMatch($aumidText)) {
            $aumidOk = $false
        }
    }
    $iconRequired = -not $script:IsUrlTarget
    $btnCreate.Enabled = ((-not $iconRequired -or $iconValid) -and $lnkFilled -and $argsOk -and $targetOk -and $targetLenOk -and $aumidOk)
    if ($btnCreate.Enabled) {
        $btnCreate.BackColor = [System.Drawing.Color]::FromArgb(0, 120, 212)
        $btnCreate.ForeColor = [System.Drawing.Color]::White
    }
    else {
        $btnCreate.BackColor = if ($script:IsDarkMode) { [System.Drawing.Color]::FromArgb(45, 60, 80) } else { [System.Drawing.Color]::FromArgb(160, 190, 220) }
        $btnCreate.ForeColor = if ($script:IsDarkMode) { [System.Drawing.Color]::FromArgb(120, 120, 120) } else { [System.Drawing.Color]::FromArgb(200, 200, 200) }
    }
    Update-PinButtonState
}

function Update-PinButtonState {
    $hasTarget = -not (Test-StringEmpty $textTarget.Text)
    $btnPin.Enabled = $hasTarget
    if ($btnPin.Enabled) {
        # Check if the shortcut is already pinned (by filename, always .lnk in taskbar)
        $taskBarDir = [IO.Path]::Combine($env:APPDATA, 'Microsoft\Internet Explorer\Quick Launch\User Pinned\TaskBar')
        $lnkPath = Get-CleanInput $textLnkPath.Text
        $fileName = if (-not (Test-StringEmpty $lnkPath)) { [IO.Path]::ChangeExtension([IO.Path]::GetFileName($lnkPath), '.lnk') }
                    else { [IO.Path]::GetFileNameWithoutExtension((Get-CleanInput $textTarget.Text)) + '.lnk' }
        $pinnedPath = [IO.Path]::Combine($taskBarDir, $fileName)
        $alreadyPinned = [IO.File]::Exists($pinnedPath)
        if (-not $alreadyPinned -and -not (Test-StringEmpty $script:LastPinnedTaskbarFile)) {
            $alreadyPinned = [IO.File]::Exists([IO.Path]::Combine($taskBarDir, $script:LastPinnedTaskbarFile))
        }
        $btnPin.Text = if ($alreadyPinned) { [char]0x2714 + " Pinned" } else { "Pin to Taskbar" }
    }
    else {
        $btnPin.Text = "Pin to Taskbar"
    }
}

# Toggle special target mode (Shell CLSID or URL) - locks fields irrelevant for those shortcut types
function Set-SpecialTargetMode {
    param(
        [ValidateSet('Shell','Url')][string]$ModeName,
        [bool]$IsActive
    )
    $currentValue = if ($ModeName -eq 'Shell') { $script:IsShellTarget } else { $script:IsUrlTarget }
    if ($currentValue -eq $IsActive) { return }
    if ($ModeName -eq 'Shell') { $script:IsShellTarget = $IsActive }
    else                       { $script:IsUrlTarget   = $IsActive }
    $textArgs.Enabled         = -not $IsActive
    $textWorkDir.Enabled      = -not $IsActive
    $btnBrowseWorkDir.Enabled = -not $IsActive
    if ($ModeName -eq 'Url') {
        $textDescription.Enabled  = -not $IsActive
        $textAumid.Enabled        = -not $IsActive
    }
    $dimColor = if ($IsActive) {
        if ($script:IsDarkMode) { [System.Drawing.Color]::FromArgb(80,80,80) } else { [System.Drawing.Color]::FromArgb(180,180,180) }
    }
    else {
        if ($script:IsDarkMode) { [System.Drawing.Color]::FromArgb(160,160,160) } else { [System.Drawing.Color]::FromArgb(100,100,100) }
    }
    $labelArgs.ForeColor    = $dimColor
    $labelWorkDir.ForeColor = $dimColor
    if ($ModeName -eq 'Url') {
        $labelDescription.ForeColor = $dimColor
        $labelAumid.ForeColor       = $dimColor
        $labelAumidHelp.ForeColor   = $dimColor
    }
    # URL-specific : force Standard icon path (no ADS embed, no Base64, no Target Default)
    if ($ModeName -eq 'Url') {
        $radioIcon_TargetDefault.Enabled = -not $IsActive
        $radioIcon_Base64.Enabled        = -not $IsActive
        $radioEmbedIcon.Enabled          = -not $IsActive
        if ($IsActive) {
            $radioIcon_AnyFile.Checked   = $true
            $radioStandardIcon.Checked   = $true
        }
        # Auto-switch .lnk <-> .url extension in shortcut location
        $lnkText = $textLnkPath.Text
        if (-not (Test-StringEmpty $lnkText)) {
            if ($IsActive -and $lnkText.EndsWith('.lnk', [System.StringComparison]::OrdinalIgnoreCase)) {
                $script:SuppressAutoFill = $true
                $textLnkPath.Text = $lnkText.Substring(0, $lnkText.Length - 4) + '.url'
                $script:SuppressAutoFill = $false
            }
            elseif (-not $IsActive -and $lnkText.EndsWith('.url', [System.StringComparison]::OrdinalIgnoreCase)) {
                $script:SuppressAutoFill = $true
                $textLnkPath.Text = $lnkText.Substring(0, $lnkText.Length - 4) + '.lnk'
                $script:SuppressAutoFill = $false
            }
        }
        $labelLnkPath.Text = if ($IsActive) { "Shortcut location (.url) :" } else { "Shortcut location (.lnk) :" }
    }
    Write-Log "$ModeName target mode : $(if ($IsActive) {'ON'} else {'OFF'})" -Level Debug
}

# Create a .url internet shortcut file
function New-UrlShortcut {
    param(
        [string]$FilePath,
        [string]$Url,
        [string]$IconFile = '',
        [int]$IconIndex = 0
    )
    $lines = New-Object System.Collections.Generic.List[string]
    $lines.Add('[InternetShortcut]')
    $lines.Add("URL=$Url")
    if (-not (Test-StringEmpty $IconFile)) {
        $lines.Add("IconFile=$IconFile")
        $lines.Add("IconIndex=$IconIndex")
    }
    [IO.File]::WriteAllLines($FilePath, $lines.ToArray())
    Write-Log "URL shortcut created : $FilePath -> $Url" -Level Debug
}

# ── Determine which drop zone a screen point falls into ──
function Get-DropZone {
    param([System.Drawing.Point]$ScreenPoint)
    # Use cached zones during drag, fallback to live computation
    $zones = if ($null -ne $script:CachedDropZones) { $script:CachedDropZones }
             else { Build-DropZoneCache }
    # Exact hit detection first
    if ($zones[0].Rect.Contains($ScreenPoint)) { return 'icon' }
    $grpScreen = $groupShortcut.RectangleToScreen($groupShortcut.ClientRectangle)
    if ($grpScreen.Contains($ScreenPoint)) {
        for ($i = 1; $i -lt $zones.Count; $i++) {
            if ($ScreenPoint.Y -ge $zones[$i].Top -and $ScreenPoint.Y -lt $zones[$i].Bottom) {
                return $zones[$i].Name
            }
        }
    }
    # No exact hit - find closest zone by vertical distance to cursor
    $formRect = $form.RectangleToScreen($form.ClientRectangle)
    if (-not $formRect.Contains($ScreenPoint)) { return $null }
    $bestZone = $null
    $bestDist = [int]::MaxValue
    # Icon zone : distance to rectangle edges
    $iconRect = $zones[0].Rect
    $distX = [Math]::Max(0, [Math]::Max($iconRect.X - $ScreenPoint.X, $ScreenPoint.X - $iconRect.Right))
    $distY = [Math]::Max(0, [Math]::Max($iconRect.Y - $ScreenPoint.Y, $ScreenPoint.Y - $iconRect.Bottom))
    $dist  = ($distX * $distX) + ($distY * $distY)
    if ($dist -lt $bestDist) { $bestDist = $dist; $bestZone = 'icon' }
    # Right-panel zones : distance to vertical band
    for ($i = 1; $i -lt $zones.Count; $i++) {
        $distX = [Math]::Max(0, [Math]::Max($grpScreen.X - $ScreenPoint.X, $ScreenPoint.X - $grpScreen.Right))
        $distY = [Math]::Max(0, [Math]::Max($zones[$i].Top - $ScreenPoint.Y, $ScreenPoint.Y - $zones[$i].Bottom))
        $dist  = ($distX * $distX) + ($distY * $distY)
        if ($dist -lt $bestDist) { $bestDist = $dist; $bestZone = $zones[$i].Name }
    }
    return $bestZone
}

# ── Set the active drop zone and invalidate for repaint ──
function Set-ActiveDropZone {
    param([string]$Zone)
    if ($script:ActiveDropZone -eq $Zone) { return }
    $script:ActiveDropZone = $Zone
    $groupIcon.Invalidate()
    $groupIcon.Update()
    $groupShortcut.Invalidate()
    $groupShortcut.Update()
}

# ── Import all fields from an existing shortcut ──
function Import-ExistingShortcut {
    param([string]$LnkPath)
    if ($script:SuppressAutoFill) { return }
    $script:SuppressAutoFill = $true
    try {
        $form.UseWaitCursor = $true
        $form.Cursor = [System.Windows.Forms.Cursors]::WaitCursor
        [System.Windows.Forms.Application]::DoEvents()
        $script:SuppressPreviewUpdate = $true
        # Force reset of special modes
        $script:IsShellTarget = $true
        Set-SpecialTargetMode 'Shell' $false
        $script:IsUrlTarget = $true
        Set-SpecialTargetMode 'Url' $false
        $importExt = [IO.Path]::GetExtension($LnkPath).ToLower()
        Write-Log "Importing existing shortcut : $LnkPath"
        if ($importExt -eq '.url') {
            # ── .url internet shortcut (INI format) ──
            $urlValue = ""; $iconFilePath = ""; $iconFileIndex = 0
            foreach ($line in [IO.File]::ReadAllLines($LnkPath)) {
                $trimmed = $line.Trim()
                if     ($trimmed.StartsWith('URL=',       [System.StringComparison]::OrdinalIgnoreCase)) { $urlValue     = $trimmed.Substring(4).Trim() }
                elseif ($trimmed.StartsWith('IconFile=',  [System.StringComparison]::OrdinalIgnoreCase)) { $iconFilePath = $trimmed.Substring(9).Trim() }
                elseif ($trimmed.StartsWith('IconIndex=', [System.StringComparison]::OrdinalIgnoreCase)) { [int]::TryParse($trimmed.Substring(10).Trim(), [ref]$iconFileIndex) | Out-Null }
            }
            Write-Log "  URL : $urlValue | IconFile : $iconFilePath | IconIndex : $iconFileIndex" -Level Debug
            # Fill target (triggers URL mode detection via TextChanged)
            $script:SuppressTargetSplit = $true
            $textTarget.Text = $urlValue
            $script:SuppressTargetSplit = $false
            # Clear fields irrelevant for .url
            $textArgs.Text        = ""
            $textWorkDir.Text     = ""
            $textDescription.Text = ""
            $textAumid.Text       = ""
            # Fill icon
            if (-not (Test-StringEmpty $iconFilePath)) {
                $expandedIcon = [System.Environment]::ExpandEnvironmentVariables($iconFilePath)
                if ([IO.File]::Exists($expandedIcon)) {
                    $iconExt = [IO.Path]::GetExtension($expandedIcon).ToLower()
                    if ($script:ExeExtensions -contains $iconExt -or $script:ImageExtensions -contains $iconExt) {
                        $script:CurrentIconIndex = $iconFileIndex
                        $radioIcon_AnyFile.Checked = $true
                        $iconPathTextbox.Text = $expandedIcon
                        Write-Log "  Icon from .url : $expandedIcon index $iconFileIndex"
                    }
                    else {
                        $iconPathTextbox.Text = ""
                        Write-Log "  Icon file unsupported extension : $expandedIcon" -Level Debug
                    }
                }
                else {
                    $iconPathTextbox.Text = ""
                    Write-Log "  Icon file not found : $expandedIcon" -Level Warning
                }
            }
            else {
                $iconPathTextbox.Text = ""
                Write-Log "  No icon specified in .url"
            }
        }
        else {
            # ── .lnk shortcut ──
            $existTarget   = [ShortcutHelper]::GetTargetPath($LnkPath)
            $existArgs     = [ShortcutHelper]::GetArguments($LnkPath)
            $existWorkDir  = [ShortcutHelper]::GetWorkingDirectory($LnkPath)
            $existIconPath = [ShortcutHelper]::GetIconPath($LnkPath)
            Write-Log "  Target : $existTarget | Args length : $($existArgs.Length) | WorkDir : $existWorkDir | IconPath : $existIconPath" -Level Debug
            # Fill right panel fields
            $textTarget.Text  = $existTarget
            $textArgs.Text    = $existArgs
            $textWorkDir.Text = $existWorkDir
            # Resolve PIDL-based shell shortcuts (Control Panel, Recycle Bin, etc.)
            if (Test-StringEmpty $existTarget) {
                try {
                    $shellPath = [ShortcutHelper]::GetParsedDisplayName($LnkPath)
                    if (-not (Test-StringEmpty $shellPath)) {
                        $script:SuppressTargetSplit = $true
                        $textTarget.Text = $shellPath
                        $script:SuppressTargetSplit = $false
                        Write-Log "  PIDL resolved to shell path : $shellPath"
                    }
                    else {
                        Write-Log "  PIDL resolution returned empty" -Level Warning
                    }
                }
                catch {
                    $script:SuppressTargetSplit = $false
                    Write-Log "  PIDL resolution failed : $($_.Exception.Message)" -Level Warning
                }
            }
            # Description
            try {
                $existDesc = [ShortcutHelper]::GetDescription($LnkPath)
                $textDescription.Text = $existDesc
                Write-Log "  Description : $existDesc" -Level Debug
            }
            catch {
                $textDescription.Text = ""
                Write-Log "  Could not read Description : $($_.Exception.Message)" -Level Debug
            }
            # AUMID
            try {
                $existAumid = [ShortcutHelper]::GetAppUserModelId($LnkPath)
                if (-not (Test-StringEmpty $existAumid)) {
                    $textAumid.Text = $existAumid
                    Write-Log "  AUMID : $existAumid" -Level Debug
                }
                else {
                    $textAumid.Text = ""
                    Write-Log "  AUMID : (empty or null)" -Level Debug
                }
            }
            catch {
                Write-Log "  Could not read AUMID : $($_.Exception.Message)" -Level Warning
            }
            # Fill left panel (icon)
            $isAdsIcon = (-not (Test-StringEmpty $existIconPath)) -and $existIconPath.EndsWith(":icon.ico", [System.StringComparison]::OrdinalIgnoreCase)
            if ($isAdsIcon) {
                $adsHostPath = $existIconPath.Substring(0, $existIconPath.Length - ":icon.ico".Length)
                $adsBytes = $null
                foreach ($adsSource in @($adsHostPath, $LnkPath)) {
                    if ((-not (Test-StringEmpty $adsSource)) -and [IO.File]::Exists($adsSource)) {
                        $adsBytes = [AdsHelper]::ReadStream($adsSource, "icon.ico")
                        if ($null -ne $adsBytes -and $adsBytes.Length -gt 0) {
                            Write-Log "  ADS icon read from : $adsSource ($($adsBytes.Length) bytes)" -Level Debug
                            break
                        }
                    }
                }
                if ($null -ne $adsBytes -and $adsBytes.Length -gt 0) {
                    $radioIcon_Base64.Checked = $true
                    $TextboxIcon_Base64.Text = [System.Convert]::ToBase64String($adsBytes)
                    Write-Log "  ADS icon imported ($($adsBytes.Length) bytes) -> Base64 mode"
                }
                else {
                    $radioIcon_AnyFile.Checked = $true
                    $iconPathTextbox.Text = ""
                    Write-Log "  ADS icon path detected but stream empty on both '$adsHostPath' and '$LnkPath' -> File mode (empty)" -Level Warning
                }
            }
            else {
                if (-not (Test-StringEmpty $existIconPath)) {
                    $cleanIconPath = [System.Environment]::ExpandEnvironmentVariables($existIconPath)
                    if ($cleanIconPath.Contains(":") -and -not [System.IO.Path]::IsPathRooted($cleanIconPath.Substring(0,2))) {
                        $cleanIconPath = ""
                    }
                    try { $importIconIndex = [ShortcutHelper]::GetIconIndex($LnkPath) }
                    catch { $importIconIndex = 0 }
                    # Detect self-reference
                    $importTarget = Get-CleanInput $textTarget.Text
                    $isSelfReference = (-not (Test-StringEmpty $cleanIconPath)) -and
                                      (-not (Test-StringEmpty $importTarget)) -and
                                      [string]::Equals($cleanIconPath, $importTarget, [System.StringComparison]::OrdinalIgnoreCase)
                    if ($isSelfReference -or (Test-StringEmpty $cleanIconPath)) {
                        $radioIcon_TargetDefault.Checked = $true
                        Write-Log "  Icon is self-reference or empty -> Target (Default) mode"
                    }
                    else {
                        $script:CurrentIconIndex = $importIconIndex
                        $radioIcon_AnyFile.Checked = $true
                        $iconPathTextbox.Text = $cleanIconPath
                        Write-Log "  Icon from file : $cleanIconPath index $($script:CurrentIconIndex)"
                    }
                }
                else {
                    $radioIcon_TargetDefault.Checked = $true
                    Write-Log "  No explicit icon -> Target (Default) mode"
                }
            }
        }
        $labelLnkInfo.Text = "Existing $importExt detected : fields auto-filled"
        $labelLnkInfo.ForeColor = if ($script:IsDarkMode) { [System.Drawing.Color]::FromArgb(100,180,255) } else { [System.Drawing.Color]::FromArgb(0,100,180) }
    }
    catch {
        Write-Log "Error importing shortcut : $($_.Exception.Message)" -Level Error
        $labelLnkInfo.Text = "Existing shortcut detected (import error)"
        $labelLnkInfo.ForeColor = [System.Drawing.Color]::Red
    }
    finally {
        $script:SuppressPreviewUpdate = $false
        $script:SuppressAutoFill = $false
        $form.UseWaitCursor = $false
        $form.Cursor = [System.Windows.Forms.Cursors]::Default
        try { Update-IconPreview } catch {
            Write-Log "Update-IconPreview failed during import : $($_.Exception.Message)" -Level Warning
        }
    }
    Update-CreateButtonState
}

# ── Apply a resolved icon path+index to the left panel (returns $true on success) ──
function Set-ResolvedIconSource {
    param([string]$Path, [int]$Index = 0, [string]$Source = '')
    if (Test-StringEmpty $Path -or -not [IO.File]::Exists($Path)) {
        if (-not (Test-StringEmpty $Source)) { Write-Log "  [$Source] no result" -Level Debug }
        return $false
    }
    $iconExt = [IO.Path]::GetExtension($Path).ToLower()
    if ($script:ExeExtensions -contains $iconExt) {
        $nativeSizes = [IcoBuilder]::GetNativeSizes($Path, [Math]::Max(0, $Index))
        $peReadable  = -not ([IcoBuilder]::LastDiagnostic -like '*LoadLibraryEx failed*')
        if (-not $peReadable) {
            Write-Log "  [$Source] rejected (PE unreadable) : $Path | $([IcoBuilder]::LastDiagnostic)" -Level Debug
            return $false
        }
        if ($Index -ge 0 -and $nativeSizes.Length -eq 0) {
            Write-Log "  [$Source] rejected (no icons at index $Index) : $Path" -Level Debug
            return $false
        }
        if ($Index -lt 0 -and -not [IcoBuilder]::CanExtractIcon($Path, $Index)) {
            Write-Log "  [$Source] rejected (resource ID $Index not extractable) : $Path" -Level Debug
            return $false
        }
        $script:SuppressAutoFill = $true
        try {
            $script:CurrentIconIndex = $Index
            $radioIcon_AnyFile.Checked = $true
            $iconPathTextbox.Text = $Path
        }
        finally { $script:SuppressAutoFill = $false }
        Write-Log "  [$Source] accepted : $Path index $Index" -Level Debug
        return $true
    }
    if ($script:ImageExtensions -contains $iconExt) {
        $script:SuppressAutoFill = $true
        try {
            $script:CurrentIconIndex = 0
            $radioIcon_AnyFile.Checked = $true
            $iconPathTextbox.Text = $Path
        }
        finally { $script:SuppressAutoFill = $false }
        Write-Log "  [$Source] accepted : $Path (image)" -Level Debug
        return $true
    }
    Write-Log "  [$Source] unsupported extension : $Path ($iconExt)" -Level Debug
    return $false
}

# ── Resolve the icon of any file and fill the left panel accordingly ──
function Resolve-FileIcon {
    param([string]$FilePath)
    if (-not [IO.File]::Exists($FilePath)) { return }
    $ext = [IO.Path]::GetExtension($FilePath).ToLower()
    # ── Direct file types (image / executable) ──
    if (Set-ResolvedIconSource $FilePath 0) { return }
    # ── .lnk shortcut ──
    if ($ext -eq '.lnk') {
        Write-Log "Resolving icon for .lnk : $FilePath" -Level Debug
        $iconPath = [ShortcutHelper]::GetIconPath($FilePath)
        Write-Log "  .lnk IconPath : '$iconPath'" -Level Debug
        # ADS-embedded icon (only case where base64 is truly unavoidable)
        if ((-not (Test-StringEmpty $iconPath)) -and $iconPath.EndsWith(':icon.ico', [System.StringComparison]::OrdinalIgnoreCase)) {
            $adsHostPath = $iconPath.Substring(0, $iconPath.Length - ":icon.ico".Length)
            $adsBytes = $null
            # Try the referenced ADS host first, then the dropped LNK itself as fallback
            foreach ($adsSource in @($adsHostPath, $FilePath)) {
                if ((-not (Test-StringEmpty $adsSource)) -and [IO.File]::Exists($adsSource)) {
                    $adsBytes = [AdsHelper]::ReadStream($adsSource, "icon.ico")
                    if ($null -ne $adsBytes -and $adsBytes.Length -gt 0) {
                        Write-Log "  ADS icon read from : $adsSource ($($adsBytes.Length) bytes)" -Level Debug
                        break
                    }
                }
            }
            if ($null -ne $adsBytes -and $adsBytes.Length -gt 0) {
                $radioIcon_Base64.Checked = $true
                $TextboxIcon_Base64.Text = [System.Convert]::ToBase64String($adsBytes)
                Write-Log "Icon resolved (.lnk ADS) : $($adsBytes.Length) bytes"
                return
            }
        }
        # Icon location from shortcut properties
        if (-not (Test-StringEmpty $iconPath)) {
            $cleanPath = [System.Environment]::ExpandEnvironmentVariables($iconPath)
            $colonIdx  = $cleanPath.LastIndexOf(':')
            if ($colonIdx -gt 2) { $cleanPath = $cleanPath.Substring(0, $colonIdx) }
            if (Set-ResolvedIconSource $cleanPath ([ShortcutHelper]::GetIconIndex($FilePath))) { return }
        }
        # Recurse into shortcut target
        $lnkTarget = [ShortcutHelper]::GetTargetPath($FilePath)
        if ((-not (Test-StringEmpty $lnkTarget)) -and [IO.File]::Exists($lnkTarget)) {
            Write-Log "Icon fallback to .lnk target : $lnkTarget" -Level Debug
            Resolve-FileIcon $lnkTarget
        }
        else {
            # PIDL-based shortcut (Control Panel, shell folders, etc.)
            try {
                $shellPath = [ShortcutHelper]::GetParsedDisplayName($FilePath)
                if (-not (Test-StringEmpty $shellPath)) {
                    Write-Log "Icon fallback to .lnk PIDL : $shellPath" -Level Debug
                    # Try standard icon resolution chain first (file-based reference)
                    $resolvedIndex = 0
                    $resolvedPath = [IconResolver]::GetIconLocation($shellPath, [ref]$resolvedIndex)
                    if (Set-ResolvedIconSource $resolvedPath $resolvedIndex 'PIDL-IconLocation') { return }
                    $resolvedPath = [IconResolver]::GetAssocDefaultIcon(([IO.Path]::GetExtension($shellPath)), [ref]$resolvedIndex)
                    if (Set-ResolvedIconSource $resolvedPath $resolvedIndex 'PIDL-AssocDefault') { return }
                    # CLSID registry DefaultIcon lookup
                    $clsidMatches = [regex]::Matches($shellPath, '\{[0-9a-fA-F]{8}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{12}\}')
                    for ($ci = $clsidMatches.Count - 1; $ci -ge 0; $ci--) {
                        $clsid = $clsidMatches[$ci].Value
                        $clsidIconKey = [Microsoft.Win32.Registry]::ClassesRoot.OpenSubKey("CLSID\$clsid\DefaultIcon")
                        if ($null -ne $clsidIconKey) {
                            try {
                                $iconVal = $clsidIconKey.GetValue($null) -as [string]
                                if (-not (Test-StringEmpty $iconVal)) {
                                    $ref = Split-IconReference $iconVal
                                    Write-Log "  [PIDL-CLSID] $clsid -> $($ref.Path) index $($ref.Index)" -Level Debug
                                    if (Set-ResolvedIconSource $ref.Path $ref.Index 'PIDL-CLSID') { return }
                                }
                            }
                            finally { $clsidIconKey.Close() }
                        }
                    }
                    # Bitmap extraction as last resort
                    $shellBmp = [ShellDropHelper]::ExtractIconFromShellPath($shellPath)
                    if ($null -ne $shellBmp) {
                        $icoBytes = [IcoBuilder]::BuildFromBitmap($shellBmp)
                        $shellBmp.Dispose()
                        $radioIcon_Base64.Checked = $true
                        $TextboxIcon_Base64.Text = [System.Convert]::ToBase64String($icoBytes)
                        Write-Log "Icon resolved (.lnk PIDL bitmap) : $shellPath ($($icoBytes.Length) bytes)"
                    }
                    else { Write-Log "Icon extraction failed for PIDL : $shellPath" -Level Warning }
                }
                else { Write-Log "PIDL resolution returned empty for : $FilePath" -Level Warning }
            }
            catch { Write-Log "PIDL icon fallback failed : $($_.Exception.Message)" -Level Warning }
        }
        return
    }
    # ── .url internet shortcut (INI format with IconFile/IconIndex) ──
    if ($ext -eq '.url') {
        Write-Log "Resolving icon for .url : $FilePath" -Level Debug
        try {
            $iconFilePath = $null; $iconFileIndex = 0
            foreach ($line in [IO.File]::ReadAllLines($FilePath)) {
                $trimmed = $line.Trim()
                if     ($trimmed.StartsWith('IconFile=',  [System.StringComparison]::OrdinalIgnoreCase)) { $iconFilePath  = $trimmed.Substring(9).Trim() }
                elseif ($trimmed.StartsWith('IconIndex=', [System.StringComparison]::OrdinalIgnoreCase)) { [int]::TryParse($trimmed.Substring(10).Trim(), [ref]$iconFileIndex) | Out-Null }
            }
            if (-not (Test-StringEmpty $iconFilePath)) {
                if (Set-ResolvedIconSource ([System.Environment]::ExpandEnvironmentVariables($iconFilePath)) $iconFileIndex) { return }
            }
        } catch { Write-Log "Error reading .url : $($_.Exception.Message)" -Level Warning }
        # .url with no IconFile falls through to generic resolution below
    }
    # ── Generic resolution chain (for .url fallthrough and all other file types) ──
    Write-Log "Resolving icon for : $FilePath ($ext)" -Level Debug
    $shellSysIndex = 0
    if (-not (Test-StringEmpty $ext)) {
        $resolvedIndex = 0
        # Step 1 : SHGetFileInfo SHGFI_ICONLOCATION
        $resolvedPath = [IconResolver]::GetIconLocation($FilePath, [ref]$resolvedIndex)
        $shellSysIndex = $resolvedIndex
        Write-Log "  [SHGetFileInfo] $([IconResolver]::LastDiagnostic)" -Level Debug
        if (Set-ResolvedIconSource $resolvedPath $resolvedIndex 'SHGetFileInfo') { return }
        # Step 2 : AssocQueryString ASSOCSTR_DEFAULTICON
        $resolvedPath = [IconResolver]::GetAssocDefaultIcon($ext, [ref]$resolvedIndex)
        Write-Log "  [AssocDefaultIcon] $([IconResolver]::LastDiagnostic)" -Level Debug
        if (Set-ResolvedIconSource $resolvedPath $resolvedIndex 'AssocDefaultIcon') { return }
        # Step 3 : Registry - primary candidates first, then shell bitmap, then secondary
        $regPrimaryCount = 0
        $regCandidates = [IconResolver]::GetRegistryDefaultIcons($ext, [ref]$regPrimaryCount)
        Write-Log "  [Registry] $([IconResolver]::LastDiagnostic)" -Level Debug
        # Step 3a : try primary candidates (default ProgID, UserChoice)
        for ($ci = 0; $ci -lt $regPrimaryCount; $ci++) {
            $ref = Split-IconReference $regCandidates[$ci]
            Write-Log "  [Registry primary] $($ref.Path) index $($ref.Index)" -Level Debug
            if (Set-ResolvedIconSource $ref.Path $ref.Index 'Registry') { return }
        }
        # Step 3b : try secondary candidates (OpenWithProgids, conventional)
        for ($ci = $regPrimaryCount; $ci -lt $regCandidates.Length; $ci++) {
            $ref = Split-IconReference $regCandidates[$ci]
            # Skip standalone .ico files from third-party apps when shell has a known icon
            if ($shellSysIndex -gt 0 -and (-not (Test-StringEmpty $ref.Path))) {
                $candExt = [IO.Path]::GetExtension($ref.Path).ToLower()
                if ($candExt -eq '.ico') {
                    Write-Log "  [Registry secondary] skipped .ico (shell sysIndex=$shellSysIndex available) : $($ref.Path)" -Level Debug
                    continue
                }
            }
            Write-Log "  [Registry secondary] $($ref.Path) index $($ref.Index)" -Level Debug
            if (Set-ResolvedIconSource $ref.Path $ref.Index 'Registry') { return }
        }
        # Step 4 : Snap-in CLSID icon for .msc console files
        if ($ext -eq '.msc') {
            try {
                $mscContent = [IO.File]::ReadAllText($FilePath)
                $clsidMatch = [regex]::Match($mscContent, 'CLSID\s*=\s*"(\{[0-9a-fA-F\-]+\})"')
                if ($clsidMatch.Success) {
                    $snapinClsid = $clsidMatch.Groups[1].Value
                    Write-Log "  [MscSnapIn] CLSID : $snapinClsid" -Level Debug
                    $clsidIconKey = [Microsoft.Win32.Registry]::ClassesRoot.OpenSubKey("CLSID\$snapinClsid\DefaultIcon")
                    if ($null -ne $clsidIconKey) {
                        try {
                            $iconVal = $clsidIconKey.GetValue($null) -as [string]
                            if (-not (Test-StringEmpty $iconVal)) {
                                $ref = Split-IconReference $iconVal
                                Write-Log "  [MscSnapIn] DefaultIcon : $($ref.Path) index $($ref.Index)" -Level Debug
                                if (Set-ResolvedIconSource $ref.Path $ref.Index 'MscSnapIn') { return }
                            }
                        } finally { $clsidIconKey.Close() }
                    }
                }
                else { Write-Log "  [MscSnapIn] no CLSID found in XML" -Level Debug }
            }
            catch { Write-Log "  [MscSnapIn] parse failed : $($_.Exception.Message)" -Level Debug }
        }
        # Step 5 : Shell had a valid sysIndex - use JUMBO bitmap
        if ($shellSysIndex -gt 0) {
            Write-Log "  [ShellDirect] sysIndex=$shellSysIndex, all registry failed, extracting JUMBO bitmap" -Level Debug
            $shellBmp = [IconResolver]::ExtractShellIcon($FilePath)
            Write-Log "  [ShellDirect] $([IconResolver]::LastDiagnostic)" -Level Debug
            if ($null -ne $shellBmp) {
                $bmpWidth = $shellBmp.Width; $bmpHeight = $shellBmp.Height
                $icoBytes = [IcoBuilder]::BuildFromBitmap($shellBmp)
                $shellBmp.Dispose()
                $radioIcon_Base64.Checked = $true
                $TextboxIcon_Base64.Text = [System.Convert]::ToBase64String($icoBytes)
                Write-Log "Icon resolved (shell direct, sysIndex=$shellSysIndex) : $FilePath (${bmpWidth}x${bmpHeight})" -Level Info
                return
            }
        }
    }
    # Step 6 : Last resort - shell bitmap without prior sysIndex
    Write-Log "  [ShellBitmap] all prior steps failed, extracting shell bitmap" -Level Debug
    $shellBmp = [IconResolver]::ExtractShellIcon($FilePath)
    Write-Log "  [ShellBitmap] $([IconResolver]::LastDiagnostic)" -Level Debug
    if ($null -ne $shellBmp) {
        $bmpWidth = $shellBmp.Width; $bmpHeight = $shellBmp.Height
        $icoBytes = [IcoBuilder]::BuildFromBitmap($shellBmp)
        $shellBmp.Dispose()
        $radioIcon_Base64.Checked = $true
        $TextboxIcon_Base64.Text = [System.Convert]::ToBase64String($icoBytes)
        Write-Log "Icon resolved (shell bitmap fallback) : $FilePath (${bmpWidth}x${bmpHeight})" -Level Info
        return
    }
    Write-Log "Could not resolve icon for : $FilePath" -Level Warning
}

# Resolve a .cpl file to its Control Panel namespace item (shell path + display name)
function Resolve-CplControlPanelItem {
    param([string]$CplFilePath)
    $CplFileName = [IO.Path]::GetFileName($CplFilePath).ToLower()
    $CplBaseName = [IO.Path]::GetFileNameWithoutExtension($CplFilePath).ToLower()
    $CplShellApp = New-Object -ComObject Shell.Application
    $CplNamespace = $CplShellApp.Namespace('shell:ControlPanelFolder')
    $MatchedResult = $null
    foreach ($ControlPanelItem in $CplNamespace.Items()) {
        $ControlPanelItemPath = $ControlPanelItem.Path
        $AllGuidsInPath = [regex]::Matches($ControlPanelItemPath, '\{[0-9A-Fa-f\-]+\}')
        foreach ($GuidMatch in $AllGuidsInPath) {
            $CandidateGuid = $GuidMatch.Value
            $InprocRegistryKey = [Microsoft.Win32.Registry]::ClassesRoot.OpenSubKey("CLSID\$CandidateGuid\InprocServer32")
            if ($InprocRegistryKey) {
                $ModulePath = $InprocRegistryKey.GetValue($null, ''); $InprocRegistryKey.Close()
                if ($ModulePath -and [IO.Path]::GetFileName($ModulePath).ToLower() -eq $CplFileName) {
                    $MatchedResult = @{ Name = $ControlPanelItem.Name; Path = $ControlPanelItemPath }; break
                }
            }
            $DefaultIconKey = [Microsoft.Win32.Registry]::ClassesRoot.OpenSubKey("CLSID\$CandidateGuid\DefaultIcon")
            if ($DefaultIconKey) {
                $IconValue = $DefaultIconKey.GetValue($null, ''); $DefaultIconKey.Close()
                if ($IconValue -and $IconValue.ToLower().Contains($CplBaseName)) {
                    $MatchedResult = @{ Name = $ControlPanelItem.Name; Path = $ControlPanelItemPath }; break
                }
            }
        }
        [void][Runtime.InteropServices.Marshal]::ReleaseComObject($ControlPanelItem)
        if ($MatchedResult) { break }
    }
    [void][Runtime.InteropServices.Marshal]::ReleaseComObject($CplNamespace)
    [void][Runtime.InteropServices.Marshal]::ReleaseComObject($CplShellApp)
    return $MatchedResult
}

# ── Split a pasted string into target path and arguments ──
function Split-TargetAndArguments {
    param([string]$RawInput)
    $raw = $RawInput.Trim()
    if ($raw.Length -eq 0) { return @{ Target = ''; Arguments = '' } }
    # ── Case 1 : Quoted path at the start ──
    $firstChar = $raw[0]
    if ($firstChar -eq '"' -or $firstChar -eq "'") {
        $closeIdx = $raw.IndexOf($firstChar, 1)
        if ($closeIdx -gt 0) {
            $target    = $raw.Substring(1, $closeIdx - 1)
            $afterQuote = $closeIdx + 1
            $remainder = if ($afterQuote -lt $raw.Length) { $raw.Substring($afterQuote).TrimStart() } else { '' }
            return @{ Target = $target; Arguments = $remainder }
        }
    }
    # ── Case 2 : Find executable extension boundaries ──
    $allMatches = $script:ExtSplitRegex.Matches($raw)
    if ($allMatches.Count -gt 0) {
        # Pass 1 : try each match, check if the candidate path exists on disk
        for ($i = 0; $i -lt $allMatches.Count; $i++) {
            $splitPos = $allMatches[$i].Index + $allMatches[$i].Length
            $candidatePath = $raw.Substring(0, $splitPos)
            # Strip surrounding quotes without allocating a Trim char array
            if ($candidatePath.Length -gt 1 -and ($candidatePath[0] -eq '"' -or $candidatePath[0] -eq "'")) {
                $candidatePath = $candidatePath.Substring(1, $candidatePath.Length - (if ($candidatePath[$candidatePath.Length - 1] -eq $candidatePath[0]) { 2 } else { 1 }))
            }
            if ([IO.File]::Exists($candidatePath)) {
                $remainder = if ($splitPos -lt $raw.Length) { $raw.Substring($splitPos).TrimStart() } else { '' }
                return @{ Target = $candidatePath; Arguments = $remainder }
            }
        }
        # Pass 2 : split at the first extension followed by whitespace or end-of-string
        for ($i = 0; $i -lt $allMatches.Count; $i++) {
            $splitPos = $allMatches[$i].Index + $allMatches[$i].Length
            if ($splitPos -ge $raw.Length -or $raw[$splitPos] -eq ' ') {
                $target    = $raw.Substring(0, $splitPos).Trim('"', "'")
                $remainder = if ($splitPos -lt $raw.Length) { $raw.Substring($splitPos).TrimStart() } else { '' }
                return @{ Target = $target; Arguments = $remainder }
            }
        }
        # Pass 3 : fallback to last extension match
        $lastM    = $allMatches[$allMatches.Count - 1]
        $splitPos = $lastM.Index + $lastM.Length
        $target    = $raw.Substring(0, $splitPos).Trim('"', "'")
        $remainder = if ($splitPos -lt $raw.Length) { $raw.Substring($splitPos).TrimStart() } else { '' }
        return @{ Target = $target; Arguments = $remainder }
    }
    # ── Case 3 : No extension found, try to resolve first token via PATH ──
    $spaceIdx = $raw.IndexOf(' ')
    if ($spaceIdx -gt 0) {
        $firstToken = $raw.Substring(0, $spaceIdx)
        $argsStart  = $spaceIdx + 1
        # Check if the bare token resolves to an executable in any PATH directory
        $foundInPath = $false
        foreach ($dir in $script:PathDirs) {
            foreach ($ext in $script:ProbeExts) {
                if ([IO.File]::Exists([IO.Path]::Combine($dir, [string]::Concat($firstToken, $ext)))) {
                    $foundInPath = $true
                    break
                }
            }
            if ($foundInPath) { break }
        }
        if ($foundInPath) {
            return @{ Target = $firstToken; Arguments = $raw.Substring($argsStart).TrimStart() }
        }
    }
    # ── Case 4 : Cannot determine split point, everything is the target ──
    return @{ Target = $raw; Arguments = '' }
}

# ── Apply target/arguments split to the textboxes ──
function Invoke-TargetSplit {
    param([string]$Text)
    $parsed = Split-TargetAndArguments $Text
    if (Test-StringEmpty $parsed.Arguments) { return }
    $existingArgs = $textArgs.Text
    $hasExistingArgs = -not (Test-StringEmpty $existingArgs)
    # If args were manually typed (not from a previous auto-split), ask before replacing
    if ($hasExistingArgs -and -not $script:ArgsFromAutoSplit) {
        $result = [System.Windows.Forms.MessageBox]::Show(
            "The Arguments field already contains manually entered text :`n`n$($existingArgs.Substring(0, [Math]::Min($existingArgs.Length, 120)))$(if ($existingArgs.Length -gt 120) {'...'})`n`nReplace with the new split arguments ?`n`n$($parsed.Arguments.Substring(0, [Math]::Min($parsed.Arguments.Length, 120)))$(if ($parsed.Arguments.Length -gt 120) {'...'})",
            "Arguments conflict",
            [System.Windows.Forms.MessageBoxButtons]::YesNo,
            [System.Windows.Forms.MessageBoxIcon]::Question)
        if ($result -ne [System.Windows.Forms.DialogResult]::Yes) {
            Write-Log "User declined auto-split argument replacement" -Level Debug
            return
        }
    }
    $script:SuppressTargetSplit = $true
    try {
        $textTarget.Text = $parsed.Target
        $textArgs.Text   = $parsed.Arguments
        $script:ArgsFromAutoSplit = $true
        Write-Log "Auto-split target : '$($parsed.Target)' | args : '$($parsed.Arguments)'" -Level Debug
    }
    finally {
        $script:SuppressTargetSplit = $false
    }
}

# ── Process a dropped file based on zone ──
function Invoke-ZoneDrop {
    param([string]$FilePath, [string]$Zone)
    $form.UseWaitCursor = $true
    $form.Cursor = [System.Windows.Forms.Cursors]::WaitCursor
    [System.Windows.Forms.Application]::DoEvents()
    if (Test-StringEmpty $Zone) { return }
    Write-Log "Drop received : '$FilePath' on zone '$Zone'"
    $ext = [System.IO.Path]::GetExtension($FilePath).ToLower()
    switch ($Zone) {
        'icon' {
            $script:SuppressPreviewUpdate = $true
            try { Resolve-FileIcon $FilePath }
            finally {
                $script:SuppressPreviewUpdate = $false
                Update-IconPreview
            }
        }
        'target' {
            if ($ext -eq '.lnk') {
                $lnkTarget = [ShortcutHelper]::GetTargetPath($FilePath)
                if (-not (Test-StringEmpty $lnkTarget)) {
                    $script:SuppressTargetSplit = $true
                    try {
                        $textTarget.Text = $lnkTarget
                        $lnkArgs = [ShortcutHelper]::GetArguments($FilePath)
                        if (-not (Test-StringEmpty $lnkArgs)) { $textArgs.Text = $lnkArgs }
                        $lnkWorkDir = [ShortcutHelper]::GetWorkingDirectory($FilePath)
                        if (-not (Test-StringEmpty $lnkWorkDir)) { $textWorkDir.Text = $lnkWorkDir }
                        try {
                            $lnkDesc = [ShortcutHelper]::GetDescription($FilePath)
                            if (-not (Test-StringEmpty $lnkDesc)) { $textDescription.Text = $lnkDesc }
                        } catch {}
                        try {
                            $lnkAumid = [ShortcutHelper]::GetAppUserModelId($FilePath)
                            if (-not (Test-StringEmpty $lnkAumid)) { $textAumid.Text = $lnkAumid }
                        } catch {}
                    }
                    finally { $script:SuppressTargetSplit = $false }
                    Write-Log "Target resolved from .lnk : $lnkTarget"
                    # Auto-fill icon : switch to Target Default if icon source is currently empty
                    if (Test-IconSourceEmpty) {
                        $radioIcon_TargetDefault.Checked = $true
                    }
                }
                else {
                    # PIDL-based shortcut
                    try {
                        $shellPath = [ShortcutHelper]::GetParsedDisplayName($FilePath)
                        if (-not (Test-StringEmpty $shellPath)) {
                            $script:SuppressTargetSplit = $true
                            $textTarget.Text = $shellPath
                            $script:SuppressTargetSplit = $false
                            Write-Log "Target resolved from .lnk (PIDL) : $shellPath"
                        }
                    }
                    catch { Write-Log "PIDL resolution failed on target drop : $($_.Exception.Message)" -Level Warning }
                    # Auto-fill icon : switch to Target Default if icon source is currently empty
                    if (Test-IconSourceEmpty) {
                        $radioIcon_TargetDefault.Checked = $true
                    }
                }
            }
            elseif ($ext -eq '.url') {
                try {
                    foreach ($line in [IO.File]::ReadAllLines($FilePath)) {
                        $trimmed = $line.Trim()
                        if ($trimmed.StartsWith('URL=', [System.StringComparison]::OrdinalIgnoreCase)) {
                            $textTarget.Text = $trimmed.Substring(4).Trim()
                            Write-Log "Target resolved from .url : $($textTarget.Text)"
                            break
                        }
                    }
                }
                catch { Write-Log "Error reading .url for target : $($_.Exception.Message)" -Level Warning }
            }
            else {
                $textTarget.Text = $FilePath
                # Auto-fill icon
                if ($ext -eq '.cpl') {
                    # .cpl : Target Default would show a generic icon, resolve actual icon into Any File mode
                    if ($radioIcon_TargetDefault.Checked -or (Test-IconSourceEmpty)) {
                        $script:SuppressPreviewUpdate = $true
                        try { Resolve-FileIcon $FilePath }
                        finally {
                            $script:SuppressPreviewUpdate = $false
                            Update-IconPreview
                        }
                    }
                }
                elseif (Test-IconSourceEmpty) {
                    $radioIcon_TargetDefault.Checked = $true
                }
                # Auto-fill working directory if empty
                if (Test-StringEmpty (Get-CleanInput $textWorkDir.Text)) {
                    if ([IO.Directory]::Exists($FilePath)) {
                        $textWorkDir.Text = $FilePath
                    }
                    else {
                        $textWorkDir.Text = [IO.Path]::GetDirectoryName($FilePath)
                    }
                }
            }
        }
        'workdir' {
            if ($ext -eq '.lnk') {
                $lnkWorkDir = [ShortcutHelper]::GetWorkingDirectory($FilePath)
                if (-not (Test-StringEmpty $lnkWorkDir)) {
                    $textWorkDir.Text = $lnkWorkDir
                    Write-Log "WorkDir resolved from .lnk property : $lnkWorkDir"
                }
                else {
                    # Fallback : derive from the shortcut's target path
                    $lnkTarget = [ShortcutHelper]::GetTargetPath($FilePath)
                    if ((-not (Test-StringEmpty $lnkTarget)) -and [IO.File]::Exists($lnkTarget)) {
                        $textWorkDir.Text = [IO.Path]::GetDirectoryName($lnkTarget)
                        Write-Log "WorkDir derived from .lnk target : $($textWorkDir.Text)"
                    }
                    else {
                        # Last resort : directory containing the .lnk itself
                        $textWorkDir.Text = [IO.Path]::GetDirectoryName($FilePath)
                        Write-Log "WorkDir fallback to .lnk location : $($textWorkDir.Text)" -Level Debug
                    }
                }
            }
            else {
                $dir = if ([IO.Directory]::Exists($FilePath)) { $FilePath }
                       else { [IO.Path]::GetDirectoryName($FilePath) }
                $textWorkDir.Text = $dir
            }
        }
        'lnk' {
            $hasExistingFields = -not (Test-StringEmpty (Get-CleanInput $textTarget.Text))
            if ($hasExistingFields) { $script:SuppressAutoFill = $true }
            try {
                if ($ext -eq '.lnk' -or $ext -eq '.url') {
                    $textLnkPath.Text = $FilePath
                }
                elseif ([System.IO.Directory]::Exists($FilePath)) {
                    $textLnkPath.Text = $FilePath
                }
                else {
                    $lnkDir  = [System.IO.Path]::GetDirectoryName($FilePath)
                    $lnkName = [System.IO.Path]::GetFileNameWithoutExtension($FilePath) + ".lnk"
                    $textLnkPath.Text = [System.IO.Path]::Combine($lnkDir, $lnkName)
                }
            }
            finally {
                if ($hasExistingFields) { $script:SuppressAutoFill = $false }
            }
        }
    }
    $form.UseWaitCursor = $false
    $form.Cursor = [System.Windows.Forms.Cursors]::Default
}

# ── Write ICO bytes to an NTFS Alternate Data Stream ──
function Write-AdsIcon {
    param([string]$FilePath, [byte[]]$IcoBytes)
    [AdsHelper]::WriteStream($FilePath, "icon.ico", $IcoBytes)
    Write-Log "ADS icon written to : ${FilePath}:icon.ico ($($IcoBytes.Length) bytes)" -Level Debug
}

# ── Verify an NTFS ADS icon stream exists and return its length ──
function Get-AdsIconLength {
    param([string]$FilePath)
    return [AdsHelper]::GetStreamLength($FilePath, "icon.ico")
}

# ── Recursive dark scrollbar applicator ──
function Set-DarkScrollbars {
    param([System.Windows.Forms.Control]$Root, [bool]$Dark)
    $types = @('TextBox','RichTextBox','FlowLayoutPanel','Panel','CheckedListBox')
    $queue = New-Object System.Collections.Queue
    $queue.Enqueue($Root)
    while ($queue.Count -gt 0) {
        $ctrl = $queue.Dequeue()
        if ($ctrl.IsHandleCreated -and ($types -contains $ctrl.GetType().Name)) {
            [DarkMode]::ApplyControl($ctrl.Handle, $Dark)
        }
        foreach ($child in $ctrl.Controls) { $queue.Enqueue($child) }
    }
}

#region ── START MENU SHORTCUT (ADS-embedded icon) ─

Update-LoadingPopup 55 "Loading..."

# Multi-size Icon from full ICO byte array (supports all DPI scaling levels)
$taskIconStream = New-Object System.IO.MemoryStream(,$script:AppIcoBytes)
$taskIcon       = New-Object System.Drawing.Icon($taskIconStream)

if (-not (Test-StringEmpty $batFile)) {
    $startMenuLnk = [IO.Path]::Combine($script:StartMenuDir, $script:LnkName)
    $needsUpdate = $true
    if ([IO.File]::Exists($startMenuLnk)) {
        try {
            $desc = [ShortcutHelper]::GetDescription($startMenuLnk)
            if ($desc.Contains('[UserPinned]')) { $script:UserPinnedStartMenu = $true; $needsUpdate = $false }
            elseif ([ShortcutHelper]::GetTargetPath($startMenuLnk) -eq $batFile) { $needsUpdate = $false }
        } catch {}
    }
    if ($needsUpdate -and -not $script:UserPinnedStartMenu) {
        try {
            [ShortcutHelper]::CreateWithEmbeddedIcon(
                $startMenuLnk, $batFile, "", $script:AppIcoBytes, $script:AppId,
                "$($script:AppName) v$($script:Version)")
            Write-AdsIcon -FilePath $startMenuLnk -IcoBytes $script:AppIcoBytes
            Write-Log "Auto-registered Start Menu shortcut with ADS icon (temporary)"
        } catch { Write-Log "Failed to auto-register shortcut : $_" -Level Warning }
    }
}

#region ── MAIN FORM ─

Update-LoadingPopup 60 "Loading..."

$titleBarHeight = 34

$form = New-Object CustomForm
$form.SuspendLayout()
$form.AutoScaleDimensions = New-Object System.Drawing.SizeF(96,96)
$form.AutoScaleMode   = [System.Windows.Forms.AutoScaleMode]::DPI
$form.ShowInTaskbar   = $true
$form.Text            = "$($script:AppName) v$($script:Version)"
$form.FormBorderStyle = 'None'
$form.BackColor       = [System.Drawing.Color]::FromArgb(243, 243, 243)
$form.Icon            = $taskIcon
$form.ShowIcon        = $true
$form.StartPosition   = [System.Windows.Forms.FormStartPosition]::CenterScreen
$form.Font            = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Regular, [System.Drawing.GraphicsUnit]::Point)
$form.ClientSize      = New-Object System.Drawing.Size(840, 540)

# Subtle border (theme-aware)
$form.Add_Paint({
    param($s, $e)
    $pen = if ($script:IsDarkMode) { $script:FormBorderPenDark } else { $script:FormBorderPenLight }
    $e.Graphics.DrawRectangle($pen, 0, 0, $s.ClientSize.Width - 1, $s.ClientSize.Height - 1)
})

#region ── TITLE BAR ─

$titleBar = gen $null "Win11TitleBar" "Dock=Top" "BackColor=240 240 240"
$titleBar.Height = $titleBarHeight

# Title bar icon with high-quality zoom (DPI-aware via custom paint)
$iconBox = gen $titleBar "PictureBox" 8 ([int](($titleBarHeight-20)/2)) 20 20 "SizeMode=Normal" "BackColor=240 240 240"
$iconBox.Tag = $iconImage
$iconBox.Add_Paint({
    param($s, $e)
    $g = $e.Graphics
    $g.InterpolationMode  = [System.Drawing.Drawing2D.InterpolationMode]::HighQualityBicubic
    $g.SmoothingMode      = [System.Drawing.Drawing2D.SmoothingMode]::HighQuality
    $g.PixelOffsetMode    = [System.Drawing.Drawing2D.PixelOffsetMode]::HighQuality
    $g.CompositingQuality = [System.Drawing.Drawing2D.CompositingQuality]::HighQuality
    $g.DrawImage($s.Tag, 0, 0, $s.Width, $s.Height)
})
Enable-HitTestPassThrough $iconBox

# Title label (explicit font : leaf control)
$titleLabel = gen $titleBar "Label" "$($script:AppName) v$($script:Version)" 34 0 145 $titleBarHeight "ForeColor=30 30 30" "Font=Arial, 10, Bold" "BackColor=240 240 240" "TextAlign=MiddleLeft"
Enable-HitTestPassThrough $titleLabel

$script:HasMDL2 = (New-Object System.Drawing.Text.InstalledFontCollection).Families | Where-Object { $_.Name -eq 'Segoe MDL2 Assets' }

# Window buttons (Dock=Right : first added = rightmost)
$btnMinimize = gen $titleBar "TitleBarButton" $(if($script:HasMDL2){[string][char]0xE921}else{[string][char]0x2015}) "Dock=Right" "Font=$(if($script:HasMDL2){'Segoe MDL2 Assets, 10'}else{'Arial, 10'})"
$btnMinimize.Add_Click({ $form.WindowState = 'Minimized' })

$btnClose = gen $titleBar "TitleBarButton" "r" "Dock=Right" "IsCloseButton=$true" "Font=Marlett, 10"
$btnClose.Add_Click({ $form.Close() })

$titleBar.Controls.AddRange(@($iconBox, $titleLabel))

# Console toggle (explicit font : leaf control)
$chkConsole = gen $titleBar "CheckBox" "Show console" 0 ([int](($titleBarHeight-18)/2)) 101 18 "Font=Arial, 8" "BackColor=240 240 240" "FlatStyle=Flat" "Anchor=Top,Right"
$chkConsole.Location = New-Object System.Drawing.Point(($btnMinimize.Left - $chkConsole.Width - 8), $chkConsole.Location.Y)
$chkConsole.Add_CheckedChanged({
    $h = [NativeMethods]::GetConsoleWindow()
    if ($h -ne [IntPtr]::Zero) {
        if ($this.Checked) { [NativeMethods]::ShowWindow($h, [NativeMethods]::SW_SHOW) | Out-Null }
        else               { [NativeMethods]::ShowWindow($h, [NativeMethods]::SW_HIDE) | Out-Null }
    }
})
$chkConsole.BringToFront()

# Dark mode toggle (explicit font : leaf control)
$chkDarkMode = gen $titleBar "CheckBox" "Dark mode" 0 ([int](($titleBarHeight-18)/2)) 85 18 "Font=Arial, 8" "BackColor=240 240 240" "FlatStyle=Flat" "Anchor=Top,Right"
$chkDarkMode.Location = New-Object System.Drawing.Point(($chkConsole.Left - $chkDarkMode.Width - 4), $chkDarkMode.Location.Y)
$chkDarkMode.BringToFront()

# Reset button
$btnReset = gen $titleBar "TitleBarButton" $(if($script:HasMDL2){[string][char]0xE72C}else{[string][char]0x21BB}) "Dock=None" "Font=$(if($script:HasMDL2){'Segoe MDL2 Assets, 8'}else{'Arial, 10'})"
$btnReset.Size = New-Object System.Drawing.Size(36, 20)
$btnReset.Location = New-Object System.Drawing.Point(($titleLabel.Right + 6), 8)
$script:ToolTip = New-Object System.Windows.Forms.ToolTip
$script:ToolTip.SetToolTip($btnReset, "Reset all fields")
$btnReset.Add_Click({
    $script:SuppressAutoFill      = $true
    $script:SuppressPreviewUpdate = $true
    $script:SuppressTargetSplit   = $true
    try {
        # Right panel
        $textTarget.Text      = ""
        $textArgs.Text        = ""
        $textWorkDir.Text     = ""
        $textDescription.Text = ""
        $textAumid.Text       = ""
        $textLnkPath.Text     = ""
        $labelLnkInfo.Text    = ""
        $labelExplorerWarn.Text = ""
        # Left panel
        $radioIcon_TargetDefault.Checked = $true
        $iconPathTextbox.Text            = ""
        $TextboxIcon_Base64.Text   = ""
        $radioEmbedIcon.Checked    = $true
        $script:CurrentIconIndex    = 0
        $script:CurrentExeIconCount = 0
        $script:LastIconMenuPath    = ""
        $script:ArgsFromAutoSplit    = $false
        $script:PreviousTargetText   = ""
        $script:LastPinnedTaskbarFile = ""
        # Reset special modes
        $script:IsShellTarget = $true
        Set-SpecialTargetMode 'Shell' $false
        $script:IsUrlTarget = $true
        Set-SpecialTargetMode 'Url' $false
    }
    finally {
        $script:SuppressAutoFill      = $false
        $script:SuppressPreviewUpdate = $false
        $script:SuppressTargetSplit   = $false
        Update-IconPreview
        Update-ArgsLength
        Update-CreateButtonState
    }
    Write-Log "Interface reset by user"
})

# About button (notch-style trapezoid)
$script:notchTopWidth    = 96
$script:notchBottomWidth = 64
$script:notchHeight      = 21

$btnAbout = gen $null "DoubleBufferedPanel" 0 0 $script:notchTopWidth $script:notchHeight "BackColor=228 228 228"
$btnAbout.Location = New-Object System.Drawing.Point(([int](($form.ClientSize.Width - $script:notchTopWidth) / 2)), 0)
$btnAbout.Cursor   = [System.Windows.Forms.Cursors]::Hand

$btnAbout.Add_Resize({
    $w = $this.Width
    $h = $this.Height
    $inset = [int](($w - ($w * $script:notchBottomWidth / $script:notchTopWidth)) / 2)
    $path = New-Object System.Drawing.Drawing2D.GraphicsPath
    $path.AddPolygon(@(
        (New-Object System.Drawing.Point(0, 0)),
        (New-Object System.Drawing.Point($w, 0)),
        (New-Object System.Drawing.Point(($w - $inset), $h)),
        (New-Object System.Drawing.Point($inset, $h))
    ))
    if ($this.Region) { $this.Region.Dispose() }
    $this.Region = New-Object System.Drawing.Region($path)
    $path.Dispose()
})
$btnAbout.GetType().GetMethod('OnResize', [System.Reflection.BindingFlags]'NonPublic,Instance').Invoke($btnAbout, @([System.EventArgs]::Empty))

$btnAbout.Add_Paint({
    param($s, $e)
    try {
        $g = $e.Graphics
        $g.SmoothingMode     = [System.Drawing.Drawing2D.SmoothingMode]::AntiAlias
        $g.TextRenderingHint = [System.Drawing.Text.TextRenderingHint]::ClearTypeGridFit
        $borderC = if ($script:IsDarkMode) { [System.Drawing.Color]::FromArgb(70,70,70) } else { [System.Drawing.Color]::FromArgb(180,180,180) }
        $borderPen = New-Object System.Drawing.Pen($borderC, 1)
        $ins = [int](($s.Width - ($s.Width * $script:notchBottomWidth / $script:notchTopWidth)) / 2)
        $g.DrawLine($borderPen, 0, 0, $ins, $s.Height-1)
        $g.DrawLine($borderPen, $ins, $s.Height-1, $s.Width-$ins, $s.Height-1)
        $g.DrawLine($borderPen, $s.Width-$ins, $s.Height-1, $s.Width, 0)
        $borderPen.Dispose()
        $sf = New-Object System.Drawing.StringFormat
        $sf.Alignment = $sf.LineAlignment = [System.Drawing.StringAlignment]::Center
        $font = New-Object System.Drawing.Font("Arial", 8)
        $textBr = if ($script:IsDarkMode) { [System.Drawing.Brushes]::White } else { [System.Drawing.Brushes]::Black }
        $g.DrawString("About", $font, $textBr, (New-Object System.Drawing.RectangleF(0,1,$s.Width,$s.Height)), $sf)
        $font.Dispose(); $sf.Dispose()
    }
    catch {
        Write-Log "About notch paint error : $($_.Exception.Message)" -Level Debug
    }
})

$btnAbout.Add_MouseEnter({ $this.BackColor = $script:AboutHoverColor;  $this.Invalidate() })
$btnAbout.Add_MouseLeave({ $this.BackColor = $script:AboutNormalColor; $this.Invalidate() })

$btnAbout.Add_Click({
    try {
        $isDk   = $script:IsDarkMode
        $bgCol  = if ($isDk) { [System.Drawing.Color]::FromArgb(45,45,48) }    else { [System.Drawing.Color]::FromArgb(243,243,243) }
        $fgCol  = if ($isDk) { [System.Drawing.Color]::White }                  else { [System.Drawing.Color]::Black }
        $hdrBg  = if ($isDk) { [System.Drawing.Color]::FromArgb(32,32,32) }    else { [System.Drawing.Color]::White }
        $subFg  = if ($isDk) { [System.Drawing.Color]::FromArgb(170,170,170) } else { [System.Drawing.Color]::FromArgb(120,120,120) }
        # About dialog
        $aboutForm = gen $null "Form" "About" 0 0 380 310 "StartPosition=CenterParent" 'MaximizeBox=$false' 'MinimizeBox=$false' 'FormBorderStyle=FixedDialog' "BackColor=$bgCol" "ForeColor=$fgCol"
        $aboutForm.SuspendLayout()
        $aboutForm.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Regular, [System.Drawing.GraphicsUnit]::Point)
        $aboutForm.AutoScaleDimensions = New-Object System.Drawing.SizeF(96, 96)
        $aboutForm.AutoScaleMode = [System.Windows.Forms.AutoScaleMode]::Dpi
        $aboutForm.ClientSize = New-Object System.Drawing.Size(380, 310)
        $headerPanel = gen $aboutForm "Panel" 0 0 380 90 "BackColor=$hdrBg"
        $headerIcon = gen $headerPanel "PictureBox" 20 13 64 64 "SizeMode=Normal"
        $headerIcon.Tag = $iconImage
        $headerIcon.Add_Paint({
            param($s, $e)
            try {
                $g = $e.Graphics
                $g.InterpolationMode  = [System.Drawing.Drawing2D.InterpolationMode]::HighQualityBicubic
                $g.SmoothingMode      = [System.Drawing.Drawing2D.SmoothingMode]::HighQuality
                $g.PixelOffsetMode    = [System.Drawing.Drawing2D.PixelOffsetMode]::HighQuality
                $g.CompositingQuality = [System.Drawing.Drawing2D.CompositingQuality]::HighQuality
                $g.DrawImage($s.Tag, 0, 0, $s.Width, $s.Height)
            }
            catch {}
        })
        gen $headerPanel "Label" "$($script:AppName)" 94 12 0 0 "Font=Arial, 18, Bold" 'AutoSize=$true' "ForeColor=$fgCol" | Out-Null
        gen $headerPanel "Label" "v$($script:Version)" 94 48 0 0 "Font=Arial, 10" 'AutoSize=$true' "ForeColor=$subFg" | Out-Null
        $btnShowLogs = gen $headerPanel "Button" "Show logs" 160 56 75 26 "FlatStyle=Flat" "Font=Arial, 8" "AutoSize=$true" "BackColor=$bgCol" "ForeColor=$fgCol" 'TabStop=$false'
        $btnShowLogs.FlatAppearance.BorderColor = if ($isDk) { [System.Drawing.Color]::FromArgb(80,80,80) } else { [System.Drawing.Color]::FromArgb(180,180,180) }
        $btnShowLogs.Add_Click({ [System.Diagnostics.Process]::Start("explorer.exe", "`"$($script:LogDir)`"") })
        # Pin Start Menu button
        $btnPinStart = gen $headerPanel "Button" "" 240 56 130 26 "FlatStyle=Flat" "Font=Arial, 8"  "AutoSize=$true" "BackColor=$bgCol" "ForeColor=$fgCol" 'TabStop=$false'
        $btnPinStart.FlatAppearance.BorderColor = if ($isDk) { [System.Drawing.Color]::FromArgb(80,80,80) } else { [System.Drawing.Color]::FromArgb(180,180,180) }
        $btnPinStart.Text = if ($script:UserPinnedStartMenu) { [char]0x2714 + " Start Menu" } else { "Keep in Start Menu" }
        $btnPinStart.Tag  = $script:UserPinnedStartMenu
        $btnPinStart.Add_Click({
            $lnk = [IO.Path]::Combine($script:StartMenuDir, $script:LnkName)
            if ($this.Tag) {
                try {
                    if ([IO.File]::Exists($lnk)) { [IO.File]::Delete($lnk) }
                    $script:UserPinnedStartMenu = $false; $this.Tag = $false; $this.Text = "Keep in Start Menu"
                    [ShortcutHelper]::CreateWithEmbeddedIcon(
                        $lnk, $batFile, "", $script:AppIcoBytes, $script:AppId,
                        "$($script:AppName) v$($script:Version)")
                    Write-AdsIcon -FilePath $lnk -IcoBytes $script:AppIcoBytes
                    Write-Log "User unpinned Start Menu shortcut"
                } catch { Write-Log "Failed to revert shortcut : $_" -Level Warning }
            } else {
                try {
                    [ShortcutHelper]::CreateWithEmbeddedIcon(
                        $lnk, $batFile, "", $script:AppIcoBytes, $script:AppId,
                        "$($script:AppName) v$($script:Version) [UserPinned]")
                    Write-AdsIcon -FilePath $lnk -IcoBytes $script:AppIcoBytes
                    $script:UserPinnedStartMenu = $true; $this.Tag = $true; $this.Text = [char]0x2714 + " Start Menu"
                    Write-Log "User pinned Start Menu shortcut"
                } catch { Write-Log "Failed to create shortcut : $_" -Level Warning }
            }
        })
        $lblMadeBy = gen $aboutForm "Label" "Made by Léo Gillet - Freenitial" 20 108 0 0 "Font=Arial, 10" "AutoSize=$true" "ForeColor=$fgCol"
        $linkGitHub = gen $aboutForm "LinkLabel" "(GitHub)" ($lblMadeBy.Location.X + $lblMadeBy.PreferredWidth + 2) 108 0 0 "Font=Arial, 10" "AutoSize=$true"
        $linkGitHub.LinkColor        = if ($isDk) { [System.Drawing.Color]::FromArgb(100,180,255) } else { [System.Drawing.Color]::FromArgb(0,102,204) }
        $linkGitHub.ActiveLinkColor  = if ($isDk) { [System.Drawing.Color]::FromArgb(140,200,255) } else { [System.Drawing.Color]::FromArgb(0,80,180) }
        $linkGitHub.VisitedLinkColor = $linkGitHub.LinkColor
        $linkGitHub.Add_LinkClicked({ [System.Diagnostics.Process]::Start("https://github.com/Freenitial") })
        gen $aboutForm "Label" "Embeds icons directly into .lnk files via NTFS ADS." 20 131 0 0 "Font=Arial, 9" "AutoSize=$true" "ForeColor=$subFg" | Out-Null
        gen $aboutForm "Panel" "" 20 160 340 2 "BorderStyle=FixedSingle" | Out-Null
        gen $aboutForm "Label" "Changelog :" 20 174 0 0 "Font=Arial, 10, Bold" "AutoSize=$true" "ForeColor=$fgCol" | Out-Null
        $aboutFormText = @"
$([char]0x2022)  v1.0 : Initial release
"@
        $changelogPanel = gen $aboutForm "Panel" "" 20 195 ($aboutForm.ClientSize.Width - 40) ($aboutForm.ClientSize.Height - 210) "AutoScroll=$true" "BackColor=$bgCol"
        $changelogLabel = gen $changelogPanel "Label" $aboutFormText 5 0 0 0 "Font=Arial, 9" "AutoSize=$true" "ForeColor=$fgCol"
        $changelogLabel.MaximumSize = New-Object System.Drawing.Size(($changelogPanel.Width - 25), 0)
        $aboutForm.ResumeLayout($true)
        $aboutForm.ShowDialog($form) | Out-Null
    }
    catch {
        Write-Log "About dialog error : $($_.Exception.Message)" -Level Error
    }
    finally {
        if ($null -ne $headerIcon) { $headerIcon.Tag = $null }
        if ($null -ne $aboutForm)  { try { $aboutForm.Dispose() } catch {} }
    }
})

$titleBar.Add_Resize({ $btnAbout.Location = New-Object System.Drawing.Point(([int](($titleBar.Width - $script:notchTopWidth) / 2)), 0) })
$titleBar.Controls.Add($btnAbout)
$btnAbout.BringToFront()

$form.Controls.Add($titleBar)

#region ── LEFT PANEL : ICON SOURCE ─

Update-LoadingPopup 70 "Loading..."

$panelLeft = gen $null "Panel" 10 ($titleBarHeight + 10) 385 ($form.ClientSize.Height - $titleBarHeight - 70) "Anchor=Top,Left,Bottom"

# GroupBox panel left
$groupIcon = gen $panelLeft "GroupBox" "Icon Source (optional)" "Dock=Fill"
$groupIcon.Add_Paint($script:GroupBoxPaintHandler)
$dr = $groupIcon.DisplayRectangle
$xLeft = $dr.X + $script:GroupPadding
$innerWidth = $dr.Width - ($script:GroupPadding * 2)

# Icon preview button (clickable for icon index selection)
$btnIconPreview = gen $groupIcon "Button" "" (([int](($innerWidth - 96) / 2) + $xLeft)) 22 96 96 "FlatStyle=Flat" "BackColor=230 230 230" "BorderColor=LightGray" "HoverColor=220 220 230"
$btnIconPreview.Padding = [System.Windows.Forms.Padding]::Empty
$btnIconPreview.TabStop = $false
$btnIconPreview.Add_GotFocus({ $form.ActiveControl = $null })
$btnIconPreview.Add_Paint({
    param($s, $e)
    try {
        $g = $e.Graphics
        $g.InterpolationMode  = [System.Drawing.Drawing2D.InterpolationMode]::HighQualityBicubic
        $g.SmoothingMode      = [System.Drawing.Drawing2D.SmoothingMode]::HighQuality
        $g.PixelOffsetMode    = [System.Drawing.Drawing2D.PixelOffsetMode]::HighQuality
        $g.CompositingQuality = [System.Drawing.Drawing2D.CompositingQuality]::HighQuality
        # Draw centered preview image from Tag
        $previewBmp = $script:CurrentPreviewBitmap
        if ($null -ne $previewBmp) {
            $borderW = $s.FlatAppearance.BorderSize
            $innerW  = $s.Width  - ($borderW * 2)
            $innerH  = $s.Height - ($borderW * 2)
            # Center at natural size if smaller, scale down if larger
            if ($previewBmp.Width -lt $innerW -and $previewBmp.Height -lt $innerH) {
                $drawX = $borderW + [int](($innerW - $previewBmp.Width)  / 2)
                $drawY = $borderW + [int](($innerH - $previewBmp.Height) / 2)
                $g.DrawImage($previewBmp, $drawX, $drawY, $previewBmp.Width, $previewBmp.Height)
            }
            else {
                $scaleX = $innerW / [double]$previewBmp.Width
                $scaleY = $innerH / [double]$previewBmp.Height
                $scale  = [Math]::Min($scaleX, $scaleY)
                $drawW  = [int]($previewBmp.Width  * $scale)
                $drawH  = [int]($previewBmp.Height * $scale)
                $drawX  = $borderW + [int](($innerW - $drawW) / 2)
                $drawY  = $borderW + [int](($innerH - $drawH) / 2)
                $g.DrawImage($previewBmp, $drawX, $drawY, $drawW, $drawH)
            }
        }
        # Dropdown triangle indicator
        if ($script:CurrentExeIconCount -gt 1 -and $radioIcon_AnyFile.Checked) {
            $g.SmoothingMode = [System.Drawing.Drawing2D.SmoothingMode]::AntiAlias
            $triSize = 6
            $triX = $s.Width - $triSize * 2 - 4
            $triY = $s.Height - $triSize - 4
            $triColor = if ($script:IsDarkMode) { [System.Drawing.Color]::FromArgb(180,180,180) } else { [System.Drawing.Color]::FromArgb(80,80,80) }
            $brush = New-Object System.Drawing.SolidBrush($triColor)
            $points = @(
                (New-Object System.Drawing.Point($triX, $triY)),
                (New-Object System.Drawing.Point(($triX + $triSize * 2), $triY)),
                (New-Object System.Drawing.Point(($triX + $triSize), ($triY + $triSize)))
            )
            $g.FillPolygon($brush, $points)
            $brush.Dispose()
        }
    }
    catch {
        Write-Log "IconPreview paint error : $($_.Exception.Message)" -Level Debug
    }
})
# Icon index context menu
$script:IconIndexMenu = New-Object System.Windows.Forms.ContextMenuStrip
$script:IconIndexMenu.ImageScalingSize = New-Object System.Drawing.Size(32, 32)
$btnIconPreview.Add_Click({
    if (-not $radioIcon_AnyFile.Checked) { return }
    if ($script:CurrentExeIconCount -le 1) { return }
    $script:IconIndexMenu.Show($this, (New-Object System.Drawing.Point(0, $this.Height)))
})
$labelPreviewIconInfo = gen $groupIcon "Label" "No icon" $xLeft 122 $innerWidth 18 "TextAlign=MiddleCenter" "Anchor=Top,Left,Right"

# Icon source radio buttons
$labelIconSource = gen $groupIcon "Label" "Get icon from :" $xLeft 149 95 20 "Font=Segoe UI, 9, Bold"
$radioIcon_TargetDefault = gen $groupIcon "RadioButton" "Target (Default)" ($xLeft + 100) 148 118 22 "Checked=$true"
$radioIcon_AnyFile = gen $groupIcon "RadioButton" "Any file" ($xLeft + 222) 148 68 22
$radioIcon_Base64 = gen $groupIcon "RadioButton" "Base64" ($xLeft + 300) 148 65 22

# Target Default explanation label (visible by default since Target Default is the default radio)
$labelTargetDefaultInfo = gen $groupIcon "Label" "No custom icon. Windows will assign the default icon`nbased on the target type, same as a normal shortcut`ncreated from Explorer." $xLeft 176 $innerWidth 44 "Font=Segoe UI, 8.25"

# File path textbox + browse button (hidden by default, Target Default is selected)
$btnBrowseIcon = gen $groupIcon "Button" "Browse..." ($xLeft + $innerWidth - 80) 176 80 20 "Anchor=Top,Right" "Visible=$false"
$iconPathTextbox = gen $groupIcon "TextBox" $xLeft 176 ($innerWidth - 90) 20 "Anchor=Top,Left,Right" "Visible=$false"

# visible in File mode (hidden by default, Target Default is selected)
$labelDropHint = gen $groupIcon "Label" "You can also drag && drop any file onto this window" $xLeft 210 $innerWidth 20 "TextAlign=TopCenter" "Anchor=Top,Left,Right" "Visible=$false"
$panelIconMethod = gen $groupIcon "Panel" $xLeft 250 $innerWidth 170 "Anchor=Top,Left,Right" "Visible=$false"
$labelIconMethod = gen $panelIconMethod "Label" "Icon reference :" 0 0 $innerWidth 18 "Font=Segoe UI, 9, Bold"
$radioEmbedIcon = gen $panelIconMethod "RadioButton" "Embed icon (ADS injection)" 0 30 $innerWidth 22 "Checked=$true"
$labelEmbedInfo = gen $panelIconMethod "Label" "Icon stored inside the .lnk via NTFS Alternate Data Streams.`nIcon lost if shortcut is copied to a non-NTFS drive.`nE.g. USB FAT32, exFAT, cloud..." 0 54 $innerWidth 44 "Font=Segoe UI, 8.25"
$radioStandardIcon = gen $panelIconMethod "RadioButton" "Standard path" 0 110 $innerWidth 22
$labelStandardInfo = gen $panelIconMethod "Label" "Shortcut points to an existing .ico / .exe / .dll on disk.`nIcon breaks if the referenced file is moved or deleted." 0 134 $innerWidth 44 "Font=Segoe UI, 8.25"

# Base64 textbox
$base64Top = 176
$base64Height = $dr.Y + $dr.Height - $base64Top - $script:GroupPadding
$TextboxIcon_Base64 = gen $groupIcon "TextBox" $xLeft $base64Top $innerWidth $base64Height "Multiline=$true" "ScrollBars=Both" "WordWrap=$false" "Font=Consolas, 8" "Anchor=Top,Left,Bottom,Right" "Visible=$false"
$TextboxIcon_Base64.Add_GotFocus({ $this.SelectAll() })
$TextboxIcon_Base64.Add_Click({ $this.SelectAll() })

$form.Controls.Add($panelLeft)

$panelRight = gen $null "Panel" 405 ($titleBarHeight + 10) ($form.ClientSize.Width - 415) ($form.ClientSize.Height - $titleBarHeight - 70) "Anchor=Top,Left,Bottom,Right"

# GroupBox panel right
$groupShortcut = gen $panelRight "GroupBox" "Shortcut Configuration" "Dock=Fill"
$groupShortcut.Add_Paint($script:GroupBoxPaintHandler)
$drR = $groupShortcut.DisplayRectangle
$xLeftR = $drR.X + $script:GroupPadding
$innerWidthR = $drR.Width - ($script:GroupPadding * 2)

# Target path
$labelTarget = gen $groupShortcut "Label" "Target (executable, file, folder, or command) :" $xLeftR 23 250 18  "AutoSize=$true"
$labelTargetLen = gen $groupShortcut "Label" "0 / $($script:MaxTargetPath)" ($xLeftR + 252) 25 ($innerWidthR - 290) 18 "TextAlign=TopRight" "Font=Segoe UI, 7.5" "Anchor=Top,Left,Right" "AutoSize=$true"
$btnBrowseTarget = gen $groupShortcut "Button" "Browse..." ($xLeftR + $innerWidthR - 80) 45 80 20 "Anchor=Top,Right"
$textTarget = gen $groupShortcut "TextBox" $xLeftR 45 ($innerWidthR - 90) 20 "Anchor=Top,Left,Right"

# Arguments
$labelArgs = gen $groupShortcut "Label" "Arguments (optional) :" $xLeftR 78 150 18 "AutoSize=$true"
$labelArgsHelp = gen $groupShortcut "Label" "Limits : 8,191 (cmd.exe) | 32,767 (CreateProcess)" ($xLeftR + 150) 80 ($innerWidthR - 150) 18 "Font=Segoe UI, 7.5" "TextAlign=TopRight" "Anchor=Top,Left,Right"
# Explicit font on multiline TextBox
$textArgs = gen $groupShortcut "TextBox" $xLeftR 100 $innerWidthR 68 "Multiline=$true" "ScrollBars=Vertical" "WordWrap=$true" "Font=Segoe UI, 9" "Anchor=Top,Left,Right" "MaxLength=0"
$labelArgsLen = gen $groupShortcut "Label" "Max length : 0 / $($script:MaxArgsCreateProcess)" $xLeftR 171 $innerWidthR 17 "TextAlign=TopRight" "Anchor=Top,Left,Right"
$labelExplorerWarn = gen $labelArgsLen "Label" "" "Dock=Fill" "Font=Segoe UI, 7.5" "TextAlign=MiddleLeft" "BackColor=Transparent"

# Working directory
$labelWorkDir = gen $groupShortcut "Label" "Working directory (optional) :" $xLeftR 196 $innerWidthR 18 "AutoSize=$true"
$btnBrowseWorkDir = gen $groupShortcut "Button" "Browse..." ($xLeftR + $innerWidthR - 80) 217 80 20 "Anchor=Top,Right"
$textWorkDir = gen $groupShortcut "TextBox" $xLeftR 217 ($innerWidthR - 90) 20 "Anchor=Top,Left,Right"

# Description
$labelDescription = gen $groupShortcut "Label" "Comment / Description (optional) :" $xLeftR 250 $innerWidthR 18 "AutoSize=$true"
$textDescription = gen $groupShortcut "TextBox" $xLeftR 272 $innerWidthR 20 "Anchor=Top,Left,Right"

# AUMID
$labelAumid = gen $groupShortcut "Label" "AUMID (optional) :" $xLeftR 305 150 18 "AutoSize=$true"
$labelAumidHelp = gen $groupShortcut "Label" "Give a name to FORCE UNIQUE TASKBAR GROUP" ($xLeftR + 100) 307 ($innerWidthR - 100) 18 "Font=Segoe UI, 7.5" "TextAlign=TopRight" "Anchor=Top,Left,Right"
$textAumid = gen $groupShortcut "TextBox" $xLeftR 327 $innerWidthR 20 "Anchor=Top,Left,Right" "MaxLength=128"

# Shortcut location (.lnk)
$labelLnkPath = gen $groupShortcut "Label" "Shortcut location (.lnk) :" $xLeftR 360 $innerWidthR 18 "Anchor=Top,Left,Right"
$btnBrowseLnk = gen $groupShortcut "Button" "Browse..." ($xLeftR + $innerWidthR - 80) 380 80 20 "Anchor=Top,Right"
$textLnkPath = gen $groupShortcut "TextBox" $xLeftR 380 ($innerWidthR - 90) 20 "Anchor=Top,Left,Right"
$labelLnkInfo = gen $groupShortcut "Label" "" $xLeftR 406 $innerWidthR 18 "Font=Segoe UI, 7.5" "Anchor=Top,Left,Right"

# Create / Pin / Test buttons
$btnPinWidth    = 140
$btnGap         = 6
$fullBtnWidth   = $form.ClientSize.Width - 20
$btnCreateWidth = $fullBtnWidth - $btnPinWidth - $btnGap
$btnRowTop      = $form.ClientSize.Height - 50
$btnCreate = gen $form "Button" "Create Shortcut" 10 $btnRowTop $btnCreateWidth 40 "Font=Segoe UI Semibold, 11" "Enabled=$false" "Anchor=Bottom,Left,Right"
$btnPin    = gen $form "Button" "Pin to Taskbar" (10 + $btnCreateWidth + $btnGap) $btnRowTop $btnPinWidth 40 "Font=Segoe UI Semibold, 11" "Enabled=$false" "Anchor=Bottom,Right"
$btnPin.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat

$form.Controls.Add($panelRight)

#region ── SET-THEME FUNCTION ─

Update-LoadingPopup 80 "Loading..."

function Set-Theme {
    param([bool]$IsDark)
    $script:IsDarkMode = $IsDark
    if ($IsDark) {
        $theme = @{
            Back        = [System.Drawing.Color]::FromArgb(45,45,48)
            Fore        = [System.Drawing.Color]::White
            ForeDim     = [System.Drawing.Color]::FromArgb(160,160,160)
            ControlBack = [System.Drawing.Color]::FromArgb(55,55,55)
            Border      = [System.Drawing.Color]::FromArgb(80,80,80)
            GroupBoxBorder = [System.Drawing.Color]::FromArgb(50,90,120)
            BtnBack     = [System.Drawing.Color]::FromArgb(60,60,60)
            TitleBar    = [System.Drawing.Color]::FromArgb(32,32,32)
            TitleBtnHov = [System.Drawing.Color]::FromArgb(55,55,55)
            AboutNormal = [System.Drawing.Color]::FromArgb(50,50,50)
            AboutHover  = [System.Drawing.Color]::FromArgb(65,65,65)
            PreviewBack = [System.Drawing.Color]::FromArgb(25,25,25)
            Accent      = [System.Drawing.Color]::FromArgb(0,120,212)
            AccentFore  = [System.Drawing.Color]::White
        }
    } else {
        $theme = @{
            Back        = [System.Drawing.Color]::FromArgb(243,243,243)
            Fore        = [System.Drawing.Color]::FromArgb(30,30,30)
            ForeDim     = [System.Drawing.Color]::FromArgb(100,100,100)
            ControlBack = [System.Drawing.Color]::White
            Border      = [System.Drawing.Color]::FromArgb(200,200,200)
            GroupBoxBorder = [System.Drawing.SystemColors]::ControlDark
            BtnBack     = [System.Drawing.Color]::FromArgb(240,240,240)
            TitleBar    = [System.Drawing.Color]::FromArgb(240,240,240)
            TitleBtnHov = [System.Drawing.Color]::FromArgb(218,218,218)
            AboutNormal = [System.Drawing.Color]::FromArgb(228,228,228)
            AboutHover  = [System.Drawing.Color]::FromArgb(210,210,210)
            PreviewBack = [System.Drawing.Color]::FromArgb(230,230,230)
            Accent      = [System.Drawing.Color]::FromArgb(0,120,212)
            AccentFore  = [System.Drawing.Color]::White
        }
    }
    $script:AboutNormalColor = $theme.AboutNormal
    $script:AboutHoverColor  = $theme.AboutHover
    # GroupBox border pen (disposed and recreated on theme switch)
    if ($null -ne $script:GroupBoxBorderPen) { $script:GroupBoxBorderPen.Dispose() }
    $script:GroupBoxBorderPen = New-Object System.Drawing.Pen($theme.GroupBoxBorder)
    $form.SuspendLayout()
    # Form
    $form.BackColor = $theme.Back
    $form.ForeColor = $theme.Fore
    # Title bar
    foreach ($c in @($titleBar, $titleLabel, $iconBox, $chkConsole, $chkDarkMode)) {
        $c.BackColor = $theme.TitleBar; $c.ForeColor = $theme.Fore
    }
    foreach ($btn in @($btnMinimize, $btnClose, $btnReset)) {
        $btn.NormalBack = $theme.TitleBar; $btn.NormalFore = $theme.Fore
        $btn.HoverBack  = $theme.TitleBtnHov
        $btn.BackColor  = $theme.TitleBar; $btn.ForeColor  = $theme.Fore
    }
    $btnAbout.BackColor = $theme.AboutNormal
    $btnAbout.Invalidate()
    # Panels and group boxes
    foreach ($p in @($panelLeft, $panelRight, $panelIconMethod)) { $p.BackColor = $theme.Back }
    foreach ($g in @($groupIcon, $groupShortcut)) { $g.BackColor = $theme.Back; $g.ForeColor = $theme.Fore }
    # Picture Box Preview button
    $btnIconPreview.BackColor = $theme.PreviewBack
    $btnIconPreview.FlatAppearance.BorderColor = $theme.Border
    Update-IconPreview
    # Labels
    foreach ($lbl in @($labelPreviewIconInfo, $labelTarget, $labelArgs, $labelArgsLen,
                       $labelLnkPath, $labelArgsHelp, $labelTargetLen, $labelDropHint,
                       $labelLnkInfo, $labelExplorerWarn, $labelWorkDir,
                       $labelEmbedInfo, $labelStandardInfo, $labelTargetDefaultInfo,
                       $labelDescription, $labelAumid, $labelAumidHelp)) {
        $lbl.ForeColor = $theme.ForeDim
    }
    # Refresh shell target mode colors
    if ($script:IsShellTarget -or $script:IsUrlTarget) {
        $dimColor = if ($IsDark) { [System.Drawing.Color]::FromArgb(80,80,80) } else { [System.Drawing.Color]::FromArgb(180,180,180) }
        $labelArgs.ForeColor        = $dimColor
        $labelWorkDir.ForeColor     = $dimColor
        $labelDescription.ForeColor = $dimColor
        $labelAumid.ForeColor       = $dimColor
        $labelAumidHelp.ForeColor   = $dimColor
    }
    # Bold section headers
    foreach ($lbl in @($labelIconSource, $labelIconMethod)) {
        $lbl.ForeColor = $theme.Fore
    }
    # TextBoxes
    foreach ($tb in @($TextboxIcon_Base64, $iconPathTextbox, $textTarget, $textArgs, $textLnkPath, $textWorkDir, $textDescription, $textAumid)) {
        $tb.BackColor   = $theme.ControlBack
        $tb.ForeColor   = $theme.Fore
        $tb.BorderStyle = [System.Windows.Forms.BorderStyle]::FixedSingle
    }
    # Radio buttons
    foreach ($rb in @($radioIcon_TargetDefault, $radioIcon_Base64, $radioIcon_AnyFile, $radioEmbedIcon, $radioStandardIcon)) { $rb.ForeColor = $theme.Fore }
    # Standard buttons
    foreach ($btn in @($btnBrowseIcon, $btnBrowseTarget, $btnBrowseLnk, $btnBrowseWorkDir, $btnPin)) {
        $btn.BackColor = $theme.BtnBack; $btn.ForeColor = $theme.Fore
        $btn.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
        $btn.FlatAppearance.BorderColor = $theme.Border
        $btn.FlatAppearance.BorderSize = 1
    }
    # Create button (style only, colors handled by Update-CreateButtonState)
    $btnCreate.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
    $btnCreate.FlatAppearance.BorderColor = $theme.Accent
    $btnCreate.FlatAppearance.BorderSize = 1
    # Refresh argument length color + button state
    Update-ArgsLength
    Update-CreateButtonState
    # Dark scrollbars
    if ($form.IsHandleCreated) {
        Set-DarkScrollbars -Root $form -Dark $IsDark
        [DarkMode]::ApplyWindowFrame($form.Handle, $IsDark)
    }
    $form.ResumeLayout($true)
    $form.Invalidate($true)
    Write-Log "Theme applied : $(if ($IsDark) {'Dark'} else {'Light'})" -Level Debug
}

$chkDarkMode.Add_CheckedChanged({ Set-Theme $chkDarkMode.Checked })

#region ── EVENT HANDLERS ─

Update-LoadingPopup 85 "Loading..."

# Drop zone highlight paint (yellow borders around active drop zone)
$groupShortcut.Add_Paint({
    param($s, $e)
    if ($null -eq $script:ActiveDropZone -or $script:ActiveDropZone -eq 'icon') { return }
    $pad = 4
    $rect = $null
    switch ($script:ActiveDropZone) {
        'target' {
            $top = $labelTarget.Top - $pad
            $bot = $labelArgs.Top - $pad
            $rect = New-Object System.Drawing.Rectangle(($xLeftR - $pad), $top, ($innerWidthR + $pad * 2), ($bot - $top))
        }
        'workdir' {
            $top = $labelWorkDir.Top - $pad
            $bot = $labelDescription.Top - $pad
            $rect = New-Object System.Drawing.Rectangle(($xLeftR - $pad), $top, ($innerWidthR + $pad * 2), ($bot - $top))
        }
        'lnk' {
            $top = $labelLnkPath.Top - $pad
            $bot = $labelLnkInfo.Top + $labelLnkInfo.Height + $pad
            $rect = New-Object System.Drawing.Rectangle(($xLeftR - $pad), $top, ($innerWidthR + $pad * 2), ($bot - $top))
        }
    }
    if ($null -ne $rect) { $e.Graphics.DrawRectangle($script:DropZonePen, $rect) }
})

# Drop zone highlight paint for icon group
$groupIcon.Add_Paint({
    param($s, $e)
    if ($script:ActiveDropZone -ne 'icon') { return }
    $e.Graphics.DrawRectangle($script:DropZonePen, 0, 0, ($s.Width - 1), ($s.Height - 1))
})

# Radio button toggle
# Shared visibility updater for icon source radio buttons
function Update-IconSourceVisibility {
    $isTargetDefault = $radioIcon_TargetDefault.Checked
    $isBase64 = $radioIcon_Base64.Checked
    $isAnyFile = $radioIcon_AnyFile.Checked
    $labelTargetDefaultInfo.Visible = $isTargetDefault
    $TextboxIcon_Base64.Visible = $isBase64
    $iconPathTextbox.Visible = $isAnyFile
    $btnBrowseIcon.Visible = $isAnyFile
    $labelDropHint.Visible = $isAnyFile
    $panelIconMethod.Visible = $isAnyFile
    if ($isBase64 -and $TextboxIcon_Base64.IsHandleCreated) {
        [DarkMode]::ApplyControl($TextboxIcon_Base64.Handle, $script:IsDarkMode)
    }
    $modeName = if ($isTargetDefault) {'Target (Default)'} elseif ($isBase64) {'Base64'} else {'File'}
    Write-Log "Icon source mode : $modeName" -Level Debug
    Update-IconPreview
    Update-NtfsWarning
}

$script:IconRadioChangedHandler = {
    if (-not $this.Checked) { return }
    $ownCursor = -not $form.UseWaitCursor
    if ($ownCursor) { [System.Windows.Forms.Cursor]::Current = [System.Windows.Forms.Cursors]::WaitCursor }
    # Auto-fill icon path from target when switching to Any File with empty icon path
    if ($this -eq $radioIcon_AnyFile -and (Test-StringEmpty (Get-CleanInput $iconPathTextbox.Text))) {
        if (-not $script:IsShellTarget -and -not $script:IsUrlTarget) {
            $targetPath = Get-CleanInput $textTarget.Text
            if (-not (Test-StringEmpty $targetPath) -and [IO.File]::Exists($targetPath)) {
                $iconPathTextbox.Text = $targetPath
            }
        }
    }
    Update-IconSourceVisibility
    if ($ownCursor) { [System.Windows.Forms.Cursor]::Current = [System.Windows.Forms.Cursors]::Default }
}
$radioIcon_TargetDefault.Add_CheckedChanged($script:IconRadioChangedHandler)
$radioIcon_AnyFile.Add_CheckedChanged($script:IconRadioChangedHandler)
$radioIcon_Base64.Add_CheckedChanged($script:IconRadioChangedHandler)

$TextboxIcon_Base64.Add_TextChanged({
    if ($script:SuppressPreviewUpdate) { return }
    $form.UseWaitCursor = $true
    $form.Cursor = [System.Windows.Forms.Cursors]::WaitCursor
    [System.Windows.Forms.Application]::DoEvents()
    Update-IconPreview
    Update-NtfsWarning
    $form.UseWaitCursor = $false
    $form.Cursor = [System.Windows.Forms.Cursors]::Default
})

# File path text changed
$iconPathTextbox.Add_TextChanged({
    # Only reset index to 0 when the user manually changes the path
    if (-not $script:SuppressAutoFill) {
        $script:CurrentIconIndex = 0
    }
    $filePath = Get-CleanInput $iconPathTextbox.Text
    if (-not (Test-StringEmpty $filePath)) {
        $ext = [System.IO.Path]::GetExtension($filePath).ToLower()
        $canUseStandard = ($script:StandardIconExt -contains $ext)
        $radioStandardIcon.Enabled = $canUseStandard
        if ($canUseStandard) {
            $radioStandardIcon.Checked = $true
        }
        elseif ($radioStandardIcon.Checked) {
            $radioEmbedIcon.Checked = $true
        }
    }
    else {
        $radioStandardIcon.Enabled = $true
    }
    Update-IconPreview
    Update-NtfsWarning
})

$radioEmbedIcon.Add_CheckedChanged({
    Write-Log "Icon method : $(if ($this.Checked) {'Embed (ADS)'} else {'Standard'})" -Level Debug
    Update-NtfsWarning
    Request-CreateButtonUpdate
})

# Browse for icon source file
$btnBrowseIcon.Add_Click({
    $ofd = New-Object System.Windows.Forms.OpenFileDialog
    $ofd.Title = "Select any file to extract icon"
    $ofd.Filter = "All supported|*.png;*.jpg;*.jpeg;*.bmp;*.gif;*.ico;*.tiff;*.exe;*.dll;*.ocx;*.cpl;*.scr|Images|*.png;*.jpg;*.jpeg;*.bmp;*.gif;*.ico;*.tiff|Executables|*.exe;*.dll;*.ocx;*.cpl;*.scr|All files|*.*"
    if ($ofd.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        $iconPathTextbox.Text = $ofd.FileName
        Write-Log "Icon source file selected : $($ofd.FileName)"
    }
    $ofd.Dispose()
})

# Browse for target
$btnBrowseTarget.Add_Click({
    $ofd = New-Object System.Windows.Forms.OpenFileDialog
    $ofd.Title = "Select target file"
    $ofd.Filter = "All files|*.*|Executables|*.exe;*.bat;*.cmd;*.ps1;*.vbs"
    $ofd.CheckFileExists = $false
    $ofd.CheckPathExists = $false
    if ($ofd.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        $textTarget.Text = $ofd.FileName
        Write-Log "Target file selected : $($ofd.FileName)"
        # Auto-fill icon
        $selectedExt = [IO.Path]::GetExtension($ofd.FileName).ToLower()
        if ($selectedExt -eq '.cpl') {
            # .cpl : Target Default would show a generic icon, resolve actual icon into Any File mode
            if ($radioIcon_TargetDefault.Checked -or (Test-IconSourceEmpty)) {
                $script:SuppressPreviewUpdate = $true
                try { Resolve-FileIcon $ofd.FileName }
                finally {
                    $script:SuppressPreviewUpdate = $false
                    Update-IconPreview
                }
            }
        }
        elseif (Test-IconSourceEmpty) {
            $radioIcon_TargetDefault.Checked = $true
        }
        # Auto-fill working directory if empty
        if (Test-StringEmpty (Get-CleanInput $textWorkDir.Text)) {
            if ([IO.File]::Exists($ofd.FileName)) {
                $textWorkDir.Text = [IO.Path]::GetDirectoryName($ofd.FileName)
            }
            elseif ([IO.Directory]::Exists($ofd.FileName)) {
                $textWorkDir.Text = $ofd.FileName
            }
        }
    }
    $ofd.Dispose()
})

# Target text changed : length tracking + paste detection with auto-split
$textTarget.Add_TextChanged({
    # Invalidate shell icon cache when target changes (unless set by shell drop which pre-caches)
    if (-not $script:SuppressTargetSplit -and $null -ne $script:ShellTargetIconCache) {
        try { $script:ShellTargetIconCache.Dispose() } catch {}
        $script:ShellTargetIconCache = $null
    }
    $currentText = $textTarget.Text
    $len = (Get-CleanInput $currentText).Length
    $labelTargetLen.Text = "$len / $($script:MaxTargetPath)"
    $labelTargetLen.ForeColor = if ($len -gt $script:MaxTargetPath) { [System.Drawing.Color]::Red } else {
        if ($script:IsDarkMode) { [System.Drawing.Color]::FromArgb(140,140,140) } else { [System.Drawing.Color]::Gray }
    }
    # Shell target detection (CLSID paths)
    $isShell = $script:ShellTargetRegex.IsMatch($currentText.TrimStart())
    Set-SpecialTargetMode 'Shell' $isShell
    if ($isShell) {
        $labelTargetLen.Text = "Shell CLSID"
        $labelTargetLen.ForeColor = if ($script:IsDarkMode) { [System.Drawing.Color]::FromArgb(100,180,255) } else { [System.Drawing.Color]::FromArgb(0,100,180) }
    }
    # URL target detection
    if (-not $isShell) {
        $isUrl = $script:UrlTargetRegex.IsMatch($currentText.TrimStart())
        Set-SpecialTargetMode 'Url' $isUrl
        if ($isUrl) {
            $labelTargetLen.Text = "URL"
            $labelTargetLen.ForeColor = if ($script:IsDarkMode) { [System.Drawing.Color]::FromArgb(100,200,100) } else { [System.Drawing.Color]::FromArgb(0,130,60) }
        }
    }
    else {
        Set-SpecialTargetMode 'Url' $false
    }
    # Paste detection : text grew by more than 1 char and contains a space
    if (-not $script:SuppressTargetSplit) {
        $prevLen = $script:PreviousTargetText.Length
        $delta   = $currentText.Length - $prevLen
        if ($delta -gt 1 -and $currentText.Contains(' ')) {
            Invoke-TargetSplit $currentText
        }
    }
    # Disable AUMID for non-executable targets (shell/URL already handled by Set-SpecialTargetMode)
    if ($script:IsShellTarget) {
        $textAumid.Enabled = $true
        $normalColor = if ($script:IsDarkMode) { [System.Drawing.Color]::FromArgb(160,160,160) } else { [System.Drawing.Color]::FromArgb(100,100,100) }
        $labelAumid.ForeColor     = $normalColor
        $labelAumidHelp.ForeColor = $normalColor
    }
    elseif (-not $script:IsUrlTarget) {
        $cleanTarget = Get-CleanInput $currentText
        $targetExt = ''
        if (-not (Test-StringEmpty $cleanTarget)) {
            try { $targetExt = [IO.Path]::GetExtension($cleanTarget).ToLower() } catch {}
        }
        $isExecutable = @('.exe','.bat','.cmd','.ps1','.vbs','.com','.msi','.wsf','.scr') -contains $targetExt
        $textAumid.Enabled = $isExecutable
        $aumidDimColor = if (-not $isExecutable) {
            if ($script:IsDarkMode) { [System.Drawing.Color]::FromArgb(80,80,80) } else { [System.Drawing.Color]::FromArgb(180,180,180) }
        }
        else {
            if ($script:IsDarkMode) { [System.Drawing.Color]::FromArgb(160,160,160) } else { [System.Drawing.Color]::FromArgb(100,100,100) }
        }
        $labelAumid.ForeColor     = $aumidDimColor
        $labelAumidHelp.ForeColor = $aumidDimColor
    }
    # Refresh icon preview when Target Default is selected (icon depends on target)
    if ($radioIcon_TargetDefault.Checked) {
        Update-IconPreview
    }
    $script:PreviousTargetText = $textTarget.Text
    Update-ArgsLength
    Request-CreateButtonUpdate
})
# Target lost focus : auto-split if user typed a target+args combo manually
$textTarget.Add_Leave({
    $currentText = $textTarget.Text
    if (-not $script:SuppressTargetSplit -and $currentText.Contains(' ')) {
        Invoke-TargetSplit $currentText
    }
})

# Arguments length tracking + manual edit detection
$textArgs.Add_TextChanged({
    if (-not $script:SuppressTargetSplit) {
        $script:ArgsFromAutoSplit = $false
    }
    Update-ArgsLength
})

# Browse for working directory
$btnBrowseWorkDir.Add_Click({
    $fbd = New-Object System.Windows.Forms.FolderBrowserDialog
    $fbd.Description = "Select working directory"
    $fbd.ShowNewFolderButton = $true
    if ($fbd.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        $textWorkDir.Text = $fbd.SelectedPath
        Write-Log "Working directory selected : $($fbd.SelectedPath)"
    }
    $fbd.Dispose()
})

$pasteCleanHandler = {
    if ($_.Control -and $_.KeyCode -eq [System.Windows.Forms.Keys]::V) {
        $clip = [System.Windows.Forms.Clipboard]::GetText()
        if ($clip.Contains("`r") -or $clip.Contains("`n")) {
            $_.Handled = $true
            $_.SuppressKeyPress = $true
            $clean = $clip -replace '[\r\n]+', ' '
            $selStart = $this.SelectionStart
            $before = $this.Text.Substring(0, $selStart)
            $after  = $this.Text.Substring($selStart + $this.SelectionLength)
            $this.Text = ($before + $clean + $after)
            $this.SelectionStart = [Math]::Min($before.Length + $clean.Length, $this.Text.Length)
        }
    }
}
$textTarget.Add_KeyDown($pasteCleanHandler)
$textWorkDir.Add_KeyDown($pasteCleanHandler)
$textLnkPath.Add_KeyDown($pasteCleanHandler)
$textDescription.Add_KeyDown($pasteCleanHandler)
$textAumid.Add_KeyDown($pasteCleanHandler)

# AUMID input filtering : block invalid characters during typing
$textAumid.Add_KeyPress({
    $c = $_.KeyChar
    # Allow control characters (backspace, delete, etc.)
    if ([char]::IsControl($c)) { return }
    # Allow only letters, digits, dots, hyphens
    if ($c -match '[a-zA-Z0-9.\-]') { return }
    $_.Handled = $true
})

# AUMID sanitization on text change (handles paste with invalid chars)
$textAumid.Add_TextChanged({
    $current = $this.Text
    $clean = $script:AumidCleanRegex.Replace($current, '')
    $clean = $script:AumidLeadDotRegex.Replace($clean, '')
    if ($clean -ne $current) {
        $pos = $this.SelectionStart
        $diff = $current.Length - $clean.Length
        $this.Text = $clean
        $this.SelectionStart = [Math]::Max(0, [Math]::Min($pos - $diff, $clean.Length))
    }
    Request-CreateButtonUpdate
})

# AUMID trim trailing dots on leave
$textAumid.Add_Leave({
    $txt = $this.Text.TrimEnd('.')
    if ($txt -ne $this.Text) { $this.Text = $txt }
})

# Description has no impact on create button but trigger state refresh for consistency
$textDescription.Add_TextChanged({ Request-CreateButtonUpdate })

# Browse for .lnk save location
$btnBrowseLnk.Add_Click({
    $sfd = New-Object System.Windows.Forms.SaveFileDialog
    $sfd.Title = "Save shortcut as"
    $sfd.Filter = "LNK Shortcut|*.lnk|URL Shortcut|*.url|All shortcuts|*.lnk;*.url"
    $sfd.FilterIndex = if ($script:IsUrlTarget) { 2 } else { 1 }
    $sfd.DefaultExt = if ($script:IsUrlTarget) { "url" } else { "lnk" }
    $sfd.OverwritePrompt = $false
    if ($sfd.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        $textLnkPath.Text = $sfd.FileName
        Write-Log "Shortcut save location selected : $($sfd.FileName)"
    }
    $sfd.Dispose()
})

# .lnk path changed : detect existing files and auto-fill
$textLnkPath.Add_TextChanged({
    $lnkPath = Get-CleanInput $textLnkPath.Text
    if ((-not (Test-StringEmpty $lnkPath)) -and [System.IO.File]::Exists($lnkPath)) {
        $lnkExt = [IO.Path]::GetExtension($lnkPath).ToLower()
        if ($lnkExt -eq '.lnk' -or $lnkExt -eq '.url') {
            Import-ExistingShortcut $lnkPath
        }
    }
    else {
        $labelLnkInfo.Text = ""
    }
    Request-CreateButtonUpdate
    Update-NtfsWarning
})

#region ── CREATE SHORTCUT LOGIC ─

$btnCreate.Add_Click({
    try {
        $lnkPath = Get-CleanInput $textLnkPath.Text
        $useEmbed = $radioIcon_Base64.Checked -or ($radioIcon_AnyFile.Checked -and $radioEmbedIcon.Checked)
        if ($script:IsUrlTarget) {
            if (-not $lnkPath.EndsWith('.url', [System.StringComparison]::OrdinalIgnoreCase)) { $lnkPath += '.url' }
        }
        else {
            if (-not $lnkPath.EndsWith('.lnk', [System.StringComparison]::OrdinalIgnoreCase)) { $lnkPath += '.lnk' }
        }
        # NTFS check for embed mode
        if ($useEmbed -and -not (Test-NtfsVolume $lnkPath)) {
            Write-Log "Shortcut creation blocked : destination not NTFS and embed mode selected" -Level Error
            [System.Windows.Forms.MessageBox]::Show(
                "The destination is not on an NTFS volume.`nADS icon embedding requires NTFS.`n`nSwitch to 'Standard icon reference' or choose an NTFS destination.",
                "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
            return
        }
        $lnkDir = [IO.Path]::GetDirectoryName($lnkPath)
        if (-not [IO.Directory]::Exists($lnkDir)) {
            [IO.Directory]::CreateDirectory($lnkDir) | Out-Null
            Write-Log "Created directory for shortcut : $lnkDir"
        }
        $existingLnk = [IO.File]::Exists($lnkPath)
        if ($existingLnk) { [IO.File]::Delete($lnkPath) }
        # Detect .cpl targets and auto-convert to shell shortcut
        $createTargetPath = Get-CleanInput $textTarget.Text
        $createTargetExt = ''
        try { $createTargetExt = [IO.Path]::GetExtension($createTargetPath).ToLower() } catch {}
        if ($createTargetExt -eq '.cpl' -and -not $script:IsShellTarget -and -not $script:IsUrlTarget) {
            $resolvedCplPath = $createTargetPath
            if (-not [IO.File]::Exists($resolvedCplPath)) {
                $system32Candidate = [IO.Path]::Combine($env:SystemRoot, 'System32', [IO.Path]::GetFileName($createTargetPath))
                if ([IO.File]::Exists($system32Candidate)) { $resolvedCplPath = $system32Candidate }
            }
            $cplItem = Resolve-CplControlPanelItem $resolvedCplPath
            if ($null -ne $cplItem) {
                $userDesc = $textDescription.Text.Trim()
                if (Test-StringEmpty $userDesc) { $userDesc = $cplItem.Name }
                $userAumid = Get-CleanInput $textAumid.Text
                if (Test-StringEmpty $userAumid) { $userAumid = $cplItem.Path }
                $useEmbed = $radioIcon_Base64.Checked -or ($radioIcon_AnyFile.Checked -and $radioEmbedIcon.Checked)
                $hasIcon = -not $radioIcon_TargetDefault.Checked
                if ($hasIcon -and $useEmbed) {
                    $icoBytes = Get-IcoBytesFromCurrentSource
                    if ($null -ne $icoBytes -and $icoBytes.Length -gt 0) {
                        [ShortcutHelper]::CreateShellWithEmbeddedIcon($lnkPath, $cplItem.Path, ([byte[]]$icoBytes), $userAumid, $userDesc)
                        Write-AdsIcon -FilePath $lnkPath -IcoBytes ([byte[]]$icoBytes)
                    }
                    else {
                        [ShortcutHelper]::CreateShellWithStandardIcon($lnkPath, $cplItem.Path, "", 0, $userAumid, $userDesc)
                    }
                }
                elseif ($hasIcon) {
                    $iconFilePath = Get-CleanInput $iconPathTextbox.Text
                    $iconIndex = $script:CurrentIconIndex
                    [ShortcutHelper]::CreateShellWithStandardIcon($lnkPath, $cplItem.Path, $iconFilePath, $iconIndex, $userAumid, $userDesc)
                }
                else {
                    [ShortcutHelper]::CreateShellWithStandardIcon($lnkPath, $cplItem.Path, "", 0, $userAumid, $userDesc)
                }
                $action = if ($existingLnk) { "updated" } else { "created" }
                Write-Log "CPL shortcut $action via PIDL : $lnkPath -> $($cplItem.Path)"
                Invoke-CreateButtonFlash ([char]0x2714 + " " + ($action.Substring(0,1).ToUpper() + $action.Substring(1)))
                return
            }
            Write-Log "CPL namespace resolution failed for '$resolvedCplPath', falling back to standard shortcut" -Level Warning
        }
        # Arguments length warning
        $argsLen = $textArgs.Text.Length
        if ($argsLen -gt $script:MaxArgsCmdExe -and $argsLen -le $script:MaxArgsCreateProcess) {
            $result = [System.Windows.Forms.MessageBox]::Show(
                "Arguments exceed 8,191 characters (cmd.exe limit).`nThe shortcut may not work if launched via cmd.exe.`n`nContinue?",
                "Warning", [System.Windows.Forms.MessageBoxButtons]::YesNo, [System.Windows.Forms.MessageBoxIcon]::Warning)
            if ($result -ne [System.Windows.Forms.DialogResult]::Yes) {
                Write-Log "User cancelled shortcut creation (argument length warning)" -Level Debug
                return
            }
        }
        # Create the shortcut
        $result = New-ShortcutFromFields $lnkPath
        if ($null -eq $result) {
            [System.Windows.Forms.MessageBox]::Show("Failed to build icon from current source.", "Error",
                [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
            return
        }
        # Build result message
        $action    = if ($existingLnk) { "updated" } else { "created" }
        $targetPath = Get-CleanInput $textTarget.Text
        $typeLabel = if ($script:IsUrlTarget) { "URL shortcut" } elseif ($script:IsShellTarget) { "Shell shortcut" } else { "Shortcut" }
        $targetLabel = if ($script:IsUrlTarget) { "URL" } elseif ($script:IsShellTarget) { "Shell path" } else { "Target" }
        $iconMsg = switch ($result.IconMode) {
            'embed' {
                $adsLength = Get-AdsIconLength -FilePath $lnkPath
                if ($adsLength -gt 0) { "Icon embedded as ADS : $adsLength bytes" }
                else {
                    Write-Log "Shortcut saved but ADS embedding may have failed : $lnkPath" -Level Warning
                    [System.Windows.Forms.MessageBox]::Show(
                        "Shortcut saved but ADS icon embedding may have failed.`nCheck : $lnkPath",
                        "Warning", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
                    return
                }
            }
            'standard' { "Icon : $($result.IconFilePath) (index $($result.IconIndex))" }
            default    { "Icon : default (Windows)" }
        }
        Write-Log "$typeLabel $action : $lnkPath ($($result.IconMode))"
        Invoke-CreateButtonFlash ([char]0x2714 + " " + ($action.Substring(0,1).ToUpper() + $action.Substring(1)))
    }
    catch {
        Write-Log "Error creating shortcut : $($_.Exception.Message)" -Level Error
        [System.Windows.Forms.MessageBox]::Show(
            "Error creating shortcut :`n$($_.Exception.Message)",
            "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
    }
})

$btnPin.Add_Click({
    $form.UseWaitCursor = $true
    $form.Cursor = [System.Windows.Forms.Cursors]::WaitCursor
    [System.Windows.Forms.Application]::DoEvents()
    $tempLnkCreated = $false
    $lnkToPin = $null
    try {
        $lnkPath = Get-CleanInput $textLnkPath.Text
        # Case 1 : existing .lnk on disk
        if ((-not (Test-StringEmpty $lnkPath)) -and $lnkPath.EndsWith('.lnk', [System.StringComparison]::OrdinalIgnoreCase) -and [IO.File]::Exists($lnkPath)) {
            $lnkToPin = $lnkPath
            Write-Log "Pin to Taskbar : using existing .lnk : $lnkPath"
        }
        else {
            # Case 2 : build temp .lnk from current fields
            $targetPath = Get-CleanInput $textTarget.Text
            if (Test-StringEmpty $targetPath) { return }
            $displayName = if (-not (Test-StringEmpty $lnkPath)) { [IO.Path]::GetFileNameWithoutExtension($lnkPath) }
                           else { [IO.Path]::GetFileNameWithoutExtension($targetPath) }
            $lnkToPin = [IO.Path]::Combine($env:TEMP, "$displayName.lnk")
            if ($script:IsUrlTarget) {
                # URL targets : create a .lnk that opens the URL via explorer.exe
                $pinAumid = Get-CleanInput $textAumid.Text
                $pinDesc  = $textDescription.Text.Trim()
                if (Test-StringEmpty $pinDesc) { $pinDesc = $displayName }
                $pinIcoBytes = Get-IcoBytesFromCurrentSource
                if ($null -ne $pinIcoBytes -and $pinIcoBytes.Length -gt 0) {
                    [ShortcutHelper]::CreateWithEmbeddedIcon($lnkToPin, "explorer.exe", $targetPath, ([byte[]]$pinIcoBytes), $pinAumid, $pinDesc)
                    Write-AdsIcon -FilePath $lnkToPin -IcoBytes ([byte[]]$pinIcoBytes)
                }
                else {
                    [ShortcutHelper]::CreateWithStandardIcon($lnkToPin, "explorer.exe", $targetPath, "", 0, $pinAumid, $pinDesc)
                }
                Write-Log "Pin to Taskbar : built URL .lnk via explorer.exe : $lnkToPin"
            }
            else {
                # Detect .cpl targets and create a PIDL shortcut via Control Panel namespace
                $pinTargetExt = ''
                try { $pinTargetExt = [IO.Path]::GetExtension($targetPath).ToLower() } catch {}
                if ($pinTargetExt -eq '.cpl') {
                    # Resolve relative .cpl paths (e.g. "main.cpl" -> "C:\Windows\System32\main.cpl")
                    $resolvedCplPath = $targetPath
                    if (-not [IO.File]::Exists($resolvedCplPath)) {
                        $system32Candidate = [IO.Path]::Combine($env:SystemRoot, 'System32', [IO.Path]::GetFileName($targetPath))
                        if ([IO.File]::Exists($system32Candidate)) { $resolvedCplPath = $system32Candidate }
                    }
                    $cplItem = Resolve-CplControlPanelItem $resolvedCplPath
                    if ($null -ne $cplItem) {
                        $safeCplName = $cplItem.Name -replace '[<>:"/\\|?*]', '_'
                        $lnkToPin = [IO.Path]::Combine($env:TEMP, "$safeCplName.lnk")
                        $pinAumid = Get-CleanInput $textAumid.Text
                        if (Test-StringEmpty $pinAumid) { $pinAumid = $cplItem.Path }
                        $pinDesc  = $textDescription.Text.Trim()
                        if (Test-StringEmpty $pinDesc) { $pinDesc = $cplItem.Name }
                        if (-not $radioIcon_TargetDefault.Checked) {
                            $pinIcoBytes = Get-IcoBytesFromCurrentSource
                            if ($null -ne $pinIcoBytes -and $pinIcoBytes.Length -gt 0) {
                                [ShortcutHelper]::CreateShellWithEmbeddedIcon($lnkToPin, $cplItem.Path, ([byte[]]$pinIcoBytes), $pinAumid, $pinDesc)
                                Write-AdsIcon -FilePath $lnkToPin -IcoBytes ([byte[]]$pinIcoBytes)
                            }
                            else {
                                [ShortcutHelper]::CreateShellWithStandardIcon($lnkToPin, $cplItem.Path, "", 0, $pinAumid, $pinDesc)
                            }
                        }
                        else {
                            [ShortcutHelper]::CreateShellWithStandardIcon($lnkToPin, $cplItem.Path, "", 0, $pinAumid, $pinDesc)
                        }
                        Write-Log "Pin to Taskbar : .cpl resolved to PIDL shortcut : $lnkToPin -> $($cplItem.Path)"
                    }
                    else {
                        Write-Log "Pin to Taskbar : .cpl namespace resolution failed, standard fallback" -Level Warning
                        $result = New-ShortcutFromFields $lnkToPin
                        if ($null -eq $result) {
                            [System.Windows.Forms.MessageBox]::Show("Failed to build shortcut for pinning.", "Pin to Taskbar",
                                [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
                            return
                        }
                    }
                }
                else {
                    $result = New-ShortcutFromFields $lnkToPin
                    if ($null -eq $result) {
                        [System.Windows.Forms.MessageBox]::Show("Failed to build shortcut for pinning.", "Pin to Taskbar",
                            [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
                        return
                    }
                    Write-Log "Pin to Taskbar : built temp .lnk : $lnkToPin"
                }
            }
            $tempLnkCreated = $true
        }
        $success = Invoke-TaskbarPin $lnkToPin
        if ($success) {
            Write-Log "Taskbar pin successful : $lnkToPin"
            Update-PinButtonState
        }
        else {
            [System.Windows.Forms.MessageBox]::Show(
                "Failed to pin shortcut to taskbar.`nCheck the log for details.",
                "Pin to Taskbar", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
        }
    }
    catch {
        Write-Log "Taskbar pin error : $($_.Exception.Message)" -Level Error
        [System.Windows.Forms.MessageBox]::Show(
            "Error pinning to taskbar :`n$($_.Exception.Message)",
            "Pin to Taskbar", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
    }
    finally {
        if ($tempLnkCreated -and $lnkToPin -and [IO.File]::Exists($lnkToPin)) {
            try { [IO.File]::Delete($lnkToPin) } catch {}
        }
        $form.UseWaitCursor = $false
        $form.Cursor = [System.Windows.Forms.Cursors]::Default
    }
})

#region ── GLOBAL EVENTS ─

Update-LoadingPopup 90 "Loading..."

# ── Drop zone cache ──
$script:CachedDropZones = $null
function Build-DropZoneCache {
    return @(
        @{ Name = 'icon';    Rect = $panelLeft.RectangleToScreen($panelLeft.ClientRectangle) }
        @{ Name = 'target';  Top = $labelTarget.PointToScreen([System.Drawing.Point]::Empty).Y;      Bottom = $labelArgs.PointToScreen([System.Drawing.Point]::Empty).Y }
        @{ Name = 'workdir'; Top = $labelWorkDir.PointToScreen([System.Drawing.Point]::Empty).Y;     Bottom = $labelDescription.PointToScreen([System.Drawing.Point]::Empty).Y }
        @{ Name = 'lnk';     Top = $labelLnkPath.PointToScreen([System.Drawing.Point]::Empty).Y;     Bottom = $btnCreate.PointToScreen([System.Drawing.Point]::Empty).Y }
    )
}

# ── .NET OLE drag-drop (yellow borders during drag, requires STA thread) ──
$oleDragDropEnabled = $false
$isSta = ([Threading.Thread]::CurrentThread.GetApartmentState() -eq [Threading.ApartmentState]::STA)
if ($isSta -and -not $isAdmin) {
    try {
        $form.AllowDrop = $true
        $oleDragDropEnabled = $true
    } catch {}
}
if ($oleDragDropEnabled) {
    $form.Add_DragEnter({
        if ($_.Data.GetDataPresent([System.Windows.Forms.DataFormats]::FileDrop) -or
            $_.Data.GetDataPresent("Shell IDList Array")) {
            $_.Effect = $_.AllowedEffect
            $script:CachedDropZones = Build-DropZoneCache
            $zone = Get-DropZone ([System.Windows.Forms.Cursor]::Position)
            Set-ActiveDropZone $zone
        }
        else { $_.Effect = [System.Windows.Forms.DragDropEffects]::None }
    })
    $form.Add_DragOver({
        if ($_.Data.GetDataPresent([System.Windows.Forms.DataFormats]::FileDrop) -or
            $_.Data.GetDataPresent("Shell IDList Array")) {
            $_.Effect = $_.AllowedEffect
            $zone = Get-DropZone ([System.Windows.Forms.Cursor]::Position)
            Set-ActiveDropZone $zone
        }
    })
    $form.Add_DragLeave({
        Set-ActiveDropZone $null
        $script:CachedDropZones = $null
    })
    $form.Add_DragDrop({
        Set-ActiveDropZone $null
        $script:CachedDropZones = $null
        if ($_.Data.GetDataPresent([System.Windows.Forms.DataFormats]::FileDrop)) {
            $files = $_.Data.GetData([System.Windows.Forms.DataFormats]::FileDrop)
            if ($files.Count -gt 0) {
                $zone = Get-DropZone ([System.Windows.Forms.Cursor]::Position)
                Invoke-ZoneDrop $files[0] $zone
            }
        }
        elseif ($_.Data.GetDataPresent("Shell IDList Array")) {
            $zone = Get-DropZone ([System.Windows.Forms.Cursor]::Position)
            Invoke-ShellItemDrop $_.Data $zone
        }
    })
    Write-Log "OLE drag-drop registered (STA + non-admin)" -Level Debug
}
else {
    Write-Log "OLE drag-drop disabled ($(if (-not $isSta) {'non-STA'} elseif ($isAdmin) {'admin UIPI'} else {'unknown'})), using WM_DROPFILES fallback" -Level Warning
}

# WM_DROPFILES + WM_SETTINGCHANGE (admin fallback via DragAcceptFiles, works on all PS versions)
$form.add_OnWindowMessage({
    $wndMsg = $args[1]
    switch ($wndMsg.Msg) {
        0x0233 {
            Set-ActiveDropZone $null
            try {
                $hDrop = $wndMsg.WParam
                if ($null -ne $hDrop -and $hDrop -ne [IntPtr]::Zero) {
                    $files = [DragDropFix]::GetDroppedFiles($hDrop)
                    if ($files.Count -gt 0) {
                        $zone = Get-DropZone ([System.Windows.Forms.Cursor]::Position)
                        Invoke-ZoneDrop $files[0] $zone
                    }
                }
            }
            catch { Write-Log "WM_DROPFILES error : $($_.Exception.Message)" -Level Error }
            return
        }
        0x001A {
            try {
                $regObj = Get-ItemProperty -Path "HKCU:\Software\Microsoft\Windows\CurrentVersion\Themes\Personalize" -Name "AppsUseLightTheme" -ErrorAction SilentlyContinue
                $isDarkNow = ($null -ne $regObj -and $regObj.AppsUseLightTheme -eq 0)
                if ($isDarkNow -ne $script:IsDarkMode) {
                    $chkDarkMode.Checked = $isDarkNow
                    Write-Log "System theme change detected : $(if ($isDarkNow) {'Dark'} else {'Light'})"
                }
            } catch {}
            return
        }
    }
})

$form.Add_Load({
    Update-LoadingPopup 95 "Finalizing..."
    # Detect and apply system theme before the form becomes visible
    $isDarkSystem = $false
    try {
        $regObj = Get-ItemProperty -Path "HKCU:\Software\Microsoft\Windows\CurrentVersion\Themes\Personalize" -Name "AppsUseLightTheme" -ErrorAction SilentlyContinue
        if ($null -ne $regObj -and $regObj.AppsUseLightTheme -eq 0) { $isDarkSystem = $true }
    } catch { }
    Write-Log "System theme detected : $(if ($isDarkSystem) {'Dark'} else {'Light'})"
    $chkDarkMode.Checked = $isDarkSystem
})

$form.Add_Shown({
    Update-LoadingPopup 100 "Finalizing..."
    Close-LoadingPopup
    $form.Activate()
    [DwmHelper]::SetRoundedCorners($form)
    [DragDropFix]::Enable($form.Handle)
    # Apply HTTRANSPARENT to all registered pass-through controls
    foreach ($passThruCtrl in $script:HitTestPassThruControls) {
        try {
            $grip = New-Object HitTestPassThrough
            $grip.AssignHandle($passThruCtrl.Handle)
            $script:HitTestNativeWindows.Add($grip)
        } catch { }
    }
    Set-DarkScrollbars -Root $form -Dark $script:IsDarkMode
    [DarkMode]::ApplyWindowFrame($form.Handle, $script:IsDarkMode)
    # AUMID placeholder (cue banner)
    [NativeMethods]::SendMessageW($textAumid.Handle, 0x1501, 0, "Letters, digits, dots, hyphens - 128 chars max. e.g. 'MyApp.MyCompany.1'") | Out-Null
    Update-ArgsLength
    Update-CreateButtonState
    $form.BringToFront()
    [System.Windows.Forms.Application]::DoEvents()
    Write-Log "Main form initialization complete"
})

$form.Add_Click({ $form.ActiveControl = $null })

#region ── CLEANUP ─

$script:CleanupDone = $false
function Invoke-ApplicationCleanup {
    if ($script:CleanupDone) { return }
    $script:CleanupDone = $true
    Write-Log "Performing application cleanup"
    # Release HitTestPassThrough native windows
    foreach ($nw in $script:HitTestNativeWindows) {
        try { $nw.ReleaseHandle() } catch {}
    }
    $script:HitTestNativeWindows.Clear()
    # Stop debounce timer
    try { $script:CreateBtnDebounceTimer.Stop(); $script:CreateBtnDebounceTimer.Dispose() } catch {}
    try { Reset-CreateButtonFlash } catch {}
    # Dispose GDI resources
    foreach ($pen in @($script:FormBorderPenLight, $script:FormBorderPenDark, $script:DropZonePen, $script:GroupBoxBorderPen)) {
        if ($null -ne $pen) { try { $pen.Dispose() } catch {} }
    }
    if ($null -ne $script:CurrentPreviewBitmap) { try { $script:CurrentPreviewBitmap.Dispose() } catch {} }
        if ($null -ne $script:ShellTargetIconCache) { try { $script:ShellTargetIconCache.Dispose() } catch {} }
    # Dispose icon index menu and its images
    foreach ($item in $script:IconIndexMenu.Items) {
        if ($item.Image) { try { $item.Image.Dispose() } catch {} }
    }
    try { $script:IconIndexMenu.Dispose() } catch {}
    # Dispose icon resources
    if ($null -ne $iconImage)      { try { $iconImage.Dispose() }      catch {} }
    if ($null -ne $taskIcon)       { try { $taskIcon.Dispose() }       catch {} }
    if ($null -ne $taskIconStream) { try { $taskIconStream.Dispose() } catch {} }
    # Start Menu shortcut lifecycle
    $taskbarLnk = [IO.Path]::Combine($script:TaskbarPinDir, $script:LnkName)
    $taskbarPinExists = [IO.File]::Exists($taskbarLnk)
    if ($taskbarPinExists) { Write-Log "Taskbar pin exists, preserving Start Menu shortcut as icon source" }
    $startMenuLnk = [IO.Path]::Combine($script:StartMenuDir, $script:LnkName)
    if ([IO.File]::Exists($startMenuLnk)) {
        $isUserPinned = $false
        try { $isUserPinned = [ShortcutHelper]::GetDescription($startMenuLnk).Contains('[UserPinned]') } catch {}
        if (-not $isUserPinned -and -not $taskbarPinExists) {
            try { [IO.File]::Delete($startMenuLnk); Write-Log "Cleaned up temporary Start Menu shortcut" } catch {}
        } elseif ($isUserPinned) { Write-Log "Start Menu shortcut kept (user pinned)" }
    }
}

[System.Windows.Forms.Application]::add_ThreadException({
    param($s, $e)
    Write-Log "Unhandled thread exception : $($e.Exception.Message)" -Level Error
})

[AppDomain]::CurrentDomain.add_UnhandledException({
    param($s, $e)
    Write-Log "Unhandled domain exception : $($e.ExceptionObject.Message)" -Level Error
    Invoke-ApplicationCleanup
})

$form.Add_FormClosed({
    Write-Log "Form closed by user"
    Invoke-ApplicationCleanup
})

#region ── RUN ─

Write-Log "ResumeLayout completed, autoscaling applied"
$form.ResumeLayout($true)
[System.Windows.Forms.Application]::Run($form)
try { $form.Dispose() } catch {}
Write-Log "═══ $($script:AppName) ended ═══"
