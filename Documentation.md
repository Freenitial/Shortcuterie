# Shortcuterie - Documentation

This document explains the techniques used by Shortcuterie to embed icons inside Windows shortcut files, and resolve icons from arbitrary file types.

---

## Table of Contents

- [ADS Icon Embedding: The Core Technique](#ads-icon-embedding-the-core-technique)
  - [Self-Referencing Icon Path](#self-referencing-icon-path)
  - [Writing the ADS Payload](#writing-the-ads-payload)
  - [Windows Shell Icon Resolution Pipeline](#windows-shell-icon-resolution-pipeline)
  - [Why Windows Accepts This](#why-windows-accepts-this)
- [ICO File Construction](#ico-file-construction)
  - [ICO Binary Format](#ico-binary-format)
  - [PE Resource Table Extraction](#pe-resource-table-extraction)
  - [Multi-Source Build Pipeline](#multi-source-build-pipeline)
- [The .LNK Binary Format](#the-lnk-binary-format)
  - [MS-SHLLINK Structure](#ms-shllink-structure)
  - [IShellLink COM Interface](#ishelllink-com-interface)
  - [AppUserModelID (AUMID)](#appusermodelid-aumid)
  - [PIDL-Based Shell Shortcuts](#pidl-based-shell-shortcuts)
- [Icon Resolution Chain](#icon-resolution-chain)
  - [Step 1: SHGetFileInfo SHGFI_ICONLOCATION](#step-1-shgetfileinfo-shgfi_iconlocation)
  - [Step 2: AssocQueryString ASSOCSTR_DEFAULTICON](#step-2-assocquerystring-assocstr_defaulticon)
  - [Step 3: Registry ProgID Cascade](#step-3-registry-progid-cascade)
  - [Step 4: CLSID DefaultIcon Lookup](#step-4-clsid-defaulticon-lookup)
  - [Step 5: System Image List (JUMBO)](#step-5-system-image-list-jumbo)
  - [Step 6: Shell Bitmap Fallback](#step-6-shell-bitmap-fallback)
- [Taskbar Pinning via Registry Blob Injection](#taskbar-pinning-via-registry-blob-injection)
- [Edge Cases and Limitations](#edge-cases-and-limitations)

---

## ADS Icon Embedding: The Core Technique

### Self-Referencing Icon Path

The entire technique hinges on a single line:

```csharp
link.SetIconLocation(lnkPath + ":icon.ico", 0);
```

`IShellLink::SetIconLocation` stores a path string in the .lnk file's StringData section. Windows Shell does not validate whether this path refers to a standalone file - it passes the string directly to `CreateFileW` during icon resolution.

`CreateFileW` natively understands NTFS Alternate Data Stream syntax:

```
C:\Users\John\Desktop\MyApp.lnk:icon.ico
```

The colon-separated `:icon.ico` suffix tells the NTFS driver to open a named data stream attached to the file, rather than the file's default `$DATA` stream. This is resolved at the filesystem driver level.

### Writing the ADS Payload

The `AdsHelper` class writes raw bytes to an alternate stream via `CreateFileW`:

```csharp
public static void WriteStream(string filePath, string streamName, byte[] data)
{
    string adsPath = filePath + ":" + streamName;
    using (SafeFileHandle h = CreateFileW(adsPath, GENERIC_WRITE, 0,
        IntPtr.Zero, CREATE_ALWAYS, FILE_ATTRIBUTE_NORMAL, IntPtr.Zero))
    {
        using (FileStream fs = new FileStream(h, FileAccess.Write))
            fs.Write(data, 0, data.Length);
    }
}
```

Key implementation details:

- **`CREATE_ALWAYS` disposition**: Creates the stream if absent, truncates if it already exists. Idempotent - safe to call repeatedly.
- **`dwShareMode = 0`**: Exclusive lock during write prevents race conditions with Explorer's icon cache reading the stream concurrently.
- **Stream name `icon.ico`**: Arbitrary. The `.ico` extension is cosmetic and has no functional impact. The stream name simply needs to match what `SetIconLocation` wrote.
- **Invisibility**: ADS data does not appear in `dir`, Explorer file size, or most file managers. Only `dir /r`, Sysinternals Streams, or forensic tools reveal it.

### Windows Shell Icon Resolution Pipeline

When Explorer needs to render a shortcut's icon:

```
Explorer
  → SHGetFileInfo / IExtractIcon
    → IShellLink::GetIconLocation → returns "file.lnk:icon.ico", index 0
      → Icon cache lookup (iconcache_*.db)
        → Cache miss → PrivateExtractIcons("file.lnk:icon.ico", 0, ...)
          → CreateFileW("file.lnk:icon.ico", GENERIC_READ, ...)
            → NTFS driver opens the ADS
              → Returns valid ICO byte stream → parsed and rendered
```

Every layer in this chain treats the ADS path as a normal file path. No Shell code has explicit ADS awareness - it's handled entirely by the NTFS filesystem driver.

### Why Windows Accepts This

This works because of a design property of NTFS and the Win32 API: Alternate Data Streams are first-class file objects. `CreateFileW` can open them, `ReadFile` can read them, and any API that ultimately calls these functions - including the Shell's icon loading code - will work transparently with ADS paths.

The icon data written to the stream is a valid multi-resolution ICO file. The Shell's icon parser sees standard ICO headers and PNG-compressed entries, indistinguishable from reading a standalone `.ico` file.

---

## ICO File Construction

### ICO Binary Format

The `IcoBuilder` class constructs valid ICO files from any image source. The binary layout:

```
Offset    Size       Field
─────────────────────────────────────────
0         2 bytes    Reserved (0x0000)
2         2 bytes    Type (0x0001 = ICO)
4         2 bytes    Image count (N)
6         16×N       Directory entries
6+16N     variable   Image data (PNG-compressed)
```

Each 16-byte directory entry:

```
Offset    Size       Field
─────────────────────────────────────────
0         1 byte     Width (0 = 256)
1         1 byte     Height (0 = 256)
2         1 byte     Color palette count (0 for 32bpp)
3         1 byte     Reserved
4         2 bytes    Color planes (1)
6         2 bytes    Bits per pixel (32)
8         4 bytes    Image data size in bytes
12        4 bytes    Absolute offset to image data
```

Shortcuterie generates all 7 standard sizes (16, 24, 32, 48, 64, 128, 256px) as **PNG-compressed** entries inside the ICO container. This is the modern ICO format supported since Windows Vista. Older BMP-based entries would also work but are significantly larger.

### PE Resource Table Extraction

For executables (.exe, .dll, .ocx, .cpl, .scr), the icon extraction process reads the PE resource table directly:

1. **`LoadLibraryEx(LOAD_LIBRARY_AS_DATAFILE)`** - Opens the PE file without executing it or loading dependencies
2. **`EnumResourceNames(RT_GROUP_ICON)`** - Walks the resource directory to find icon groups (type 14)
3. **`FindResource` + `LoadResource` + `LockResource`** - Reads the `GRPICONDIR` structure to discover available native sizes
4. **`PrivateExtractIcons`** - Extracts each size at its exact native resolution (no upscaling artifacts)

The `GRPICONDIR` structure is similar to the ICO header but uses resource IDs instead of file offsets:

```
Offset    Size       Field
─────────────────────────────────────────
0         2 bytes    Reserved
2         2 bytes    Type (1 = ICO)
4         2 bytes    Entry count
6         14×N       GRPICONDIRENTRY records
```

Each 14-byte GRPICONDIRENTRY contains `width`, `height`, and a resource ID that maps to an `RT_ICON` resource.

When PE table reading fails (access denied, corrupted resources, negative resource IDs), `BuildFromExecutableEx` falls back to probing `PrivateExtractIcons` at standard sizes in descending order, stopping when it detects upscaled duplicates.

### Multi-Source Build Pipeline

| Source | Method | Flow |
|--------|--------|------|
| Bitmap / PNG / JPG | `BuildFromBitmap` | Scale to 7 sizes → PNG-compress each → assemble ICO |
| Base64 string | `BuildFromBase64` | Decode → detect ICO vs image → extract largest entry or load bitmap → `BuildFromBitmap` |
| Executable (PE) | `BuildFromExecutable` | Read PE resource table → extract each native size → assemble ICO |
| Executable (fallback) | `BuildFromExecutableEx` | Probe PrivateExtractIcons at 256→16 → assemble ICO |
| Shell item (UWP/CLSID) | PIDL → SHGetImageList | Extract JUMBO bitmap from system image list → `BuildFromBitmap` |

---

## The .LNK Binary Format

### MS-SHLLINK Structure

Shortcut files follow the [MS-SHLLINK](https://docs.microsoft.com/en-us/openspecs/windows_protocols/ms-shllink/) specification:

```
┌─────────────────────────┐
│   ShellLinkHeader       │  76 bytes - magic, flags, file attributes, timestamps
├─────────────────────────┤
│   LinkTargetIDList      │  Target as a PIDL (shell item ID list)
├─────────────────────────┤
│   LinkInfo              │  Volume serial, local/network path info
├─────────────────────────┤
│   StringData            │  Unicode strings: name, arguments, icon location, working dir
├─────────────────────────┤
│   ExtraData             │  Optional blocks:
│   ├ TrackerDataBlock    │    DROID for distributed link tracking
│   ├ PropertyStoreBlock  │    Serialized IPropertyStore (contains AUMID)
│   └ TerminalBlock       │    Empty block (4 zero bytes) marking end
└─────────────────────────┘
```

The icon location string (e.g., `C:\path\file.lnk:icon.ico`) is stored in the **StringData** section. `IShellLink::SetIconLocation` writes this string, and `IPersistFile::Save` serializes the entire structure to disk.

### IShellLink COM Interface

Shortcuterie creates and modifies shortcuts through the `IShellLink` COM interface (CLSID `{00021401-0000-0000-C000-000000000046}`):

```csharp
IShellLink link = (IShellLink)new ShellLink();
link.SetPath(targetPath);           // Target executable or file
link.SetArguments(arguments);        // Command-line arguments
link.SetIconLocation(path, index);   // Icon source (the ADS path goes here)
link.SetDescription(description);    // Comment field
link.SetWorkingDirectory(workDir);   // Start-in directory
((IPersistFile)link).Save(lnkPath, true);  // Serialize to disk
```

For existing shortcuts (`UpdateIconOnly`), the approach is:

```csharp
((IPersistFile)link).Load(lnkPath, 0);          // Deserialize
link.SetIconLocation(lnkPath + ":icon.ico", 0); // Overwrite icon only
((IPersistFile)link).Save(lnkPath, true);        // Re-serialize
```

This preserves target, arguments, working directory, hotkey, window state - only the icon location field is mutated. However, `IPersistFile::Save` rewrites the entire file, which destroys any existing ADS. That's why the ADS icon must be re-written *after* every `Save` call.

### AppUserModelID (AUMID)

The AUMID is a string property stored in the .lnk's **ExtraData** PropertyStore block, not in the ADS. It controls taskbar grouping:

```
ExtraData
  └ PropertyStoreDataBlock
      └ {9F4C2855-9F79-4B39-A8D0-E1D42DE1D5F3}, pid=5
          └ VT_LPWSTR: "MyApp.MyCompany.1"
```

Windows groups taskbar buttons by AUMID, not by process. Without a custom AUMID, all shortcuts launching `powershell.exe` would group together. The AUMID gives each shortcut its own identity.

Shortcuterie writes the AUMID via `IPropertyStore::SetValue`:

```csharp
IPropertyStore store = (IPropertyStore)link;
PROPVARIANT pv = PROPVARIANT.FromString(appId);
store.SetValue(ref PKEY_AppUserModel_ID, ref pv);
store.Commit();
```

The AUMID is **independent from the ADS icon mechanism**. It is optional and only relevant for shortcuts that need unique taskbar grouping.

### PIDL-Based Shell Shortcuts

For targets that don't have filesystem paths (Control Panel items, Recycle Bin, UWP apps, shell folders), the shortcut stores a PIDL (Pointer to an Item IDentifier List) instead of a path string:

```csharp
IntPtr pidl = SHParseDisplayName(shellPath, ...);
link.SetIDList(pidl);
```

`SHParseDisplayName` converts a shell path (e.g., `shell:AppsFolder\Microsoft.WindowsCalculator_8wekyb3d8bbwe!App` or `::{CLSID}`) into a binary PIDL. `SetIDList` stores this PIDL in the **LinkTargetIDList** section of the .lnk.

The PIDL is an opaque binary blob - a chain of variable-length `SHITEMID` structures terminated by a 2-byte zero. Each SHITEMID represents one level in the shell namespace hierarchy.

---

## Icon Resolution Chain

When a file is dropped onto Shortcuterie, it attempts to resolve the highest-quality icon through a 6-step cascade. Each step tries progressively broader techniques until one succeeds.

### Step 1: SHGetFileInfo SHGFI_ICONLOCATION

```csharp
SHGetFileInfo(filePath, 0, ref shfi, ..., SHGFI_ICONLOCATION);
```

Returns the icon source file path and index that the shell would use. Works well for executables and common file types with direct icon handlers. Fails for types where the shell uses indirect resolution.

### Step 2: AssocQueryString ASSOCSTR_DEFAULTICON

```csharp
AssocQueryString(ASSOCF_NONE, ASSOCSTR_DEFAULTICON, extension, null, sb, ref size);
```

Queries the file association database for the default icon reference string (e.g., `"C:\Windows\System32\imageres.dll,-67"`). Handles most registered file types.

### Step 3: Registry ProgID Cascade

Walks the registry to find all icon references for a file extension:

```
HKCR\.ext                           → direct DefaultIcon
HKCR\.ext → (default value)         → ProgID\DefaultIcon
HKCU\...\FileExts\.ext\UserChoice   → UserChoice ProgID\DefaultIcon
HKCR\.ext\OpenWithProgids           → secondary ProgID\DefaultIcon
```

Primary candidates (default ProgID, UserChoice) are tried first. Secondary candidates (OpenWithProgids, conventional `extfile`) are tried next, with standalone .ico files from third-party apps deprioritized when the shell already has a valid system icon index.

### Step 4: CLSID DefaultIcon Lookup

For shell namespace objects and .msc console files, extracts the CLSID from the path or file content and checks:

```
HKCR\CLSID\{GUID}\DefaultIcon → icon reference string
```

For .msc files specifically, the XML content is parsed to find the snap-in CLSID.

### Step 5: System Image List (JUMBO)

```csharp
SHGetImageList(SHIL_JUMBO, ref IID_IImageList, out imageList);
ImageList_GetIcon(imageList, sysIndex, 0);
```

Extracts a 256×256 (JUMBO) or 48×48 (EXTRALARGE) bitmap from the Windows system image list. This handles every file type the shell can display, including those with no direct icon file reference.

### Step 6: Shell Bitmap Fallback

```csharp
SHGetFileInfo(filePath, 0, ref shfi, ..., SHGFI_ICON | SHGFI_LARGEICON);
```

Last resort - extracts the basic 32×32 shell icon. Low resolution but universal.

---

## Taskbar Pinning via Registry Blob Injection 

Refer to this project : https://github.com/Freenitial/Pin-Taskbar

---

## Edge Cases and Limitations

| Concern | Behavior |
|---------|----------|
| **Non-NTFS drive (FAT, Cloud...)** | ADS doesn't exist. Icon is silently stripped when copying. Shortcuterie probes with a test write before creation and warns the user. |
| **Copy to ZIP/archive** | ADS is stripped by all archive formats. The .lnk becomes icon-less. |
| **IPersistFile::Save destroys ADS** | Every COM save rewrites the default stream, which erases any existing ADS. Shortcuterie always writes the ADS *after* saving the .lnk. |
| **Icon cache staleness** | Explorer caches icons aggressively. Updating an existing shortcut's ADS icon may require `ie4uinit.exe -show` or deleting `iconcache_*.db`. |
| **Explorer Properties dialog** | Shows the self-referencing `file.lnk:icon.ico` path in the icon field. Functional but looks unusual to users. |
| **Target + Args > 260 chars** | Explorer's property sheet truncates the combined string, making it uneditable via Properties. The shortcut itself works fine - `CreateProcess` supports 32,767 chars. |
| **Negative icon resource IDs** | Some executables use resource IDs (negative integers) instead of sequential indices. `BuildFromExecutableEx` handles this by probing `PrivateExtractIcons` directly. |
| **Elevated processes (admin)** | UIPI blocks OLE drag-drop from non-elevated Explorer. Shortcuterie falls back to `WM_DROPFILES` via `ChangeWindowMessageFilterEx`. |

---
