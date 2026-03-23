# Shortcuterie

A Windows shortcut (.lnk) creation tool that can embeds custom icons directly inside shortcut files using **NTFS Alternate Data Streams**.
> **How does it work?** See the [Documentation](Documentation.md) of the ADS injection mechanism, ICO construction, PE resource extraction...

![Windows 7](https://img.shields.io/badge/Windows-7%2B-blue?logo=windows)
![PowerShell 2.0](https://img.shields.io/badge/PowerShell-2.0%2B-blue?logo=powershell)

<img width="840" height="540" alt="image" src="https://github.com/user-attachments/assets/c5863998-a9de-480e-8960-602036c1be6f" />

## Features

- **ADS Icon Embedding** - Icon stored __inside__ the .lnk
- **Standard Icon Reference** - Traditional mode pointing to an .ico/.exe/.dll on disk
- **Special Targets Support** - Create shortcuts to UWP apps, .cpl, .msc, URL, with their correct (or custom) icon
- **Drag & Drop Zones** - Drag & drop files onto specific zones with visual highlighting
- **Taskbar Pinning** - Pin any shortcut directly to the taskbar
- **Auto Target/Args Split** - Paste a full command line and it splits target from arguments automatically
- **Single .bat File** - No installation, no dependencies.

## Requirements

- Windows 7+
- PowerShell 2.0+
- NTFS filesystem (for injected icons)

## Usage

1. Download `Shortcuterie.bat`
2. Double-click to run
3. Set your icon source (drag a file, paste Base64, or keep the target's default icon)
4. Fill in the target path and optional fields
5. Choose where to save the .lnk (*shortcut location*)
6. Click **Create Shortcut**

### Drag & Drop

The entire window is a drop target with zone-based routing:

| Drop Zone | Behavior |
|-----------|----------|
| **Left panel** (Icon Source) | Choose an icon from any file |
| **Target field** | Sets the target path. Dropping a .lnk extracts its target, args, working dir, and description |
| **Working Directory field** | Sets the working directory from the dropped file/folder location |
| **Shortcut Location field** | Sets the save path. Dropping an existing .lnk imports all its fields for editing |

### Icon Embedding Modes

| Mode | How it works | Trade-off |
|------|-------------|-----------|
| **Embed (ADS)** | Icon binary stored inside the .lnk as an NTFS Alternate Data Stream | Lost if copied to FAT32/exFAT/cloud storage |
| **Standard Path** | Shortcut points to an .ico/.exe/.dll on disk | Breaks if the referenced file is moved or deleted |
| **Target Default** | No custom icon - Windows assigns based on the target type | Same behavior as a normal Explorer-created shortcut |

### Taskbar Pinning

Click **Pin to Taskbar** to add the shortcut into the taskbar. This works for standard executables and UWP apps.

## Limitations

- ADS icon embedding requires NTFS. Icons are silently broken when copying to FAT32, exFAT, or most cloud-synced folders.
- Normal copy/cut-paste of a shortcut will preserve injected icon. But if a tool handles the operation, the icon can break depending on how this tool works.

## License

Partially open source - enterprise or commercial usage of the Taskbar Pinning Module requires a paid license. See [LICENSE](LICENSE) for details.
