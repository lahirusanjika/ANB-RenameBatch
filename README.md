# ANB Rename Batch

Batch rename, move, and log files safely with a PowerShell CLI and a Windows GUI.

Developed by Lahiru Sanjika.

## Features
- Safe two‑pass renaming with rollback on errors.
- Rename in place or move to a destination folder.
- Preserve original subfolder structure when moving.
- Stream mode for huge sets (constant memory usage).
- Sort order: Time, Name, Size, or None.
- Descending order option.
- Sequence controls: prefix, start index, auto or fixed padding, minimum padding.
- Per‑extension numbering (separate sequence per extension).
- Filters: include/exclude patterns, include/exclude extensions, include hidden files, recurse.
- Conflict handling: Error, Skip, or Overwrite.
- Dry‑run preview with optional full listing.
- CSV logging of renames with optional hashing.
- Restore (undo) from a log, with optional hash verification.
- JSON presets: load and save settings.
- Interactive mode for terminal prompts.

## Requirements
- Windows
- PowerShell 5.1+ (or PowerShell 7+)
- .NET 8 SDK (only if you want to build the GUI)

## Quick Start (CLI)
Rename all files in a folder (and subfolders) by time order:

```powershell
pwsh -File tools\ANB-RenameBatch.ps1 -Path "C:\MyFolder" -Recurse -Order Time
```

Move and preserve folders:

```powershell
pwsh -File tools\ANB-RenameBatch.ps1 -Path "C:\MyFolder" -Destination "D:\Sorted" -PreserveFolders -Recurse
```

## Quick Start (GUI)
Build the GUI EXE:

```powershell
dotnet publish tools\ANB-RenameBatch-GUI\ANB.RenameBatch.GUI.csproj -c Release -o tools\ANB-RenameBatch-GUI\dist
```

Run:

```
tools\ANB-RenameBatch-GUI\dist\ANB.RenameBatch.GUI.exe
```

The GUI expects `ANB-RenameBatch.ps1` to be next to the EXE (the publish command copies it there).

## CLI Options

**Core**
- `-Path` Root folder (default: current directory).
- `-Destination` Optional destination folder.
- `-PreserveFolders` Keep subfolder structure when using `-Destination`.
- `-Prefix` Prefix before the number (example: `photo_`).
- `-Start` Starting index (default 1).
- `-Pad` Fixed zero‑padding width. Use `0` for auto.
- `-MinPad` Minimum padding when `-Pad 0` (default 4).
- `-Order` Sort order: `Time`, `Name`, `Size`, `None`.
- `-Descending` Reverse sort order.
- `-Recurse` Include subfolders.
- `-PerExtension` Restart numbering per extension.
- `-Stream` Constant‑memory mode (requires `-Order None` and no `-PerExtension`).

**Filters**
- `-Include` Include patterns (wildcards). Default `*`.
- `-Exclude` Exclude patterns (wildcards).
- `-Extensions` Only include these extensions (dot optional), example: `jpg,mp4`.
- `-ExcludeExtensions` Exclude extensions (dot optional).
- `-IncludeHidden` Include hidden/system files.

**Conflicts / Preview**
- `-OnConflict` `Error`, `Skip`, or `Overwrite`.
- `-DryRun` Preview only, no changes.
- `-ShowAll` Show full preview list (default shows first 50).
- `-ProgressEvery` Progress interval (0 disables). Default 1000.

**Logging / Restore**
- `-LogPath` CSV log path.
- `-NoLog` Disable automatic logging.
- `-HashAlgorithm` `SHA256` (default), `SHA1`, `SHA384`, `SHA512`, `MD5`.
- `-NoHash` Disable hashing.
- `-RestoreFrom` Restore (undo) from a CSV log.
- `-SkipHashMismatch` Skip files with hash mismatches during restore.

**Safety / Presets**
- `-ConfirmThreshold` Ask for confirmation over this count (default 10000).
- `-Force` Skip confirmation (GUI already prompts).
- `-Preset` Load a JSON preset.
- `-SavePreset` Save a JSON preset and exit.
- `-Interactive` Prompt for options.

## Examples
Rename in place with a prefix:

```powershell
pwsh -File tools\ANB-RenameBatch.ps1 -Path . -Prefix "img_" -Recurse
```

Only JPG and MP4, exclude TMP:

```powershell
pwsh -File tools\ANB-RenameBatch.ps1 -Path . -Extensions jpg,mp4 -ExcludeExtensions tmp -Recurse
```

Stream mode for huge sets:

```powershell
pwsh -File tools\ANB-RenameBatch.ps1 -Path . -Order None -Stream -Recurse
```

Restore (undo) from a log:

```powershell
pwsh -File tools\ANB-RenameBatch.ps1 -RestoreFrom C:\path\ANB_RenameLog_YYYYMMDD_HHMMSS.csv
```

## Notes
- Stream mode requires `-Order None` and does not support `-PerExtension`.
- When `-Pad 0`, padding is auto‑computed from the file count (minimum `-MinPad`).
- On Windows, keep the script and EXE in the same folder for GUI runs.

## Credits
Developed by Lahiru Sanjika.
\n## Build EXE (Batch)\n\nYou can also build the GUI EXE with the included batch file:\n\n`at\nbuild-exe.bat\n`\n\nThis creates:\n\n`
tools\\ANB-RenameBatch-GUI\\dist\\ANB.RenameBatch.GUI.exe\n`\n
