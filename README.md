# SyMenu Orphan Icons Cleaner

A utility for [SyMenu](https://www.ugmfree.it) portable menu users that identifies and removes orphaned icon files — icons that exist in the `Icons` folder but are no longer referenced by any menu item in the SyMenu configuration.

Over time, as you add and remove portable apps from your SyMenu setup, unused `.ico` files can accumulate in the `Icons` folder. This tool cleans them up, keeping your SyMenu installation tidy.

## NOTE:
There are three versions:
- v1 - AutoHotkey was the original source code by VVV_Easy_SyMenu.
- v2 - AutoHotkey was updated and enhanced by me.
- v3 - VB.NET also by me, is the active version and was created to avoid the incessant false positives by windows defender.

## Features

- **Orphan Detection** — Scans all `.ico` files in `SyMenu\Icons` against the `SyMenuItem.xml` configuration to find unreferenced icons
- **Move or Delete** — Choose to move orphans to a `_Trash\_OrphanIcons` folder for review, or permanently delete them
- **Exclusion List** — Protect specific icons from being flagged as orphans via an editable exclusion list (`Exclusions.txt`)
- **Inline Editing** — Edit and save the exclusion list directly within the GUI
- **Detailed Logging** — Activity log displayed in the GUI and saved to a `.log` file, with optional verbose mode
- **Resizable GUI** — Side-by-side layout with a draggable splitter between the exclusion list and the log window
- **Persistent Settings** — Window size, splitter position, SyMenu path, and all options are saved between sessions via `Settings.ini`

## Requirements

- [AutoHotkey v2.0](https://www.autohotkey.com/) or later
- A working [SyMenu](https://www.ugmfree.it) installation

## Installation

1. Download or clone this repository
2. Place the script files in a folder of your choice (e.g. inside your SyMenu `ProgramFiles` directory)
3. Ensure AutoHotkey v2.0+ is installed, or use a compiled `.exe` version

## Usage

### First Run

1. Run `SyMenu_Orphan_Icons v2.ahk` (or the compiled `.exe`)
2. Click **Browse...** to select your SyMenu root folder (e.g. `D:\SyMenu`)
3. The path is saved automatically for future runs

### Scanning for Orphans

1. Click **Go!** to start the scan
2. The script will:
   - Extract `SyMenuItem.xml` from `SyMenu\Config\SyMenuItem.zip`
   - Scan every `.ico` file in `SyMenu\Icons`
   - Check each icon filename against the configuration
   - Move (or delete) any icons not found in the configuration
3. Results are shown in the log window and a summary message box

### Options

| Option | Description |
|--------|-------------|
| **Delete?** | If checked, orphan icons are permanently deleted. If unchecked (default), they are moved to `SyMenu\ProgramFiles\SPSSuite\SyMenuSuite\_Trash\_OrphanIcons` for manual review |
| **Delete present Log File and start new?** | If checked (default), creates a fresh log file each run. If unchecked, appends to the existing log |
| **Full information log (verbose mode)?** | If checked, logs every icon scanned (active, excluded, and orphan). If unchecked (default), only logs orphaned icons |

### Managing Exclusions

Some icons may not be referenced in `SyMenuItem.xml` but you still want to keep them. Use the exclusion list to protect them:

1. Click **Edit** to unlock the exclusion list
2. Add one icon filename per line (e.g. `MyCustomIcon.ico`)
3. Lines starting with `;` are treated as comments and ignored
4. Click **Save** to apply your changes

The exclusion list is stored in `Exclusions.txt` alongside the script.

### GUI Layout

The window is split into two panels:

- **Left panel** — Exclusion list editor
- **Right panel** — Log output window

Drag the grey splitter bar between the panels to resize them. The window can also be freely resized. Your layout preferences are saved automatically.

## File Structure

```
SyMenu_Orphan_Icons v2.ahk   # Main script
Settings.ini                   # Auto-generated settings (path, options, window layout)
Exclusions.txt                 # Icon exclusion list (one filename per line)
SyMenu_Orphan_Icons v2.log    # Activity log
```

## How It Works

SyMenu stores all menu item configuration in a single file: `Config\SyMenuItem.zip`, which contains `SyMenuItem.xml`. This XML file references icon filenames for every item across all suites (SyMenuSuite, NirSoftSuite, custom suites, etc.).

The script extracts this XML, then performs a simple text search for each `.ico` filename found in the `Icons` folder. If an icon's filename doesn't appear anywhere in the XML, it's considered orphaned.

## Version History

| Version | Date | Changes |
|---------|------|---------|
| v2.7 | 2026-04-01 | Config read from SyMenuItem.xml, exclusion list support |
| v2.6 | 2026-04-01 | Fixed splitter drag via WM messages, anchored buttons |
| v2.5 | 2026-04-01 | Resizable window, draggable splitter, saved layout |
| v2.4 | 2026-04-01 | Side-by-side layout (exclusions left, log right) |
| v2.3 | 2026-04-01 | Settings persistence via Settings.ini |
| v2.2 | 2026-04-01 | Inline edit/save for exclusions, path fixes |
| v2.1 | 2026-04-01 | Added exclusion list support via Exclusions.txt |
| v2.0 | 2026-04-01 | Converted to AHK v2 |
| v1.0 | 2017-01-27 | First version (AHK v1) |

## Credits

- Original v1 script by the VVV_Easy_SyMenu
- v2 conversion and enhancements with assistance from GitHub Copilot by sl23

## License

This project is provided as-is for use with SyMenu. See [SyMenu's website](https://www.ugmfree.it) for more information about SyMenu itself.
