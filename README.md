# Window Layout

Save and restore complex Windows desktop layouts from a simple folder, with no installation required.

This project uses PowerShell plus built-in Windows APIs to:

- Capture the current monitor-aware window arrangement
- Restore named layouts from a text file
- Reorganize windows across multiple virtual desktops without dragging them around manually
- Work across different monitor numbering by matching monitors by position
- Support multiple monitor setups such as `3 monitors`, `ultrawide`, or `laptop`
- Support both per-window rules and cascade rules for multi-instance apps

## Why It Is Useful

- Windows can renumber monitors after events such as screensaver/sleep/wake, but this tool identifies monitors by position rather than number
- Similar home and office monitor layouts can still match even when positions and sizes are close but not identical
- Different monitor resolutions still work because window placement uses percentages instead of fixed pixel coordinates
- Virtual desktops are tedious to organize by hand when you first open a group of windows; this tool can place them on the right desktops for you
- You can also use partial layouts, for example a `MyProject.cmd` script that opens a few apps and then runs `WindowLayout.cmd -ApplyLayout "MyProject"` to arrange only those windows
- Configurations can be copied from computer to computer

## What It Looks Like

Setup examples:

- `3 monitors - developer`: IDE large, CAD smaller
- `3 monitors - engineer`: CAD large, IDE smaller
- `ultrawide - developer`: IDE on the left
- `ultrawide - engineer`: CAD on the left
- `laptop - default`: mostly full-screen windows on a single monitor

You can also apply a layout directly from an interactive terminal or from a shortcut script:

```cmd
WindowLayout.cmd -ApplyLayout "3 monitors - developer"
```

When `WindowLayout\window_layouts.txt` is reformatted, the script also regenerates one `.cmd` file per saved layout inside `WindowLayout\`, for example `WindowLayout\WindowLayout - 3 monitors - developer.cmd`.

## Features

- `No dependencies`: runs with built-in Windows PowerShell and .NET only
- `Portable`: runs directly from a folder such as the desktop
- `Monitor setups`: separate physical monitor arrangements from window layouts
- `Relative monitor matching`: matches the current monitors to a saved setup by relative position, not Windows numbering
- `Capture workflow`: saves the current layout to `current_layout.txt` with a timestamped comment header, and you can trim and copy parts into `window_layouts.txt`
- `Readable config`: uses a pipe-delimited text format instead of JSON
- `Regex titles`: supports `regex` title matching, including `|` inside regex patterns
- `Cascade support`: capture emits both a cascade row and individual rows for multi-instance apps
- `Virtual desktops`: captures and restores a manually editable desktop number for each window
- `Editable percentages`: capture writes integer `x`, `y`, `w`, and `h` values, but you can manually change them to decimals for finer positioning
- `Ignore list`: processes whose captured windows never resolve to a valid desktop number are written to `processes_to_ignore.txt` and excluded from future captures and restores

## Folder Layout

Folder contents:

```text
WindowLayout.cmd
WindowLayout/
  WindowLayout.ps1
  readme.txt
```

`window_layouts.txt` is created on first run with a comment header and a blank line before the content.

Per-layout `.cmd` shortcut scripts inside `WindowLayout\` are generated from `window_layouts.txt` when it is reformatted.

`current_layout.txt` is generated only when you capture a layout. It includes a timestamped comment header and a blank line before the captured content.

`processes_to_ignore.txt` is updated automatically when capture finds processes without a valid desktop number.

## Quick Start

1. Download or copy the folder.
2. Keep `WindowLayout.cmd` next to the `WindowLayout` folder.
3. Double-click `WindowLayout.cmd`.
4. Choose a saved layout, or capture the current layout.
5. After capture, copy the parts you want from `current_layout.txt` into `window_layouts.txt`.

When you create a new monitor setup, capture first lists the detected monitors, then asks you to name each monitor, then asks for the monitor configuration name.

Capture usually sees more windows than you want to keep, including helper windows and some windows that are not obvious at first glance. Individual windows without a valid desktop number are not written to `current_layout.txt`; if all captured windows for a process lack a valid desktop number, that process name is added to `WindowLayout\processes_to_ignore.txt`.

## Configuration Model

The config separates:

- `MonitorSetup`: the physical monitor arrangement
- `Layout`: the window arrangement for that monitor setup

Example:

```text
[MonitorSetup 3 monitors]
role       | x     | y
primary    | 0     | 0
upper-left | -1600 | 1250
lower-left | -1600 | 2160

[MonitorSetup ultrawide]
role    | x     | y
primary | 0     | 0
left    | -1600 | 1250

[MonitorSetup laptop]
role    | x | y
primary | 0 | 0

[Layout developer]
monitorSetup | 3 monitors
processName | title           | match    | x  | y  | w  | h   | monitorRole | cascade | desktop
chrome.exe  | Main            | contains | 20 | 20 | 60 | 70  | primary     | no      | 1
chrome.exe  | YouTube         | contains | 0  | 0  | 30 | 100 | lower-left  | no      | 2
notepad.exe |                 | contains | 45 | 0  | 55 | 100 | upper-left  | no      | 2
```

That shows the main idea:

- monitor setups describe the physical arrangement
- Layouts describe where apps go inside that setup
- A row with a title targets a specific window pattern
- A row with an empty title can target any window from that process
- `desktop` is a 1-based virtual desktop number and can be edited manually to place windows across multiple virtual desktops
- Captured `x`, `y`, `w`, and `h` values are written as integers, but you can manually edit them to decimal values when you want finer control

Excel example:

```text
[Layout engineer]
monitorSetup | 3 monitors
processName | title        | match    | x  | y  | w  | h   | monitorRole | cascade | desktop
excel.exe   | Visual Basic | contains | 2  | 0  | 60 | 100 | primary     | no      | 1
excel.exe   | - Excel$     | regex    | 60 | 10 | 35 | 40  | primary     | yes     | 2
```

That shows a useful pattern:

- The Excel VBA editor (`Visual Basic`) is pinned on the left
- All normal workbook windows (`- Excel$`) are handled together with `cascade = yes`
- In regex mode, `- Excel$` means the title ends with `- Excel`; `$` marks the end of the title
- Those workbook windows are moved to desktop `2` during restore
- Changing the `desktop` numbers in `window_layouts.txt` is the normal way to place windows across multiple desktops

## Capture Notes

- Capture learns ignored processes automatically and stores them in `WindowLayout\processes_to_ignore.txt`
- Individual windows without a valid desktop number are omitted from `current_layout.txt`
- If all captured windows for a process have no valid desktop number, that process is blacklisted
- `WindowLayout.cmd -CaptureCurrent -IgnoreBlacklist` captures the full visible window list for troubleshooting, including blacklisted processes and windows without a usable desktop number
- Capture still writes integer percentages, but manual decimal edits in `window_layouts.txt` are preserved when the file is reformatted

## Restore Notes

- Restore moves each window to its saved desktop number before applying size and position
- If a window is already on the target desktop, the script skips the desktop move and only applies size/position
- If a saved desktop number is missing during restore, the script leaves that window on its current desktop and continues with the rest of the layout
- Restore also uses `WindowLayout\processes_to_ignore.txt`, so previously blacklisted processes are skipped

## Commands

```cmd
WindowLayout.cmd
WindowLayout.cmd -ListLayouts
WindowLayout.cmd -CaptureCurrent
WindowLayout.cmd -CaptureCurrent -IgnoreBlacklist
WindowLayout.cmd -ApplyLayout "3 monitors - developer"
```

When you run `-CaptureCurrent` from the command line, the script automatically picks the best saved monitor setup that matches the currently detected monitors.

## License

MIT. See `LICENSE`.
