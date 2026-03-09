# Window Layout

Save and restore complex Windows desktop layouts from a simple folder, with no installation beyond Windows itself.

This project uses PowerShell plus built-in Windows APIs to:

- Capture the current monitor-aware window arrangement
- Restore named layouts from a text file
- Work across different monitor numbering by matching monitors by position
- Support multiple monitor setups such as `3 monitors`, `ultrawide`, or `laptop`
- Support both per-window rules and cascade rules for multi-instance apps

## Why It Is Useful

- Windows can renumber monitors after events such as screensaver/sleep/wake, but this tool identifies monitors by position rather than number
- Similar home and office monitor layouts can still match even when positions and sizes are close but not identical
- Different monitor resolutions still work because window placement uses percentages instead of fixed pixel coordinates
- Configurations can be copied from computer to computer

## What It Looks Like

The launcher shows layouts in the form:

- `3 monitors - developer`: IDE large, CAD smaller
- `3 monitors - engineer`: CAD large, IDE smaller
- `ultrawide - developer`: IDE on the left
- `ultrawide - engineer`: CAD on the left
- `laptop - laptop`: mostly full-screen windows on a single monitor

You can also apply a layout directly from an interactive terminal or from a shortcut script:

```cmd
WindowLayout.cmd -ApplyLayout "3 monitors - developer"
```

## Features

- `No dependencies`: runs with built-in Windows PowerShell and .NET only
- `Portable`: works from a folder such as the desktop
- `Monitor setups`: separate physical monitor arrangements from window layouts
- `Relative monitor matching`: matches the current monitors to a saved setup by relative position, not Windows numbering
- `Capture workflow`: captures the current layout into `current_layout.txt`, which you can trim and copy into `window_layouts.txt`
- `Readable config`: uses a pipe-delimited text format instead of JSON
- `Regex titles`: supports `regex` title matching, including `|` inside regex patterns
- `Cascade support`: capture emits both a cascade row and individual rows for multi-instance apps

## Folder Layout

Release zip contents:

```text
WindowLayout.cmd
WindowLayout/
  WindowLayout.ps1
  readme.txt
```

`window_layouts.txt` is created on first run.

`current_layout.txt` is generated only when you capture a layout.

## Quick Start

1. Download the release zip and extract it.
2. Keep `WindowLayout.cmd` next to the `WindowLayout` folder.
3. Double-click `WindowLayout.cmd`.
4. Choose a saved layout, or capture the current layout.
5. After capture, copy the parts you want from `current_layout.txt` into `window_layouts.txt`.

Capture usually sees more windows than you want to keep, including helper windows and some windows that are not obvious to the user. Trim the captured file before keeping it.

## Configuration Model

The config separates:

- `MonitorSetup`: the physical monitor arrangement
- `Layout`: the app/window arrangement for that monitor setup

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
processName | title           | match    | x  | y  | w  | h   | monitorRole | cascade
chrome.exe  | Main            | contains | 20 | 20 | 60 | 70  | primary     | no
chrome.exe  | Personal        | contains | 0  | 0  | 30 | 100 | lower-left  | no
ms-teams.exe|                 | contains | 45 | 0  | 55 | 100 | upper-left  | no
```

That shows the main idea:

- monitor setups describe the physical arrangement
- Layouts describe where apps go inside that setup
- A row with a title targets a specific window pattern
- A row with an empty title can target any window from that process

Excel example:

```text
[Layout engineer]
monitorSetup | 3 monitors
processName | title        | match    | x  | y  | w  | h   | monitorRole | cascade
excel.exe   | Visual Basic | contains | 2  | 0  | 60 | 100 | primary     | no
excel.exe   | - Excel$     | regex    | 60 | 10 | 35 | 40  | primary     | yes
```

That shows a useful pattern:

- The Excel VBA editor (`Visual Basic`) is pinned on the left
- All normal workbook windows (`- Excel$`) are handled together with `cascade = yes`
- In regex mode, `- Excel$` means the title ends with `- Excel`; `$` marks the end of the title

## Commands

```cmd
WindowLayout.cmd
WindowLayout.cmd -ListLayouts
WindowLayout.cmd -CaptureCurrent
WindowLayout.cmd -ApplyLayout "3 monitors - developer"
```

## Build A Release Zip

This repo includes a packaging script:

```powershell
.\build-release.ps1
```

It creates `dist\WindowLayout.zip` with the deployable files only.

## License

MIT. See `LICENSE`.
