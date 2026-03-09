Window Layout

This file explains how the deployed folder works and how to edit the configuration.

Why this is useful
- Windows may renumber monitors after screensaver, sleep or wake. This tool identifies monitors by position rather than by monitor number.
- Similar monitor layouts at home and in the office can still match even when positions and sizes are not exactly the same.
- Different monitor resolutions still work because window placement uses percentages rather than fixed pixel coordinates.
- Configurations can be copied from computer to computer.

Files
- WindowLayout.cmd must stay next to the WindowLayout folder.
- WindowLayout\WindowLayout.ps1 is the main script.
- WindowLayout\window_layouts.txt is the editable layouts file. It is created automatically on first run.
- WindowLayout\current_layout.txt is generated only when you capture the current layout.

Typical setups
- 3 monitors - developer: IDE large, CAD smaller.
- 3 monitors - engineer: CAD large, IDE smaller.
- ultrawide - developer: IDE on the left.
- ultrawide - engineer: CAD on the left.
- laptop - laptop: mostly full-screen windows on a single monitor.

How to use it
1. Double-click WindowLayout.cmd.
2. Choose a layout from the menu in the form "monitor setup - layout".
3. Or choose to capture the current layout.
4. During capture, choose which monitor setup to map the current monitors to.
5. After capture, open current_layout.txt, copy the monitor setup or layout you want into window_layouts.txt, and remove the rows you do not need.
6. Capture usually includes more windows than you want, including helper windows and some windows that are not obvious to the user.

Direct shortcut script
- You can create a .cmd file containing:
  WindowLayout.cmd -ApplyLayout "3 monitors - developer"
- Double-clicking that file will apply that layout directly.

How the config works
- The file is pipe-delimited and is reformatted by the script only when needed.
- A layout is identified by the pair: monitor setup + layout name.
- monitorRole values are local to the selected monitor setup.
- x, y, w and h are percentages of the monitor work area.
- match can be contains, exact or regex.
- cascade can be yes or no.

Sections
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
processName | title    | match    | x  | y  | w  | h   | monitorRole | cascade
chrome.exe  | Grafana  | contains | 0  | 0  | 30 | 100 | lower-left  | no
chrome.exe  | Main     | contains | 20 | 20 | 60 | 70  | primary     | no
ms-teams.exe|          | contains | 45 | 0  | 55 | 100 | upper-left  | no

Excel example
- To keep the Excel VBA editor fixed on the left while cascading ordinary workbook windows, use rows like these:
  excel.exe | Visual Basic | contains | 2  | 0  | 60 | 100 | primary | no
  excel.exe | - Excel$     | regex    | 60 | 10 | 35 | 40  | primary | yes
- In regex mode, - Excel$ means the title ends with - Excel; $ marks the end of the title.

With title vs without title
- These two rows show the difference:
  chrome.exe  | Grafana | contains | 0  | 0 | 30 | 100 | lower-left | no
  ms-teams.exe|         | contains | 45 | 0 | 55 | 100 | upper-left | no
- With a title, only matching windows from that process are targeted.
- With an empty title, any window from that process can be targeted.

Important parser note
- The title column may contain the | character only for regex patterns.
- The parser reads the first column and the last 7 columns by position, then treats everything in between as the title column.

Monitor setups
- A monitor setup name must be unique.
- Within a monitor setup, each monitor role must be unique.
- A layout must reference an existing monitor setup.
- A layout row must reference an existing monitorRole from that monitor setup.
- The script matches the currently connected monitors to the selected monitor setup by relative monitor positions, not by Windows monitor numbering.

Capture behavior
- If the config has no monitor setups yet, capture asks for a name and creates the first monitor setup from the current monitors.
- If monitor setups already exist, capture asks which setup to use, or lets you create a new one.
- The monitor match is tolerant: it does not require a perfect coordinate or size match.
- For multi-instance applications, capture emits both:
  - one cascade row with an empty title
  - the individual per-window rows
- Keep whichever version you prefer and delete the others.

Menu behavior
- Press Enter on a numbered choice prompt to exit or cancel instead of getting an invalid-choice message.
- If the config file is reformatted, the script prints a message telling you it was reformatted and saved.

Restart from scratch
- Delete WindowLayout\window_layouts.txt and WindowLayout\current_layout.txt.
- On the next run, the script recreates a fresh empty window_layouts.txt.

Command line examples
- WindowLayout.cmd
- WindowLayout.cmd -ListLayouts
- WindowLayout.cmd -CaptureCurrent
- WindowLayout.cmd -ApplyLayout "3 monitors - developer"
