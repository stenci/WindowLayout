Window Layout

This file explains how the folder works and how to edit the configuration.

Why this is useful
- Windows may renumber monitors after screensaver, sleep or wake. This tool identifies monitors by position rather than by monitor number.
- Similar monitor layouts at home and in the office can still match even when positions and sizes are not exactly the same.
- Different monitor resolutions still work because window placement uses percentages rather than fixed pixel coordinates.
- Configurations can be copied from computer to computer.

Files
- WindowLayout.cmd must stay next to the WindowLayout folder.
- WindowLayout\WindowLayout.ps1 is the main script.
- WindowLayout\window_layouts.txt is the editable layouts file. It is created automatically on first run.
- WindowLayout\WindowLayout - <monitor setup> - <layout>.cmd files are generated from window_layouts.txt when that file is reformatted.
- WindowLayout\current_layout.txt is generated only when you capture the current layout.
- WindowLayout\processes_to_ignore.txt is maintained automatically for processes without a valid desktop number.

Setup examples:
- 3 monitors - developer: IDE large, CAD smaller.
- 3 monitors - engineer: CAD large, IDE smaller.
- ultrawide - developer: IDE on the left.
- ultrawide - engineer: CAD on the left.
- laptop - default: mostly full-screen windows on a single monitor.

How to use it
1. Double-click WindowLayout.cmd.
2. Choose a layout from the menu in the form "monitor setup - layout".
3. Or choose to capture the current layout.
4. During capture, choose which monitor setup to map the current monitors to.
5. After capture, open current_layout.txt, copy the monitor setup or layout you want into window_layouts.txt, and remove the rows you do not need.
6. Capture usually includes more windows than you want, including helper windows and some windows that are not obvious at first glance.
7. Processes without a valid desktop number are added to processes_to_ignore.txt instead of being written to current_layout.txt.

Direct shortcut script
- You can also run a layout directly from the command line:
  WindowLayout.cmd -ApplyLayout "3 monitors - developer"
- When window_layouts.txt is reformatted, the script also regenerates one shortcut .cmd file per saved layout inside the WindowLayout folder.
- Those generated files are one-liners such as:
  ..\WindowLayout.cmd -ApplyLayout "3 monitors - developer"
- Double-clicking one of those files applies that layout directly.

How the config works
- The file is pipe-delimited and is reformatted by the script only when needed.
- A layout is identified by the pair `monitor setup + layout name`.
- monitorRole values are local to the selected monitor setup.
- x, y, w and h are percentages of the monitor work area.
- Capture writes x, y, w and h as integers, but you can manually edit them to decimal values for finer placement.
- match can be contains, exact or regex.
- cascade can be yes or no.
- desktop is the virtual desktop number and is manually editable, so you can place windows across multiple virtual desktops.

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
processName | title    | match    | x  | y  | w  | h   | monitorRole | cascade | desktop
chrome.exe  | YouTube  | contains | 0  | 0  | 30 | 100 | lower-left  | no      | 2
chrome.exe  | Main     | contains | 20 | 20 | 60 | 70  | primary     | no      | 1
notepad.exe |          | contains | 45 | 0  | 55 | 100 | upper-left  | no      | 2

Excel example
- To keep the Excel VBA editor fixed on the left while cascading ordinary workbook windows, use rows like these:
  excel.exe | Visual Basic | contains | 2  | 0  | 60 | 100 | primary | no  | 1
  excel.exe | - Excel$     | regex    | 60 | 10 | 35 | 40  | primary | yes | 2
- In regex mode, - Excel$ means the title ends with - Excel; $ marks the end of the title.
- Those workbook windows are moved to desktop 2 before positioning.

With title vs without title
- These two rows show the difference:
  chrome.exe  | YouTube | contains | 0  | 0 | 30 | 100 | lower-left | no
  notepad.exe |         | contains | 45 | 0 | 55 | 100 | upper-left | no
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
- Capture records the virtual desktop number for each included window.
- Processes without a valid desktop number are added to processes_to_ignore.txt and left out of current_layout.txt.
- For multi-instance applications, capture emits both:
  - one cascade row with an empty title
  - the individual per-window rows
- Keep whichever version you prefer and delete the others.
- Capture writes integer percentages, but decimal edits in window_layouts.txt are preserved when that file is reformatted.

Virtual desktop restore behavior
- The script moves each window to its saved desktop number and then applies its size and position.
- If a saved desktop number is missing during restore, the script leaves that window on its current desktop and still continues the rest of the layout.
- If the window is already on the requested desktop, the script skips the desktop move and only reapplies size and position.
- Editing the desktop column in window_layouts.txt is the normal way to place windows across multiple desktops.
- Restore also skips any process listed in processes_to_ignore.txt.

Menu behavior
- Press Enter on a numbered choice prompt to exit or cancel instead of getting an invalid-choice message.
- If the config file is reformatted, the script prints a message telling you it was reformatted and saved.
- After applying a layout from the menu, the script also tells you the matching generated shortcut script name.

Restart from scratch
- Delete WindowLayout\window_layouts.txt and WindowLayout\current_layout.txt.
- On the next run, the script recreates a fresh empty window_layouts.txt.

Command line examples
- WindowLayout.cmd
- WindowLayout.cmd -ListLayouts
- WindowLayout.cmd -CaptureCurrent
- WindowLayout.cmd -ApplyLayout "3 monitors - developer"
