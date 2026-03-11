[CmdletBinding()]
param(
    [string]$ApplyLayout,
    [switch]$CaptureCurrent,
    [switch]$ListLayouts
)

Set-StrictMode -Version 3.0
$ErrorActionPreference = 'Stop'

$script:RootDirectory = Split-Path -Parent $MyInvocation.MyCommand.Path
$script:ConfigPath = Join-Path $script:RootDirectory 'window_layouts.txt'
$script:CurrentLayoutPath = Join-Path $script:RootDirectory 'current_layout.txt'
$script:HorizontalGap = 7
$script:VerticalGap = 6
$script:PendingMessages = @()

Add-Type -AssemblyName System.Windows.Forms

$nativeSource = @"
using System;
using System.Collections.Generic;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using System.Text;

public static class WindowLayoutNative
{
    public delegate bool EnumWindowsProc(IntPtr hWnd, IntPtr lParam);

    [StructLayout(LayoutKind.Sequential)]
    public struct RECT
    {
        public int Left;
        public int Top;
        public int Right;
        public int Bottom;
    }

    [DllImport("user32.dll")]
    public static extern bool EnumWindows(EnumWindowsProc lpEnumFunc, IntPtr lParam);

    [DllImport("user32.dll")]
    public static extern bool IsWindowVisible(IntPtr hWnd);

    [DllImport("user32.dll")]
    public static extern int GetWindowTextLength(IntPtr hWnd);

    [DllImport("user32.dll", CharSet = CharSet.Unicode)]
    public static extern int GetWindowText(IntPtr hWnd, StringBuilder lpString, int nMaxCount);

    [DllImport("user32.dll", CharSet = CharSet.Unicode)]
    public static extern int GetClassName(IntPtr hWnd, StringBuilder lpClassName, int nMaxCount);

    [DllImport("user32.dll")]
    public static extern bool GetWindowRect(IntPtr hWnd, out RECT lpRect);

    [DllImport("user32.dll")]
    public static extern bool IsIconic(IntPtr hWnd);

    [DllImport("user32.dll")]
    public static extern bool MoveWindow(IntPtr hWnd, int X, int Y, int nWidth, int nHeight, bool bRepaint);

    [DllImport("user32.dll")]
    public static extern bool ShowWindowAsync(IntPtr hWnd, int nCmdShow);

    [DllImport("user32.dll")]
    public static extern IntPtr GetForegroundWindow();

    [DllImport("user32.dll")]
    public static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint lpdwProcessId);

    [DllImport("dwmapi.dll")]
    public static extern int DwmGetWindowAttribute(IntPtr hwnd, int dwAttribute, out int pvAttribute, int cbAttribute);

    public static List<IntPtr> GetTopLevelWindows()
    {
        List<IntPtr> windows = new List<IntPtr>();
        EnumWindows(delegate (IntPtr hWnd, IntPtr lParam)
        {
            windows.Add(hWnd);
            return true;
        }, IntPtr.Zero);
        return windows;
    }

    public static string ReadWindowText(IntPtr hWnd)
    {
        int length = GetWindowTextLength(hWnd);
        StringBuilder builder = new StringBuilder(length + 1);
        GetWindowText(hWnd, builder, builder.Capacity);
        return builder.ToString();
    }

    public static string ReadClassName(IntPtr hWnd)
    {
        StringBuilder builder = new StringBuilder(256);
        GetClassName(hWnd, builder, builder.Capacity);
        return builder.ToString();
    }

    public static int GetWindowCloakedValue(IntPtr hWnd)
    {
        int cloaked = 0;
        int result = DwmGetWindowAttribute(hWnd, 14, out cloaked, Marshal.SizeOf(typeof(int)));
        if (result != 0)
        {
            return 0;
        }

        return cloaked;
    }
}

[ComImport]
[Guid("A5CD92FF-29BE-454C-8D04-D82879FB3F1B")]
[InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
public interface IVirtualDesktopManager
{
    [PreserveSig]
    int IsWindowOnCurrentVirtualDesktop(IntPtr topLevelWindow, out bool onCurrentDesktop);

    [PreserveSig]
    int GetWindowDesktopId(IntPtr topLevelWindow, out Guid desktopId);

    [PreserveSig]
    int MoveWindowToDesktop(IntPtr topLevelWindow, [MarshalAs(UnmanagedType.LPStruct)] Guid desktopId);
}

[ComImport]
[Guid("AA509086-5CA9-4C25-8F95-589D3C07B48A")]
public class VirtualDesktopManagerCom
{
}

public static class VirtualDesktopHelper
{
    private static readonly Lazy<IVirtualDesktopManager> _manager = new Lazy<IVirtualDesktopManager>(() =>
        (IVirtualDesktopManager)new VirtualDesktopManagerCom(), LazyThreadSafetyMode.ExecutionAndPublication);

    public static Guid? GetWindowDesktopId(IntPtr hWnd)
    {
        try
        {
            Guid desktopId;
            int hr = _manager.Value.GetWindowDesktopId(hWnd, out desktopId);
            if (hr != 0)
            {
                return null;
            }

            return desktopId;
        }
        catch
        {
            return null;
        }
    }

    public static bool MoveWindowToDesktop(IntPtr hWnd, Guid desktopId)
    {
        try
        {
            return _manager.Value.MoveWindowToDesktop(hWnd, desktopId) == 0;
        }
        catch
        {
            return false;
        }
    }
}
"@

if (-not ('WindowLayoutNative' -as [type])) {
    Add-Type -TypeDefinition $nativeSource
}

function New-DefaultConfig {
    [pscustomobject]@{
        monitorSetups = @()
        layouts       = @()
    }
}

function Add-PendingMessage {
    param([Parameter(Mandatory = $true)][string]$Message)

    $script:PendingMessages += $Message
}

function Show-PendingMessages {
    if ($script:PendingMessages.Count -eq 0) {
        return
    }

    foreach ($message in $script:PendingMessages) {
        Write-Host $message
    }

    Write-Host ''
    $script:PendingMessages = @()
}

function Show-MenuHeader {
    param([Parameter(Mandatory = $true)][string]$Message)

    Write-Host ''
    Show-PendingMessages
    Write-Host $Message
}

function Ensure-ConfigFile {
    if (-not (Test-Path -LiteralPath $script:ConfigPath)) {
        [System.IO.File]::WriteAllText($script:ConfigPath, '', (New-Object System.Text.UTF8Encoding($true)))
    }
}

function Get-OptionalPropertyValue {
    param(
        [Parameter(Mandatory = $true)]$Object,
        [Parameter(Mandatory = $true)][string]$Name,
        $Default = $null
    )

    if ($null -eq $Object) {
        return $Default
    }

    $property = $Object.PSObject.Properties[$Name]
    if ($null -eq $property) {
        return $Default
    }

    return $property.Value
}

function Split-TableLine {
    param(
        [Parameter(Mandatory = $true)][string]$Line,
        [int]$ExpectedColumnCount = 0,
        [switch]$AllowPipeInTitle
    )

    $parts = @($Line -split '\|')
    $trimmed = @($parts | ForEach-Object { $_.Trim() })

    if (-not $AllowPipeInTitle) {
        if ($ExpectedColumnCount -gt 0 -and $trimmed.Count -ne $ExpectedColumnCount) {
            throw "Invalid table row: '$Line'"
        }
        return ,$trimmed
    }

    if ($trimmed.Count -lt 9) {
        throw "Invalid window row: '$Line'"
    }

    # Title is the only column that may contain '|', for regex patterns like (A|B).
    # Parse the first column and the last 7 columns by position, then join the middle back into the title column.
    $tailStart = $trimmed.Count - 7
    $title = ($trimmed[1..($tailStart - 1)] -join '|').Trim()
    $result = @($trimmed[0], $title)
    $result += $trimmed[$tailStart..($trimmed.Count - 1)]
    return ,$result
}

function Test-HeaderRow {
    param(
        [string[]]$Columns,
        [string[]]$Expected
    )

    if ($Columns.Count -ne $Expected.Count) {
        return $false
    }

    for ($i = 0; $i -lt $Expected.Count; $i++) {
        if ($Columns[$i].ToLowerInvariant() -ne $Expected[$i].ToLowerInvariant()) {
            return $false
        }
    }

    return $true
}

function Parse-YesNo {
    param([string]$Value)

    switch ($Value.Trim().ToLowerInvariant()) {
        'yes' { return $true }
        'y' { return $true }
        'true' { return $true }
        '1' { return $true }
        default { return $false }
    }
}

function ConvertTo-ConfigFromText {
    param([string]$Text)

    $config = New-DefaultConfig
    $section = ''
    $currentMonitorSetup = $null
    $currentLayout = $null
    $lineNumber = 0
    $expectLayoutMonitorSetupRow = $false

    foreach ($rawLine in ($Text -split "`r?`n")) {
        $lineNumber++
        $line = $rawLine.Trim()

        if ($line -eq '' -or $line.StartsWith(';') -or $line.StartsWith('#')) {
            continue
        }

        $monitorSetupMatch = [regex]::Match($line, '^\[MonitorSetup\s+(.+?)\]$', 'IgnoreCase')
        $layoutMatch = [regex]::Match($line, '^\[Layout\s+(.+?)\]$', 'IgnoreCase')

        if ($monitorSetupMatch.Success) {
            $section = 'monitorSetup'
            $currentLayout = $null
            $currentMonitorSetup = [pscustomobject]@{
                name     = $monitorSetupMatch.Groups[1].Value.Trim()
                monitors = @()
            }
            $config.monitorSetups += $currentMonitorSetup
            continue
        }

        if ($layoutMatch.Success) {
            $section = 'layout'
            $currentMonitorSetup = $null
            $currentLayout = [pscustomobject]@{
                name         = $layoutMatch.Groups[1].Value.Trim()
                monitorSetup = ''
                windows      = @()
            }
            $config.layouts += $currentLayout
            $expectLayoutMonitorSetupRow = $true
            continue
        }

        try {
            if ($section -eq 'monitorSetup' -and $null -ne $currentMonitorSetup) {
                $columns = Split-TableLine -Line $line -ExpectedColumnCount 3
                if (Test-HeaderRow -Columns $columns -Expected @('role', 'x', 'y')) {
                    continue
                }

                $currentMonitorSetup.monitors += [pscustomobject]@{
                    role = $columns[0]
                    x    = [int]$columns[1]
                    y    = [int]$columns[2]
                }
                continue
            }

            if ($section -eq 'layout' -and $null -ne $currentLayout) {
                if ($expectLayoutMonitorSetupRow) {
                    $columns = Split-TableLine -Line $line -ExpectedColumnCount 2
                    if (-not (Test-HeaderRow -Columns $columns -Expected @('monitorSetup', $columns[1]))) {
                        if ($columns[0].ToLowerInvariant() -ne 'monitorsetup') {
                            throw 'First row after a layout header must be monitorSetup | <name>.'
                        }
                    }
                    $currentLayout.monitorSetup = $columns[1]
                    $expectLayoutMonitorSetupRow = $false
                    continue
                }

                if ($line -eq '...') {
                    $currentLayout.windows += [pscustomobject]@{ rawLine = $line }
                    continue
                }

                $columns = Split-TableLine -Line $line -AllowPipeInTitle
                if (Test-HeaderRow -Columns $columns -Expected @('processName', 'title', 'match', 'x', 'y', 'w', 'h', 'monitorRole', 'cascade')) {
                    continue
                }

                if ($line -eq '...' -or $columns[3] -eq '...' -or $columns[4] -eq '...' -or $columns[5] -eq '...' -or $columns[6] -eq '...') {
                    $currentLayout.windows += [pscustomobject]@{ rawLine = $line }
                    continue
                }

                $currentLayout.windows += [pscustomobject]@{
                    processName = $columns[0]
                    title       = $columns[1]
                    titleMatch  = if ([string]::IsNullOrWhiteSpace($columns[2])) { 'contains' } else { $columns[2] }
                    x           = [int]$columns[3]
                    y           = [int]$columns[4]
                    w           = [int]$columns[5]
                    h           = [int]$columns[6]
                    monitorRole = $columns[7]
                    cascade     = Parse-YesNo -Value $columns[8]
                }
                continue
            }

            throw 'Row is outside a known section.'
        }
        catch {
            throw "Error in '$script:ConfigPath' at line ${lineNumber}: $($_.Exception.Message)"
        }
    }

    return $config
}

function New-TableValueRow {
    param([string[]]$Values)
    return [pscustomobject]@{ kind = 'values'; values = $Values }
}

function New-TableRawRow {
    param([string]$Text)
    return [pscustomobject]@{ kind = 'raw'; text = $Text }
}

function Get-PaddedRow {
    param(
        [string[]]$Values,
        [int[]]$Widths
    )

    $cells = for ($i = 0; $i -lt $Values.Count; $i++) {
        $value = if ($null -eq $Values[$i]) { '' } else { [string]$Values[$i] }
        $value.PadRight($Widths[$i])
    }
    return ($cells -join ' | ').TrimEnd()
}

function Format-Table {
    param(
        [string[]]$Headers,
        [object[]]$Rows
    )

    $widths = @()
    for ($i = 0; $i -lt $Headers.Count; $i++) {
        $maxWidth = $Headers[$i].Length
        foreach ($row in $Rows) {
            if ($row.kind -ne 'values') {
                continue
            }

            $text = if ($null -eq $row.values[$i]) { '' } else { [string]$row.values[$i] }
            if ($text.Length -gt $maxWidth) {
                $maxWidth = $text.Length
            }
        }
        $widths += $maxWidth
    }

    $lines = @((Get-PaddedRow -Values $Headers -Widths $widths))
    foreach ($row in $Rows) {
        if ($row.kind -eq 'raw') {
            $lines += $row.text
        }
        else {
            $lines += Get-PaddedRow -Values $row.values -Widths $widths
        }
    }
    return $lines
}

function ConvertTo-ConfigText {
    param([Parameter(Mandatory = $true)]$Config)

    $lines = @()

    foreach ($setup in $Config.monitorSetups) {
        if ($lines.Count -gt 0) {
            $lines += ''
        }

        $lines += "[MonitorSetup $($setup.name)]"
        $rows = foreach ($monitor in $setup.monitors) {
            New-TableValueRow -Values @(
                [string](Get-OptionalPropertyValue -Object $monitor -Name 'role' -Default ''),
                [string](Get-OptionalPropertyValue -Object $monitor -Name 'x' -Default 0),
                [string](Get-OptionalPropertyValue -Object $monitor -Name 'y' -Default 0)
            )
        }
        $lines += Format-Table -Headers @('role', 'x', 'y') -Rows $rows
    }

    foreach ($layout in $Config.layouts) {
        if ($lines.Count -gt 0) {
            $lines += ''
        }

        $lines += "[Layout $($layout.name)]"
        $lines += "monitorSetup | $($layout.monitorSetup)"

        $rows = foreach ($window in $layout.windows) {
            $rawLine = Get-OptionalPropertyValue -Object $window -Name 'rawLine' -Default ''
            if ($rawLine) {
                New-TableRawRow -Text $rawLine
            }
            else {
                New-TableValueRow -Values @(
                    [string](Get-OptionalPropertyValue -Object $window -Name 'processName' -Default ''),
                    [string](Get-OptionalPropertyValue -Object $window -Name 'title' -Default ''),
                    [string](Get-OptionalPropertyValue -Object $window -Name 'titleMatch' -Default 'contains'),
                    [string](Get-OptionalPropertyValue -Object $window -Name 'x' -Default 0),
                    [string](Get-OptionalPropertyValue -Object $window -Name 'y' -Default 0),
                    [string](Get-OptionalPropertyValue -Object $window -Name 'w' -Default 100),
                    [string](Get-OptionalPropertyValue -Object $window -Name 'h' -Default 100),
                    [string](Get-OptionalPropertyValue -Object $window -Name 'monitorRole' -Default ''),
                    $(if ([bool](Get-OptionalPropertyValue -Object $window -Name 'cascade' -Default $false)) { 'yes' } else { 'no' })
                )
            }
        }
        $lines += Format-Table -Headers @('processName', 'title', 'match', 'x', 'y', 'w', 'h', 'monitorRole', 'cascade') -Rows $rows
    }

    if ($lines.Count -eq 0) {
        return ''
    }

    return ($lines -join [Environment]::NewLine).TrimEnd() + [Environment]::NewLine
}

function Validate-Config {
    param([Parameter(Mandatory = $true)]$Config)

    $errors = New-Object System.Collections.Generic.List[string]
    $setupNames = @{}
    $layoutKeys = @{}

    foreach ($setup in $Config.monitorSetups) {
        if ([string]::IsNullOrWhiteSpace($setup.name)) {
            $errors.Add('Each monitor setup must have a name.')
            continue
        }

        if ($setupNames.ContainsKey($setup.name)) {
            $errors.Add("Duplicate monitor setup '$($setup.name)'.")
        }
        else {
            $setupNames[$setup.name] = $setup
        }

        $roles = @{}
        foreach ($monitor in $setup.monitors) {
            $role = [string](Get-OptionalPropertyValue -Object $monitor -Name 'role' -Default '')
            if ([string]::IsNullOrWhiteSpace($role)) {
                $errors.Add("Monitor setup '$($setup.name)' has a row without a role.")
                continue
            }

            if ($roles.ContainsKey($role)) {
                $errors.Add("Monitor setup '$($setup.name)' has duplicate role '$role'.")
            }
            else {
                $roles[$role] = $true
            }
        }
    }

    foreach ($layout in $Config.layouts) {
        if ([string]::IsNullOrWhiteSpace($layout.name)) {
            $errors.Add('Each layout must have a name.')
            continue
        }

        if ([string]::IsNullOrWhiteSpace($layout.monitorSetup)) {
            $errors.Add("Layout '$($layout.name)' must declare monitorSetup.")
            continue
        }

        $layoutKey = "$($layout.monitorSetup)`n$($layout.name)"
        if ($layoutKeys.ContainsKey($layoutKey)) {
            $errors.Add("Duplicate layout '$($layout.name)' for monitor setup '$($layout.monitorSetup)'.")
        }
        else {
            $layoutKeys[$layoutKey] = $true
        }

        $setup = $setupNames[$layout.monitorSetup]
        if ($null -eq $setup) {
            $errors.Add("Layout '$($layout.name)' references unknown monitor setup '$($layout.monitorSetup)'.")
            continue
        }

        $roleNames = @{}
        foreach ($monitor in $setup.monitors) {
            $roleNames[[string]$monitor.role] = $true
        }

        foreach ($window in $layout.windows) {
            if (Get-OptionalPropertyValue -Object $window -Name 'rawLine' -Default '') {
                continue
            }

            $role = [string](Get-OptionalPropertyValue -Object $window -Name 'monitorRole' -Default '')
            if ([string]::IsNullOrWhiteSpace($role)) {
                $errors.Add("Layout '$($layout.name)' has a row without monitorRole.")
            }
            elseif (-not $roleNames.ContainsKey($role)) {
                $errors.Add("Layout '$($layout.name)' references unknown monitor role '$role' for monitor setup '$($layout.monitorSetup)'.")
            }
        }
    }

    if ($errors.Count -gt 0) {
        throw ($errors -join [Environment]::NewLine)
    }
}

function Write-FormattedConfigIfNeeded {
    param(
        [Parameter(Mandatory = $true)][string]$Path,
        [AllowEmptyString()][string]$FormattedText
    )

    $existing = if (Test-Path -LiteralPath $Path) { [System.IO.File]::ReadAllText($Path, [System.Text.Encoding]::UTF8) } else { '' }
    if ($existing -ne $FormattedText) {
        [System.IO.File]::WriteAllText($Path, $FormattedText, (New-Object System.Text.UTF8Encoding($true)))
        return $true
    }

    return $false
}

function ConvertTo-ShortcutFileName {
    param([Parameter(Mandatory = $true)][string]$Name)

    $safeName = $Name
    foreach ($invalidChar in [System.IO.Path]::GetInvalidFileNameChars()) {
        $safeName = $safeName.Replace([string]$invalidChar, '_')
    }

    return $safeName.TrimEnd(' ', '.')
}

function Get-LayoutShortcutContent {
    param([Parameter(Mandatory = $true)][string]$LayoutLabel)

    $escapedLayoutLabel = $LayoutLabel.Replace('"', '""')
    return ('..\WindowLayout.cmd -ApplyLayout "{0}"' -f $escapedLayoutLabel) + [Environment]::NewLine
}

function Sync-LayoutShortcutScripts {
    param([Parameter(Mandatory = $true)]$Config)

    $expectedFiles = @{}
    foreach ($layout in $Config.layouts) {
        $layoutLabel = "$($layout.monitorSetup) - $($layout.name)"
        $fileName = ConvertTo-ShortcutFileName -Name ("WindowLayout - {0}.cmd" -f $layoutLabel)
        $expectedFiles[$fileName] = Get-LayoutShortcutContent -LayoutLabel $layoutLabel
    }

    $existingFiles = @{}
    foreach ($file in @(Get-ChildItem -LiteralPath $script:RootDirectory -Filter '*.cmd' -File)) {
        $existingFiles[$file.Name] = Get-Content -LiteralPath $file.FullName -Raw
    }

    $needsSync = $existingFiles.Count -ne $expectedFiles.Count
    if (-not $needsSync) {
        foreach ($fileName in $expectedFiles.Keys) {
            if (-not $existingFiles.ContainsKey($fileName) -or $existingFiles[$fileName] -ne $expectedFiles[$fileName]) {
                $needsSync = $true
                break
            }
        }
    }

    if (-not $needsSync) {
        return $false
    }

    foreach ($file in @(Get-ChildItem -LiteralPath $script:RootDirectory -Filter '*.cmd' -File)) {
        Remove-Item -LiteralPath $file.FullName -Force
    }

    foreach ($fileName in ($expectedFiles.Keys | Sort-Object)) {
        $filePath = Join-Path $script:RootDirectory $fileName
        Set-Content -LiteralPath $filePath -Value $expectedFiles[$fileName] -Encoding ASCII
    }

    return $true
}

function Get-Config {
    Ensure-ConfigFile
    $raw = Get-Content -LiteralPath $script:ConfigPath -Raw -Encoding UTF8
    $config = ConvertTo-ConfigFromText -Text $raw
    Validate-Config -Config $config
    $wasReformatted = Write-FormattedConfigIfNeeded -Path $script:ConfigPath -FormattedText (ConvertTo-ConfigText -Config $config)
    if ($wasReformatted) {
        [void](Sync-LayoutShortcutScripts -Config $config)
        Add-PendingMessage -Message "Reformatted and saved '$script:ConfigPath'."
        Add-PendingMessage -Message "Updated layout shortcut scripts in '$script:RootDirectory'."
    }
    return $config
}

function Get-ActualMonitors {
    $screens = [System.Windows.Forms.Screen]::AllScreens
    $index = 0
    foreach ($screen in $screens) {
        $index++
        [pscustomobject]@{
            Index  = $index
            X      = $screen.WorkingArea.Left
            Y      = $screen.WorkingArea.Top
            Width  = $screen.WorkingArea.Width
            Height = $screen.WorkingArea.Height
        }
    }
}

function Get-NormalizedPoints {
    param([Parameter(Mandatory = $true)]$Items)

    $xs = @($Items | ForEach-Object { [double]$_.x })
    $ys = @($Items | ForEach-Object { [double]$_.y })
    $minX = ($xs | Measure-Object -Minimum).Minimum
    $maxX = ($xs | Measure-Object -Maximum).Maximum
    $minY = ($ys | Measure-Object -Minimum).Minimum
    $maxY = ($ys | Measure-Object -Maximum).Maximum
    $spanX = [Math]::Max($maxX - $minX, 1.0)
    $spanY = [Math]::Max($maxY - $minY, 1.0)

    $normalized = @()
    foreach ($item in $Items) {
        $normalized += [pscustomobject]@{
            original = $item
            x        = ([double]$item.x - $minX) / $spanX
            y        = ([double]$item.y - $minY) / $spanY
        }
    }

    return $normalized
}

function Get-Permutations {
    param([int[]]$Numbers)

    if ($Numbers.Count -le 1) {
        return ,@($Numbers)
    }

    $results = New-Object System.Collections.Generic.List[object]
    for ($i = 0; $i -lt $Numbers.Count; $i++) {
        $head = $Numbers[$i]
        $tail = @()
        for ($j = 0; $j -lt $Numbers.Count; $j++) {
            if ($j -ne $i) {
                $tail += $Numbers[$j]
            }
        }

        foreach ($perm in (Get-Permutations -Numbers $tail)) {
            $results.Add(@($head) + @($perm))
        }
    }

    return $results
}

function Get-MonitorMapping {
    param(
        [Parameter(Mandatory = $true)]$MonitorSetup,
        [Parameter(Mandatory = $true)]$ActualMonitors
    )

    $expected = @($MonitorSetup.monitors)
    $actual = @($ActualMonitors)
    if ($expected.Count -ne $actual.Count) {
        throw "Monitor setup '$($MonitorSetup.name)' expects $($expected.Count) monitors but $($actual.Count) are currently detected."
    }

    $expectedNormalized = @(Get-NormalizedPoints -Items $expected)
    $actualPoints = foreach ($monitor in $actual) {
        [pscustomobject]@{ x = $monitor.X; y = $monitor.Y; actual = $monitor }
    }
    $actualNormalized = @(Get-NormalizedPoints -Items $actualPoints)

    $indexes = @()
    for ($i = 0; $i -lt $actualNormalized.Count; $i++) {
        $indexes += $i
    }

    $bestPermutation = $null
    $bestCost = [double]::PositiveInfinity
    foreach ($perm in (Get-Permutations -Numbers $indexes)) {
        $cost = 0.0
        for ($i = 0; $i -lt $expectedNormalized.Count; $i++) {
            $dx = $expectedNormalized[$i].x - $actualNormalized[$perm[$i]].x
            $dy = $expectedNormalized[$i].y - $actualNormalized[$perm[$i]].y
            $cost += [Math]::Sqrt(($dx * $dx) + ($dy * $dy))
        }

        if ($cost -lt $bestCost) {
            $bestCost = $cost
            $bestPermutation = $perm
        }
    }

    $mapping = @{}
    for ($i = 0; $i -lt $expected.Count; $i++) {
        $mapping[[string]$expected[$i].role] = $actual[$bestPermutation[$i]]
    }

    return $mapping
}

function Get-WindowProcessName {
    param([IntPtr]$Handle)

    $processId = [uint32]0
    [void][WindowLayoutNative]::GetWindowThreadProcessId($Handle, [ref]$processId)
    if ($processId -eq 0) {
        return $null
    }

    try {
        return (Get-Process -Id $processId -ErrorAction Stop).ProcessName + '.exe'
    }
    catch {
        return $null
    }
}

function Get-WindowRectObject {
    param([IntPtr]$Handle)

    $rect = New-Object WindowLayoutNative+RECT
    if (-not [WindowLayoutNative]::GetWindowRect($Handle, [ref]$rect)) {
        return $null
    }

    [pscustomobject]@{
        Left   = $rect.Left
        Top    = $rect.Top
        Right  = $rect.Right
        Bottom = $rect.Bottom
        Width  = $rect.Right - $rect.Left
        Height = $rect.Bottom - $rect.Top
    }
}

function Get-OpenWindows {
    $skipClasses = @('Progman', 'WorkerW', 'Shell_TrayWnd')

    foreach ($handle in [WindowLayoutNative]::GetTopLevelWindows()) {
        if (-not [WindowLayoutNative]::IsWindowVisible($handle)) {
            continue
        }

        $title = [WindowLayoutNative]::ReadWindowText($handle)
        if ([string]::IsNullOrWhiteSpace($title)) {
            continue
        }

        $className = [WindowLayoutNative]::ReadClassName($handle)
        if ($skipClasses -contains $className) {
            continue
        }

        $processName = Get-WindowProcessName -Handle $handle
        if ([string]::IsNullOrWhiteSpace($processName)) {
            continue
        }

        $rect = Get-WindowRectObject -Handle $handle
        if ($null -eq $rect -or $rect.Width -le 0 -or $rect.Height -le 0) {
            continue
        }

        [pscustomobject]@{
            Handle      = $handle
            Title       = $title
            ProcessName = $processName
            Rect        = $rect
            IsMinimized = [WindowLayoutNative]::IsIconic($handle)
        }
    }
}

function Test-TitleMatch {
    param(
        [string]$ActualTitle,
        [string]$ExpectedTitle,
        [string]$MatchMode
    )

    if ([string]::IsNullOrWhiteSpace($ExpectedTitle)) {
        return $true
    }

    $effectiveMatchMode = if ([string]::IsNullOrWhiteSpace($MatchMode)) { 'contains' } else { $MatchMode }
    switch ($effectiveMatchMode.ToLowerInvariant()) {
        'exact' { return $ActualTitle -eq $ExpectedTitle }
        'regex' { return $ActualTitle -match $ExpectedTitle }
        default { return $ActualTitle -like "*$ExpectedTitle*" }
    }
}

function Find-MatchingWindows {
    param(
        [Parameter(Mandatory = $true)]$WindowDefinition,
        [Parameter(Mandatory = $true)]$AvailableWindows
    )

    $expectedProcess = [string](Get-OptionalPropertyValue -Object $WindowDefinition -Name 'processName' -Default '')
    $expectedTitle = [string](Get-OptionalPropertyValue -Object $WindowDefinition -Name 'title' -Default '')
    $matchMode = [string](Get-OptionalPropertyValue -Object $WindowDefinition -Name 'titleMatch' -Default 'contains')

    $matches = @(
        $AvailableWindows | Where-Object {
            ($expectedProcess -eq '' -or $_.ProcessName -ieq $expectedProcess) -and
            (Test-TitleMatch -ActualTitle $_.Title -ExpectedTitle $expectedTitle -MatchMode $matchMode)
        }
    )

    if ($matches.Count -eq 0) {
        return @()
    }

    if ([bool](Get-OptionalPropertyValue -Object $WindowDefinition -Name 'cascade' -Default $false)) {
        return $matches
    }

    $foreground = [WindowLayoutNative]::GetForegroundWindow()
    $activeMatch = $matches | Where-Object { $_.Handle -eq $foreground } | Select-Object -First 1
    if ($activeMatch) {
        return @($activeMatch)
    }

    return @($matches[0])
}

function Move-WindowToLayout {
    param(
        [Parameter(Mandatory = $true)]$Window,
        [Parameter(Mandatory = $true)]$WindowDefinition,
        [Parameter(Mandatory = $true)]$Monitor,
        [int]$OffsetX = 0,
        [int]$OffsetY = 0
    )

    $xPct = [int](Get-OptionalPropertyValue -Object $WindowDefinition -Name 'x' -Default 0)
    $yPct = [int](Get-OptionalPropertyValue -Object $WindowDefinition -Name 'y' -Default 0)
    $wPct = [int](Get-OptionalPropertyValue -Object $WindowDefinition -Name 'w' -Default 100)
    $hPct = [int](Get-OptionalPropertyValue -Object $WindowDefinition -Name 'h' -Default 100)

    $left = $Monitor.X + [Math]::Floor($Monitor.Width * $xPct / 100)
    $top = $Monitor.Y + [Math]::Floor($Monitor.Height * $yPct / 100)
    $right = $Monitor.X + [Math]::Floor($Monitor.Width * ($xPct + $wPct) / 100)
    $bottom = $Monitor.Y + [Math]::Floor($Monitor.Height * ($yPct + $hPct) / 100)

    $x = $left - $script:HorizontalGap + $OffsetX
    $y = $top - $script:VerticalGap + $OffsetY
    $width = ($right - $left) + ($script:HorizontalGap * 2)
    $height = ($bottom - $top) + ($script:VerticalGap * 2)

    [void][WindowLayoutNative]::ShowWindowAsync($Window.Handle, 9)
    [void][WindowLayoutNative]::MoveWindow($Window.Handle, $x, $y, $width, $height, $true)

    if ($Window.IsMinimized) {
        [void][WindowLayoutNative]::ShowWindowAsync($Window.Handle, 6)
    }
}

function Get-MonitorSetupByName {
    param(
        [Parameter(Mandatory = $true)]$Config,
        [Parameter(Mandatory = $true)][string]$Name
    )

    return $Config.monitorSetups | Where-Object { $_.name -eq $Name } | Select-Object -First 1
}

function Apply-Layout {
    param(
        [Parameter(Mandatory = $true)]$Layout,
        [Parameter(Mandatory = $true)]$Config
    )

    $setup = Get-MonitorSetupByName -Config $Config -Name $Layout.monitorSetup
    if ($null -eq $setup) {
        throw "Monitor setup '$($Layout.monitorSetup)' was not found."
    }

    $mapping = Get-MonitorMapping -MonitorSetup $setup -ActualMonitors (Get-ActualMonitors)
    $availableWindows = @(Get-OpenWindows)

    foreach ($windowDefinition in $Layout.windows) {
        if (Get-OptionalPropertyValue -Object $windowDefinition -Name 'rawLine' -Default '') {
            continue
        }

        $role = [string](Get-OptionalPropertyValue -Object $windowDefinition -Name 'monitorRole' -Default '')
        $monitor = $mapping[$role]
        $matches = @(Find-MatchingWindows -WindowDefinition $windowDefinition -AvailableWindows $availableWindows)
        if ($matches.Count -eq 0) {
            continue
        }

        $cascadeOffset = 30
        $index = 0
        foreach ($window in $matches) {
            $offsetX = 0
            $offsetY = 0
            if ([bool](Get-OptionalPropertyValue -Object $windowDefinition -Name 'cascade' -Default $false) -and $matches.Count -gt 1) {
                $offsetX = $cascadeOffset * $index
                $offsetY = -1 * $cascadeOffset * $index
            }

            Move-WindowToLayout -Window $window -WindowDefinition $windowDefinition -Monitor $monitor -OffsetX $offsetX -OffsetY $offsetY
            $index++
        }
    }
}

function Get-MonitorForWindow {
    param(
        [Parameter(Mandatory = $true)]$Window,
        [Parameter(Mandatory = $true)]$ActualMonitors
    )

    $centerX = $Window.Rect.Left + [Math]::Floor($Window.Rect.Width / 2)
    $centerY = $Window.Rect.Top + [Math]::Floor($Window.Rect.Height / 2)

    foreach ($monitor in $ActualMonitors) {
        if ($centerX -ge $monitor.X -and $centerX -lt ($monitor.X + $monitor.Width) -and $centerY -ge $monitor.Y -and $centerY -lt ($monitor.Y + $monitor.Height)) {
            return $monitor
        }
    }

    return $ActualMonitors | Select-Object -First 1
}

function Get-Percent {
    param(
        [int]$Value,
        [int]$Minimum,
        [int]$Maximum
    )

    if ($Maximum -le $Minimum) {
        return 0
    }

    $percent = [Math]::Round((($Value - $Minimum) * 100.0) / ($Maximum - $Minimum), 0)
    if ($percent -lt 0) { return 0 }
    if ($percent -gt 100) { return 100 }
    return [int]$percent
}

function Get-SizePercent {
    param(
        [int]$Size,
        [int]$TotalSize
    )

    if ($TotalSize -le 0) {
        return 0
    }

    $percent = [Math]::Round(($Size * 100.0) / $TotalSize, 0)
    if ($percent -lt 0) { return 0 }
    if ($percent -gt 100) { return 100 }
    return [int]$percent
}

function New-RoleFromCoordinates {
    param(
        [int]$X,
        [int]$Y
    )

    return "x$X`_y$Y"
}

function New-MonitorSetupFromCurrentMonitors {
    param(
        [Parameter(Mandatory = $true)][string]$Name,
        [Parameter(Mandatory = $true)]$ActualMonitors
    )

    $monitors = foreach ($monitor in $ActualMonitors) {
        [pscustomobject]@{
            role = New-RoleFromCoordinates -X $monitor.X -Y $monitor.Y
            x    = $monitor.X
            y    = $monitor.Y
        }
    }

    return [pscustomobject]@{
        name     = $Name
        monitors = @($monitors)
    }
}

function Read-ChoiceNumber {
    param(
        [Parameter(Mandatory = $true)][string]$Prompt,
        [Parameter(Mandatory = $true)][int]$Minimum,
        [Parameter(Mandatory = $true)][int]$Maximum,
        [switch]$AllowBlank,
        [int]$BlankValue = 0
    )

    while ($true) {
        $choice = Read-Host $Prompt
        if ($AllowBlank -and [string]::IsNullOrWhiteSpace($choice)) {
            return $BlankValue
        }

        $number = 0
        if ([int]::TryParse($choice, [ref]$number) -and $number -ge $Minimum -and $number -le $Maximum) {
            return $number
        }

        Write-Host 'Invalid choice.'
    }
}

function Select-MonitorSetupForCapture {
    param([Parameter(Mandatory = $true)]$Config)

    $actualMonitors = @(Get-ActualMonitors)
    if ($Config.monitorSetups.Count -eq 0) {
        $name = Read-Host "No monitor setups exist yet. Enter a name for the new monitor setup"
        if ([string]::IsNullOrWhiteSpace($name)) {
            $name = 'Current setup'
        }

        return [pscustomobject]@{
            setup = New-MonitorSetupFromCurrentMonitors -Name $name -ActualMonitors $actualMonitors
            isNew = $true
        }
    }

    Show-MenuHeader -Message 'Choose the monitor setup to use for this capture:'
    $index = 1
    foreach ($setup in $Config.monitorSetups) {
        Write-Host ("{0}. {1}" -f $index, $setup.name)
        $index++
    }
    Write-Host ("{0}. create new monitor setup" -f $index)

    $selection = Read-ChoiceNumber -Prompt 'Choose an option' -Minimum 1 -Maximum $index -AllowBlank -BlankValue 0
    if ($selection -eq 0) {
        return $null
    }

    if ($selection -eq $index) {
        $name = Read-Host 'Enter the new monitor setup name'
        if ([string]::IsNullOrWhiteSpace($name)) {
            $name = 'Current setup'
        }

        return [pscustomobject]@{
            setup = New-MonitorSetupFromCurrentMonitors -Name $name -ActualMonitors $actualMonitors
            isNew = $true
        }
    }

    return [pscustomobject]@{
        setup = $Config.monitorSetups[$selection - 1]
        isNew = $false
    }
}

function Capture-CurrentLayout {
    param(
        [Parameter(Mandatory = $true)]$Config,
        [Parameter(Mandatory = $true)]$CaptureTarget
    )

    $actualMonitors = @(Get-ActualMonitors)
    $mapping = Get-MonitorMapping -MonitorSetup $CaptureTarget.setup -ActualMonitors $actualMonitors
    $windows = @(Get-OpenWindows | Where-Object { -not $_.IsMinimized } | Sort-Object ProcessName, Title)

    $capturedRows = foreach ($window in $windows) {
        $actualMonitor = Get-MonitorForWindow -Window $window -ActualMonitors $actualMonitors
        $monitorRole = $null
        foreach ($role in $mapping.Keys) {
            if ($mapping[$role].Index -eq $actualMonitor.Index) {
                $monitorRole = $role
                break
            }
        }

        $innerLeft = $window.Rect.Left + $script:HorizontalGap
        $innerTop = $window.Rect.Top + $script:VerticalGap
        $innerRight = $window.Rect.Right - $script:HorizontalGap
        $innerBottom = $window.Rect.Bottom - $script:VerticalGap

        [pscustomobject]@{
            processName = $window.ProcessName
            title       = $window.Title
            titleMatch  = 'contains'
            x           = Get-Percent -Value $innerLeft -Minimum $actualMonitor.X -Maximum ($actualMonitor.X + $actualMonitor.Width)
            y           = Get-Percent -Value $innerTop -Minimum $actualMonitor.Y -Maximum ($actualMonitor.Y + $actualMonitor.Height)
            w           = Get-SizePercent -Size ($innerRight - $innerLeft) -TotalSize $actualMonitor.Width
            h           = Get-SizePercent -Size ($innerBottom - $innerTop) -TotalSize $actualMonitor.Height
            monitorRole = $monitorRole
            cascade     = $false
        }
    }

    $layoutWindows = @()
    $groups = $capturedRows | Group-Object -Property processName
    foreach ($group in $groups) {
        if ($group.Count -gt 1) {
            $first = $group.Group | Select-Object -First 1
            $layoutWindows += [pscustomobject]@{
                processName = $first.processName
                title       = ''
                titleMatch  = 'contains'
                x           = $first.x
                y           = $first.y
                w           = $first.w
                h           = $first.h
                monitorRole = $first.monitorRole
                cascade     = $true
            }
        }

        foreach ($item in $group.Group) {
            $layoutWindows += $item
        }
    }

    $captureConfig = New-DefaultConfig
    $captureConfig.monitorSetups = @($CaptureTarget.setup)
    $captureConfig.layouts = @(
        [pscustomobject]@{
            name         = 'Current'
            monitorSetup = $CaptureTarget.setup.name
            windows      = @($layoutWindows)
        }
    )

    return $captureConfig
}

function Save-CapturedLayout {
    param(
        [Parameter(Mandatory = $true)]$Config
    )

    $text = ConvertTo-ConfigText -Config $Config
    Set-Content -LiteralPath $script:CurrentLayoutPath -Value $text -Encoding UTF8
}

function Show-Menu {
    param([Parameter(Mandatory = $true)]$Config)

    while ($true) {
        Clear-Host
        Write-Host 'Window Layout'
        Write-Host '============='
        Show-MenuHeader -Message 'Choose one of the following options:'

        $options = @()
        $index = 1
        foreach ($layout in $Config.layouts) {
            $label = "$($layout.monitorSetup) - $($layout.name)"
            $options += [pscustomobject]@{ Number = $index; Action = 'apply'; Layout = $layout; Label = $label }
            Write-Host ("{0}. set window layout '{1}'" -f $index, $label)
            $index++
        }

        $captureNumber = $index
        $options += [pscustomobject]@{ Number = $captureNumber; Action = 'capture' }
        Write-Host ("{0}. get current window layout" -f $captureNumber)
        Write-Host '0. exit'
        Write-Host ''

        $choice = Read-ChoiceNumber -Prompt 'Choose an option' -Minimum 0 -Maximum $captureNumber -AllowBlank -BlankValue 0
        if ($choice -eq 0) {
            return
        }

        $selected = $options | Where-Object { $_.Number -eq $choice } | Select-Object -First 1
        try {
            if ($selected.Action -eq 'apply') {
                Apply-Layout -Layout $selected.Layout -Config $Config
                Write-Host ''
                Write-Host ("Applied layout '{0}'. Press Enter to continue." -f $selected.Label)
                Write-Host ('Tip: you can create a shortcut script containing: WindowLayout.cmd -ApplyLayout "{0}"' -f $selected.Label)
                Write-Host ('Shortcut script available: "WindowLayout - {0}.cmd" in "{1}".' -f (ConvertTo-ShortcutFileName -Name $selected.Label), $script:RootDirectory)
                [void][Console]::ReadLine()
            }
            else {
                $target = Select-MonitorSetupForCapture -Config $Config
                if ($null -eq $target) {
                    return
                }

                $snapshot = Capture-CurrentLayout -Config $Config -CaptureTarget $target
                Save-CapturedLayout -Config $snapshot
                Write-Host ''
                Write-Host ("Saved current layout to '{0}'." -f $script:CurrentLayoutPath)
                Write-Host ("Copy the monitor setup and layout you want into '{0}', remove the rows you do not need, and see '{1}' for details." -f $script:ConfigPath, (Join-Path $script:RootDirectory 'readme.txt'))
                Write-Host 'Press Enter to continue.'
                [void][Console]::ReadLine()
            }
        }
        catch {
            Write-Host ''
            Write-Host $_.Exception.Message
            Write-Host 'Press Enter to continue.'
            [void][Console]::ReadLine()
        }

        $Config = Get-Config
    }
}

Ensure-ConfigFile
$config = Get-Config

if ($ListLayouts) {
    Show-PendingMessages
    foreach ($layout in $config.layouts) {
        "$($layout.monitorSetup) - $($layout.name)"
    }
    exit 0
}

if ($CaptureCurrent) {
    Show-PendingMessages
    $target = Select-MonitorSetupForCapture -Config $config
    if ($null -eq $target) {
        exit 0
    }

    $snapshot = Capture-CurrentLayout -Config $config -CaptureTarget $target
    Save-CapturedLayout -Config $snapshot
    Write-Host "Saved current layout to '$script:CurrentLayoutPath'."
    Write-Host "Copy the monitor setup and layout you want into '$script:ConfigPath', remove the rows you do not need, and see '$(Join-Path $script:RootDirectory 'readme.txt')' for details."
    exit 0
}

if (-not [string]::IsNullOrWhiteSpace($ApplyLayout)) {
    Show-PendingMessages
    $layout = $config.layouts | Where-Object { "$($_.monitorSetup) - $($_.name)" -eq $ApplyLayout -or $_.name -eq $ApplyLayout } | Select-Object -First 1
    if ($null -eq $layout) {
        throw "Layout '$ApplyLayout' was not found."
    }

    Apply-Layout -Layout $layout -Config $config
    Write-Host "Applied layout '$($layout.monitorSetup) - $($layout.name)'."
    Write-Host ('Tip: you can create a shortcut script containing: WindowLayout.cmd -ApplyLayout "{0}"' -f "$($layout.monitorSetup) - $($layout.name)")
    exit 0
}

Show-Menu -Config $config
