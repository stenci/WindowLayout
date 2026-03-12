[CmdletBinding()]
param(
    [string]$ApplyLayout,
    [switch]$CaptureCurrent,
    [switch]$ListLayouts,
    [switch]$IgnoreBlacklist
)

Set-StrictMode -Version 3.0
$ErrorActionPreference = 'Stop'

$script:RootDirectory = Split-Path -Parent $MyInvocation.MyCommand.Path
$script:ConfigPath = Join-Path $script:RootDirectory 'window_layouts.txt'
$script:CurrentLayoutPath = Join-Path $script:RootDirectory 'current_layout.txt'
$script:IgnoredProcessesPath = Join-Path $script:RootDirectory 'processes_to_ignore.txt'
$script:PendingMessages = @()

Add-Type -AssemblyName System.Windows.Forms
[System.Threading.Thread]::CurrentThread.CurrentCulture = [System.Globalization.CultureInfo]::InvariantCulture
[System.Threading.Thread]::CurrentThread.CurrentUICulture = [System.Globalization.CultureInfo]::InvariantCulture

$nativeSource = @"
using System;
using System.Collections.Generic;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using System.Threading;
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
    int MoveWindowToDesktop(IntPtr topLevelWindow, ref Guid desktopId);
}

[ComImport]
[InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
[Guid("372E1D3B-38D3-42E4-A15B-8AB2B178F513")]
public interface IApplicationView
{
}

[ComImport]
[InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
[Guid("1841C6D7-4F9D-42C0-AF41-8747538F10E5")]
public interface IApplicationViewCollection
{
    int GetViews(out IObjectArray array);
    int GetViewsByZOrder(out IObjectArray array);
    int GetViewsByAppUserModelId(string id, out IObjectArray array);
    int GetViewForHwnd(IntPtr hwnd, out IApplicationView view);
}

[ComImport]
[InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
[Guid("FF72FFDD-BE7E-43FC-9C03-AD81681E88E4")]
public interface IVirtualDesktop
{
    bool IsViewVisible(IApplicationView view);
    Guid GetId();
}

[ComImport]
[InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
[Guid("F31574D6-B682-4CDC-BD56-1827860ABEC6")]
public interface IVirtualDesktopManagerInternal
{
    int GetCount();
    void MoveViewToDesktop(IApplicationView view, IVirtualDesktop desktop);
    bool CanViewMoveDesktops(IApplicationView view);
    IVirtualDesktop GetCurrentDesktop();
    void GetDesktops(out IObjectArray desktops);
    int GetAdjacentDesktop(IVirtualDesktop from, int direction, out IVirtualDesktop desktop);
    void SwitchDesktop(IVirtualDesktop desktop);
    IVirtualDesktop CreateDesktop();
    void RemoveDesktop(IVirtualDesktop desktop, IVirtualDesktop fallback);
    IVirtualDesktop FindDesktop(ref Guid desktopId);
}

[ComImport]
[InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
[Guid("92CA9DCD-5622-4BBA-A805-5E9F541BD8C9")]
public interface IObjectArray
{
    void GetCount(out int count);
    void GetAt(int index, ref Guid iid, [MarshalAs(UnmanagedType.Interface)] out object obj);
}

[ComImport]
[InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
[Guid("6D5140C1-7436-11CE-8034-00AA006009FA")]
public interface IServiceProvider10
{
    [return: MarshalAs(UnmanagedType.IUnknown)]
    object QueryService(ref Guid service, ref Guid riid);
}

[ComImport]
[Guid("AA509086-5CA9-4C25-8F95-589D3C07B48A")]
public class VirtualDesktopManagerCom
{
}

public static class VirtualDesktopHelper
{
    private static readonly Guid CLSID_ImmersiveShell = new Guid("C2F03A33-21F5-47FA-B4BB-156362A2F239");
    private static readonly Guid CLSID_VirtualDesktopManagerInternal = new Guid("C5E0CDCA-7B6E-41B2-9FC4-D93975CC467B");
    private static readonly Lazy<IVirtualDesktopManager> _manager = new Lazy<IVirtualDesktopManager>(() =>
        (IVirtualDesktopManager)new VirtualDesktopManagerCom(), LazyThreadSafetyMode.ExecutionAndPublication);
    private static readonly Lazy<IVirtualDesktopManagerInternal> _internalManager = new Lazy<IVirtualDesktopManagerInternal>(() =>
    {
        IServiceProvider10 shell = (IServiceProvider10)Activator.CreateInstance(Type.GetTypeFromCLSID(CLSID_ImmersiveShell));
        Guid service = CLSID_VirtualDesktopManagerInternal;
        Guid iid = typeof(IVirtualDesktopManagerInternal).GUID;
        return (IVirtualDesktopManagerInternal)shell.QueryService(ref service, ref iid);
    }, LazyThreadSafetyMode.ExecutionAndPublication);
    private static readonly Lazy<IApplicationViewCollection> _viewCollection = new Lazy<IApplicationViewCollection>(() =>
    {
        IServiceProvider10 shell = (IServiceProvider10)Activator.CreateInstance(Type.GetTypeFromCLSID(CLSID_ImmersiveShell));
        Guid iid = typeof(IApplicationViewCollection).GUID;
        return (IApplicationViewCollection)shell.QueryService(ref iid, ref iid);
    }, LazyThreadSafetyMode.ExecutionAndPublication);

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
            return MoveWindowToDesktopWithResult(hWnd, desktopId) == 0;
        }
        catch
        {
            return false;
        }
    }

    public static int MoveWindowToDesktopWithResult(IntPtr hWnd, Guid desktopId)
    {
        try
        {
            IApplicationView view;
            int hr = _viewCollection.Value.GetViewForHwnd(hWnd, out view);
            if (hr != 0 || view == null)
            {
                return hr != 0 ? hr : unchecked((int)0x80004005);
            }

            IVirtualDesktop desktop = _internalManager.Value.FindDesktop(ref desktopId);
            if (desktop == null)
            {
                return unchecked((int)0x80070057);
            }

            _internalManager.Value.MoveViewToDesktop(view, desktop);
            return 0;
        }
        catch (COMException ex)
        {
            return ex.HResult;
        }
        catch
        {
            return unchecked((int)0x80004005);
        }
    }

    public static string[] GetDesktopIds()
    {
        try
        {
            IObjectArray desktops;
            _internalManager.Value.GetDesktops(out desktops);
            int count;
            desktops.GetCount(out count);
            List<string> ids = new List<string>(count);
            Guid iid = typeof(IVirtualDesktop).GUID;
            for (int i = 0; i < count; i++)
            {
                object desktopObject;
                desktops.GetAt(i, ref iid, out desktopObject);
                IVirtualDesktop desktop = (IVirtualDesktop)desktopObject;
                ids.Add(desktop.GetId().ToString("D").ToLowerInvariant());
                Marshal.ReleaseComObject(desktopObject);
            }

            Marshal.ReleaseComObject(desktops);
            return ids.ToArray();
        }
        catch
        {
            return new string[0];
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

function Get-IgnoredProcesses {
    $map = @{}
    if (-not (Test-Path -LiteralPath $script:IgnoredProcessesPath)) {
        return $map
    }

    foreach ($line in (Get-Content -LiteralPath $script:IgnoredProcessesPath -Encoding UTF8)) {
        $name = $line.Trim()
        if ([string]::IsNullOrWhiteSpace($name)) {
            continue
        }

        if ($name.StartsWith('#')) {
            continue
        }

        $map[$name.ToLowerInvariant()] = $true
    }

    return $map
}

function Save-IgnoredProcesses {
    param([Parameter(Mandatory = $true)]$IgnoredProcesses)

    $names = @($IgnoredProcesses.Keys | Sort-Object)
    $lines = @(
        '# Processes whose windows did not appear on real virtual desktops during a full scan.',
        '# Delete a line if you want the script to try that process again.',
        ''
    ) + $names

    Set-Content -LiteralPath $script:IgnoredProcessesPath -Value $lines -Encoding UTF8
}

function Update-IgnoredProcessesFromObservations {
    param(
        [Parameter(Mandatory = $true)]$ExistingIgnored,
        [Parameter(Mandatory = $true)]$Observations,
        [hashtable]$ProtectedProcesses = @{}
    )

    $updated = @{}
    foreach ($key in $ExistingIgnored.Keys) {
        $updated[$key] = $true
    }

    foreach ($key in $ProtectedProcesses.Keys) {
        if ($updated.ContainsKey($key)) {
            $updated.Remove($key)
        }
    }

    foreach ($entry in $Observations.GetEnumerator()) {
        $processName = $entry.Key
        $stats = $entry.Value
        if ($ProtectedProcesses.ContainsKey($processName)) {
            if ($updated.ContainsKey($processName)) {
                $updated.Remove($processName)
            }
            continue
        }

        if ($stats.RealDesktopCount -gt 0) {
            if ($updated.ContainsKey($processName)) {
                $updated.Remove($processName)
            }
            continue
        }

        if ($stats.UnknownDesktopCount -gt 0) {
            $updated[$processName] = $true
        }
    }

    Save-IgnoredProcesses -IgnoredProcesses $updated
    return $updated
}

function Add-IgnoredProcesses {
    param([string[]]$ProcessNames)

    if ($null -eq $ProcessNames -or $ProcessNames.Count -eq 0) {
        return Get-IgnoredProcesses
    }

    $updated = Get-IgnoredProcesses
    foreach ($name in $ProcessNames) {
        if ([string]::IsNullOrWhiteSpace($name)) {
            continue
        }

        $updated[$name.Trim().ToLowerInvariant()] = $true
    }

    Save-IgnoredProcesses -IgnoredProcesses $updated
    return $updated
}

function Get-ProtectedProcessesFromConfig {
    param($Config)

    $result = @{}
    if ($null -eq $Config) {
        return $result
    }

    foreach ($layout in @($Config.layouts)) {
        foreach ($window in @($layout.windows)) {
            $processName = [string](Get-OptionalPropertyValue -Object $window -Name 'processName' -Default '')
            if ([string]::IsNullOrWhiteSpace($processName)) {
                continue
            }

            $result[$processName.Trim().ToLowerInvariant()] = $true
        }
    }

    return $result
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

function Parse-DesktopNumber {
    param([string]$Value)

    $text = if ($null -eq $Value) { '' } else { $Value.Trim() }
    if ([string]::IsNullOrWhiteSpace($text)) {
        return $null
    }

    $number = 0
    if (-not [int]::TryParse($text, [ref]$number) -or $number -lt 1) {
        throw "Desktop value '$Value' must be blank or a positive whole number."
    }

    return $number
}

function Parse-LayoutNumber {
    param([string]$Value)

    $text = if ($null -eq $Value) { '' } else { $Value.Trim() }
    if ([string]::IsNullOrWhiteSpace($text)) {
        throw 'Layout numeric values cannot be blank.'
    }

    $number = 0.0
    if (-not [double]::TryParse($text, [System.Globalization.NumberStyles]::Float, [System.Globalization.CultureInfo]::InvariantCulture, [ref]$number)) {
        throw "Invalid numeric value '$Value'."
    }

    return $number
}

function Format-LayoutNumber {
    param($Value)

    if ($null -eq $Value) {
        return ''
    }

    if ($Value -is [double] -or $Value -is [float] -or $Value -is [decimal]) {
        return ([double]$Value).ToString('0.###############', [System.Globalization.CultureInfo]::InvariantCulture)
    }

    return [string]$Value
}

function Get-NormalizedDesktopId {
    param($Value)

    if ($null -eq $Value) {
        return ''
    }

    if ($Value -is [guid]) {
        return $Value.ToString('D').ToLowerInvariant()
    }

    $text = [string]$Value
    if ([string]::IsNullOrWhiteSpace($text)) {
        return ''
    }

    $guid = [guid]::Empty
    if ([guid]::TryParse($text, [ref]$guid)) {
        return $guid.ToString('D').ToLowerInvariant()
    }

    return $text.Trim().ToLowerInvariant()
}

function ConvertTo-WindowColumns {
    param([Parameter(Mandatory = $true)][string]$Line)

    $parts = @($Line -split '\|')
    $trimmed = @($parts | ForEach-Object { $_.Trim() })
    if ($trimmed.Count -lt 9) {
        throw "Invalid window row: '$Line'"
    }

    $candidateTailLengths = @(8, 7)
    foreach ($tailLength in $candidateTailLengths) {
        if ($trimmed.Count -lt ($tailLength + 2)) {
            continue
        }

        $tailStart = $trimmed.Count - $tailLength
        if ($tailStart -lt 2) {
            continue
        }

        $title = ($trimmed[1..($tailStart - 1)] -join '|').Trim()
        $candidate = @($trimmed[0], $title)
        $candidate += $trimmed[$tailStart..($trimmed.Count - 1)]

        $matchValue = $candidate[2].ToLowerInvariant()
        $cascadeValue = $candidate[8].ToLowerInvariant()
        $desktopValue = if ($candidate.Count -ge 10) { $candidate[9] } else { '' }
        $desktopIsValid = ($candidate.Count -lt 10) -or [string]::IsNullOrWhiteSpace($desktopValue) -or ($desktopValue -match '^\d+$') -or ($desktopValue.ToLowerInvariant() -eq 'desktop')

        if (($matchValue -in @('contains', 'exact', 'regex', 'match')) -and ($cascadeValue -in @('yes', 'no', 'true', 'false', '1', '0', 'cascade')) -and $desktopIsValid) {
            return ,$candidate
        }
    }

    throw "Invalid window row: '$Line'"
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

                $columns = ConvertTo-WindowColumns -Line $line
                if (Test-HeaderRow -Columns $columns -Expected @('processName', 'title', 'match', 'x', 'y', 'w', 'h', 'monitorRole', 'cascade')) {
                    continue
                }

                if (Test-HeaderRow -Columns $columns -Expected @('processName', 'title', 'match', 'x', 'y', 'w', 'h', 'monitorRole', 'cascade', 'desktop')) {
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
                    x           = Parse-LayoutNumber -Value $columns[3]
                    y           = Parse-LayoutNumber -Value $columns[4]
                    w           = Parse-LayoutNumber -Value $columns[5]
                    h           = Parse-LayoutNumber -Value $columns[6]
                    monitorRole = $columns[7]
                    cascade     = Parse-YesNo -Value $columns[8]
                    desktop     = if ($columns.Count -ge 10) { Parse-DesktopNumber -Value $columns[9] } else { $null }
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
                    (Format-LayoutNumber (Get-OptionalPropertyValue -Object $window -Name 'x' -Default 0)),
                    (Format-LayoutNumber (Get-OptionalPropertyValue -Object $window -Name 'y' -Default 0)),
                    (Format-LayoutNumber (Get-OptionalPropertyValue -Object $window -Name 'w' -Default 100)),
                    (Format-LayoutNumber (Get-OptionalPropertyValue -Object $window -Name 'h' -Default 100)),
                    [string](Get-OptionalPropertyValue -Object $window -Name 'monitorRole' -Default ''),
                    $(if ([bool](Get-OptionalPropertyValue -Object $window -Name 'cascade' -Default $false)) { 'yes' } else { 'no' }),
                    [string](Get-OptionalPropertyValue -Object $window -Name 'desktop' -Default '')
                )
            }
        }
        $lines += Format-Table -Headers @('processName', 'title', 'match', 'x', 'y', 'w', 'h', 'monitorRole', 'cascade', 'desktop') -Rows $rows
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

            $desktop = Get-OptionalPropertyValue -Object $window -Name 'desktop' -Default $null
            if ($null -ne $desktop -and [int]$desktop -lt 1) {
                $errors.Add("Layout '$($layout.name)' has invalid desktop '$desktop'.")
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

function ConvertTo-GuidList {
    param([byte[]]$Bytes)

    if ($null -eq $Bytes -or $Bytes.Length -lt 16) {
        return @()
    }

    $guids = @()
    for ($offset = 0; $offset -le ($Bytes.Length - 16); $offset += 16) {
        $chunk = New-Object byte[] 16
        [Array]::Copy($Bytes, $offset, $chunk, 0, 16)
        $guids += ([guid]$chunk).ToString('D').ToLowerInvariant()
    }

    return $guids
}

function Get-VirtualDesktopIds {
    return @([VirtualDesktopHelper]::GetDesktopIds())
}

function Get-WindowDesktopId {
    param([IntPtr]$Handle)

    $desktopId = [VirtualDesktopHelper]::GetWindowDesktopId($Handle)
    if ($null -eq $desktopId) {
        return ''
    }

    $valueProperty = $desktopId.PSObject.Properties['Value']
    if ($null -ne $valueProperty) {
        $normalizedValue = Get-NormalizedDesktopId -Value $valueProperty.Value
        if ($normalizedValue -eq '00000000-0000-0000-0000-000000000000') {
            return ''
        }

        return $normalizedValue
    }

    $normalized = Get-NormalizedDesktopId -Value $desktopId
    if ($normalized -eq '00000000-0000-0000-0000-000000000000') {
        return ''
    }

    return $normalized
}

function Get-DesktopOrdinalFromState {
    param(
        [AllowEmptyString()][string]$DesktopId,
        [Parameter(Mandatory = $true)]$State
    )

    $normalized = Get-NormalizedDesktopId -Value $DesktopId
    if ([string]::IsNullOrWhiteSpace($normalized)) {
        return $null
    }

    if ($State.IndexById.ContainsKey($normalized)) {
        return $State.IndexById[$normalized]
    }

    return $null
}

function New-VirtualDesktopState {
    $orderedIds = New-Object System.Collections.Generic.List[string]
    $indexById = @{}

    foreach ($desktopId in @(Get-VirtualDesktopIds)) {
        $normalized = Get-NormalizedDesktopId -Value $desktopId
        if ([string]::IsNullOrWhiteSpace($normalized) -or $indexById.ContainsKey($normalized)) {
            continue
        }

        $orderedIds.Add($normalized)
        $indexById[$normalized] = $orderedIds.Count
    }

    return [pscustomobject]@{
        OrderedIds = $orderedIds
        IndexById  = $indexById
    }
}

function Resolve-VirtualDesktopId {
    param(
        [int]$DesktopNumber,
        [Parameter(Mandatory = $true)]$State
    )

    if ($DesktopNumber -lt 1 -or $DesktopNumber -gt $State.OrderedIds.Count) {
        return ''
    }

    return $State.OrderedIds[$DesktopNumber - 1]
}

function Test-IsRealDesktopId {
    param(
        [string]$DesktopId,
        [Parameter(Mandatory = $true)]$State
    )

    $normalized = Get-NormalizedDesktopId -Value $DesktopId
    if ([string]::IsNullOrWhiteSpace($normalized)) {
        return $false
    }

    return $State.IndexById.ContainsKey($normalized)
}

function Move-WindowToVirtualDesktopNumber {
    param(
        [Parameter(Mandatory = $true)]$Window,
        [int]$DesktopNumber,
        [Parameter(Mandatory = $true)]$State
    )

    if ($DesktopNumber -lt 1) {
        return $false
    }

    $desktopIdText = Resolve-VirtualDesktopId -DesktopNumber $DesktopNumber -State $State
    if ([string]::IsNullOrWhiteSpace($desktopIdText)) {
        return $false
    }

    $desktopId = [guid]::Empty
    if (-not [guid]::TryParse($desktopIdText, [ref]$desktopId)) {
        return $false
    }

    $result = [VirtualDesktopHelper]::MoveWindowToDesktopWithResult($Window.Handle, $desktopId)
    if ($result -ne 0) {
        return $false
    }

    return $true
}

function Test-WindowOnVirtualDesktopNumber {
    param(
        [Parameter(Mandatory = $true)]$Window,
        [int]$DesktopNumber,
        [Parameter(Mandatory = $true)]$State
    )

    if ($DesktopNumber -lt 1) {
        return $true
    }

    $targetDesktopId = Resolve-VirtualDesktopId -DesktopNumber $DesktopNumber -State $State
    if ([string]::IsNullOrWhiteSpace($targetDesktopId)) {
        return $false
    }

    $currentDesktopId = Ensure-WindowDesktopId -Window $Window
    if ([string]::IsNullOrWhiteSpace($currentDesktopId)) {
        return $false
    }

    return $currentDesktopId -eq $targetDesktopId
}

function Ensure-WindowDesktopId {
    param([Parameter(Mandatory = $true)]$Window)

    $currentDesktopId = Get-NormalizedDesktopId -Value (Get-OptionalPropertyValue -Object $Window -Name 'DesktopId' -Default '')
    if (-not [string]::IsNullOrWhiteSpace($currentDesktopId)) {
        return $currentDesktopId
    }

    $resolvedDesktopId = Get-WindowDesktopId -Handle $Window.Handle
    $Window | Add-Member -NotePropertyName DesktopId -NotePropertyValue $resolvedDesktopId -Force
    return $resolvedDesktopId
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

    $match = Get-MonitorSetupMatch -MonitorSetup $MonitorSetup -ActualMonitors $ActualMonitors
    return $match.Mapping
}

function Get-MonitorSetupMatch {
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

    return [pscustomobject]@{
        MonitorSetup = $MonitorSetup
        Mapping      = $mapping
        Cost         = $bestCost
    }
}

function Get-WindowProcessName {
    param(
        [IntPtr]$Handle,
        [hashtable]$Cache
    )

    $processId = [uint32]0
    [void][WindowLayoutNative]::GetWindowThreadProcessId($Handle, [ref]$processId)
    if ($processId -eq 0) {
        return $null
    }

    if ($null -ne $Cache -and $Cache.ContainsKey($processId)) {
        return $Cache[$processId]
    }

    $processName = $null
    try {
        $processName = (Get-Process -Id $processId -ErrorAction Stop).ProcessName + '.exe'
    }
    catch {
        $processName = $null
    }

    if ($null -ne $Cache) {
        $Cache[$processId] = $processName
    }

    return $processName
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
    param(
        [bool]$IncludeDesktopId = $true,
        [bool]$IgnoreBlacklist = $false,
        [hashtable]$ProtectedProcesses = @{}
    )

    $skipClasses = @('Progman', 'WorkerW', 'Shell_TrayWnd')
    $ignoredProcesses = if ($IgnoreBlacklist) { @{} } else { Get-IgnoredProcesses }
    $observations = @{}
    $desktopState = if ($IncludeDesktopId) { New-VirtualDesktopState } else { $null }
    $handles = [WindowLayoutNative]::GetTopLevelWindows()
    $processNameCache = @{}
    $current = 0
    $result = New-Object System.Collections.Generic.List[object]

    foreach ($handle in $handles) {
        $current++
        $processName = Get-WindowProcessName -Handle $handle -Cache $processNameCache
        if ([string]::IsNullOrWhiteSpace($processName)) {
            continue
        }

        $processKey = $processName.ToLowerInvariant()
        $isIgnoredProcess = $ignoredProcesses.ContainsKey($processKey) -and -not $ProtectedProcesses.ContainsKey($processKey)
        if (-not $IncludeDesktopId -and $isIgnoredProcess) {
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

        $rect = Get-WindowRectObject -Handle $handle
        if ($null -eq $rect -or $rect.Width -le 0 -or $rect.Height -le 0) {
            continue
        }

        $isVisible = [WindowLayoutNative]::IsWindowVisible($handle)
        $isCloaked = [WindowLayoutNative]::GetWindowCloakedValue($handle) -ne 0
        if (-not $isVisible -and -not $isCloaked) {
            continue
        }

        $desktopId = ''
        if ($IncludeDesktopId) {
            $desktopId = Get-WindowDesktopId -Handle $handle

            if (-not $observations.ContainsKey($processKey)) {
                $observations[$processKey] = [pscustomobject]@{
                    ProcessName        = $processName
                    RealDesktopCount   = 0
                    UnknownDesktopCount = 0
                }
            }

            if (-not (Test-IsRealDesktopId -DesktopId $desktopId -State $desktopState)) {
                $observations[$processKey].UnknownDesktopCount++
            }
            else {
                $observations[$processKey].RealDesktopCount++
            }
        }

        $window = [pscustomobject]@{
            Handle      = $handle
            Title       = $title
            ProcessName = $processName
            Rect        = $rect
            IsMinimized = [WindowLayoutNative]::IsIconic($handle)
            DesktopId   = $desktopId
        }

        if (-not $isIgnoredProcess) {
            $result.Add($window)
        }

    }

    if ($IncludeDesktopId -and -not $IgnoreBlacklist) {
        $ignoredProcesses = Update-IgnoredProcessesFromObservations -ExistingIgnored $ignoredProcesses -Observations $observations -ProtectedProcesses $ProtectedProcesses
    }

    if ($IgnoreBlacklist) {
        return @($result.ToArray())
    }

    return @($result.ToArray() | Where-Object {
        $processKey = $_.ProcessName.ToLowerInvariant()
        (-not $ignoredProcesses.ContainsKey($processKey)) -or $ProtectedProcesses.ContainsKey($processKey)
    })
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

    $xPct = [double](Get-OptionalPropertyValue -Object $WindowDefinition -Name 'x' -Default 0)
    $yPct = [double](Get-OptionalPropertyValue -Object $WindowDefinition -Name 'y' -Default 0)
    $wPct = [double](Get-OptionalPropertyValue -Object $WindowDefinition -Name 'w' -Default 100)
    $hPct = [double](Get-OptionalPropertyValue -Object $WindowDefinition -Name 'h' -Default 100)

    $left = $Monitor.X + [int][Math]::Round($Monitor.Width * $xPct / 100.0, 0)
    $top = $Monitor.Y + [int][Math]::Round($Monitor.Height * $yPct / 100.0, 0)
    $right = $Monitor.X + [int][Math]::Round($Monitor.Width * ($xPct + $wPct) / 100.0, 0)
    $bottom = $Monitor.Y + [int][Math]::Round($Monitor.Height * ($yPct + $hPct) / 100.0, 0)

    $x = $left + $OffsetX
    $y = $top + $OffsetY
    $width = $right - $left
    $height = $bottom - $top

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
    $availableWindows = @(Get-OpenWindows -IncludeDesktopId $false)
    $desktopState = New-VirtualDesktopState

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
            $desktopNumber = Get-OptionalPropertyValue -Object $windowDefinition -Name 'desktop' -Default $null
            if ([bool](Get-OptionalPropertyValue -Object $windowDefinition -Name 'cascade' -Default $false) -and $matches.Count -gt 1) {
                $offsetX = $cascadeOffset * $index
                $offsetY = -1 * $cascadeOffset * $index
            }

            if ($null -ne $desktopNumber) {
                if (-not (Test-WindowOnVirtualDesktopNumber -Window $window -DesktopNumber ([int]$desktopNumber) -State $desktopState)) {
                    if (Move-WindowToVirtualDesktopNumber -Window $window -DesktopNumber ([int]$desktopNumber) -State $desktopState) {
                        $window.DesktopId = Resolve-VirtualDesktopId -DesktopNumber ([int]$desktopNumber) -State $desktopState
                    }
                }
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

function Resolve-CaptureTarget {
    param([Parameter(Mandatory = $true)]$Config)

    $actualMonitors = @(Get-ActualMonitors)
    if ($Config.monitorSetups.Count -eq 0) {
        return Select-MonitorSetupForCapture -Config $Config
    }

    $matches = New-Object System.Collections.Generic.List[object]
    foreach ($setup in $Config.monitorSetups) {
        try {
            $match = Get-MonitorSetupMatch -MonitorSetup $setup -ActualMonitors $actualMonitors
            $matches.Add($match)
        }
        catch {
        }
    }

    if ($matches.Count -eq 0) {
        throw "No saved monitor setup matches the currently detected monitors. Run interactive capture to choose or create a monitor setup."
    }

    $bestMatch = @($matches | Sort-Object Cost, @{ Expression = { $_.MonitorSetup.name } })[0]
    Add-PendingMessage ("Using monitor setup '{0}' for capture." -f $bestMatch.MonitorSetup.name)
    return [pscustomobject]@{
        setup = $bestMatch.MonitorSetup
        isNew = $false
    }
}

function Capture-CurrentLayout {
    param(
        [Parameter(Mandatory = $true)]$Config,
        [Parameter(Mandatory = $true)]$CaptureTarget,
        [bool]$IgnoreBlacklist = $false
    )

    $actualMonitors = @(Get-ActualMonitors)
    $mapping = Get-MonitorMapping -MonitorSetup $CaptureTarget.setup -ActualMonitors $actualMonitors
    $protectedProcesses = Get-ProtectedProcessesFromConfig -Config $Config
    $windows = @(Get-OpenWindows -IgnoreBlacklist:$IgnoreBlacklist -ProtectedProcesses $protectedProcesses | Where-Object { -not $_.IsMinimized } | Sort-Object ProcessName, Title)
    $desktopState = New-VirtualDesktopState

    $capturedRows = foreach ($window in $windows) {
        $actualMonitor = Get-MonitorForWindow -Window $window -ActualMonitors $actualMonitors
        $monitorRole = $null
        foreach ($role in $mapping.Keys) {
            if ($mapping[$role].Index -eq $actualMonitor.Index) {
                $monitorRole = $role
                break
            }
        }

        [pscustomobject]@{
            processName = $window.ProcessName
            title       = $window.Title
            titleMatch  = 'contains'
            x           = Get-Percent -Value $window.Rect.Left -Minimum $actualMonitor.X -Maximum ($actualMonitor.X + $actualMonitor.Width)
            y           = Get-Percent -Value $window.Rect.Top -Minimum $actualMonitor.Y -Maximum ($actualMonitor.Y + $actualMonitor.Height)
            w           = Get-SizePercent -Size $window.Rect.Width -TotalSize $actualMonitor.Width
            h           = Get-SizePercent -Size $window.Rect.Height -TotalSize $actualMonitor.Height
            monitorRole = $monitorRole
            cascade     = $false
            desktop     = Get-DesktopOrdinalFromState -DesktopId $window.DesktopId -State $desktopState
        }
    }

    $processesWithBlankDesktop = @(
        $capturedRows |
            Group-Object -Property processName |
            Where-Object {
                $processName = [string]$_.Name
                $processKey = $processName.ToLowerInvariant()
                $hasDesktop = @($_.Group | Where-Object { $null -ne $_.desktop }).Count -gt 0
                (-not $protectedProcesses.ContainsKey($processKey)) -and (-not $hasDesktop)
            } |
            ForEach-Object { $_.Name } |
            Sort-Object -Unique
    )

    if (-not $IgnoreBlacklist -and $processesWithBlankDesktop.Count -gt 0) {
        [void](Add-IgnoredProcesses -ProcessNames $processesWithBlankDesktop)
        $capturedRows = @($capturedRows | Where-Object { $processesWithBlankDesktop -notcontains $_.processName })
        Add-PendingMessage ("Ignored processes with no usable desktop: {0}" -f ($processesWithBlankDesktop -join ', '))
    }
    elseif ($IgnoreBlacklist) {
        Add-PendingMessage 'Capture ignored processes blacklist and kept windows even when their desktop number was unavailable.'
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
                desktop     = $first.desktop
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

                $snapshot = Capture-CurrentLayout -Config $Config -CaptureTarget $target -IgnoreBlacklist:$IgnoreBlacklist
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
    $target = Resolve-CaptureTarget -Config $config
    if ($null -eq $target) {
        exit 0
    }

    $snapshot = Capture-CurrentLayout -Config $config -CaptureTarget $target -IgnoreBlacklist:$IgnoreBlacklist
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
