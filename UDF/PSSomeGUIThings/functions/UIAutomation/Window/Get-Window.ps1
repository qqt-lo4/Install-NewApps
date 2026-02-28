function Get-Window {
    <#
    .SYNOPSIS
        Gets windows in the current Windows session

    .DESCRIPTION
        Enumerates top-level windows with optional filtering by process name,
        process ID, window class name, or title. Can also include windows from
        child and grandchild processes using -IncludeChildren.
        Returns window handle, process ID, process name, class name, title,
        visibility state, and position/size.

    .PARAMETER IncludeHidden
        If specified, includes hidden windows in the results.

    .PARAMETER ProcessName
        Filters windows by process name (supports wildcards).

    .PARAMETER ProcessId
        Filters windows by process ID. Defaults to all processes.

    .PARAMETER ClassName
        Filters windows by class name (supports wildcards, e.g. "Chrome_WidgetWin_*").

    .PARAMETER Title
        Filters windows by title (supports wildcards, e.g. "*Google*").

    .PARAMETER ExcludeEmptyTitle
        If specified, excludes windows with empty titles.

    .PARAMETER IncludeChildren
        If specified with -ProcessId, also includes windows from child/grandchild processes.

    .PARAMETER ChildDepth
        How many levels of child processes to include. Default is 2. Only used with -IncludeChildren.

    .OUTPUTS
        Array of PSCustomObject with properties: Handle, ProcessId, ProcessName, ClassName, Title, IsVisible, IsMinimized, IsMaximized, IsChildProcess, Left, Top, Width, Height.

    .EXAMPLE
        Get-Window
        Gets all visible windows.

    .EXAMPLE
        Get-Window -IncludeHidden
        Gets all windows including hidden ones.

    .EXAMPLE
        Get-Window -ProcessName chrome
        Gets all Chrome windows.

    .EXAMPLE
        Get-Window -ProcessName *cord*
        Gets all windows from processes containing 'cord' (e.g., Discord).

    .EXAMPLE
        Get-Window -Title "*Notepad*"
        Gets all windows with 'Notepad' in the title.

    .EXAMPLE
        Get-Window -ExcludeEmptyTitle
        Gets all visible windows that have a title.

    .EXAMPLE
        Get-Window -ProcessId $PID -IncludeChildren
        Gets windows from the current PowerShell process and its child processes.

    .EXAMPLE
        Get-Window -ProcessId $PID -IncludeChildren -ClassName "Chrome_WidgetWin_*"
        Gets Chrome windows launched from this PowerShell session.

    .NOTES
        Author  : Loïc Ade
        Version : 1.1.0
        Uses P/Invoke to call Windows API functions for window enumeration.
        Requires Get-ProcessChildProcess (PSSomeSystemThings) when using -IncludeChildren.

        Version History:
            1.0.0 - Initial version
            1.1.0 - Added ClassName, ProcessId, Title, ExcludeEmptyTitle, IncludeChildren, ChildDepth parameters and IsVisible, IsMinimized, IsMaximized, IsChildProcess, Left, Top, Width, Height output
    #>

    [CmdletBinding()]
    Param(
        [Parameter(Mandatory=$false)]
        [switch]$IncludeHidden,

        [Parameter(Mandatory=$false)]
        [string]$ProcessName,

        [Parameter(Mandatory=$false)]
        [uint32]$ProcessId,

        [Parameter(Mandatory=$false)]
        [string]$ClassName,

        [Parameter(Mandatory=$false)]
        [string]$Title,

        [Parameter(Mandatory=$false)]
        [switch]$ExcludeEmptyTitle,

        [Parameter(Mandatory=$false)]
        [switch]$IncludeChildren,

        [Parameter(Mandatory=$false)]
        [int]$ChildDepth = 2
    )

    # Define Windows API functions
    Add-Type @"
using System;
using System.Runtime.InteropServices;
using System.Text;
using System.Collections.Generic;

public class WindowEnumerator {
    [DllImport("user32.dll")]
    [return: MarshalAs(UnmanagedType.Bool)]
    public static extern bool EnumWindows(EnumWindowsProc lpEnumFunc, IntPtr lParam);

    [DllImport("user32.dll")]
    [return: MarshalAs(UnmanagedType.Bool)]
    public static extern bool IsWindowVisible(IntPtr hWnd);

    [DllImport("user32.dll", CharSet = CharSet.Auto, SetLastError = true)]
    public static extern int GetWindowText(IntPtr hWnd, StringBuilder lpString, int nMaxCount);

    [DllImport("user32.dll", SetLastError = true)]
    public static extern int GetWindowTextLength(IntPtr hWnd);

    [DllImport("user32.dll", SetLastError = true)]
    public static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint lpdwProcessId);

    [DllImport("user32.dll", CharSet = CharSet.Auto)]
    public static extern int GetClassName(IntPtr hWnd, StringBuilder lpClassName, int nMaxCount);

    [DllImport("user32.dll")]
    [return: MarshalAs(UnmanagedType.Bool)]
    public static extern bool GetWindowRect(IntPtr hWnd, out RECT lpRect);

    [DllImport("user32.dll")]
    [return: MarshalAs(UnmanagedType.Bool)]
    public static extern bool IsIconic(IntPtr hWnd);

    [DllImport("user32.dll")]
    [return: MarshalAs(UnmanagedType.Bool)]
    public static extern bool IsZoomed(IntPtr hWnd);

    [StructLayout(LayoutKind.Sequential)]
    public struct RECT {
        public int Left;
        public int Top;
        public int Right;
        public int Bottom;
    }

    public delegate bool EnumWindowsProc(IntPtr hWnd, IntPtr lParam);

    public static List<WindowInfo> GetWindows(bool includeHidden) {
        List<WindowInfo> windows = new List<WindowInfo>();

        EnumWindows(delegate(IntPtr hWnd, IntPtr lParam) {
            bool isVisible = IsWindowVisible(hWnd);

            // Check if window is visible (unless includeHidden is true)
            if (!includeHidden && !isVisible) {
                return true;
            }

            // Get window text length
            int length = GetWindowTextLength(hWnd);

            // Get window title
            StringBuilder sb = new StringBuilder(length + 1);
            GetWindowText(hWnd, sb, sb.Capacity);
            string title = sb.ToString();

            // Get class name
            StringBuilder classSb = new StringBuilder(256);
            GetClassName(hWnd, classSb, classSb.Capacity);
            string className = classSb.ToString();

            // Get process ID
            uint processId;
            GetWindowThreadProcessId(hWnd, out processId);

            // Get window rectangle
            RECT rect;
            GetWindowRect(hWnd, out rect);

            // Get window state
            bool isMinimized = IsIconic(hWnd);
            bool isMaximized = IsZoomed(hWnd);

            // Add to list
            windows.Add(new WindowInfo {
                Handle = hWnd,
                ProcessId = processId,
                ClassName = className,
                Title = title,
                IsVisible = isVisible,
                IsMinimized = isMinimized,
                IsMaximized = isMaximized,
                Left = rect.Left,
                Top = rect.Top,
                Width = rect.Right - rect.Left,
                Height = rect.Bottom - rect.Top
            });

            return true;
        }, IntPtr.Zero);

        return windows;
    }
}

public class WindowInfo {
    public IntPtr Handle { get; set; }
    public uint ProcessId { get; set; }
    public string ClassName { get; set; }
    public string Title { get; set; }
    public bool IsVisible { get; set; }
    public bool IsMinimized { get; set; }
    public bool IsMaximized { get; set; }
    public int Left { get; set; }
    public int Top { get; set; }
    public int Width { get; set; }
    public int Height { get; set; }
}
"@

    # Build target PIDs map if ProcessId filtering is requested
    # Key = PID, Value = $true if child process, $false if direct match
    $targetPids = $null
    if ($ProcessId) {
        $targetPids = @{ $ProcessId = $false }

        if ($IncludeChildren) {
            $childProcesses = Get-ProcessChildProcess -ProcessId $ProcessId -Depth $ChildDepth
            foreach ($proc in $childProcesses) {
                $targetPids[$proc.Id] = $true
                Write-Verbose "Child process: PID=$($proc.Id), Name=$($proc.ProcessName)"
            }
        }
    }

    # Get all windows
    $windows = [WindowEnumerator]::GetWindows($IncludeHidden.IsPresent)

    # Convert to PowerShell objects with process information
    $result = @()
    foreach ($window in $windows) {
        try {
            # Filter by target PIDs if specified
            if ($targetPids -and -not $targetPids.ContainsKey($window.ProcessId)) {
                continue
            }

            # Filter by class name if specified
            if ($ClassName -and $window.ClassName -notlike $ClassName) {
                continue
            }

            # Filter by title if specified
            if ($Title -and $window.Title -notlike $Title) {
                continue
            }

            # Exclude empty titles if requested
            if ($ExcludeEmptyTitle -and [string]::IsNullOrEmpty($window.Title)) {
                continue
            }

            $process = Get-Process -Id $window.ProcessId -ErrorAction SilentlyContinue
            $procName = if ($process) { $process.ProcessName } else { "Unknown" }

            # Filter by process name if specified
            if ($ProcessName -and $procName -notlike $ProcessName) {
                continue
            }

            # Determine IsChildProcess
            $isChild = if ($targetPids) { $targetPids[$window.ProcessId] } else { $false }

            $result += [PSCustomObject]@{
                Handle         = $window.Handle
                ProcessId      = $window.ProcessId
                ProcessName    = $procName
                ClassName      = $window.ClassName
                Title          = $window.Title
                IsVisible      = $window.IsVisible
                IsMinimized    = $window.IsMinimized
                IsMaximized    = $window.IsMaximized
                IsChildProcess = $isChild
                Left           = $window.Left
                Top            = $window.Top
                Width          = $window.Width
                Height         = $window.Height
            }
        }
        catch {
            # Skip windows from processes we can't access
            Write-Verbose "Could not get process info for PID $($window.ProcessId): $_"
        }
    }

    return $result
}
