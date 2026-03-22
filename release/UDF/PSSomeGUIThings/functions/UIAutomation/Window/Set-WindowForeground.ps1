function Set-WindowForeground {
    <#
    .SYNOPSIS
        Activates a window and brings it to the foreground

    .DESCRIPTION
        Forces a window to the foreground using the Windows API SetForegroundWindow.
        Uses the AttachThreadInput trick to bypass Windows restrictions that prevent
        background processes from stealing focus.

    .PARAMETER Handle
        The window handle (IntPtr) to bring to the foreground.

    .OUTPUTS
        [bool]. True if the window was successfully brought to the foreground.

    .EXAMPLE
        Set-WindowForeground -Handle $hwnd
        Brings the window to the foreground and gives it focus.

    .EXAMPLE
        Get-Window -Title "*Notepad*" | ForEach-Object { Set-WindowForeground -Handle $_.Handle }
        Brings the first Notepad window to the foreground.

    .NOTES
        Author  : Loïc Ade
        Version : 1.0.0
    #>

    [CmdletBinding()]
    param(
        [Parameter(Mandatory, ValueFromPipelineByPropertyName)]
        [IntPtr]$Handle
    )

    process {
        # Import Windows API if not already loaded
        if (-not ([System.Management.Automation.PSTypeName]'Win32.WindowForeground').Type) {
            Add-Type -TypeDefinition @"
using System;
using System.Runtime.InteropServices;

namespace Win32 {
    public class WindowForeground {
        [DllImport("user32.dll")]
        [return: MarshalAs(UnmanagedType.Bool)]
        public static extern bool SetForegroundWindow(IntPtr hWnd);

        [DllImport("user32.dll")]
        public static extern IntPtr GetForegroundWindow();

        [DllImport("user32.dll")]
        public static extern uint GetWindowThreadProcessId(IntPtr hWnd, IntPtr ProcessId);

        [DllImport("user32.dll")]
        [return: MarshalAs(UnmanagedType.Bool)]
        public static extern bool AttachThreadInput(uint idAttach, uint idAttachTo, bool fAttach);

        [DllImport("kernel32.dll")]
        public static extern uint GetCurrentThreadId();
    }
}
"@
        }

        # Get current foreground window and thread
        $foregroundWindow = [Win32.WindowForeground]::GetForegroundWindow()
        $currentThreadId = [Win32.WindowForeground]::GetCurrentThreadId()
        $foregroundThreadId = [Win32.WindowForeground]::GetWindowThreadProcessId($foregroundWindow, [IntPtr]::Zero)

        # Attach to foreground thread to bypass SetForegroundWindow restrictions
        $attached = $false
        if ($foregroundThreadId -ne 0 -and $foregroundThreadId -ne $currentThreadId) {
            $attached = [Win32.WindowForeground]::AttachThreadInput($currentThreadId, $foregroundThreadId, $true)
        }

        try {
            $success = [Win32.WindowForeground]::SetForegroundWindow($Handle)

            if ($success) {
                Write-Verbose "Window $Handle brought to foreground"
            } else {
                Write-Verbose "Window $Handle`: SetForegroundWindow failed"
            }

            return $success
        }
        finally {
            if ($attached) {
                [Win32.WindowForeground]::AttachThreadInput($currentThreadId, $foregroundThreadId, $false) | Out-Null
            }
        }
    }
}
