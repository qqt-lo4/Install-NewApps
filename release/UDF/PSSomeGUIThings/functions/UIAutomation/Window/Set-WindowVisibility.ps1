function Set-WindowVisibility {
    <#
    .SYNOPSIS
        Shows, hides, minimizes, maximizes, or restores a window by its handle

    .DESCRIPTION
        Controls the visibility of a window using the Windows User32 ShowWindow API.
        Accepts a window handle (IntPtr) and applies the specified visibility action.

    .PARAMETER Handle
        The window handle (IntPtr) to control.

    .PARAMETER Hide
        Hides the window completely.

    .PARAMETER Show
        Shows the window (restores if minimized/hidden).

    .PARAMETER Minimize
        Minimizes the window to the taskbar.

    .PARAMETER Maximize
        Maximizes the window.

    .PARAMETER Restore
        Restores the window to its previous size and position (before minimize/maximize).

    .OUTPUTS
        [bool]. True if the action succeeded, False otherwise.

    .EXAMPLE
        Set-WindowVisibility -Handle $hwnd -Hide
        Hides the window.

    .EXAMPLE
        Set-WindowVisibility -Handle $hwnd -Show
        Shows the window.

    .EXAMPLE
        Set-WindowVisibility -Handle $hwnd -Restore
        Restores a minimized or maximized window to its normal state.

    .EXAMPLE
        Get-Window -ProcessName notepad | ForEach-Object { Set-WindowVisibility -Handle $_.Handle -Minimize }
        Minimizes all Notepad windows.

    .NOTES
        Author  : Loïc Ade
        Version : 1.0.0
    #>

    [CmdletBinding()]
    param(
        [Parameter(Mandatory, ValueFromPipelineByPropertyName)]
        [IntPtr]$Handle,

        [Parameter(Mandatory, ParameterSetName = 'Hide')]
        [switch]$Hide,

        [Parameter(Mandatory, ParameterSetName = 'Show')]
        [switch]$Show,

        [Parameter(Mandatory, ParameterSetName = 'Minimize')]
        [switch]$Minimize,

        [Parameter(Mandatory, ParameterSetName = 'Maximize')]
        [switch]$Maximize,

        [Parameter(Mandatory, ParameterSetName = 'Restore')]
        [switch]$Restore
    )

    process {
        # Windows API constants
        $SW_HIDE = 0
        $SW_MAXIMIZE = 3
        $SW_SHOW = 5
        $SW_MINIMIZE = 6
        $SW_RESTORE = 9

        # Import Windows API if not already loaded
        if (-not ([System.Management.Automation.PSTypeName]'Win32.WindowVisibility').Type) {
            Add-Type -TypeDefinition @"
using System;
using System.Runtime.InteropServices;

namespace Win32 {
    public class WindowVisibility {
        [DllImport("user32.dll")]
        public static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);
    }
}
"@
        }

        $cmdShow = switch ($PSCmdlet.ParameterSetName) {
            'Hide'     { $SW_HIDE }
            'Show'     { $SW_SHOW }
            'Minimize' { $SW_MINIMIZE }
            'Maximize' { $SW_MAXIMIZE }
            'Restore'  { $SW_RESTORE }
        }

        $success = [Win32.WindowVisibility]::ShowWindow($Handle, $cmdShow)

        if ($success) {
            Write-Verbose "Window $Handle`: $($PSCmdlet.ParameterSetName) succeeded"
        } else {
            Write-Verbose "Window $Handle`: $($PSCmdlet.ParameterSetName) failed"
        }

        return $success
    }
}
