function Set-WindowOnTop {
    <#
    .SYNOPSIS
        Brings a window to the top of the Z-order

    .DESCRIPTION
        Moves a window to the top of the Z-order using the Windows API BringWindowToTop.
        Unlike Set-WindowForeground, this does not activate the window or give it focus.

    .PARAMETER Handle
        The window handle (IntPtr) to bring to the top.

    .OUTPUTS
        [bool]. True if the window was successfully brought to the top.

    .EXAMPLE
        Set-WindowOnTop -Handle $hwnd
        Brings the window to the top of the Z-order.

    .EXAMPLE
        Get-Window -ProcessName notepad | ForEach-Object { Set-WindowOnTop -Handle $_.Handle }
        Brings all Notepad windows to the top.

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
        if (-not ([System.Management.Automation.PSTypeName]'Win32.WindowZOrder').Type) {
            Add-Type -TypeDefinition @"
using System;
using System.Runtime.InteropServices;

namespace Win32 {
    public class WindowZOrder {
        [DllImport("user32.dll")]
        [return: MarshalAs(UnmanagedType.Bool)]
        public static extern bool BringWindowToTop(IntPtr hWnd);
    }
}
"@
        }

        $success = [Win32.WindowZOrder]::BringWindowToTop($Handle)

        if ($success) {
            Write-Verbose "Window $Handle brought to top"
        } else {
            Write-Verbose "Window $Handle`: BringWindowToTop failed"
        }

        return $success
    }
}
