function Get-ConsoleWindowHandle {
    <#
    .SYNOPSIS
        Returns the handle of the current console window

    .DESCRIPTION
        Uses the Windows kernel32 GetConsoleWindow API to retrieve the handle
        of the console window attached to the current process.
        Unlike MainWindowHandle, this works reliably when PowerShell
        is launched from an external EXE or has no main window.

    .OUTPUTS
        [IntPtr]. The console window handle, or [IntPtr]::Zero if no console is attached.

    .EXAMPLE
        $handle = Get-ConsoleWindowHandle
        Set-WindowVisibility -Handle $handle -Hide

    .NOTES
        Author  : Loïc Ade
        Version : 1.0.0

        History :
        1.0.0 - 2026-03-22 - Initial version
    #>

    [CmdletBinding()]
    param()

    if (-not ([System.Management.Automation.PSTypeName]'Win32.ConsoleWindow').Type) {
        Add-Type -Name ConsoleWindow -Namespace Win32 -MemberDefinition @"
[DllImport("kernel32.dll")]
public static extern IntPtr GetConsoleWindow();
"@
    }

    return [Win32.ConsoleWindow]::GetConsoleWindow()
}
