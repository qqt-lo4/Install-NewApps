function Wait-MainWindowFromProcessClose {
    <#
    .SYNOPSIS
        Waits for a process's main window to close

    .DESCRIPTION
        Monitors a process and blocks execution until its main window handle becomes zero (window closed).
        Polls the process every 50 milliseconds to check if the main window is closed.

    .PARAMETER processName
        The name of the process to monitor (without .exe extension).

    .OUTPUTS
        None

    .EXAMPLE
        Wait-MainWindowFromProcessClose -processName "notepad"

    .NOTES
        Author  : Loïc Ade
        Version : 1.0.0
    #>

    Param(
        [string]$processName
    )
    [System.Diagnostics.Process]$process = Get-Process -Name $processName
    while ($process.MainWindowHandle -ne [IntPtr]::Zero) {
        $process = Get-Process -Name $processName
        Start-Sleep -Milliseconds 50
    }
}
