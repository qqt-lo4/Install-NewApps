function Close-LoadingWindow {
    <#
    .SYNOPSIS
        Closes the loading window

    .DESCRIPTION
        Closes a loading window created by Show-LoadingWindow and cleans up resources.

    .PARAMETER Window
        The synchronized hashtable returned by Show-LoadingWindow

    .EXAMPLE
        Close-LoadingWindow -Window $loadingWindow

    .NOTES
        Author  : Loïc Ade
        Version : 1.0.0
    #>

    [CmdletBinding()]
    Param(
        [Parameter(Mandatory=$false)]
        $Window
    )

    if ($Window) {
        # Close the WPF window
        if ($Window.Window) {
            try {
                $Window.Window.Dispatcher.Invoke([Action]{
                    $Window.Window.Close()
                }, [System.Windows.Threading.DispatcherPriority]::Normal)

                Start-Sleep -Milliseconds 100
            }
            catch {
                Write-Warning "Failed to close loading window: $_"
            }
        }

        # Dispose PowerShell command
        if ($Window.PowerShell) {
            try {
                if ($Window.Handle) {
                    $Window.PowerShell.EndInvoke($Window.Handle)
                }
                $Window.PowerShell.Dispose()
            }
            catch {
                Write-Verbose "Failed to dispose PowerShell command: $_"
            }
        }

        # Close and dispose runspace
        if ($Window.Runspace) {
            try {
                $Window.Runspace.Close()
                $Window.Runspace.Dispose()
            }
            catch {
                Write-Verbose "Failed to dispose runspace: $_"
            }
        }
    }
}
