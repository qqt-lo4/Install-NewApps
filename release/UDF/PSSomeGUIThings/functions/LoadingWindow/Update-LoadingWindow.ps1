function Update-LoadingWindow {
    <#
    .SYNOPSIS
        Updates the loading window message and progress

    .DESCRIPTION
        Updates the message text and/or progress bar value of a loading window
        created by Show-LoadingWindow.

    .PARAMETER Window
        The synchronized hashtable returned by Show-LoadingWindow

    .PARAMETER Message
        New message to display

    .PARAMETER Progress
        Progress value (0-100). If not specified, keeps indeterminate mode

    .EXAMPLE
        Update-LoadingWindow -Window $loadingWindow -Message "Downloading package 1/10" -Progress 10

    .NOTES
        Author  : Loïc Ade
        Version : 1.0.0
    #>

    [CmdletBinding()]
    Param(
        [Parameter(Mandatory=$true)]
        $Window,

        [Parameter(Mandatory=$false)]
        [string]$Message,

        [Parameter(Mandatory=$false)]
        [int]$Progress = -1,

        [Parameter(Mandatory=$false)]
        [string]$Details,

        [Parameter(Mandatory=$false)]
        [bool]$AutoScrollDetails = $true
    )

    if (-not $Window -or -not $Window.Window) {
        return
    }

    try {
        $Window.Window.Dispatcher.Invoke([Action]{
            if ($Message) {
                $Window.MessageText.Text = $Message
            }

            if ($Progress -ge 0) {
                $Window.ProgressBar.IsIndeterminate = $false
                $Window.ProgressBar.Value = $Progress
            }

            # Update details text if provided (only if changed to avoid scroll flicker)
            if ($Details -and $Window.DetailsTextBox) {
                if ($Window.DetailsTextBox.Text -ne $Details) {
                    $Window.DetailsTextBox.Text = $Details

                    # Auto-scroll to bottom if enabled
                    if ($AutoScrollDetails) {
                        $Window.DetailsTextBox.ScrollToEnd()
                    }
                }
            }
        }, [System.Windows.Threading.DispatcherPriority]::Normal)
    }
    catch {
        Write-Warning "Failed to update loading window: $_"
    }
}
