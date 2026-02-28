function Wait-Window {
    <#
    .SYNOPSIS
        Waits for a window to appear for a specific process

    .DESCRIPTION
        Polls until a visible window appears for the specified process name.
        Uses Get-Window to check for visible windows.

    .PARAMETER ProcessName
        The name of the process to wait for (without .exe extension)

    .PARAMETER TimeoutSeconds
        Maximum time to wait in seconds (default: 30)

    .OUTPUTS
        Returns the window object when found, or $null if timeout is reached

    .EXAMPLE
        Wait-Window -ProcessName Discord
        Waits for a Discord window to appear

    .EXAMPLE
        Wait-Window -ProcessName chrome -TimeoutSeconds 60
        Waits up to 60 seconds for a Chrome window

    .NOTES
        Author  : Loïc Ade
        Version : 1.0.0
        Requires Get-Window function.
    #>

    [CmdletBinding()]
    Param(
        [Parameter(Mandatory=$true)]
        [string]$ProcessName,

        [Parameter(Mandatory=$false)]
        [int]$TimeoutSeconds = 30
    )

    # Import Get-Window if not already loaded
    $getWindowPath = Join-Path (Split-Path $PSScriptRoot -Parent) "Window\Get-WindowList.ps1"
    if (Test-Path $getWindowPath) {
        . $getWindowPath
    }

    $startTime = Get-Date
    $timeout = New-TimeSpan -Seconds $TimeoutSeconds

    while ((Get-Date) - $startTime -lt $timeout) {
        # Check if a visible window exists for this process
        $window = Get-Window -ProcessName $ProcessName

        if ($window) {
            Write-Verbose "Window found for process '$ProcessName': $($window.Title)"
            return $window | Select-Object -First 1
        }

        Start-Sleep -Milliseconds 100
    }

    Write-Warning "Timeout reached waiting for window from process '$ProcessName'"
    return $null
}