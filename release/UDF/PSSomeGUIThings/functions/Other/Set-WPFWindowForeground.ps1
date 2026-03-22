function Set-WPFWindowForeground {
    <#
    .SYNOPSIS
        Brings a WPF window to the foreground using Windows API

    .DESCRIPTION
        This function activates a WPF window and forces it to the foreground using the Windows API SetForegroundWindow.
        Unlike setting Topmost = $true, this provides a one-time activation without keeping the window always on top.

    .PARAMETER Window
        The WPF Window object to bring to foreground

    .EXAMPLE
        $window = New-Object System.Windows.Window
        $window.Show()
        Set-WPFWindowForeground -Window $window

    .EXAMPLE
        # Use in Window.Add_Loaded event
        $window.Add_Loaded({
            param($win, $e)
            Set-WPFWindowForeground -Window $win
        })

    .NOTES
        Author  : Loïc Ade
        Version : 1.0.0
        This function uses P/Invoke to call the Windows API SetForegroundWindow function.
        Falls back gracefully to WPF's Activate() method if the API call fails.
    #>

    [CmdletBinding()]
    Param(
        [Parameter(Mandatory=$true, ValueFromPipeline=$true)]
        [System.Windows.Window]$Window
    )

    Process {
        try {
            # First, use WPF's built-in activation
            $Window.Activate()

            # Then force to foreground using Windows API
            $windowHandle = (New-Object System.Windows.Interop.WindowInteropHelper($Window)).Handle

            if ($windowHandle -ne [System.IntPtr]::Zero) {
                # Show and restore window
                Set-WindowVisibility -Handle $windowHandle -Show | Out-Null
                Set-WindowVisibility -Handle $windowHandle -Restore | Out-Null

                # Bring to top and set foreground
                Set-WindowOnTop -Handle $windowHandle | Out-Null
                Set-WindowForeground -Handle $windowHandle | Out-Null

                Write-Verbose "Window successfully brought to foreground"
            }
            else {
                Write-Warning "Unable to get window handle, falling back to Activate() only"
            }
        }
        catch {
            Write-Warning "Error bringing window to foreground: $_. Falling back to Activate() only"
        }
    }
}
