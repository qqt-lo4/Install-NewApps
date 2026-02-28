function Set-DarkTitleBar {
    <#
    .SYNOPSIS
        Applies dark title bar to a WPF window

    .DESCRIPTION
        Sets the dark immersive mode for a WPF window's title bar using the Desktop Window Manager API.
        This makes the title bar use dark theme colors on Windows 10/11.

    .PARAMETER Window
        The WPF Window object to apply dark title bar to.

    .OUTPUTS
        None

    .EXAMPLE
        $window = New-Object System.Windows.Window
        $window.Show()
        Set-DarkTitleBar -Window $window

    .NOTES
        Author  : Loïc Ade
        Version : 1.0.0
        Uses P/Invoke to call DwmSetWindowAttribute from dwmapi.dll.
    #>

    Param(
        [Parameter(Mandatory)]
        [System.Windows.Window]$Window
    )

    try {
        $windowHelper = New-Object System.Windows.Interop.WindowInteropHelper($Window)
        $hwnd = $windowHelper.Handle

        # DWMWA_USE_IMMERSIVE_DARK_MODE = 20
        $DWMWA_USE_IMMERSIVE_DARK_MODE = 20
        $useImmersiveDarkMode = 1

        $dwmapi = Add-Type -MemberDefinition @"
[DllImport("dwmapi.dll", PreserveSig = true)]
public static extern int DwmSetWindowAttribute(IntPtr hwnd, int attr, ref int attrValue, int attrSize);
"@ -Name "DwmApi" -Namespace "Win32" -PassThru -ErrorAction SilentlyContinue

        if ($dwmapi) {
            [Win32.DwmApi]::DwmSetWindowAttribute($hwnd, $DWMWA_USE_IMMERSIVE_DARK_MODE, [ref]$useImmersiveDarkMode, 4) | Out-Null
        }
    }
    catch {
        Write-Verbose "Failed to apply dark title bar: $_"
    }
}
