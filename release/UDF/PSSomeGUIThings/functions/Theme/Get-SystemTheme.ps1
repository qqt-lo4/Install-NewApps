function Get-SystemTheme {
    <#
    .SYNOPSIS
        Detects if Windows is using dark theme

    .DESCRIPTION
        Reads the Windows registry to determine if the system is using dark theme mode.
        Checks the AppsUseLightTheme registry value in the Personalize settings.

    .OUTPUTS
        [Boolean]. Returns $true if dark theme is enabled, $false otherwise.

    .EXAMPLE
        $isDarkTheme = Get-SystemTheme

    .NOTES
        Author  : Loïc Ade
        Version : 1.0.0
    #>

    try {
        $regPath = "HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\Themes\Personalize"
        $appsUseLightTheme = Get-ItemProperty -Path $regPath -Name "AppsUseLightTheme" -ErrorAction SilentlyContinue

        if ($null -eq $appsUseLightTheme) {
            return $false  # Default to light theme if registry key doesn't exist
        }

        return $appsUseLightTheme.AppsUseLightTheme -eq 0
    }
    catch {
        return $false  # Default to light theme on error
    }
}
