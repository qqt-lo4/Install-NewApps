function New-ThemedWPFWindow {
    <#
    .SYNOPSIS
        Creates a WPF window with automatic theming support

    .DESCRIPTION
        Parses XAML and applies theme colors dynamically. Automatically detects system theme
        and applies appropriate colors, styles, and dark title bar. Simplifies creating themed
        WPF windows by handling style injection and window configuration.

    .PARAMETER XAML
        The XAML string defining the window structure.

    .PARAMETER ApplyDarkTitleBar
        If true, applies dark title bar when dark theme is detected (default: true).

    .PARAMETER Colors
        Hashtable of theme colors (default: from Get-ThemedColors).

    .OUTPUTS
        [System.Windows.Window]. The loaded WPF window with theming applied.

    .EXAMPLE
        $xaml = @"
        <Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
                Title="My Window" Height="400" Width="600">
            <Grid>
                <Button Content="Click Me" Style="{StaticResource PrimaryButtonStyle}"/>
            </Grid>
        </Window>
        "@
        $window = New-ThemedWPFWindow -XAML $xaml
        $window.ShowDialog()

    .NOTES
        Author  : Loïc Ade
        Version : 1.0.0
        Automatically injects theme styles (buttons, scrollbars, combobox, textbox, passwordbox, listbox, groupbox, label) into Window.Resources.
    #>

    [CmdletBinding()]
    Param(
        [Parameter(Mandatory=$true)]
        [string]$XAML,

        [Parameter(Mandatory=$false)]
        [bool]$ApplyDarkTitleBar = $true,

        [Parameter(Mandatory=$false)]
        [hashtable]$Colors = (Get-ThemedColors)
    )

    # Get all required styles
    $primaryButtonStyle = Get-WPFPrimaryButtonStyle -Size "Large" -Colors $Colors
    $primaryButtonSmallStyle = Get-WPFPrimaryButtonStyle -Size "Small" -Colors $Colors
    $primaryButtonInlineStyle = Get-WPFPrimaryButtonStyle -Size "Inline" -Colors $Colors
    $scrollBarStyle = Get-WPFScrollBarStyle -Colors $Colors
    $comboBoxStyle = Get-WPFComboBoxStyle -Colors $Colors
    $textBoxStyle = Get-WPFTextBoxStyle -Colors $Colors
    $passwordBoxStyle = Get-WPFPasswordBoxStyle -Colors $Colors
    $listBoxStyle = Get-WPFListBoxStyle -Colors $Colors
    $groupBoxStyle = Get-WPFGroupBoxStyle -Colors $Colors
    $labelStyle = Get-WPFLabelStyle -Colors $Colors

    $allStyles = @"
$primaryButtonStyle
$primaryButtonSmallStyle
$primaryButtonInlineStyle
$scrollBarStyle
$comboBoxStyle
$textBoxStyle
$passwordBoxStyle
$listBoxStyle
$groupBoxStyle
$labelStyle
"@

    # Replace color placeholders using literal string replacement (not regex)
    # to avoid issues with { } being interpreted as regex quantifiers
    $themedXaml = $XAML
    foreach ($key in $Colors.Keys) {
        $themedXaml = $themedXaml.Replace("{ThemeColor:$key}", $Colors[$key])
    }

    # Inject styles into Window.Resources using literal string operations
    if ($themedXaml.Contains('<Window.Resources>')) {
        # Resources section exists — inject styles right after the opening tag
        $themedXaml = $themedXaml.Replace('<Window.Resources>', "<Window.Resources>`n$allStyles")
    }
    else {
        # No Resources section — insert one right after the closing > of the <Window> tag.
        # In valid XML/XAML, attribute values cannot contain unescaped >, so the first >
        # after <Window is always the end of the opening tag.
        $windowTagStart = $themedXaml.IndexOf('<Window')
        $windowTagEnd   = $themedXaml.IndexOf('>', $windowTagStart) + 1
        $resourcesBlock = @"

    <Window.Resources>
$allStyles
    </Window.Resources>
"@
        $themedXaml = $themedXaml.Substring(0, $windowTagEnd) + $resourcesBlock + $themedXaml.Substring($windowTagEnd)
    }

    # Load the themed XAML
    Add-Type -AssemblyName PresentationFramework
    Add-Type -AssemblyName PresentationCore
    Add-Type -AssemblyName WindowsBase

    try {
        $reader = New-Object System.Xml.XmlNodeReader([xml]$themedXaml)
        $window = [Windows.Markup.XamlReader]::Load($reader)

        # Bring window to foreground on load
        $window.Add_Loaded({
            param($win, $e)
            Set-WPFWindowForeground -Window $win
        })

        # Apply dark title bar if requested and dark theme is active
        if ($ApplyDarkTitleBar -and (Get-SystemTheme)) {
            $window.Add_SourceInitialized({
                param($win, $e)
                Set-DarkTitleBar -Window $win
            })
        }

        return $window
    }
    catch {
        Write-Error "Failed to create themed WPF window: $_"
        throw
    }
}
