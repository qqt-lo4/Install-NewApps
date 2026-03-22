function Get-WPFLabelStyle {
    <#
    .SYNOPSIS
        Returns WPF XAML style for themed Label controls

    .DESCRIPTION
        Generates XAML style definition for Label controls with custom theming.
        Ensures consistent text color across the application.

    .PARAMETER Colors
        Hashtable of theme colors (default: from Get-ThemedColors).

    .OUTPUTS
        [String]. XAML style definition for Label.

    .EXAMPLE
        $labelStyle = Get-WPFLabelStyle

    .NOTES
        Author  : Loïc Ade
        Version : 1.0.0
    #>

    Param(
        [hashtable]$Colors = (Get-ThemedColors)
    )

    return @"
        <Style TargetType="Label">
            <Setter Property="Foreground" Value="$($Colors.TextPrimary)"/>
        </Style>
"@
}
