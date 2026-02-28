function Get-WPFPasswordBoxStyle {
    <#
    .SYNOPSIS
        Returns WPF XAML style for themed PasswordBox controls

    .DESCRIPTION
        Wrapper around Get-WPFTextBoxStyle -TargetType "PasswordBox".
        Both controls share the same visual template.

    .PARAMETER Colors
        Hashtable of theme colors (default: from Get-ThemedColors).

    .OUTPUTS
        [String]. XAML style definition for PasswordBox.

    .EXAMPLE
        $passwordBoxStyle = Get-WPFPasswordBoxStyle

    .NOTES
        Author  : Loïc Ade
        Version : 1.0.0
    #>

    Param(
        [hashtable]$Colors = (Get-ThemedColors)
    )

    return Get-WPFTextBoxStyle -TargetType "PasswordBox" -Colors $Colors
}
