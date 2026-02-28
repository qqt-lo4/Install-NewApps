function Get-WPFPrimaryButtonStyle {
    <#
    .SYNOPSIS
        Returns WPF XAML style for themed primary button controls

    .DESCRIPTION
        Generates XAML style definition for primary Button controls with custom theming.
        Supports two sizes: Small (13pt font, 15x8 padding) and Large (16pt font, 30x12 padding).

    .PARAMETER Size
        Button size: "Small" or "Large" (default: Large).

    .PARAMETER Colors
        Hashtable of theme colors (default: from Get-ThemedColors).

    .OUTPUTS
        [String]. XAML style definition for primary Button.

    .EXAMPLE
        $buttonStyle = Get-WPFPrimaryButtonStyle

    .EXAMPLE
        $smallButtonStyle = Get-WPFPrimaryButtonStyle -Size "Small"

    .NOTES
        Author  : Loïc Ade
        Version : 1.0.0
    #>

    Param(
        [ValidateSet("Small", "Large", "Inline")]
        [string]$Size = "Large",
        [hashtable]$Colors = (Get-ThemedColors)
    )

    if ($Size -eq "Small") {
        $styleName = "PrimaryButtonSmallStyle"
        $fontSize = "13"
        $fontWeight = "Normal"
        $padding = "15,8"
    } elseif ($Size -eq "Inline") {
        $styleName = "PrimaryButtonInlineStyle"
        $fontSize = "13"
        $fontWeight = "Normal"
        $padding = "12,3"
    } else {
        $styleName = "PrimaryButtonStyle"
        $fontSize = "16"
        $fontWeight = "SemiBold"
        $padding = "30,12"
    }

    return @"
        <Style x:Key="$styleName" TargetType="Button">
            <Setter Property="Background" Value="$($Colors.AccentColor)"/>
            <Setter Property="Foreground" Value="White"/>
            <Setter Property="FontSize" Value="$fontSize"/>
            <Setter Property="FontWeight" Value="$fontWeight"/>
            <Setter Property="Padding" Value="$padding"/>
            <Setter Property="BorderThickness" Value="0"/>
            <Setter Property="Cursor" Value="Hand"/>
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="Button">
                        <Border x:Name="border"
                                Background="{TemplateBinding Background}"
                                CornerRadius="6"
                                Padding="{TemplateBinding Padding}">
                            <ContentPresenter HorizontalAlignment="Center" VerticalAlignment="Center"/>
                        </Border>
                        <ControlTemplate.Triggers>
                            <Trigger Property="IsMouseOver" Value="True">
                                <Setter TargetName="border" Property="Background" Value="$($Colors.AccentHover)"/>
                            </Trigger>
                            <Trigger Property="IsEnabled" Value="False">
                                <Setter TargetName="border" Property="Background" Value="$($Colors.ButtonDisabledBg)"/>
                                <Setter Property="Foreground" Value="$($Colors.ButtonDisabledFg)"/>
                            </Trigger>
                        </ControlTemplate.Triggers>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
        </Style>
"@
}
