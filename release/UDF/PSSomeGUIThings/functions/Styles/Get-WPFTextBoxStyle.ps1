function Get-WPFTextBoxStyle {
    <#
    .SYNOPSIS
        Returns WPF XAML style for themed TextBox or PasswordBox controls

    .DESCRIPTION
        Generates XAML style definition for TextBox or PasswordBox controls with custom theming.
        Both controls share the same visual template (rounded border, hover/focus highlight).
        Includes AccentHover on mouse-over and AccentColor on focus.

    .PARAMETER TargetType
        Control type to style: "TextBox" or "PasswordBox" (default: TextBox).

    .PARAMETER Colors
        Hashtable of theme colors (default: from Get-ThemedColors).

    .OUTPUTS
        [String]. XAML style definition for the specified control type.

    .EXAMPLE
        $textBoxStyle = Get-WPFTextBoxStyle

    .EXAMPLE
        $passwordBoxStyle = Get-WPFTextBoxStyle -TargetType "PasswordBox"

    .NOTES
        Author  : Loïc Ade
        Version : 1.0.0
    #>

    Param(
        [ValidateSet("TextBox", "PasswordBox")]
        [string]$TargetType = "TextBox",

        [hashtable]$Colors = (Get-ThemedColors)
    )

    $baseStyle = @"
        <Style TargetType="$TargetType">
            <Setter Property="Background" Value="$($Colors.CardBackground)"/>
            <Setter Property="Foreground" Value="$($Colors.TextPrimary)"/>
            <Setter Property="BorderBrush" Value="$($Colors.BorderColor)"/>
            <Setter Property="BorderThickness" Value="1"/>
            <Setter Property="Padding" Value="5,5"/>
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="$TargetType">
                        <Border x:Name="border"
                                Background="{TemplateBinding Background}"
                                BorderBrush="{TemplateBinding BorderBrush}"
                                BorderThickness="{TemplateBinding BorderThickness}"
                                CornerRadius="4">
                            <ScrollViewer x:Name="PART_ContentHost"
                                         Focusable="false"
                                         HorizontalScrollBarVisibility="Hidden"
                                         VerticalScrollBarVisibility="Hidden"
                                         Margin="0"
                                         Padding="{TemplateBinding Padding}"/>
                        </Border>
                        <ControlTemplate.Triggers>
                            <Trigger Property="IsMouseOver" Value="True">
                                <Setter TargetName="border" Property="BorderBrush" Value="$($Colors.InputBorderHover)"/>
                            </Trigger>
                            <Trigger Property="IsFocused" Value="True">
                                <Setter TargetName="border" Property="BorderBrush" Value="$($Colors.InputBorderFocus)"/>
                            </Trigger>
                        </ControlTemplate.Triggers>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
        </Style>
"@

    if ($TargetType -eq "TextBox") {
        $squareTopVariant = @"

        <Style x:Key="TextBoxStyleSquareTop" TargetType="TextBox">
            <Setter Property="Background" Value="$($Colors.CardBackground)"/>
            <Setter Property="Foreground" Value="$($Colors.TextPrimary)"/>
            <Setter Property="BorderBrush" Value="$($Colors.BorderColor)"/>
            <Setter Property="BorderThickness" Value="1"/>
            <Setter Property="Padding" Value="5,5"/>
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="TextBox">
                        <Border x:Name="border"
                                Background="{TemplateBinding Background}"
                                BorderBrush="{TemplateBinding BorderBrush}"
                                BorderThickness="{TemplateBinding BorderThickness}"
                                CornerRadius="0,0,4,4">
                            <ScrollViewer x:Name="PART_ContentHost"
                                         Focusable="false"
                                         HorizontalScrollBarVisibility="Hidden"
                                         VerticalScrollBarVisibility="Hidden"
                                         Margin="0"
                                         Padding="{TemplateBinding Padding}"/>
                        </Border>
                        <ControlTemplate.Triggers>
                            <Trigger Property="IsMouseOver" Value="True">
                                <Setter TargetName="border" Property="BorderBrush" Value="$($Colors.InputBorderHover)"/>
                            </Trigger>
                            <Trigger Property="IsFocused" Value="True">
                                <Setter TargetName="border" Property="BorderBrush" Value="$($Colors.InputBorderFocus)"/>
                            </Trigger>
                        </ControlTemplate.Triggers>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
        </Style>
"@
        return $baseStyle + $squareTopVariant
    }

    return $baseStyle
}
