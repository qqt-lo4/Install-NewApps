function Get-WPFListBoxStyle {
    <#
    .SYNOPSIS
        Returns WPF XAML style for themed ListBox controls

    .DESCRIPTION
        Generates XAML style definitions for ListBox and ListBoxItem controls with custom theming.
        Includes hover and selection effects with smooth visual transitions.

    .PARAMETER Colors
        Hashtable of theme colors (default: from Get-ThemedColors).

    .OUTPUTS
        [String]. XAML style definitions for ListBox and ListBoxItem.

    .EXAMPLE
        $listBoxStyle = Get-WPFListBoxStyle

    .NOTES
        Author  : Loïc Ade
        Version : 1.0.0
    #>

    Param(
        [hashtable]$Colors = (Get-ThemedColors)
    )

    return @"
        <Style TargetType="ListBox">
            <Setter Property="Background" Value="$($Colors.CardBackground)"/>
            <Setter Property="Foreground" Value="$($Colors.TextPrimary)"/>
            <Setter Property="BorderBrush" Value="$($Colors.BorderColor)"/>
            <Setter Property="BorderThickness" Value="1"/>
            <Setter Property="Padding" Value="2"/>
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="ListBox">
                        <Border Background="{TemplateBinding Background}"
                                BorderBrush="{TemplateBinding BorderBrush}"
                                BorderThickness="{TemplateBinding BorderThickness}"
                                CornerRadius="4"
                                Padding="{TemplateBinding Padding}">
                            <ScrollViewer Focusable="false">
                                <ItemsPresenter/>
                            </ScrollViewer>
                        </Border>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
        </Style>

        <Style TargetType="ListBoxItem">
            <Setter Property="Background" Value="Transparent"/>
            <Setter Property="Foreground" Value="$($Colors.TextPrimary)"/>
            <Setter Property="Padding" Value="5,3"/>
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="ListBoxItem">
                        <Border x:Name="Border"
                                Background="{TemplateBinding Background}"
                                Padding="{TemplateBinding Padding}"
                                CornerRadius="3"
                                Margin="2">
                            <ContentPresenter/>
                        </Border>
                        <ControlTemplate.Triggers>
                            <Trigger Property="IsMouseOver" Value="True">
                                <Setter TargetName="Border" Property="Background" Value="$($Colors.TileHover)"/>
                            </Trigger>
                            <Trigger Property="IsSelected" Value="True">
                                <Setter TargetName="Border" Property="Background" Value="$($Colors.AccentColor)"/>
                                <Setter Property="Foreground" Value="White"/>
                            </Trigger>
                        </ControlTemplate.Triggers>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
        </Style>

        <Style x:Key="ListBoxStyleSquareBottom" TargetType="ListBox">
            <Setter Property="Background" Value="$($Colors.CardBackground)"/>
            <Setter Property="Foreground" Value="$($Colors.TextPrimary)"/>
            <Setter Property="BorderBrush" Value="$($Colors.BorderColor)"/>
            <Setter Property="BorderThickness" Value="1"/>
            <Setter Property="Padding" Value="2"/>
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="ListBox">
                        <Border Background="{TemplateBinding Background}"
                                BorderBrush="{TemplateBinding BorderBrush}"
                                BorderThickness="{TemplateBinding BorderThickness}"
                                CornerRadius="4,4,0,0"
                                Padding="{TemplateBinding Padding}">
                            <ScrollViewer Focusable="false">
                                <ItemsPresenter/>
                            </ScrollViewer>
                        </Border>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
        </Style>
"@
}
