function Get-WPFScrollBarStyle {
    <#
    .SYNOPSIS
        Returns WPF XAML style for themed ScrollBar controls

    .DESCRIPTION
        Generates XAML style definition for ScrollBar controls with custom theming.
        Includes hover effect on scroll thumb.

    .PARAMETER Colors
        Hashtable of theme colors (default: from Get-ThemedColors).

    .OUTPUTS
        [String]. XAML style definition for ScrollBar.

    .EXAMPLE
        $scrollBarStyle = Get-WPFScrollBarStyle

    .NOTES
        Author  : Loïc Ade
        Version : 1.0.0
    #>

    Param(
        [hashtable]$Colors = (Get-ThemedColors)
    )

    return @"
        <Style TargetType="ScrollBar">
            <Setter Property="Background" Value="$($Colors.ScrollBarBackground)"/>
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="ScrollBar">
                        <Grid>
                            <Grid.RowDefinitions>
                                <RowDefinition Height="*"/>
                            </Grid.RowDefinitions>
                            <Border Background="{TemplateBinding Background}" CornerRadius="6"/>
                            <Track x:Name="PART_Track" IsDirectionReversed="True">
                                <Track.Thumb>
                                    <Thumb>
                                        <Thumb.Template>
                                            <ControlTemplate TargetType="Thumb">
                                                <Border x:Name="ThumbBorder"
                                                        Background="$($Colors.ScrollBarThumb)"
                                                        CornerRadius="6"
                                                        Margin="2">
                                                </Border>
                                                <ControlTemplate.Triggers>
                                                    <Trigger Property="IsMouseOver" Value="True">
                                                        <Setter TargetName="ThumbBorder" Property="Background" Value="$($Colors.ScrollBarThumbHover)"/>
                                                    </Trigger>
                                                </ControlTemplate.Triggers>
                                            </ControlTemplate>
                                        </Thumb.Template>
                                    </Thumb>
                                </Track.Thumb>
                            </Track>
                        </Grid>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
        </Style>
"@
}
