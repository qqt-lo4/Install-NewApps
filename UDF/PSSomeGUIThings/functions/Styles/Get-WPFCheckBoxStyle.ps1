function Get-WPFCheckBoxStyle {
    <#
    .SYNOPSIS
        Returns WPF XAML style for themed CheckBox controls

    .DESCRIPTION
        Generates XAML style definition for CheckBox controls with custom theming.

    .PARAMETER Colors
        Hashtable of theme colors (default: from Get-ThemedColors).

    .OUTPUTS
        [String]. XAML style definition for CheckBox.

    .EXAMPLE
        $checkboxStyle = Get-WPFCheckBoxStyle

    .NOTES
        Author  : Loïc Ade
        Version : 1.0.0
    #>
    Param(
        [hashtable]$Colors = (Get-ThemedColors)
    )

    return @"
        <Style TargetType="CheckBox">
            <Setter Property="Foreground" Value="$($Colors.TextPrimary)"/>
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="CheckBox">
                        <Grid>
                            <Border x:Name="CheckBoxBorder"
                                    Width="18"
                                    Height="18"
                                    Background="$($Colors.CheckBoxBackground)"
                                    BorderBrush="$($Colors.CheckBoxBorder)"
                                    BorderThickness="2"
                                    CornerRadius="3"
                                    HorizontalAlignment="Center"
                                    VerticalAlignment="Center">
                                <Viewbox Width="10" Height="10">
                                    <Path x:Name="CheckMark"
                                          Data="M 0,5 L 4,9 L 10,0"
                                          Stroke="White"
                                          StrokeThickness="2"
                                          Visibility="Collapsed"/>
                                </Viewbox>
                            </Border>
                        </Grid>
                        <ControlTemplate.Triggers>
                            <Trigger Property="IsMouseOver" Value="True">
                                <Setter TargetName="CheckBoxBorder" Property="Background" Value="$($Colors.CheckBoxHover)"/>
                            </Trigger>
                            <Trigger Property="IsChecked" Value="True">
                                <Setter TargetName="CheckBoxBorder" Property="Background" Value="$($Colors.CheckBoxChecked)"/>
                                <Setter TargetName="CheckBoxBorder" Property="BorderBrush" Value="$($Colors.CheckBoxChecked)"/>
                                <Setter TargetName="CheckMark" Property="Visibility" Value="Visible"/>
                            </Trigger>
                        </ControlTemplate.Triggers>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
        </Style>
"@
}
