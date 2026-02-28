function Get-WPFGroupBoxStyle {
    <#
    .SYNOPSIS
        Returns WPF XAML style for themed GroupBox controls

    .DESCRIPTION
        Generates XAML style definition for GroupBox controls with custom theming.
        Creates a modern look with rounded corners and themed colors.

    .PARAMETER Colors
        Hashtable of theme colors (default: from Get-ThemedColors).

    .OUTPUTS
        [String]. XAML style definition for GroupBox.

    .EXAMPLE
        $groupBoxStyle = Get-WPFGroupBoxStyle

    .NOTES
        Author  : Loïc Ade
        Version : 1.0.0
    #>

    Param(
        [hashtable]$Colors = (Get-ThemedColors)
    )

    return @"
        <Style TargetType="GroupBox">
            <Setter Property="Foreground" Value="$($Colors.TextPrimary)"/>
            <Setter Property="BorderBrush" Value="$($Colors.BorderColor)"/>
            <Setter Property="BorderThickness" Value="1"/>
            <Setter Property="Padding" Value="5"/>
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="GroupBox">
                        <Grid>
                            <Grid.RowDefinitions>
                                <RowDefinition Height="Auto"/>
                                <RowDefinition Height="*"/>
                            </Grid.RowDefinitions>
                            <Border Grid.Row="0"
                                    Background="$($Colors.TileBackground)"
                                    BorderBrush="{TemplateBinding BorderBrush}"
                                    BorderThickness="1,1,1,0"
                                    CornerRadius="6,6,0,0"
                                    Padding="10,5">
                                <ContentPresenter ContentSource="Header"
                                                 RecognizesAccessKey="True"/>
                            </Border>
                            <Border Grid.Row="1"
                                    BorderBrush="{TemplateBinding BorderBrush}"
                                    BorderThickness="1,0,1,1"
                                    CornerRadius="0,0,6,6"
                                    Padding="{TemplateBinding Padding}">
                                <ContentPresenter/>
                            </Border>
                        </Grid>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
        </Style>
"@
}
