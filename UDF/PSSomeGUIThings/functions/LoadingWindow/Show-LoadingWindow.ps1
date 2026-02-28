function Show-LoadingWindow {
    <#
    .SYNOPSIS
        Displays a loading window with progress bar in a separate runspace

    .DESCRIPTION
        Creates a themed loading window that runs in a separate runspace, allowing
        the main script to continue executing while the window is displayed.
        The window can be updated with progress information and closed when done.

    .PARAMETER Title
        Window title text

    .PARAMETER Message
        Loading message to display

    .PARAMETER IconFile
        Path to the icon file (.ico) for the window

    .OUTPUTS
        Returns a synchronized hashtable containing:
        - Window: The WPF Window object
        - MessageText: The TextBlock control for the message
        - ProgressBar: The ProgressBar control

    .EXAMPLE
        $loadingWindow = Show-LoadingWindow -Title "Loading" -Message "Please wait..."
        # Do work
        Update-LoadingWindow -Window $loadingWindow -Message "Step 1/3" -Progress 33
        # More work
        Close-LoadingWindow -Window $loadingWindow

    .NOTES
        Author  : Loïc Ade
        Version : 1.0.0
        Requires the following functions to be available:
        - Get-ThemedColors
        - Get-FunctionCode
        - Set-WPFWindowForeground
        - Set-WindowIcon
    #>

    [CmdletBinding()]
    Param(
        [Parameter(Mandatory=$false)]
        [string]$Title = "Loading",

        [Parameter(Mandatory=$false)]
        [string]$Message = "Please wait...",

        [Parameter(Mandatory=$false)]
        [string]$IconFile = "",

        [Parameter(Mandatory=$false)]
        [string]$ShowDetailsText = "Show Details",

        [Parameter(Mandatory=$false)]
        [string]$HideDetailsText = "Hide Details",

        [Parameter(Mandatory=$false)]
        [switch]$ShowDetailsSection
    )

    # Create synchronized hashtable for communication
    $syncHash = [hashtable]::Synchronized(@{})
    $syncHash.Title = $Title
    $syncHash.Message = $Message
    $syncHash.IconFile = $IconFile
    $syncHash.ShowDetailsText = $ShowDetailsText
    $syncHash.HideDetailsText = $HideDetailsText
    $syncHash.DetailsExpanded = $false
    $syncHash.ShowDetailsSection = $ShowDetailsSection.IsPresent

    # Get theme colors
    $themeColors = Get-ThemedColors

    # Get function definitions for the runspace
    $setForegroundFunctionCode = Get-FunctionCode -FunctionName "Set-WPFWindowForeground"
    $setWindowIconFunctionCode = Get-FunctionCode -FunctionName "Set-WindowIcon"

    # Create runspace
    $runspace = [runspacefactory]::CreateRunspace()
    $runspace.ApartmentState = "STA"
    $runspace.ThreadOptions = "ReuseThread"
    $runspace.Open()
    $runspace.SessionStateProxy.SetVariable("syncHash", $syncHash)
    $runspace.SessionStateProxy.SetVariable("themeColors", $themeColors)
    $runspace.SessionStateProxy.SetVariable("setForegroundFunctionCode", $setForegroundFunctionCode)
    $runspace.SessionStateProxy.SetVariable("setWindowIconFunctionCode", $setWindowIconFunctionCode)

    # PowerShell command to create window
    $psCmd = [PowerShell]::Create().AddScript({
        param($syncHash, $themeColors, $setForegroundFunctionCode, $setWindowIconFunctionCode)

        Add-Type -AssemblyName PresentationFramework
        Add-Type -AssemblyName PresentationCore
        Add-Type -AssemblyName WindowsBase

        # Load the functions into this runspace
        Invoke-Expression $setForegroundFunctionCode
        Invoke-Expression $setWindowIconFunctionCode

        $windowHeight = if ($syncHash.ShowDetailsSection) { 200 } else { 160 }
        $xaml = @"
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="$($syncHash.Title)"
        Height="$windowHeight" Width="500"
        MinHeight="$windowHeight"
        WindowStartupLocation="CenterScreen"
        WindowStyle="None"
        AllowsTransparency="True"
        Background="Transparent"
        ResizeMode="NoResize">

    <Window.Resources>
        <!-- Transparent button style without hover effect -->
        <Style x:Key="TransparentButtonStyle" TargetType="Button">
            <Setter Property="Background" Value="Transparent"/>
            <Setter Property="BorderThickness" Value="0"/>
            <Setter Property="Cursor" Value="Hand"/>
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="Button">
                        <Border Background="{TemplateBinding Background}" Padding="{TemplateBinding Padding}">
                            <ContentPresenter HorizontalAlignment="Left" VerticalAlignment="Center"/>
                        </Border>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
        </Style>
        <!-- Dark ScrollBar Thumb -->
        <Style x:Key="DarkScrollBarThumb" TargetType="Thumb">
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="Thumb">
                        <Border Background="$($themeColors.TextSecondary)" CornerRadius="4" Margin="2"/>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
        </Style>
        <!-- Dark Vertical ScrollBar -->
        <Style x:Key="DarkScrollBar" TargetType="ScrollBar">
            <Setter Property="Background" Value="$($themeColors.WindowBackground)"/>
            <Setter Property="Width" Value="12"/>
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="ScrollBar">
                        <Grid Background="{TemplateBinding Background}">
                            <Track x:Name="PART_Track" IsDirectionReversed="True">
                                <Track.Thumb>
                                    <Thumb Style="{StaticResource DarkScrollBarThumb}"/>
                                </Track.Thumb>
                            </Track>
                        </Grid>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
            <Style.Triggers>
                <Trigger Property="Orientation" Value="Horizontal">
                    <Setter Property="Width" Value="Auto"/>
                    <Setter Property="Height" Value="12"/>
                    <Setter Property="Template">
                        <Setter.Value>
                            <ControlTemplate TargetType="ScrollBar">
                                <Grid Background="{TemplateBinding Background}">
                                    <Track x:Name="PART_Track" IsDirectionReversed="False">
                                        <Track.Thumb>
                                            <Thumb Style="{StaticResource DarkScrollBarThumb}"/>
                                        </Track.Thumb>
                                    </Track>
                                </Grid>
                            </ControlTemplate>
                        </Setter.Value>
                    </Setter>
                </Trigger>
            </Style.Triggers>
        </Style>
        <!-- Dark ScrollViewer with styled corner -->
        <Style x:Key="DarkScrollViewer" TargetType="ScrollViewer">
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="ScrollViewer">
                        <Grid Background="{TemplateBinding Background}">
                            <Grid.ColumnDefinitions>
                                <ColumnDefinition Width="*"/>
                                <ColumnDefinition Width="Auto"/>
                            </Grid.ColumnDefinitions>
                            <Grid.RowDefinitions>
                                <RowDefinition Height="*"/>
                                <RowDefinition Height="Auto"/>
                            </Grid.RowDefinitions>
                            <ScrollContentPresenter Grid.Column="0" Grid.Row="0" Margin="{TemplateBinding Padding}"/>
                            <ScrollBar x:Name="PART_VerticalScrollBar" Grid.Column="1" Grid.Row="0"
                                      Style="{StaticResource DarkScrollBar}"
                                      Value="{TemplateBinding VerticalOffset}"
                                      Maximum="{TemplateBinding ScrollableHeight}"
                                      ViewportSize="{TemplateBinding ViewportHeight}"
                                      Visibility="{TemplateBinding ComputedVerticalScrollBarVisibility}"/>
                            <ScrollBar x:Name="PART_HorizontalScrollBar" Grid.Column="0" Grid.Row="1"
                                      Style="{StaticResource DarkScrollBar}"
                                      Orientation="Horizontal"
                                      Value="{TemplateBinding HorizontalOffset}"
                                      Maximum="{TemplateBinding ScrollableWidth}"
                                      ViewportSize="{TemplateBinding ViewportWidth}"
                                      Visibility="{TemplateBinding ComputedHorizontalScrollBarVisibility}"/>
                            <Rectangle Grid.Column="1" Grid.Row="1" Fill="$($themeColors.WindowBackground)"/>
                        </Grid>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
        </Style>
        <!-- Console-style TextBox with dark scrollbars -->
        <Style x:Key="ConsoleTextBoxStyle" TargetType="TextBox">
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="TextBox">
                        <Border Background="{TemplateBinding Background}"
                                BorderBrush="{TemplateBinding BorderBrush}"
                                BorderThickness="{TemplateBinding BorderThickness}">
                            <ScrollViewer x:Name="PART_ContentHost"
                                         Style="{StaticResource DarkScrollViewer}"
                                         Padding="{TemplateBinding Padding}"/>
                        </Border>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
        </Style>
    </Window.Resources>

    <Border Background="$($themeColors.SidebarBackground)"
            BorderBrush="$($themeColors.AccentColor)"
            BorderThickness="2"
            CornerRadius="8"
            Padding="30">
        <StackPanel>
            <TextBlock x:Name="TitleText"
                      Text="$($syncHash.Title)"
                      FontSize="18"
                      FontWeight="Bold"
                      Foreground="$($themeColors.TextPrimary)"
                      HorizontalAlignment="Center"
                      Margin="0,0,0,20"/>

            <TextBlock x:Name="MessageText"
                      Text="$($syncHash.Message)"
                      FontSize="14"
                      Foreground="$($themeColors.TextSecondary)"
                      HorizontalAlignment="Center"
                      Margin="0,0,0,20"/>

            <ProgressBar x:Name="ProgressBar"
                        Height="20"
                        Minimum="0"
                        Maximum="100"
                        Value="0"
                        IsIndeterminate="True"/>

            <!-- Details Toggle Button -->
            <Button x:Name="DetailsToggleButton"
                    Style="{StaticResource TransparentButtonStyle}"
                    HorizontalAlignment="Left"
                    Margin="0,15,0,0"
                    Padding="0"
                    Visibility="$(if ($syncHash.ShowDetailsSection) { 'Visible' } else { 'Collapsed' })">
                <StackPanel Orientation="Horizontal">
                    <TextBlock x:Name="DetailsArrow"
                              Text="&#x25B6;"
                              Foreground="$($themeColors.TextSecondary)"
                              VerticalAlignment="Center"
                              Margin="0,0,8,0"
                              FontSize="10"/>
                    <TextBlock x:Name="DetailsToggleText"
                              Text="$($syncHash.ShowDetailsText)"
                              FontSize="12"
                              Foreground="$($themeColors.AccentColor)"/>
                </StackPanel>
            </Button>

            <!-- Collapsible Details Section -->
            <Border x:Name="DetailsPanel"
                    Visibility="Collapsed"
                    Margin="0,10,0,0"
                    Background="$($themeColors.WindowBackground)"
                    BorderBrush="$($themeColors.BorderColor)"
                    BorderThickness="1"
                    CornerRadius="4">
                <TextBox x:Name="DetailsTextBox"
                        Style="{StaticResource ConsoleTextBoxStyle}"
                        IsReadOnly="True"
                        Height="230"
                        Background="$($themeColors.WindowBackground)"
                        BorderThickness="0"
                        Foreground="$($themeColors.TextSecondary)"
                        FontFamily="Consolas"
                        FontSize="11"
                        TextWrapping="NoWrap"
                        VerticalScrollBarVisibility="Auto"
                        HorizontalScrollBarVisibility="Auto"
                        AcceptsReturn="True"
                        Padding="8"/>
            </Border>
        </StackPanel>
    </Border>
</Window>
"@

        $reader = New-Object System.Xml.XmlNodeReader ([xml]$xaml)
        $syncHash.Window = [Windows.Markup.XamlReader]::Load($reader)

        # Set window icon (title bar + taskbar)
        Set-WindowIcon -Window $syncHash.Window -IconPath $syncHash.IconFile

        # Get controls
        $syncHash.MessageText = $syncHash.Window.FindName("MessageText")
        $syncHash.ProgressBar = $syncHash.Window.FindName("ProgressBar")

        # Get details controls
        $syncHash.DetailsToggleButton = $syncHash.Window.FindName("DetailsToggleButton")
        $syncHash.DetailsToggleText = $syncHash.Window.FindName("DetailsToggleText")
        $syncHash.DetailsArrow = $syncHash.Window.FindName("DetailsArrow")
        $syncHash.DetailsPanel = $syncHash.Window.FindName("DetailsPanel")
        $syncHash.DetailsTextBox = $syncHash.Window.FindName("DetailsTextBox")

        # Toggle button click handler
        $syncHash.DetailsToggleButton.Add_Click({
            $syncHash.DetailsExpanded = -not $syncHash.DetailsExpanded

            if ($syncHash.DetailsExpanded) {
                $syncHash.DetailsPanel.Visibility = [System.Windows.Visibility]::Visible
                $syncHash.DetailsToggleText.Text = $syncHash.HideDetailsText
                $syncHash.DetailsArrow.Text = [char]0x25BC  # Down arrow
                $syncHash.Window.Height = 450
            }
            else {
                $syncHash.DetailsPanel.Visibility = [System.Windows.Visibility]::Collapsed
                $syncHash.DetailsToggleText.Text = $syncHash.ShowDetailsText
                $syncHash.DetailsArrow.Text = [char]0x25B6  # Right arrow
                $syncHash.Window.Height = 200
            }
        })

        # Allow dragging the window by clicking anywhere
        $syncHash.Window.Add_MouseLeftButtonDown({
            param($win, $e)
            $win.DragMove()
        })

        # Show window
        $syncHash.Window.Add_Loaded({
            param($win, $e)
            Set-WPFWindowForeground -Window $win
        })
        $syncHash.Window.ShowDialog() | Out-Null
    })

    $psCmd.AddArgument($syncHash)
    $psCmd.AddArgument($themeColors)
    $psCmd.AddArgument($setForegroundFunctionCode)
    $psCmd.AddArgument($setWindowIconFunctionCode)
    $psCmd.Runspace = $runspace

    # Store runspace and command for cleanup
    $syncHash.Runspace = $runspace
    $syncHash.PowerShell = $psCmd

    # Start async
    $handle = $psCmd.BeginInvoke()
    $syncHash.Handle = $handle

    # Wait for window to be created
    $timeout = 0
    while (-not $syncHash.Window -and $timeout -lt 100) {
        Start-Sleep -Milliseconds 50
        $timeout++
    }

    # Wait a bit more for window to be fully shown
    Start-Sleep -Milliseconds 100

    # Force window to foreground from main thread
    if ($syncHash.Window) {
        try {
            $syncHash.Window.Dispatcher.Invoke([Action]{
                Set-WPFWindowForeground -Window $syncHash.Window
            }, [System.Windows.Threading.DispatcherPriority]::Normal)
        }
        catch {
            Write-Verbose "Could not bring loading window to foreground: $_"
        }
    }

    # Return synchronized hash
    return $syncHash
}
