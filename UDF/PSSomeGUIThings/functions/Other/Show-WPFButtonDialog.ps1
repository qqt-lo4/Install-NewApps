function Show-WPFButtonDialog {
    <#
    .SYNOPSIS
        Displays a themed button dialog with custom buttons

    .DESCRIPTION
        Shows a WPF dialog with custom buttons and automatic theming support.
        Runs in a separate runspace to avoid blocking the main thread.
        Supports custom icons, vertical or horizontal button layout.

    .PARAMETER Title
        Dialog title.

    .PARAMETER Message
        Message text to display. Use PowerShell's `n for newlines.

    .PARAMETER Buttons
        Array of button objects with 'text' and 'value' properties.

    .PARAMETER Icon
        Icon type: Information, Warning, Error, or Question (default: Information).

    .PARAMETER Vertical
        Use vertical button layout instead of horizontal.

    .PARAMETER TextBoxContent
        Optional text to display in a read-only scrollable TextBox between the message and the buttons.
        Useful for showing a list to validate or a summary of operations.

    .OUTPUTS
        Returns the 'value' of the clicked button, or $null if closed.

    .EXAMPLE
        $buttons = @(
            @{text="Yes"; value="yes"},
            @{text="No"; value="no"}
        )
        $result = Show-WPFButtonDialog -Title "Confirm" -Message "Continue?" -Buttons $buttons

    .EXAMPLE
        Show-WPFButtonDialog -Title "Error" -Message "Operation failed!" -Buttons @(@{text="OK"; value="ok"}) -Icon Error

    .NOTES
        Author  : Loïc Ade
        Version : 1.0.0
    #>
    [CmdletBinding()]
    Param(
        [Parameter(Mandatory=$true)]
        [string]$Title,

        [Parameter(Mandatory=$true)]
        [string]$Message,

        [Parameter(Mandatory=$true)]
        [array]$Buttons,

        [Parameter(Mandatory=$false)]
        [ValidateSet("Information", "Warning", "Error", "Question")]
        [string]$Icon = "Information",

        [Parameter(Mandatory=$false)]
        [switch]$Vertical,

        [Parameter(Mandatory=$false)]
        [string]$TextBoxContent
    )

    # Get Theme colors and style before runspace (module functions may not be available in runspace)
    $hColors = Get-ThemedColors
    $primaryButtonStyleXaml = Get-WPFPrimaryButtonStyle -Size Small -Colors $hColors
    $textBoxStyleXaml = Get-WPFTextBoxStyle -Colors $hColors
    $scrollBarStyleXaml = Get-WPFScrollBarStyle -Colors $hColors

    # Create synchronized hashtable for communication with runspace
    $syncHash = [hashtable]::Synchronized(@{})
    $syncHash.Title = $Title
    $syncHash.Message = $Message
    $syncHash.Buttons = $Buttons
    $syncHash.Icon = $Icon
    $syncHash.Vertical = $Vertical.IsPresent
    $syncHash.TextBoxContent = $TextBoxContent
    $syncHash.Colors = $hColors
    $syncHash.PrimaryButtonStyleXaml = $primaryButtonStyleXaml
    $syncHash.TextBoxStyleXaml = $textBoxStyleXaml
    $syncHash.ScrollBarStyleXaml = $scrollBarStyleXaml
    $syncHash.Result = $null

    # Get function definitions for the runspace
    $setForegroundFunctionCode = ""
    if (Get-Command Get-FunctionCode -ErrorAction SilentlyContinue) {
        try {
            $setForegroundFunctionCode = Get-FunctionCode -FunctionName "Set-WPFWindowForeground"
        }
        catch {
            Write-Verbose "Could not get Set-WPFWindowForeground function code"
        }
    }

    # Create runspace
    $runspace = [runspacefactory]::CreateRunspace()
    $runspace.ApartmentState = "STA"
    $runspace.ThreadOptions = "ReuseThread"
    $runspace.Open()
    $runspace.SessionStateProxy.SetVariable("syncHash", $syncHash)
    $runspace.SessionStateProxy.SetVariable("setForegroundFunctionCode", $setForegroundFunctionCode)

    # PowerShell command to create and show dialog
    $psCmd = [PowerShell]::Create().AddScript({
        param($syncHash, $setForegroundFunctionCode)

        Add-Type -AssemblyName PresentationFramework
        Add-Type -AssemblyName PresentationCore
        Add-Type -AssemblyName WindowsBase

        # Load the function into this runspace if available
        if ($setForegroundFunctionCode) {
            try {
                Invoke-Expression $setForegroundFunctionCode
            }
            catch {
                # Function not available, will use fallback
            }
        }

        # Extract parameters from syncHash
        $Title = $syncHash.Title
        $Message = $syncHash.Message
        $Buttons = $syncHash.Buttons
        $Icon = $syncHash.Icon
        $Vertical = $syncHash.Vertical
        $hColors = $syncHash.Colors

        $windowBg = $hColors.WindowBackground
        $borderColor = $hColors.BorderColor
        $textPrimary = $hColors.TextPrimary
        $textSecondary = $hColors.TextSecondary
        $titleBg = $hColors.TitleBarBackground
        $primaryButtonStyleXaml = $syncHash.PrimaryButtonStyleXaml
        $textBoxStyleXaml = $syncHash.TextBoxStyleXaml
        $scrollBarStyleXaml = $syncHash.ScrollBarStyleXaml
        $cardBg = $hColors.CardBackground

        # Icon unicode characters
        $iconChar = switch ($Icon) {
            "Information" { [char]0x2139 }  # ℹ
            "Warning"     { [char]0x26A0 }  # ⚠
            "Error"       { [char]0x2716 }  # ✖
            "Question"    { "?" }
            default       { [char]0x2139 }  # ℹ
        }

        $iconColor = switch ($Icon) {
            "Information" { $hColors.IconInformation }
            "Warning"     { $hColors.IconWarning }
            "Error"       { $hColors.IconError }
            "Question"    { $hColors.IconQuestion }
            default       { $hColors.IconInformation }
        }

        # Determine panel type and button width based on orientation
        if ($Vertical) {
            $panelType = "StackPanel"
            $panelOrientation = 'Orientation="Vertical"'
            $buttonWidth = 'MinWidth="200"'
        }
        else {
            $panelType = "WrapPanel"
            $panelOrientation = ""
            $buttonWidth = 'MinWidth="120"'
        }

        # Extract TextBox parameters from syncHash
        $TextBoxContent = $syncHash.TextBoxContent
        $hasTextBox = -not [string]::IsNullOrEmpty($TextBoxContent)

        # Escape basic XML entities for safe attribute injection
        function ConvertTo-XmlAttrSafe([string]$s) {
            if ($null -eq $s) { return "" }
            return $s.Replace('&','&amp;').Replace('<','&lt;').Replace('>','&gt;').Replace('"','&quot;').Replace("'","&apos;")
        }

        $titleXamlSafe = ConvertTo-XmlAttrSafe $Title

        # Build buttons XAML
        $buttonsXaml = ""
        for ($i = 0; $i -lt $Buttons.Count; $i++) {
            $btnTextSafe = ConvertTo-XmlAttrSafe ([string]$Buttons[$i].text)
            $buttonsXaml += @"
        <Button Content="$btnTextSafe"
                Tag="$i"
                $buttonWidth
                Height="32"
                Margin="5"
                Style="{StaticResource PrimaryButtonSmallStyle}"/>
"@
        }

        # Build optional TextBox XAML (inserted between message and buttons)
        if ($hasTextBox) {
            $extraRowDef = '                <RowDefinition Height="Auto"/>'
            $textboxXaml = @"
            <!-- TextBox -->
            <Border Grid.Row="2"
                    Margin="20,0,20,10"
                    Background="$cardBg"
                    BorderBrush="$borderColor"
                    BorderThickness="1"
                    CornerRadius="4">
                <ScrollViewer MaxHeight="200"
                              VerticalScrollBarVisibility="Auto"
                              HorizontalScrollBarVisibility="Disabled">
                    <TextBox x:Name="ContentTextBox"
                             IsReadOnly="True"
                             Background="Transparent"
                             BorderThickness="0"
                             Padding="8"
                             FontSize="12"
                             FontFamily="Consolas"
                             TextWrapping="Wrap"/>
                </ScrollViewer>
            </Border>
"@
            $buttonsRow = "3"
        } else {
            $extraRowDef = ""
            $textboxXaml = ""
            $buttonsRow  = "2"
        }

        $xaml = @"
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="$titleXamlSafe"
        WindowStartupLocation="CenterScreen"
        SizeToContent="WidthAndHeight"
        MinWidth="450"
        MaxWidth="600"
        ResizeMode="NoResize"
        Background="Transparent"
        WindowStyle="None"
        AllowsTransparency="True">

    <Window.Resources>
        $primaryButtonStyleXaml
        $textBoxStyleXaml
        $scrollBarStyleXaml
    </Window.Resources>

    <Border Background="$windowBg" BorderBrush="$borderColor" BorderThickness="1" CornerRadius="8">
        <Grid>
            <Grid.RowDefinitions>
                <RowDefinition Height="Auto"/>
                <RowDefinition Height="Auto"/>
                <RowDefinition Height="Auto"/>
$extraRowDef
            </Grid.RowDefinitions>

            <!-- Title Bar -->
            <Border Grid.Row="0" Background="$titleBg" CornerRadius="8,8,0,0" Padding="20,15">
                <Grid>
                    <TextBlock x:Name="TitleText"
                              FontSize="16"
                              FontWeight="SemiBold"
                              Foreground="$textPrimary"
                              VerticalAlignment="Center"/>
                    <Button x:Name="CloseButton"
                            Content="&#x2715;"
                            HorizontalAlignment="Right"
                            VerticalAlignment="Center"
                            Width="30"
                            Height="30"
                            Background="Transparent"
                            BorderThickness="0"
                            FontSize="16"
                            Foreground="$textSecondary"
                            Cursor="Hand">
                        <Button.Style>
                            <Style TargetType="Button">
                                <Setter Property="Template">
                                    <Setter.Value>
                                        <ControlTemplate TargetType="Button">
                                            <Border x:Name="border"
                                                    Background="{TemplateBinding Background}"
                                                    CornerRadius="4">
                                                <ContentPresenter HorizontalAlignment="Center" VerticalAlignment="Center"/>
                                            </Border>
                                            <ControlTemplate.Triggers>
                                                <Trigger Property="IsMouseOver" Value="True">
                                                    <Setter TargetName="border" Property="Background" Value="#E81123"/>
                                                    <Setter Property="Foreground" Value="White"/>
                                                </Trigger>
                                            </ControlTemplate.Triggers>
                                        </ControlTemplate>
                                    </Setter.Value>
                                </Setter>
                            </Style>
                        </Button.Style>
                    </Button>
                </Grid>
            </Border>

            <!-- Content with icon and message -->
            <StackPanel Grid.Row="1" Orientation="Horizontal" Margin="20,20,20,10">
                <TextBlock x:Name="IconText"
                          FontSize="32"
                          Foreground="$iconColor"
                          Margin="0,0,15,0"
                          VerticalAlignment="Top"/>
                <TextBlock x:Name="MessageText"
                          FontSize="13"
                          Foreground="$textPrimary"
                          TextWrapping="Wrap"
                          MaxWidth="500"
                          VerticalAlignment="Center"
                          LineStackingStrategy="BlockLineHeight"/>
            </StackPanel>

$textboxXaml
            <!-- Buttons -->
            <$panelType x:Name="ButtonPanel"
                      Grid.Row="$buttonsRow"
                      $panelOrientation
                      HorizontalAlignment="Center"
                      Margin="20,10,20,20">
$buttonsXaml
            </$panelType>
        </Grid>
    </Border>
</Window>
"@

        try {
            $reader = New-Object System.Xml.XmlNodeReader([xml]$xaml)
            $window = [Windows.Markup.XamlReader]::Load($reader)

            # Set Title
            $titleText = $window.FindName("TitleText")
            if ($titleText) {
                $titleText.Text = $Title
            }

            # Set Icon
            $iconText = $window.FindName("IconText")
            if ($iconText) {
                $iconText.Text = $iconChar
            }

            # Set Message
            $messageText = $window.FindName("MessageText")
            if ($messageText) {
                $messageText.Text = [string]$Message
            }

            # Set TextBox content
            if ($hasTextBox) {
                $contentTextBox = $window.FindName("ContentTextBox")
                if ($contentTextBox) {
                    $contentTextBox.Text = $TextBoxContent
                }
            }

            # Find the close button and attach handler
            $closeButton = $window.FindName("CloseButton")
            if ($closeButton) {
                $closeButton.Add_Click({
                    param($s, $e)
                    $syncHash.Result = $null
                    $window.DialogResult = $false
                }.GetNewClosure())
            }

            # Find the ButtonPanel and attach handlers
            $buttonPanel = $window.FindName("ButtonPanel")
            if ($buttonPanel) {
                $buttonControls = $buttonPanel.Children | Where-Object { $_ -is [System.Windows.Controls.Button] }

                foreach ($btn in $buttonControls) {
                    $buttonIndex = [int]$btn.Tag
                    $buttonValue = $Buttons[$buttonIndex].value

                    $btn.Add_Click({
                        param($s, $e)
                        $syncHash.Result = $buttonValue
                        $window.DialogResult = $true
                    }.GetNewClosure())
                }
            }

            # Handle Escape key
            $window.Add_KeyDown({
                param($s, $e)
                if ($e.Key -eq [System.Windows.Input.Key]::Escape) {
                    $syncHash.Result = $null
                    $window.DialogResult = $false
                }
            }.GetNewClosure())

            # Bring window to foreground when loaded
            $window.Add_Loaded({
                param($win, $e)
                if (Get-Command Set-WPFWindowForeground -ErrorAction SilentlyContinue) {
                    Set-WPFWindowForeground -Window $win
                }
                else {
                    $win.Activate()
                    $win.Topmost = $true
                    $win.Topmost = $false
                    $win.Focus()
                }
            }.GetNewClosure())

            # Show modal dialog
            $window.ShowDialog() | Out-Null
        }
        catch {
            Write-Error "Failed to show dialog: $_"
            $syncHash.Result = $null
        }
    })

    $psCmd.AddArgument($syncHash) | Out-Null
    $psCmd.AddArgument($setForegroundFunctionCode) | Out-Null
    $psCmd.Runspace = $runspace

    # Run and wait for completion using polling loop
    try {
        $handle = $psCmd.BeginInvoke()

        # Poll for completion instead of blocking WaitOne
        while (-not $handle.IsCompleted) {
            Start-Sleep -Milliseconds 50
        }

        # End invoke to get any errors
        $psCmd.EndInvoke($handle) | Out-Null
    }
    catch {
        Write-Error "Dialog execution failed: $_"
    }
    finally {
        # Cleanup
        $psCmd.Dispose()
        $runspace.Close()
        $runspace.Dispose()
    }

    return $syncHash.Result
}
