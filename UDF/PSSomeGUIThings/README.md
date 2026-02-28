# PSSomeGUIThings

PowerShell module providing comprehensive WPF GUI utilities including themed controls, loading windows, message boxes, theming, and UI automation.

## Author
**Loïc Ade**

## Version
1.0.0

## Description
PSSomeGUIThings is a PowerShell module that provides a complete toolkit for building modern WPF-based graphical user interfaces. It includes automatic theme detection, styled controls, loading windows with progress tracking, credential dialogs, message boxes, and UI automation utilities.

## Installation

```powershell
# Import the module
Import-Module PSSomeGUIThings
```

## Requirements
- PowerShell 5.1 or higher
- Windows Presentation Foundation (WPF)
- Windows 10 or higher (for theme detection)

## Functions

### Loading Window Functions (3)

| Function | Description |
|----------|-------------|
| `Show-LoadingWindow` | Displays a loading window with progress bar in separate runspace |
| `Update-LoadingWindow` | Updates message, progress, and details of loading window |
| `Close-LoadingWindow` | Closes and disposes a loading window and its runspace |

### GUI Dialog Functions (6)

| Function | Description |
|----------|-------------|
| `Get-CredentialGUI` | Displays a graphical credential prompt dialog |
| `New-ThemedWPFWindow` | Creates a themed WPF window with automatic style and color injection |
| `New-WPFMessageBox` | Creates customizable WPF message boxes with icons and buttons |
| `Set-WindowIcon` | Sets the icon for a WPF window (title bar and taskbar) |
| `Set-WPFWindowForeground` | Brings a WPF window to the foreground using Windows API |
| `Show-WPFButtonDialog` | Shows a themed button dialog with configurable options |

### Style Functions (10)

| Function | Description |
|----------|-------------|
| `Get-WPFCheckBoxStyle` | Returns XAML style for themed CheckBox controls |
| `Get-WPFComboBoxStyle` | Returns XAML style for themed ComboBox controls |
| `Get-WPFGroupBoxStyle` | Returns XAML style for themed GroupBox controls |
| `Get-WPFLabelStyle` | Returns XAML style for themed Label controls |
| `Get-WPFListBoxStyle` | Returns XAML style for themed ListBox controls |
| `Get-WPFPasswordBoxStyle` | Returns XAML style for themed PasswordBox controls |
| `Get-WPFPrimaryButtonStyle` | Returns XAML style for themed primary Button controls |
| `Get-WPFScrollBarStyle` | Returns XAML style for themed ScrollBar controls |
| `Get-WPFSidebarButtonStyle` | Returns XAML style for sidebar navigation RadioButtons |
| `Get-WPFTextBoxStyle` | Returns XAML style for themed TextBox controls |

### Theme Functions (3)

| Function | Description |
|----------|-------------|
| `Get-SystemTheme` | Detects if Windows is using dark or light theme |
| `Set-DarkTitleBar` | Applies dark title bar to a WPF window |
| `Get-ThemedColors` | Returns color scheme based on system theme |

### UI Automation Functions (6)

| Function | Description |
|----------|-------------|
| `Get-Window` | Enumerates windows with filtering by process, class name, title, and child process support |
| `Set-WindowVisibility` | Shows, hides, minimizes, maximizes, or restores a window by its handle |
| `Set-WindowForeground` | Activates a window and brings it to the foreground (with AttachThreadInput) |
| `Set-WindowOnTop` | Brings a window to the top of the Z-order without giving it focus |
| `Wait-Window` | Waits for a window to appear for a specific process |
| `Wait-MainWindowFromProcessClose` | Waits for a process's main window to close |

## Usage Examples

### Creating a Themed Loading Window

```powershell
# Show loading window
$loadingWindow = Show-LoadingWindow -Title "Installation" `
                                    -Message "Preparing installation..." `
                                    -IconFile "C:\app\icon.ico" `
                                    -ShowDetailsSection

# Update progress
Update-LoadingWindow -Window $loadingWindow `
                     -Message "Installing component 1/3" `
                     -Progress 33 `
                     -Details "Extracting files..."

# Add more details
Update-LoadingWindow -Window $loadingWindow `
                     -Message "Installing component 2/3" `
                     -Progress 66 `
                     -Details "Configuring settings..." `
                     -AppendDetails

# Complete and close
Update-LoadingWindow -Window $loadingWindow `
                     -Message "Installation complete!" `
                     -Progress 100

Start-Sleep -Seconds 2
Close-LoadingWindow -Window $loadingWindow
```

### Creating a Themed WPF Window

```powershell
Add-Type -AssemblyName PresentationFramework

# Get theme colors
$colors = Get-ThemedColors

# Create window with themed XAML
$xaml = @"
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        Title="My App" Height="400" Width="600"
        Background="$($colors.WindowBackground)">
    <Window.Resources>
        $(Get-WPFCheckBoxStyle -Colors $colors)
        $(Get-WPFComboBoxStyle -Colors $colors)
        $(Get-WPFPrimaryButtonStyle -Size Large -Colors $colors)
    </Window.Resources>

    <Grid Margin="20">
        <StackPanel>
            <TextBlock Text="Welcome" FontSize="24"
                      Foreground="$($colors.TextPrimary)" Margin="0,0,0,20"/>

            <CheckBox Content="Enable feature" Margin="0,10"/>

            <ComboBox Margin="0,10">
                <ComboBoxItem Content="Option 1"/>
                <ComboBoxItem Content="Option 2"/>
                <ComboBoxItem Content="Option 3"/>
            </ComboBox>

            <Button Content="Continue" Style="{StaticResource PrimaryButtonStyle}"
                   HorizontalAlignment="Right" Margin="0,20,0,0"/>
        </StackPanel>
    </Grid>
</Window>
"@

$reader = [System.Xml.XmlNodeReader]::new([xml]$xaml)
$window = [Windows.Markup.XamlReader]::Load($reader)

# Apply dark title bar if dark theme
if (Get-SystemTheme) {
    Set-DarkTitleBar -Window $window
}

# Set window icon
Set-WindowIcon -Window $window -IconPath "C:\app\icon.ico"

# Bring to foreground when loaded
$window.Add_Loaded({
    param($win, $e)
    Set-WPFWindowForeground -Window $win
})

$window.ShowDialog()
```

### Using Message Boxes and Dialogs

```powershell
# Show information message
New-WPFMessageBox -Title "Success" `
                  -Message "Operation completed successfully!" `
                  -Icon Information

# Show warning with Yes/No buttons
$result = New-WPFMessageBox -Title "Confirm" `
                             -Message "Are you sure you want to continue?" `
                             -Icon Warning `
                             -Buttons YesNo

if ($result -eq "Yes") {
    # User confirmed
}

# Show custom button dialog
$result = Show-WPFButtonDialog -Title "Choose Action" `
                               -Message "What would you like to do?" `
                               -Buttons @("Install", "Repair", "Uninstall", "Cancel")

switch ($result) {
    "Install" { # Handle install }
    "Repair"  { # Handle repair }
    "Uninstall" { # Handle uninstall }
}
```

### Getting Credentials

```powershell
# Show credential dialog
$cred = Get-CredentialGUI -TargetName "Remote Server" `
                          -Message "Enter your credentials to connect"

if ($cred) {
    # Use credentials
    Invoke-Command -ComputerName "server01" -Credential $cred -ScriptBlock {
        # Remote commands
    }
}
```

### UI Automation

```powershell
# Get all Chrome windows with position and state
$chromeWindows = Get-Window -ProcessName chrome
foreach ($window in $chromeWindows) {
    Write-Host "$($window.Title) - ${$window.Width}x$($window.Height) Visible=$($window.IsVisible)"
}

# Find windows by title
Get-Window -Title "*Notepad*" -ExcludeEmptyTitle

# Find windows from a process and its children
Get-Window -ProcessId $PID -IncludeChildren -ClassName "Chrome_WidgetWin_*"

# Hide/show/minimize/maximize/restore a window
$hwnd = (Get-Window -ProcessName notepad)[0].Handle
Set-WindowVisibility -Handle $hwnd -Minimize
Set-WindowVisibility -Handle $hwnd -Restore

# Bring a window to the foreground
Set-WindowForeground -Handle $hwnd

# Bring a window to the top without focus
Set-WindowOnTop -Handle $hwnd

# Launch application and wait for window
Start-Process "notepad.exe"
$window = Wait-Window -ProcessName notepad -TimeoutSeconds 10

if ($window) {
    Write-Host "Notepad opened: $($window.Title)"
}

# Wait for window to close
Wait-MainWindowFromProcessClose -processName notepad
Write-Host "Notepad closed"
```

## Common Workflows

### Long-Running Operation with Loading Window

```powershell
# Start loading window
$loading = Show-LoadingWindow -Title "Processing Data" `
                              -Message "Initializing..." `
                              -ShowDetailsSection

try {
    $items = @(1..100)
    $total = $items.Count

    foreach ($i in 0..($total-1)) {
        # Perform work
        Start-Sleep -Milliseconds 100

        # Update progress
        $percent = [int](($i + 1) / $total * 100)
        Update-LoadingWindow -Window $loading `
                            -Message "Processing item $($i+1) of $total" `
                            -Progress $percent `
                            -Details "Item $($i+1): Processing..." `
                            -AppendDetails
    }

    # Success
    Update-LoadingWindow -Window $loading `
                        -Message "Processing complete!" `
                        -Progress 100
    Start-Sleep -Seconds 1
}
finally {
    Close-LoadingWindow -Window $loading
}
```

### Building a Modern Settings Dialog

```powershell
# Get themed colors
$colors = Get-ThemedColors

# Build XAML with sidebar navigation
$xaml = @"
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        Title="Settings" Height="500" Width="700"
        Background="$($colors.WindowBackground)">
    <Window.Resources>
        $(Get-WPFSidebarButtonStyle -Colors $colors)
        $(Get-WPFPrimaryButtonStyle -Size Small -Colors $colors)
        $(Get-WPFCheckBoxStyle -Colors $colors)
    </Window.Resources>

    <Grid>
        <Grid.ColumnDefinitions>
            <ColumnDefinition Width="200"/>
            <ColumnDefinition Width="*"/>
        </Grid.ColumnDefinitions>

        <!-- Sidebar -->
        <Border Grid.Column="0" Background="$($colors.SidebarBackground)"
                BorderBrush="$($colors.BorderColor)" BorderThickness="0,0,1,0">
            <StackPanel Margin="0,20,0,0">
                <RadioButton Content="General" IsChecked="True"
                            Style="{StaticResource SidebarButtonStyle}"/>
                <RadioButton Content="Appearance"
                            Style="{StaticResource SidebarButtonStyle}"/>
                <RadioButton Content="Advanced"
                            Style="{StaticResource SidebarButtonStyle}"/>
            </StackPanel>
        </Border>

        <!-- Content -->
        <StackPanel Grid.Column="1" Margin="30">
            <TextBlock Text="General Settings" FontSize="20" FontWeight="Bold"
                      Foreground="$($colors.TextPrimary)" Margin="0,0,0,20"/>

            <CheckBox Content="Start on Windows startup" Margin="0,10"/>
            <CheckBox Content="Minimize to system tray" Margin="0,10"/>
            <CheckBox Content="Check for updates automatically" Margin="0,10"/>

            <Button Content="Save" Style="{StaticResource PrimaryButtonSmallStyle}"
                   HorizontalAlignment="Right" Margin="0,30,0,0"/>
        </StackPanel>
    </Grid>
</Window>
"@

$reader = [System.Xml.XmlNodeReader]::new([xml]$xaml)
$window = [Windows.Markup.XamlReader]::Load($reader)

# Apply dark title bar if needed
if (Get-SystemTheme) {
    Set-DarkTitleBar -Window $window
}

# Show foreground when loaded
$window.Add_Loaded({ Set-WPFWindowForeground -Window $args[0] })

$window.ShowDialog()
```

### Application Installer with Progress

```powershell
# Show loading window
$installer = Show-LoadingWindow -Title "App Installer" `
                                -Message "Preparing installation..." `
                                -IconFile "setup.ico" `
                                -ShowDetailsSection

try {
    # Check prerequisites
    Update-LoadingWindow -Window $installer `
                        -Message "Checking prerequisites..." `
                        -Progress 10 `
                        -Details "[1/5] Checking system requirements"

    Start-Sleep -Seconds 2

    # Extract files
    Update-LoadingWindow -Window $installer `
                        -Message "Extracting files..." `
                        -Progress 30 `
                        -Details "[2/5] Extracting application files" `
                        -AppendDetails

    Start-Sleep -Seconds 3

    # Install dependencies
    Update-LoadingWindow -Window $installer `
                        -Message "Installing dependencies..." `
                        -Progress 60 `
                        -Details "[3/5] Installing required components" `
                        -AppendDetails

    Start-Sleep -Seconds 2

    # Configure
    Update-LoadingWindow -Window $installer `
                        -Message "Configuring application..." `
                        -Progress 80 `
                        -Details "[4/5] Configuring settings" `
                        -AppendDetails

    Start-Sleep -Seconds 2

    # Finalize
    Update-LoadingWindow -Window $installer `
                        -Message "Finalizing installation..." `
                        -Progress 95 `
                        -Details "[5/5] Creating shortcuts and registry entries" `
                        -AppendDetails

    Start-Sleep -Seconds 1

    # Complete
    Update-LoadingWindow -Window $installer `
                        -Message "Installation complete!" `
                        -Progress 100 `
                        -Details "Installation finished successfully" `
                        -AppendDetails

    Start-Sleep -Seconds 2
}
finally {
    Close-LoadingWindow -Window $installer
}

# Show success message
New-WPFMessageBox -Title "Installation Complete" `
                  -Message "The application has been installed successfully!" `
                  -Icon Information
```

## Use Cases

1. **Modern Application GUIs**: Build professional WPF interfaces with automatic theme detection
2. **Installation Wizards**: Create installers with progress tracking and detailed logging
3. **System Tools**: Build administration tools with credential prompts and confirmations
4. **Data Processing**: Display progress for long-running operations with cancellation support
5. **Settings Dialogs**: Create configuration interfaces with sidebar navigation
6. **User Notifications**: Show themed message boxes and confirmation dialogs
7. **UI Automation**: Monitor and control application windows programmatically
8. **Theme-Aware Apps**: Automatically adapt to Windows dark/light theme settings

## Theming

The module automatically detects the Windows theme (dark/light) and provides appropriate colors.

For a complete guide on the theming system — including color placeholders, the `New-ThemedWPFWindow` function, best practices, and migration steps — see [README-Theming.md](doc/README-Theming.md).

### Color Palette

Both dark and light themes include colors for:
- **Window**: Background, borders
- **Text**: Primary, secondary
- **Controls**: Buttons, checkboxes, comboboxes, scrollbars
- **Interactive**: Hover states, selected states, disabled states
- **Icons**: Information, warning, error, question
- **Accent**: Primary accent color and hover state

### Custom Theming

```powershell
# Get current theme colors
$colors = Get-ThemedColors

# Override specific colors
$colors.AccentColor = "#FF6B00"  # Custom orange
$colors.AccentHover = "#FF8C00"

# Use in styles
$buttonStyle = Get-WPFPrimaryButtonStyle -Colors $colors
```

## Notes

- **Theme Detection**: Automatically reads Windows registry to detect dark/light theme preference
- **Runspace Windows**: Loading windows run in separate runspaces for non-blocking operation
- **Window Foreground**: Uses Windows API to reliably bring windows to foreground
- **Icon Support**: Supports .ico files for window icons (title bar and taskbar)
- **UI Automation**: Window enumeration works across all processes (requires appropriate permissions)
- **Styled Controls**: All style functions return XAML as strings for embedding in Window.Resources
- **Thread-Safe**: Loading window updates are dispatched to the UI thread automatically

## License

This module is licensed under the **PolyForm Noncommercial License 1.0.0**.

See the [LICENSE](LICENSE) file for full license text.

**Required Notice**: Copyright Loïc Ade (https://github.com/qqt-lo4)

## Support

For issues, questions, or contributions, please contact the author.

## Version History

- **1.0.0** (Initial Release)
  - Loading windows with progress tracking and details panel
  - Credential and message box dialogs
  - Themed WPF control styles (CheckBox, ComboBox, Button, ScrollBar, Sidebar)
  - Automatic Windows theme detection
  - Dark title bar support
  - Window foreground activation
  - Window icon management
  - UI automation for window enumeration and monitoring
