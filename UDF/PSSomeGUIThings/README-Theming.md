# WPF Theme System Usage Guide

## Overview

The **PSSomeGUIThings** module provides a comprehensive theming system for WPF applications in PowerShell. It automatically detects the Windows system theme (light/dark) and applies consistent styles to all WPF controls.

## Core Functions

### Theme Detection and Management

#### `Get-SystemTheme`
Detects whether Windows is using the dark theme.

```powershell
$isDark = Get-SystemTheme
# Returns: $true if dark theme, $false otherwise
```

#### `Get-ThemedColors`
Returns a hashtable of colors based on the system theme.

```powershell
$colors = Get-ThemedColors
# Accessing colors:
# $colors.WindowBackground
# $colors.TextPrimary
# $colors.AccentColor
# etc.
```

#### `Set-DarkTitleBar`
Applies a dark title bar to a WPF window.

```powershell
Set-DarkTitleBar -Window $window
```

### Creating Themed Windows

#### `New-ThemedWPFWindow`
Creates a WPF window with automatic theming and style injection.

```powershell
$xaml = @"
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="My Application"
        Height="400"
        Width="600"
        Background="{ThemeColor:WindowBackground}">
    <Grid>
        <Button Content="Click me" Style="{StaticResource PrimaryButtonStyle}"/>
    </Grid>
</Window>
"@

$window = New-ThemedWPFWindow -XAML $xaml
$window.ShowDialog()
```

### Available Control Styles

The following functions generate XAML styles for various controls:

- `Get-WPFPrimaryButtonStyle` - Buttons with accent color
- `Get-WPFScrollBarStyle` - Styled scrollbars
- `Get-WPFComboBoxStyle` - Dropdown lists
- `Get-WPFTextBoxStyle` - Text input fields
- `Get-WPFPasswordBoxStyle` - Password input fields
- `Get-WPFListBoxStyle` - List boxes
- `Get-WPFGroupBoxStyle` - Group boxes
- `Get-WPFLabelStyle` - Labels

### Dialogs and Messages

#### `New-WPFMessageBox`
Creates a customizable dialog box with theming.

```powershell
$result = New-WPFMessageBox -Content "Do you want to continue?" `
                             -Title "Confirmation" `
                             -ButtonType "Yes-No" `
                             -Icon "Question"
```

#### `Show-WPFButtonDialog`
Displays a dialog box with custom buttons.

```powershell
$buttons = @(
    @{text="Option 1"; value="opt1"},
    @{text="Option 2"; value="opt2"}
)
$result = Show-WPFButtonDialog -Title "Choose" `
                                -Message "Select an option" `
                                -Buttons $buttons `
                                -Icon "Information"
```

## Using Color Placeholders

In your XAML, you can use placeholders that will be automatically replaced by `New-ThemedWPFWindow`:

```xml
<Window Background="{ThemeColor:WindowBackground}">
    <Grid>
        <TextBlock Foreground="{ThemeColor:TextPrimary}" Text="Hello"/>
        <Border BorderBrush="{ThemeColor:BorderColor}"/>
    </Grid>
</Window>
```

### Available Colors

| Placeholder | Description |
|-------------|-------------|
| `WindowBackground` | Window background |
| `TitleBarBackground` | Title bar background |
| `CardBackground` | Card/control background |
| `BorderColor` | Border color |
| `TextPrimary` | Primary text |
| `TextSecondary` | Secondary text |
| `AccentColor` | Accent color (blue) |
| `AccentHover` | Accent on hover |
| `TileHover` | Background on hover |
| `ScrollBarBackground` | Scrollbar background |
| `ScrollBarThumb` | Scrollbar thumb |
| `IconInformation` | Information icon |
| `IconWarning` | Warning icon |
| `IconError` | Error icon |

## Complete Example

The **Install-NewApps** project is a real-world example of an application built with this theming system.

### Key points of the example:

1. Import the PSSomeGUIThings module
2. Use `{ThemeColor:...}` placeholders in XAML
3. Create the window with `New-ThemedWPFWindow`
4. Automatic style application
5. Use `New-WPFMessageBox` instead of standard MessageBox

## Best Practices

1. **Always use placeholders**: Use `{ThemeColor:...}` instead of hard-coded colors
2. **Apply styles**: Use `Style="{StaticResource ...}"` for buttons
3. **Use New-ThemedWPFWindow**: Let the function handle style injection
4. **Replace MessageBox**: Use `New-WPFMessageBox` for a consistent interface

## Migrating an Existing Application

To migrate an existing WPF application:

1. Replace hard-coded colors with `{ThemeColor:...}` placeholders
2. Replace `[Windows.Markup.XamlReader]::Load()` with `New-ThemedWPFWindow`
3. Add `Style="{StaticResource PrimaryButtonSmallStyle}"` to buttons
4. Replace `[System.Windows.MessageBox]::Show()` with `New-WPFMessageBox`
5. Remove redundant properties (they will be applied by the styles)

---

**Author**: Loïc Ade
**Version**: 1.0.0
**Module**: PSSomeGUIThings
