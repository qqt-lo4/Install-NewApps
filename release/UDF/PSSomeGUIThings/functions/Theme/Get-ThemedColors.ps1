function Get-ThemedColors {
    <#
    .SYNOPSIS
        Returns color scheme based on system theme

    .DESCRIPTION
        Automatically detects if Windows is using dark or light theme and returns
        a hashtable of color values for all UI elements. The color scheme includes
        backgrounds, text colors, borders, buttons, scrollbars, icons, and accent colors.

    .OUTPUTS
        [Hashtable]. Color values for UI theming (dark or light theme).

    .EXAMPLE
        $colors = Get-ThemedColors
        $window.Background = $colors.WindowBackground

    .NOTES
        Author  : Loïc Ade
        Version : 1.0.0
        Requires Get-SystemTheme function to detect theme mode.
    #>

    $isDark = Get-SystemTheme

    if ($isDark) {
        return @{
            WindowBackground = "#1E1E1E"
            TitleBarBackground = "#252526"
            SidebarBackground = "#252526"
            BorderColor = "#3F3F46"
            TextPrimary = "#E0E0E0"
            TextSecondary = "#A0A0A0"
            TileBackground = "#2D2D30"
            TileHover = "#3E3E42"
            AccentColor = "#0078D4"
            AccentHover = "#1084D8"
            InputBorderHover = "#4A7FA5"
            InputBorderFocus = "#3AABF8"
            ButtonBg = "#0078D4"
            ButtonHover = "#005A9E"
            ButtonDisabledBg = "#3F3F46"
            ButtonDisabledFg = "#6E6E6E"
            CategoryCheckedBg = "#0078D4"
            CategoryHoverBg = "#3E3E42"
            ScrollBarBackground = "#1E1E1E"
            ScrollBarThumb = "#686868"
            ScrollBarThumbHover = "#9E9E9E"
            CheckBoxBorder = "#686868"
            CheckBoxBackground = "#2D2D30"
            CheckBoxChecked = "#0078D4"
            CheckBoxHover = "#3E3E42"
            CardBackground = "#2D2D30"
            IconInformation = "#0078D4"
            IconWarning = "#FFA500"
            IconError = "#E81123"
            IconSuccess = "#4CAF50"
            IconQuestion = "#0078D4"
        }
    }
    else {
        return @{
            WindowBackground = "#FFFFFF"
            TitleBarBackground = "#F5F5F5"
            SidebarBackground = "White"
            BorderColor = "#E0E0E0"
            TextPrimary = "#333333"
            TextSecondary = "#666666"
            TileBackground = "White"
            TileHover = "#F8F8F8"
            AccentColor = "#0078D4"
            AccentHover = "#005A9E"
            InputBorderHover = "#90B8D8"
            InputBorderFocus = "#0063B1"
            ButtonBg = "#0078D4"
            ButtonHover = "#005A9E"
            ButtonDisabledBg = "#CCCCCC"
            ButtonDisabledFg = "#666666"
            CategoryCheckedBg = "#0078D4"
            CategoryHoverBg = "#E8E8E8"
            ScrollBarBackground = "#F0F0F0"
            ScrollBarThumb = "#CDCDCD"
            ScrollBarThumbHover = "#A6A6A6"
            CheckBoxBorder = "#ADADAD"
            CheckBoxBackground = "White"
            CheckBoxChecked = "#0078D4"
            CheckBoxHover = "#F0F0F0"
            CardBackground = "White"
            IconInformation = "#0078D4"
            IconWarning = "#FFA500"
            IconError = "#E81123"
            IconSuccess = "#107C10"
            IconQuestion = "#0078D4"
        }
    }
}
