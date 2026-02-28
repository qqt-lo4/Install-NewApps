# Install-NewApps

A powerful PowerShell-based software installation manager with a modern WPF graphical interface. Install applications from multiple sources including WinGet, Microsoft Store, and Office Deployment Tool with a single UAC prompt.

![PowerShell](https://img.shields.io/badge/PowerShell-5.1%2B-blue)
![Windows](https://img.shields.io/badge/Windows-10%2F11-0078D6)
![License](https://img.shields.io/badge/License-CC%20BY--NC%204.0-lightgrey)

## Features

- **Multi-Source Support**: Install packages from WinGet, Microsoft Store, and Office Deployment Tool
- **Modern WPF Interface**: Clean, theme-aware GUI with dark/light mode support
- **Single UAC Prompt**: Batch install multiple machine-scoped packages with one elevation
- **Dependency Resolution**: Automatic detection and installation of package dependencies
- **Portable Package Support**: Install and configure portable applications with PATH management
- **Localization**: Full English and French language support
- **Installation Detection**: Intelligent detection of already-installed software
- **Category Filtering**: Organize packages by category (Office, Development, Internet, etc.)

## Requirements

- Windows 10/11
- PowerShell 5.1 or later
- WinGet (Windows Package Manager)
- Administrator privileges for machine-scoped installations

## Quick Start

1. Clone the repository:
   ```powershell
   git clone https://github.com/qqt-lo4/Install-NewApps.git
   ```

2. Run the script:
   ```powershell
   .\Install-NewApps.ps1
   ```

3. Select the applications you want to install from the GUI

4. Click "Install" and approve the UAC prompt

## Documentation

- [Installation Guide](INSTALLATION.md) - Detailed installation instructions
- [Configuration Guide](CONFIGURATION.md) - How to configure and customize packages
- [Architecture](ARCHITECTURE.md) - Technical architecture and design
- [Function Reference](FUNCTIONS.md) - API documentation for main functions
- [Localization](LOCALIZATION.md) - Adding new languages

## Package Sources

### WinGet
Standard Windows Package Manager packages with silent installation support. Supports `.exe`, `.msi`, `.zip`, `.msix`, and `.appx` installers.

### Microsoft Store
Windows Store applications installed via MSA token authentication. Includes support for Win32 apps distributed through the Store.

### Office Deployment Tool (ODT)
Microsoft Office products with customizable XML configuration. Supports multiple products, languages, and deployment channels.

## Project Structure

```
Install-NewApps/
├── Install-NewApps.ps1          # Main application script
├── input/
│   ├── apps.json                # Package definitions
│   ├── apps_custom.json         # Custom package overrides
│   ├── Install-NewApps.ico      # Application icon
│   ├── icons/                   # Package icons (PNG)
│   └── lang/
│       ├── en-US.json           # English translations
│       └── fr-FR.json           # French translations
├── UDF/                         # Reusable function modules
│   ├── PSSomeAppsThings/        # WinGet, Store, ODT, program detection
│   ├── PSSomeCoreThings/        # Localization, script configuration
│   ├── PSSomeGUIThings/         # WPF interface functions
│   ├── PSSomeSystemThings/      # System info, environment management
│   └── ...                      # Other utility modules
├── ARCHITECTURE.md
├── CONFIGURATION.md
├── FUNCTIONS.md
├── INSTALLATION.md
└── LOCALIZATION.md
```

## Usage

### Basic Usage
```powershell
# Launch the GUI
.\Install-NewApps.ps1
```

### With Verbose Output
```powershell
.\Install-NewApps.ps1 -Verbose
```

## Supported Applications

The default configuration includes more than 40 applications across categories:

| Category | Examples |
|----------|----------|
| Office | LibreOffice, draw.io, Microsoft Office 2024 |
| Development | Git, VS Code, Visual Studio, AutoIt |
| Internet | Chrome, Firefox, Telegram, Discord |
| System Tools | 7-Zip, Notepad++, PowerShell, VirtualBox |
| Audio/Video | Audacity, OBS Studio, VLC, Kdenlive |
| Photo | GIMP, PhotoDemon |
| Games | Minecraft, Epic Games Launcher |

## License

This project is licensed under **Creative Commons Attribution-NonCommercial 4.0 International (CC BY-NC 4.0)**.

You are free to:
- **Share** — copy and redistribute the material in any medium or format
- **Adapt** — remix, transform, and build upon the material

Under the following terms:
- **Attribution** — You must give appropriate credit, provide a link to the license, and indicate if changes were made
- **NonCommercial** — You may not use the material for commercial purposes

Full license: https://creativecommons.org/licenses/by-nc/4.0/
