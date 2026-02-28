# Installation Guide

This guide covers the installation and setup of Install-NewApps.

## Prerequisites

### System Requirements

| Requirement | Minimum |
|-------------|---------|
| Operating System | Windows 10 (1809+) or Windows 11 |
| PowerShell | 5.1 or later |
| Display | 1024x768 resolution |

### Required Components

#### WinGet (Windows Package Manager)

WinGet is required for installing packages from the WinGet source.

**Check if installed:**
```powershell
winget --version
```

**Install via Microsoft Store:**
Search for "App Installer" in the Microsoft Store, or use this link:
https://www.microsoft.com/store/productId/9NBLGGH4NNS1

**Install via PowerShell (Windows 10):**
```powershell
Add-AppxPackage -RegisterByFamilyName -MainPackage Microsoft.DesktopAppInstaller_8wekyb3d8bbwe
```

## Installation

### Method 1: Git Clone

```powershell
# Clone the repository
git clone https://github.com/qqt-lo4/Install-NewApps.git

# Navigate to the directory
cd Install-NewApps
```

### Method 2: Download ZIP

1. Download the ZIP file from the GitHub repository
2. Extract to your desired location
3. Open PowerShell and navigate to the extracted folder

## Directory Structure

After installation, ensure the following structure exists:

```
Install-NewApps/
├── Install-NewApps.ps1
├── input/
│   ├── apps.json
│   ├── apps_custom.json
│   ├── Install-NewApps.ico
│   ├── icons/
│   │   └── *.png
│   └── lang/
│       ├── en-US.json
│       └── fr-FR.json
└── UDF/
    ├── PSSomeAppsThings/
    ├── PSSomeCoreThings/
    ├── PSSomeGUIThings/
    ├── PSSomeSystemThings/
    └── ...
```

## PowerShell Execution Policy

If you encounter execution policy errors, you may need to adjust PowerShell's execution policy:

```powershell
# Check current policy
Get-ExecutionPolicy

# Set policy to allow local scripts (requires admin)
Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Scope CurrentUser
```

## First Run

1. Open PowerShell
2. Navigate to the Install-NewApps directory
3. Run the script:

```powershell
.\Install-NewApps.ps1
```

4. The GUI will display available applications
5. Select the applications you want to install
6. Click "Install"
7. Approve the UAC prompt if installing machine-scoped applications

## Troubleshooting

### "Script cannot be loaded" Error

**Cause:** PowerShell execution policy blocks the script.

**Solution:**
```powershell
Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Scope CurrentUser
```

### "WinGet not found" Error

**Cause:** WinGet is not installed or not in PATH.

**Solution:**
1. Install App Installer from Microsoft Store
2. Restart PowerShell
3. Verify: `winget --version`

### GUI Does Not Appear

**Cause:** WPF assemblies not loaded properly.

**Solution:**
1. Ensure you are running Windows 10 1809 or later
2. Try running PowerShell as Administrator
3. Check for errors in the console output

### Missing Package Icons

**Cause:** Icons folder not found or icons missing.

**Solution:**
Icons are optional. The application will work without them, displaying default icons instead.

## Updating

To update Install-NewApps:

### Git Method
```powershell
cd Install-NewApps
git pull origin main
```

### Manual Method
1. Download the latest release
2. Extract and replace existing files
3. Preserve your custom `apps.json` if modified

## Uninstallation

Simply delete the Install-NewApps folder. The application does not install system-wide components.

Note: Applications installed via Install-NewApps remain installed and must be uninstalled separately.
