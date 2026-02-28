# Function Reference

This document describes the main functions used in Install-NewApps.

## Main Script Functions

### Install-OfficeWithODT

Installs Microsoft Office using Office Deployment Tool.

```powershell
Install-OfficeWithODT -Package <hashtable> [-TempPath <string>]
```

| Parameter | Type | Description |
|-----------|------|-------------|
| `Package` | hashtable | Package with ODT configuration |
| `TempPath` | string | Temp directory (default: `$env:TEMP`) |

### Install-Package

Main dispatcher for package installation based on source type.

```powershell
Install-Package -Package <hashtable> [-Credential <PSCredential>] [-TempPath <string>]
```

### Test-PackageInstalled

Checks if a package is already installed.

```powershell
Test-PackageInstalled -Package <hashtable>
```

Returns `$true` if installed, `$false` otherwise.

### Get-SortedPackagesByDependencies

Resolves dependencies and returns packages in installation order.

```powershell
Get-SortedPackagesByDependencies -Packages <array> -AllAvailablePackages <array>
```

Returns hashtable with:
- `SortedPackages`: Ordered package list
- `MissingDependencies`: Unresolved dependencies

### Show-PackageManagerUI

Displays the main WPF interface for package selection.

```powershell
Show-PackageManagerUI -Packages <array> [-IconFolder <string>] [-IconFile <string>]
```

Returns array of selected packages.

### Install-SelectedPackagesWithUI

Orchestrates installation with progress UI.

```powershell
Install-SelectedPackagesWithUI -Packages <array> [-IconFile <string>]
```

Returns array of installation results.

---

## GUI Functions (PSSomeGUIThings)

### Show-LoadingWindow

Displays a progress window with message and progress bar.

```powershell
Show-LoadingWindow -Title <string> -Message <string> [-IconFile <string>]
```

Returns window object for later updates.

### Update-LoadingWindow

Updates the loading window message and progress.

```powershell
Update-LoadingWindow -Window <object> -Message <string> [-Progress <int>]
```

### Close-LoadingWindow

Closes the loading window and cleans up resources.

```powershell
Close-LoadingWindow -Window <object>
```

### Show-WPFButtonDialog

Shows a modal dialog with custom buttons.

```powershell
Show-WPFButtonDialog -Title <string> -Message <string> -Buttons <array> [-Icon <string>]
```

| Parameter | Type | Description |
|-----------|------|-------------|
| `Title` | string | Dialog title |
| `Message` | string | Dialog message |
| `Buttons` | array | Array of @{text="Label"; value="return_value"} |
| `Icon` | string | Icon type: `Information`, `Warning`, `Error` |

Returns the value of the clicked button.

### Get-SystemTheme

Detects the current Windows theme (light/dark).

```powershell
Get-SystemTheme
```

Returns `"Light"` or `"Dark"`.

### Get-ThemedColors

Returns theme-appropriate colors for UI elements.

```powershell
Get-ThemedColors [-Theme <string>]
```

Returns hashtable with color values.

---

## WinGet Functions (PSSomeAppsThings)

### Get-WingetPackageCatalog

Initializes connection to WinGet SQLite database.

```powershell
Get-WingetPackageCatalog [-Source <string>]
```

### Get-WingetPackageInstaller

Retrieves installer information for a package.

```powershell
Get-WingetPackageInstaller -PackageId <string> [-Architecture <string>] [-Scope <string>]
```

Returns hashtable with:
- `Url`: Download URL
- `InstallerType`: exe, msi, msix, zip, etc.
- `SilentArgs`: Silent installation arguments

### Get-WingetPackageManifest

Gets the full manifest for a package.

```powershell
Get-WingetPackageManifest -PackageId <string>
```

### Search-WingetPackages

Searches packages in WinGet catalog.

```powershell
Search-WingetPackages -Query <string> [-Source <string>]
```

---

## Microsoft Store Functions (PSSomeAppsThings)

### Get-StoreAppInfo

Retrieves app information from Microsoft Store.

```powershell
Get-StoreAppInfo -ProductId <string>
```

### Install-MSStoreApp

Installs a Microsoft Store application.

```powershell
Install-MSStoreApp -ProductId <string> [-Scope <string>]
```

### Get-DeviceMSAToken

Retrieves MSA authentication token for Store APIs.

```powershell
Get-DeviceMSAToken [-Force]
```

---

## Office Functions (PSSomeAppsThings)

### Get-OfficeDeploymentToolPath

Locates the Office Deployment Tool executable.

```powershell
Get-OfficeDeploymentToolPath
```

Returns path to setup.exe or `$null` if not found.

### New-OfficeDeploymentConfiguration

Generates ODT XML configuration file.

```powershell
New-OfficeDeploymentConfiguration -Products <array> -OutputPath <string> `
    [-Language <string>] [-OfficeClientEdition <string>] [-Channel <string>] `
    [-ExcludeApps <array>] [-DisplayLevel <string>]
```

### Test-OfficeDeploymentTool

Checks if ODT is installed.

```powershell
Test-OfficeDeploymentTool
```

Returns `$true` if installed.

---

## Localization Functions (PSSomeCoreThings)

### Get-CurrentLocale

Gets the current application locale.

```powershell
Get-CurrentLocale
```

Returns locale code (e.g., `en-US`, `fr-FR`).

### Set-CurrentLocale

Changes the application language.

```powershell
Set-CurrentLocale -Locale <string>
```

### Get-Translations

Loads translations from JSON file.

```powershell
Get-Translations [-Locale <string>]
```

### Get-LocalizedString (alias: tr)

Retrieves a translated string.

```powershell
tr <string> [-Parameters <array>]
```

Example:
```powershell
tr 'UI.Installing' -Parameters @($packageName)
# Returns: "Installing Git..."
```

---

## Program Detection Functions (PSSomeAppsThings)

### Get-InstalledPrograms

Enumerates installed programs from registry.

```powershell
Get-InstalledPrograms [-Name <string>]
```

Returns array of program information.

---

## System Functions (PSSomeSystemThings)

### Get-SystemArchitecture

Detects system processor architecture.

```powershell
Get-SystemArchitecture
```

Returns `x64`, `x86`, or `ARM64`.

### Invoke-AsSystem

Executes a script block as SYSTEM account.

```powershell
Invoke-AsSystem -ScriptBlock <scriptblock>
```

---

## Script Utility Functions (PSSomeCoreThings)

### Get-FunctionCode

Extracts the source code of a PowerShell function.

```powershell
Get-FunctionCode -FunctionName <string>
```

Returns the function definition as string.

### Get-ScriptConfig

Loads script configuration from JSON file.

```powershell
Get-ScriptConfig [-ConfigPath <string>]
```

---

## Environment Functions (PSSomeSystemThings)

### Add-PathToEnvironment

Adds a directory to PATH environment variable.

```powershell
Add-PathToEnvironment -Path <string> [-Scope <string>]
```

| Parameter | Type | Description |
|-----------|------|-------------|
| `Path` | string | Directory to add |
| `Scope` | string | `User` or `Machine` |
