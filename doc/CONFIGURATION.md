# Configuration Guide

This guide explains how to configure and customize the package catalog.

## Configuration File

The package catalog is defined in `input/apps.json`. This JSON file contains an array of package definitions. Custom overrides can be placed in `input/apps_custom.json`.

## Package Structure

### Basic Package (WinGet)

```json
{
    "Category": "Development",
    "Name": "Git",
    "Id": "Git.Git",
    "Scope": "machine",
    "Source": "winget"
}
```

### Package with Dependencies

```json
{
    "Category": "Development",
    "Name": "SciTE4AutoIt3",
    "Id": "AutoIt.SciTE4AutoIt3",
    "Scope": "machine",
    "Source": "winget",
    "Requires": ["AutoIt.AutoIt"]
}
```

### Package with Custom Detection

```json
{
    "Category": "Office",
    "Name": "Office Deployment Tool",
    "Id": "Microsoft.OfficeDeploymentTool",
    "Scope": "machine",
    "Source": "winget",
    "DetectionScript": "Test-OfficeDeploymentTool"
}
```

### Package with Detection Pattern

```json
{
    "Category": "Development",
    "Name": "SciTE4AutoIt3",
    "Id": "AutoIt.SciTE4AutoIt3",
    "Scope": "machine",
    "Source": "winget",
    "Detection": {
        "Name": "SciTE4AutoIt3*"
    }
}
```

### Microsoft Store Package

```json
{
    "Category": "Office",
    "Name": "Microsoft Sticky Notes",
    "PackageName": "Microsoft.MicrosoftStickyNotes",
    "Id": "9nblggh4qghw",
    "Scope": "machine",
    "Source": "msstore"
}
```

### Office Deployment Tool Package

```json
{
    "Category": "Office",
    "Name": "Microsoft Office 2024",
    "Id": "Microsoft.Office2024",
    "Scope": "machine",
    "Source": "odt",
    "ODT": {
        "Products": ["ProPlus2024Volume"],
        "Language": "fr-fr",
        "OfficeClientEdition": "64",
        "Channel": "PerpetualVL2024"
    },
    "Requires": ["Microsoft.OfficeDeploymentTool"],
    "Detection": {
        "Name": "Microsoft Office LTSC * 2024 - *"
    }
}
```

## Property Reference

### Required Properties

| Property | Type | Description |
|----------|------|-------------|
| `Category` | string | Package category for UI grouping |
| `Name` | string | Display name shown in the GUI |
| `Id` | string | Package identifier (WinGet ID or Store Product ID) |
| `Scope` | string | Installation scope: `machine` or `user` |
| `Source` | string | Package source: `winget`, `msstore`, or `odt` |

### Optional Properties

| Property | Type | Description |
|----------|------|-------------|
| `Requires` | array | List of package IDs that must be installed first |
| `Detection` | object | Custom detection configuration |
| `DetectionScript` | string | Name of a PowerShell function for custom detection |
| `PackageName` | string | AppX package family name (for Store apps) |
| `ODT` | object | Office Deployment Tool configuration |

## Categories

Default categories used in the project:

| Category | Description |
|----------|-------------|
| `Office` | Office suites, document editors, PDF tools |
| `Development` | IDEs, programming tools, version control |
| `Internet` | Browsers, communication apps, download tools |
| `SystemTools` | Utilities, archivers, system administration |
| `AudioVideo` | Media players, editors, codecs |
| `Photo` | Image editors, viewers |
| `Games` | Games and gaming platforms |

You can create custom categories by using any string value.

## Scopes

### Machine Scope

- Installs for all users on the computer
- Requires administrator privileges (UAC)
- Installed to `Program Files` or system locations

```json
"Scope": "machine"
```

### User Scope

- Installs for current user only
- No administrator privileges required
- Installed to user's AppData folder

```json
"Scope": "user"
```

## Sources

### WinGet

Standard Windows Package Manager. Use the WinGet package ID.

```powershell
# Find WinGet package IDs
winget search <name>
```

### Microsoft Store (msstore)

Windows Store applications. Use the Store Product ID (found in the Store URL).

Example URL: `https://www.microsoft.com/store/productId/9NBLGGH4QGH`
Product ID: `9NBLGGH4QGH`

### Office Deployment Tool (odt)

Microsoft Office products with custom configuration.

#### ODT Properties

| Property | Type | Description |
|----------|------|-------------|
| `Products` | array | Office product IDs (e.g., `ProPlus2024Volume`) |
| `Language` | string | Language code (e.g., `fr-fr`, `en-us`) |
| `OfficeClientEdition` | string | Architecture: `32` or `64` |
| `Channel` | string | Update channel (e.g., `PerpetualVL2024`, `Current`) |
| `ExcludeApps` | array | Apps to exclude (e.g., `["Access", "Publisher"]`) |
| `DisplayLevel` | string | Installation UI: `None` or `Full` |
| `AcceptEULA` | boolean | Auto-accept EULA |

## Detection Configuration

### Registry Name Pattern

Use wildcards to match program names in the registry:

```json
"Detection": {
    "Name": "Microsoft Office LTSC * 2024 - *"
}
```

### Custom Detection Script

Reference a PowerShell function that returns `$true` if installed:

```json
"DetectionScript": "Test-OfficeDeploymentTool"
```

The function must be available in the UDF folder.

## Dependencies

Use `Requires` to specify packages that must be installed before this package:

```json
{
    "Name": "Microsoft Office 2024",
    "Id": "Microsoft.Office2024",
    "Requires": ["Microsoft.OfficeDeploymentTool"]
}
```

The installer will:
1. Detect missing dependencies
2. Prompt the user to add them
3. Install dependencies before the main package

## Adding a New Package

1. Find the package ID:
   - WinGet: `winget search <name>`
   - Store: Check the Microsoft Store URL

2. Add the package to `apps.json`:
   ```json
   {
       "Category": "SystemTools",
       "Name": "My Application",
       "Id": "Publisher.AppName",
       "Scope": "machine",
       "Source": "winget"
   }
   ```

3. (Optional) Add an icon:
   - Save a PNG image to `input/icons/`
   - Name it `Publisher.AppName.png` (matching the Id)

4. Test the installation

## Package Icons

Icons are displayed in the GUI for each package.

- **Location**: `input/icons/`
- **Format**: PNG
- **Naming**: `{PackageId}.png` (e.g., `Git.Git.png`)
- **Recommended size**: 48x48 or 64x64 pixels

Icons are optional. Packages without icons display a default icon.
