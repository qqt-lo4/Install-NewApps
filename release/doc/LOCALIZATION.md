# Localization Guide

This guide explains how to add new languages to Install-NewApps.

## Supported Languages

| Language | Code | File |
|----------|------|------|
| English | `en-US` | `lang/en-US.json` |
| French | `fr-FR` | `lang/fr-FR.json` |

## File Structure

Translation files are located in:
```
input/lang/
├── en-US.json
└── fr-FR.json
```

## JSON Structure

Each translation file contains a JSON object with the following sections:

```json
{
  "LanguageName": "English",
  "UI": { ... },
  "Categories": { ... },
  "Settings": { ... }
}
```

### Sections

| Section | Description |
|---------|-------------|
| `LanguageName` | Display name of the language |
| `UI` | GUI window texts, buttons, messages |
| `Categories` | Package category names |
| `Settings` | Settings menu labels |

## Adding a New Language

### Step 1: Create the Translation File

Copy an existing file and rename it with the appropriate locale code:

```
lang/de-DE.json  # German
lang/es-ES.json  # Spanish
lang/pt-BR.json  # Brazilian Portuguese
```

### Step 2: Translate All Strings

Edit the JSON file and translate all values:

```json
{
  "LanguageName": "Deutsch",
  "UI": {
    "WindowTitle": "Software-Installationsmanager",
    "Categories": "Kategorien",
    ...
  }
}
```

### Step 3: Test the Translation

Run the script and change the language in settings to verify translations.

## Translation Syntax

### Simple Strings

```json
"WindowTitle": "Software Installation Manager"
```

### Strings with Parameters

Use `{0}`, `{1}`, etc. for parameter placeholders:

```json
"AllInstalledSuccessfully": "All software has been installed successfully! ({0}/{1})"
```

Usage in code:
```powershell
tr 'UI.AllInstalledSuccessfully' -Parameters @($successCount, $totalCount)
```

### Pluralization

Use pipe `|` to separate singular and plural forms:

```json
"InstallButton": "Install {0} software|Install {0} software"
```

The format is: `singular|plural`

Note: In English, both forms may be identical. In other languages like French:
```json
"InstallButton": "Installer {0} logiciel|Installer {0} logiciels"
```

### Escaped Characters

Use `\\n` for newlines in JSON:

```json
"LanguageChangeMessage": "The application needs to restart.\\n\\nRestart now?"
```

## Translation Key Reference

### UI Section

| Key | Description |
|-----|-------------|
| `WindowTitle` | Main window title |
| `Categories` | "Categories" label |
| `AllSoftware` | "All software" filter option |
| `NoSoftwareSelected` | Empty selection message |
| `InstallButton` | Install button text (with count) |
| `SoftwareSelected` | Selection counter text |
| `AlreadyInstalled` | Installed status label |
| `InstallationInProgress` | Progress window title |
| `PreparingInstallation` | Preparation message |
| `InstallationCompleted` | Completion dialog title |
| `AllInstalledSuccessfully` | Success message |
| `InstallationCompletedWithFailures` | Partial success message |
| `InstallationError` | Error message template |
| `Error` | "Error" label |
| `Initialization` | Initialization phase label |
| `MissingDependenciesHeader` | Dependencies dialog header |
| `AddAutomatically` | Button to add dependencies |
| `ContinueWithoutAdding` | Button to skip dependencies |
| `CancelInstallation` | Cancel button |
| `Language` | Language menu label |
| `LanguageChangeMessage` | Restart confirmation message |
| `Yes` / `No` | Yes/No button labels |

### Categories Section

| Key | Description |
|-----|-------------|
| `AudioVideo` | Audio & Video category |
| `Office` | Office category |
| `Development` | Development category |
| `Internet` | Internet category |
| `Games` | Games category |
| `SystemTools` | System Tools category |
| `Photo` | Photo category |

## Using Translations in Code

### Import the Module

```powershell
Import-Module $PSScriptRoot\UDF\PSSomeCoreThings
```

### Get a Translation

```powershell
# Simple string
$title = tr 'UI.WindowTitle'

# With parameters
$summary = tr 'UI.AllInstalledSuccessfully' -Parameters @($successCount, $totalCount)
```

## Language Detection

The application automatically detects the system language on first run:

1. Gets system locale via `Get-CurrentLocale`
2. Checks if a matching translation file exists
3. Falls back to `en-US` if not found

## Language Switching

Users can change the language via the settings menu in the GUI. The application will:

1. Prompt for restart confirmation
2. Save the language preference
3. Restart with the new language

## Best Practices

1. **Keep translations consistent** - Use the same terminology throughout
2. **Test all screens** - Verify translations appear correctly in all dialogs
3. **Handle long strings** - Some languages produce longer text; ensure UI accommodates
4. **Preserve placeholders** - Don't translate `{0}`, `{1}` etc.
5. **Maintain structure** - Keep the same JSON structure as the original file
