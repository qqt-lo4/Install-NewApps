# Architecture

This document describes the technical architecture of Install-NewApps.

## Overview

Install-NewApps is built as a modular PowerShell application with:

- A main orchestration script
- Reusable function modules (UDF)
- WPF-based graphical interface
- Multi-source package management
- Localization framework

## Architecture Diagram

```
┌─────────────────────────────────────────────────────────────────┐
│                     Install-NewApps.ps1                         │
│                    (Main Orchestrator)                          │
├─────────────────────────────────────────────────────────────────┤
│  ┌─────────────┐  ┌─────────────┐  ┌─────────────────────────┐  │
│  │   WPF GUI   │  │ Localization│  │  Package Installation   │  │
│  │             │  │             │  │                         │  │
│  │ - Selection │  │ - en-US     │  │ - WinGet                │  │
│  │ - Progress  │  │ - fr-FR     │  │ - Microsoft Store       │  │
│  │ - Dialogs   │  │             │  │ - ODT                   │  │
│  └─────────────┘  └─────────────┘  └─────────────────────────┘  │
├─────────────────────────────────────────────────────────────────┤
│                         UDF Modules                             │
│  ┌──────────┐ ┌──────────┐ ┌──────────┐ ┌──────────┐            │
│  │   GUI    │ │  Winget  │ │  Store   │ │  Office  │  ...       │
│  └──────────┘ └──────────┘ └──────────┘ └──────────┘            │
└─────────────────────────────────────────────────────────────────┘
```

## Components

### Main Script (Install-NewApps.ps1)

The main script handles:

1. **Initialization**
   - Load configuration
   - Import UDF modules
   - Initialize localization

2. **Package Discovery**
   - Load package catalog from JSON
   - Check installation status
   - Resolve dependencies

3. **User Interface**
   - Display package selection GUI
   - Handle user interactions
   - Show progress during installation

4. **Installation Orchestration**
   - Sort packages by dependencies
   - Batch machine-scope installations
   - Execute user-scope installations

### UDF Modules

Reusable functions organized by domain:

| Module | Purpose |
|--------|---------|
| `PSSomeAppsThings` | WinGet, Store, ODT, program detection |
| `PSSomeCoreThings` | Localization, script configuration |
| `PSSomeGUIThings` | WPF windows, dialogs, theme management |
| `PSSomeSystemThings` | System information, architecture, environment |
| `PSSomeDataThings` | Data operations |
| `PSSomeEngineThings` | Engine functions |
| `PSSomeFileThings` | File operations |
| `PSSqlite` | SQLite database access |
| `powershell-yaml` | YAML parsing |

## Installation Flow

```
┌──────────────────┐
│   Start Script   │
└────────┬─────────┘
         │
         ▼
┌──────────────────┐
│  Load Config &   │
│  Check Installed │
└────────┬─────────┘
         │
         ▼
┌──────────────────┐
│   Show GUI for   │
│ Package Selection│
└────────┬─────────┘
         │
         ▼
┌──────────────────┐
│    Resolve       │
│  Dependencies    │
└────────┬─────────┘
         │
         ▼
┌──────────────────┐
│  Sort by Scope   │
│ & Dependencies   │
└────────┬─────────┘
         │
         ▼
┌──────────────────────────────────────┐
│         Installation Phase           │
├──────────────────┬───────────────────┤
│  Machine Scope   │    User Scope     │
│  (Single UAC)    │  (No elevation)   │
│                  │                   │
│  - WinGet pkgs   │  - WinGet pkgs    │
│  - Store apps    │  - Store apps     │
│  - ODT products  │                   │
└──────────────────┴───────────────────┘
         │
         ▼
┌──────────────────┐
│  Show Results    │
│    Summary       │
└──────────────────┘
```

## Elevation Strategy

### Single UAC Prompt

Machine-scope packages are installed via a generated PowerShell script executed with elevation:

1. Generate a temporary script containing all installation commands
2. Execute with `Start-Process -Verb RunAs`
3. Single UAC prompt for all machine packages
4. Progress communicated via file-based IPC

```
User Process                    Elevated Process
     │                               │
     │  ──── UAC Prompt ────►        │
     │                               │
     │                        ┌──────┴──────┐
     │                        │ Install     │
     │                        │ Package 1   │
     │  ◄─── Progress ────    │             │
     │       (via file)       │ Install     │
     │                        │ Package 2   │
     │  ◄─── Progress ────    │             │
     │                        │ ...         │
     │                        └──────┬──────┘
     │  ◄─── Results ─────           │
     │       (via file)              │
```

### User Scope

User-scope packages are installed directly without elevation in the current process.

## Package Sources

### WinGet Integration

```
┌─────────────────┐     ┌─────────────────┐
│  WinGet SQLite  │────►│  Get Installer  │
│    Database     │     │     Info        │
└─────────────────┘     └────────┬────────┘
                                 │
                                 ▼
                        ┌─────────────────┐
                        │ Download & Run  │
                        │   Installer     │
                        └─────────────────┘
```

- Access WinGet's SQLite database for package metadata
- Download installers directly from source URLs
- Execute with appropriate silent switches

### Microsoft Store Integration

```
┌─────────────────┐     ┌───────────────────┐
│  Display        │────►│   Get Package     │
│  Catalog API    │     │   Manifest        │
└─────────────────┘     └─────────┬─────────┘
                                  │
                                  ▼
                        ┌───────────────────┐
                        │   FE3 API for     │
                        │   Download URLs   │
                        └─────────┬─────────┘
                                  │
                                  ▼
                        ┌───────────────────┐
                        │  Add-AppxPackage  │
                        │  or Provisioning  │
                        └───────────────────┘
```

- Query Display Catalog for app metadata
- Use FE3 API to get download URLs
- Install via AppX commands

### Office Deployment Tool

```
┌─────────────────┐     ┌─────────────────┐
│  ODT Config     │────►│  Generate XML   │
│  from JSON      │     │  Configuration  │
└─────────────────┘     └────────┬────────┘
                                 │
                                 ▼
                        ┌─────────────────┐
                        │  Execute ODT    │
                        │  setup.exe      │
                        └─────────────────┘
```

- Generate XML configuration from package definition
- Execute Office Deployment Tool with configuration

### Chocolatey

```
┌─────────────────┐     ┌─────────────────┐
│  Chocolatey     │────►│  choco install  │
│  (via WinGet)   │     │  <package> -y   │
└─────────────────┘     └─────────────────┘
```

- Chocolatey is installed automatically via WinGet as a prerequisite
- Packages use `choco install` with silent flags
- Used for packages not available on WinGet or the Microsoft Store

### Windows Features & Capabilities

```
┌─────────────────────┐     ┌────────────────────────────────┐
│  windowsfeature     │────►│  Enable-WindowsOptionalFeature │
│  (Hyper-V, WSL...)  │     │  -Online -All -NoRestart       │
└─────────────────────┘     └────────────────────────────────┘

┌─────────────────────┐     ┌──────────────────────────────────────────────┐
│  windowscapability  │────►│  dism.exe /Online /Add-Capability /NoRestart │
│  (RSAT tools...)    │     │    /CapabilityName:X /CapabilityName:Y ...   │
└─────────────────────┘     └──────────────────────────────────────────────┘
```

- Optional features: Hyper-V, Windows Sandbox, WSL, .NET Framework 3.5, Telnet Client
- On-demand capabilities: RSAT tools (Active Directory, DNS, DHCP, etc.)
- Supports prerequisite checks (edition, architecture, build number)
- All requested `windowscapability` packages are installed in **one** DISM call (single CBS transaction). `Get-WindowsCapability -Online` is queried once to resolve the full capability names. If the batch DISM call fails, the script falls back to per-capability `Add-WindowsCapability` so the failing item can be identified.

### Script

```
┌─────────────────┐     ┌─────────────────┐
│  InstallScript  │────►│  ScriptBlock    │
│  (from JSON)    │     │  execution      │
└─────────────────┘     └─────────────────┘
```

- Executes arbitrary PowerShell code defined in the `InstallScript` property
- Used for packages that need custom installation logic (e.g., Discord via `winget install`)
- Works in both user and machine scope (elevated script)

## GUI Architecture

### WPF in PowerShell

The GUI uses Windows Presentation Foundation (WPF) via PowerShell:

- XAML defined inline in PowerShell
- Event handlers in PowerShell script blocks
- Theme detection for light/dark mode

### Window Types

| Window | Purpose |
|--------|---------|
| `Show-PackageManagerUI` | Main package selection window |
| `Show-LoadingWindow` | Progress indicator during operations |
| `Show-WPFButtonDialog` | Modal dialogs for confirmations |

### Runspace Isolation

Dialogs use separate PowerShell runspaces to prevent UI blocking:

```
Main Thread                   Dialog Runspace
     │                             │
     │  ──── Create ────►          │
     │                        ┌────┴────┐
     │                        │  Show   │
     │                        │ Dialog  │
     │  ◄─── Poll ────        │         │
     │       Status           │  Wait   │
     │                        │  User   │
     │  ◄─── Result ────      │         │
     │                        └────┬────┘
     │                             │
```

## Dependency Resolution

### Topological Sort

Packages are sorted using topological ordering:

1. Build dependency graph
2. Detect circular dependencies
3. Sort: dependencies first, then dependents
4. Group by scope (machine before user)

```
Package A ──────► Package C
    │                 │
    ▼                 ▼
Package B ──────► Package D

Installation order: A, B, C, D
```

## Localization

### Translation Flow

```
┌─────────────────┐     ┌─────────────────┐
│  Get System     │────►│  Load Language  │
│  Locale         │     │  JSON File      │
└─────────────────┘     └────────┬────────┘
                                 │
                                 ▼
                        ┌─────────────────┐
                        │  tr('key')      │
                        │  Returns String │
                        └─────────────────┘
```

### File Structure

```
lang/
├── en-US.json    # English translations
└── fr-FR.json    # French translations
```

## Error Handling

### Installation Errors

- Caught per-package, don't stop other installations
- Results collected and displayed in summary
- Failed packages marked in results

### UAC Cancellation

- Detected via process exit code
- Returns special result to caller
- Displays user-friendly error dialog

## File-Based IPC

Communication between user and elevated processes:

| File | Purpose |
|------|---------|
| `progress.json` | Current installation progress |
| `results.json` | Final installation results |
| `script.ps1` | Generated installation script |

All temporary files are cleaned up after installation.
