#region Includes
Import-Module $PSScriptRoot\UDF\PSSomeAppsThings -WarningAction SilentlyContinue
Import-Module $PSScriptRoot\UDF\PSSomeCoreThings
Import-Module $PSScriptRoot\UDF\PSSomeDataThings
Import-Module $PSScriptRoot\UDF\PSSomeEngineThings -WarningAction SilentlyContinue
Import-Module $PSScriptRoot\UDF\PSSomeFileThings
Import-Module $PSScriptRoot\UDF\PSSomeGUIThings
Import-Module $PSScriptRoot\UDF\PSSomeSystemThings -WarningAction SilentlyContinue
Import-Module $PSScriptRoot\UDF\PSSQLite
Import-Module $PSScriptRoot\UDF\powershell-yaml
#endregion Includes

#region script info
#scriptType=standard
#scriptVersion=1.0
#outputMode=multiple
#outputMultipleChoices=qqt
#endregion script info

#region export script qqt
#New-OutputFolder "output\%scriptName%\%scriptVersion%\%target%\"
#New-ContentFolder "output\%scriptName%\%scriptVersion%\%target%\input\"
#Copy-ScriptContent "input\%scriptName%\*.json" "%outputDir%\input\"
#Copy-ScriptContent "input\%scriptName%\Install-NewApps.ico" "%outputDir%\input\"
#Copy-ScriptContent "input\%scriptName%\app-wide.png" "%outputDir%\input\"
#New-ContentFolder "%outputDir%\input\icons\"
#Copy-ScriptContent "input\%scriptName%\icons\*.*" "%outputDir%\input\icons\"
#New-ContentFolder "%outputDir%\input\lang\"
#Copy-ScriptContent "input\%scriptName%\lang\*.*" "%outputDir%\input\lang\"
#Write-OutputScript "%scriptFile%" "%outputDir%"
#New-PowershellScriptRunner "%scriptRoot%\%outputDir%\%scriptFileName%" "%scriptRoot%\%outputDir%\%scriptName%.exe" -Icon "%scriptRoot%\%outputDir%\input\%scriptName%.ico" -X64
#endregion export script qqt

function Install-OfficeWithODT {
    <#
    .SYNOPSIS
        Installs Microsoft Office using Office Deployment Tool

    .DESCRIPTION
        Creates an ODT configuration file from package ODT settings and executes
        Office Deployment Tool to install Office.

    .PARAMETER Package
        Package hashtable with ODT property containing configuration

    .PARAMETER TempPath
        Temporary directory for configuration file (default: $env:TEMP)

    .EXAMPLE
        Install-OfficeWithODT -Package $package
    #>

    [CmdletBinding()]
    Param(
        [Parameter(Mandatory=$true)]
        [hashtable]$Package,

        [Parameter(Mandatory=$false)]
        [string]$TempPath = $env:TEMP
    )

    # Get Office Deployment Tool path
    $odtPath = Get-OfficeDeploymentToolPath
    if (-not $odtPath) {
        throw "Office Deployment Tool not found. Please install it first."
    }

    Write-Verbose "Office Deployment Tool found at: $odtPath"

    # Create configuration file path
    $configFileName = "office-config-$($Package.Id).xml"
    $configPath = Join-Path $TempPath $configFileName

    Write-Host "  Creating ODT configuration file..." -ForegroundColor Gray

    # Extract ODT configuration parameters
    $odtConfig = $Package.ODT
    $products = $odtConfig.Products
    $language = if ($odtConfig.Language) { $odtConfig.Language } else { "fr-fr" }
    $officeEdition = if ($odtConfig.OfficeClientEdition) { $odtConfig.OfficeClientEdition } else { "64" }
    $channel = if ($odtConfig.Channel) { $odtConfig.Channel } else { "Current" }
    $excludeApps = $odtConfig.ExcludeApps
    $displayLevel = if ($odtConfig.DisplayLevel) { $odtConfig.DisplayLevel } else { "None" }
    $acceptEULA = if ($null -ne $odtConfig.AcceptEULA) { $odtConfig.AcceptEULA } else { $true }
    $pinIcons = if ($null -ne $odtConfig.PinIconsToTaskbar) { $odtConfig.PinIconsToTaskbar } else { $false }
    $autoActivate = if ($null -ne $odtConfig.AutoActivate) { $odtConfig.AutoActivate } else { $true }

    # Validate required parameters
    if (-not $products -or $products.Count -eq 0) {
        throw "ODT configuration must specify at least one product"
    }

    # Build New-OfficeDeploymentConfiguration parameters
    $configParams = @{
        Products = $products
        Language = $language
        OfficeClientEdition = $officeEdition
        Channel = $channel
        OutputPath = $configPath
        DisplayLevel = $displayLevel
        AcceptEULA = $acceptEULA
        PinIconsToTaskbar = $pinIcons
        AutoActivate = $autoActivate
    }

    # Add optional parameters
    if ($excludeApps -and $excludeApps.Count -gt 0) {
        $configParams.ExcludeApps = $excludeApps
    }

    # Generate configuration file
    try {
        $null = New-OfficeDeploymentConfiguration @configParams
        Write-Verbose "Configuration file created: $configPath"
    }
    catch {
        throw "Failed to create ODT configuration file: $_"
    }

    # Display configuration details
    Write-Host "  Product(s): $($products -join ', ')" -ForegroundColor Gray
    Write-Host "  Language: $language" -ForegroundColor Gray
    Write-Host "  Edition: ${officeEdition}-bit" -ForegroundColor Gray
    Write-Host "  Channel: $channel" -ForegroundColor Gray

    # Execute Office Deployment Tool
    Write-Host "  Running Office Deployment Tool..." -ForegroundColor Gray

    try {
        $arguments = @('/configure', $configPath)
        Write-Verbose "Executing: $odtPath /configure `"$configPath`""

        $process = Start-Process -FilePath $odtPath -ArgumentList $arguments -PassThru -WindowStyle Hidden

        # Wait for ODT process to complete
        while (-not $process.HasExited) {
            Start-Sleep -Milliseconds 500
        }

        if ($process.ExitCode -eq 0) {
            Write-Host "  Office installation completed successfully" -ForegroundColor Green
        }
        else {
            throw "Office Deployment Tool exited with code: $($process.ExitCode)"
        }
    }
    catch {
        throw "Failed to execute Office Deployment Tool: $_"
    }
    finally {
        # Cleanup configuration file
        if (Test-Path $configPath) {
            Remove-Item -Path $configPath -Force -ErrorAction SilentlyContinue
        }
    }
}

function Install-Package {
    <#
    .SYNOPSIS
        Downloads and installs a package
    
    .PARAMETER Package
        Package object with Source, Id, Name, Scope, and Installer properties
    
    .PARAMETER Credential
        Credentials for elevated installation (machine scope)
    
    .PARAMETER TempPath
        Temporary directory for downloads (default: $env:TEMP\PackageInstall)
    
    .EXAMPLE
        Install-Package -Package $pkg -Credential $cred
    #>
    
    [CmdletBinding()]
    Param(
        [Parameter(Mandatory=$true)]
        [hashtable]$Package,
        
        [Parameter(Mandatory=$false)]
        [System.Management.Automation.PSCredential]$Credential,
        
        [Parameter(Mandatory=$false)]
        [string]$TempPath = "$env:TEMP\PackageInstall"
    )
    
    Write-Host "Installing: $($Package.Name)" -ForegroundColor Cyan

    try {
        $script:LastInstallExitCode = 0
        switch ($Package.Source) {
            "winget" {
                if (-not $Package.Installer) {
                    throw "Installer information not found for package $($Package.Name)"
                }

                Install-WingetPackageFromInstaller -Package $Package -Credential $Credential -TempPath $TempPath
            }
            "msstore" {
                Install-MSStorePackage -Package $Package
            }
            "odt" {
                if (-not $Package.ODT) {
                    throw "ODT configuration not found for package $($Package.Name)"
                }

                Install-OfficeWithODT -Package $Package -TempPath $TempPath
            }
            default {
                throw "Unknown package source: $($Package.Source)"
            }
        }

        $reboot = ($script:LastInstallExitCode -eq 3010)
        if ($reboot) {
            Write-Host "✓ $($Package.Name) installed successfully (reboot required)" -ForegroundColor Yellow
        } else {
            Write-Host "✓ $($Package.Name) installed successfully" -ForegroundColor Green
        }
        return @{ Success = $true; RebootRequired = $reboot }
    }
    catch {
        Write-Error "✗ Failed to install $($Package.Name): $_"
        return @{ Success = $false; RebootRequired = $false }
    }
}

function Install-WingetPackageFromInstaller {
    <#
    .SYNOPSIS
        Downloads and installs a winget package using the Installer information
    
    .PARAMETER Package
        Package with Installer property containing URL and Silent arguments
    
    .PARAMETER Credential
        Credentials for elevated installation
    
    .PARAMETER TempPath
        Temporary directory for downloads
    #>
    
    [CmdletBinding()]
    Param(
        [Parameter(Mandatory=$true)]
        [hashtable]$Package,
        
        [Parameter(Mandatory=$false)]
        [System.Management.Automation.PSCredential]$Credential,
        
        [Parameter(Mandatory=$false)]
        [string]$TempPath = "$env:TEMP\PackageInstall"
    )
    
    # Create temp directory
    if (-not (Test-Path $TempPath)) {
        New-Item -Path $TempPath -ItemType Directory -Force | Out-Null
    }
    
    # Get installer info
    $installerUrl = $Package.Installer.URL
    $silentArgs = $Package.Installer.Silent
    $scope = $Package.Installer.Scope
    
    if (-not $installerUrl) {
        throw "Installer URL not found"
    }
    
    Write-Verbose "Installer URL: $installerUrl"
    Write-Verbose "Silent args: $silentArgs"
    Write-Verbose "Scope: $scope"
    
    # Determine file extension from URL
    $uri = [System.Uri]$installerUrl
    $fileName = [System.IO.Path]::GetFileName($uri.LocalPath)
    $extension = [System.IO.Path]::GetExtension($fileName).ToLower()
    
    # Download installer
    $installerPath = Join-Path $TempPath $fileName
    Write-Host "  Downloading installer..." -ForegroundColor Gray
    
    try {
        # Use WebClient for progress (Invoke-WebRequest is slow for large files)
        $webClient = New-Object System.Net.WebClient
        $webClient.DownloadFile($installerUrl, $installerPath)
        $webClient.Dispose()
    }
    catch {
        throw "Failed to download installer: $_"
    }
    
    if (-not (Test-Path $installerPath)) {
        throw "Downloaded file not found: $installerPath"
    }
    
    Write-Host "  Downloaded: $fileName" -ForegroundColor Gray

    # Check if this is a portable package BEFORE any extraction
    # A package is portable if InstallerType or NestedInstallerType is "portable"
    # These properties can be at different levels in the structure
    $installerType = $Package.Installer.InstallerType
    $nestedInstallerType = $Package.Installer.NestedInstallerType

    # Also check on Installers sub-object if not found
    if (-not $nestedInstallerType -and $Package.Installer.Installers) {
        $nestedInstallerType = $Package.Installer.Installers.NestedInstallerType
    }

    Write-Verbose "InstallerType: $installerType, NestedInstallerType: $nestedInstallerType"

    $isPortable = ($installerType -eq "portable") -or ($nestedInstallerType -eq "portable")

    if ($isPortable) {
        Write-Host "  Detected portable package (ZIP with portable content)" -ForegroundColor Cyan

        # Use dedicated portable installation function
        $installResult = Install-PortablePackage -Package $Package -InstallerPath $installerPath -TempPath $TempPath

        if (-not $installResult) {
            throw "Portable package installation failed"
        }

        # Clean up downloaded installer
        Write-Verbose "Cleaning up temporary installer..."
        if (Test-Path $installerPath) {
            Remove-Item -Path $installerPath -Force -ErrorAction SilentlyContinue
        }

        return
    }

    # NON-PORTABLE PACKAGES: Handle ZIP files containing installers
    if ($extension -eq ".zip") {
        Write-Host "  Extracting ZIP archive..." -ForegroundColor Gray

        $extractPath = Join-Path $TempPath "$($Package.Id)_extracted"

        if (Test-Path $extractPath) {
            Remove-Item -Path $extractPath -Recurse -Force
        }

        Expand-Archive -Path $installerPath -DestinationPath $extractPath -Force

        # Find installer in extracted files
        $installerExtensions = @(".exe", ".msi", ".msix", ".appx")
        $foundInstaller = Get-ChildItem -Path $extractPath -Recurse -File |
                         Where-Object { $installerExtensions -contains $_.Extension.ToLower() } |
                         Select-Object -First 1

        if (-not $foundInstaller) {
            throw "No installer found in ZIP archive"
        }

        $installerPath = $foundInstaller.FullName
        $extension = $foundInstaller.Extension.ToLower()
        Write-Host "  Found installer: $($foundInstaller.Name)" -ForegroundColor Gray
    }

    # Prepare installation command based on installer type
    $installCommand = $null
    $installArgs = @()
    
    switch ($extension) {
        ".msi" {
            $installCommand = "msiexec.exe"
            $installArgs = @("/i", "`"$installerPath`"")
            
            # Add silent args from package if provided
            if ($silentArgs) {
                $silentArgsList = $silentArgs -split ' ' | Where-Object { $_ -ne '' }
                $installArgs += $silentArgsList
            }
        }
        ".exe" {
            $installCommand = $installerPath
            
            # Use silent args from package if provided
            if ($silentArgs) {
                $installArgs = $silentArgs -split ' ' | Where-Object { $_ -ne '' }
            } else {
                # No silent args - run without arguments
                $installArgs = @()
            }
        }
        ".msix" {
            # Use Add-AppxPackage for MSIX
            Write-Host "  Installing MSIX package..." -ForegroundColor Gray
            Add-AppxPackage -Path $installerPath -ErrorAction Stop
            return
        }
        ".appx" {
            # Use Add-AppxPackage for APPX
            Write-Host "  Installing APPX package..." -ForegroundColor Gray
            Add-AppxPackage -Path $installerPath -ErrorAction Stop
            return
        }
        default {
            throw "Unsupported installer type: $extension"
        }
    }
    
    
    Write-Host "  Running installer..." -ForegroundColor Gray
    Write-Verbose "Command: $installCommand $($installArgs -join ' ')"

    # Run installer based on scope
    if ($scope -eq "machine") {
        if ($Credential) {
            # Install with provided credentials
            Write-Host "  Installing with provided credentials..." -ForegroundColor Gray
            Install-WithCredentials -Command $installCommand -Arguments $installArgs -Credential $Credential
        }
        else {
            # No credential - ALWAYS use UAC for machine scope
            Write-Host "  Installing with UAC (elevation required)..." -ForegroundColor Gray
            
            try {
                $process = Start-Process -FilePath $installCommand `
                                        -ArgumentList $installArgs `
                                        -Verb RunAs `
                                        -Wait `
                                        -PassThru `
                                        -WindowStyle Hidden
                
                if ($process.ExitCode -ne 0 -and $process.ExitCode -ne 3010) {
                    throw "Installation failed with exit code: $($process.ExitCode)"
                }
                $script:LastInstallExitCode = $process.ExitCode
            }
            catch {
                throw "Installation with UAC failed: $_"
            }
        }
    }
    else {
        # User scope installation - no elevation needed
        Write-Host "  Installation user scope..." -ForegroundColor Gray
        
        $process = Start-Process -FilePath $installCommand `
                                -ArgumentList $installArgs `
                                -Wait `
                                -PassThru `
                                -WindowStyle Hidden
        
        if ($process.ExitCode -ne 0 -and $process.ExitCode -ne 3010) {
            throw "Installation failed with exit code: $($process.ExitCode)"
        }
        $script:LastInstallExitCode = $process.ExitCode
    }

    # Clean up
    Write-Verbose "Cleaning up temporary files..."
    try {
        if (Test-Path $installerPath) {
            Remove-Item -Path $installerPath -Force -ErrorAction SilentlyContinue
        }
        
        # Remove extracted folder if it exists
        $extractPath = Join-Path $TempPath "$($Package.Id)_extracted"
        if (Test-Path $extractPath) {
            Remove-Item -Path $extractPath -Recurse -Force -ErrorAction SilentlyContinue
        }
    }
    catch {
        Write-Verbose "Failed to clean up: $_"
    }
}

function Update-Progress {
    <#
    .SYNOPSIS
        Updates progress information to a file for communication between processes

    .PARAMETER PackageName
        Name of the package being installed

    .PARAMETER Current
        Current package number

    .PARAMETER Total
        Total number of packages
    #>

    [CmdletBinding()]
    param(
        [string]$PackageName,
        [int]$Current,
        [int]$Total
    )

    $progress = @{
        PackageName = $PackageName
        Current = $Current
        Total = $Total
    }

    try {
        $progress | ConvertTo-Json | Set-Content -Path $progressFile -Force
    } catch {
        Write-Warning "Failed to update progress: $_"
    }
}

function Install-AllPackagesWithSingleUAC {
    <#
    .SYNOPSIS
        Install all packages with a single UAC prompt and progress updates
    
    .PARAMETER Packages
        Array of packages to install
    
    .PARAMETER LoadingWindow
        Loading window to update progress
    #>
    
    [CmdletBinding()]
    Param(
        [Parameter(Mandatory=$true)]
        [array]$Packages,
        
        [Parameter(Mandatory=$false)]
        $LoadingWindow
    )
    
    # Helper function to check if MSA token is authentic (not fallback)
    function Test-AuthenticMSAToken {
        $cacheFile = Join-Path $env:ProgramData "StoreLib\MSAToken.dat"

        if (-not (Test-Path $cacheFile)) {
            Write-Verbose "MSA token cache file does not exist"
            return $false
        }

        try {
            $cachedToken = Get-Content -Path $cacheFile -Raw -Encoding UTF8 -ErrorAction Stop
            $cachedToken = $cachedToken.Trim()
            $fallbackToken = Get-DefaultMSAToken

            if ($cachedToken -eq $fallbackToken) {
                Write-Verbose "MSA token is the fallback token (not authentic)"
                return $false
            }

            Write-Verbose "MSA token is authentic (not fallback)"
            return $true
        }
        catch {
            Write-Verbose "Failed to read MSA token cache: $_"
            return $false
        }
    }

    # Separate packages by scope
    $machinePackages = @($Packages | Where-Object { $_.Scope -eq "machine" })
    $userPackages = @($Packages | Where-Object { $_.Scope -ne "machine" })
    $hasUserStorePackages = @($userPackages | Where-Object { $_.Source -eq "msstore" }).Count -gt 0

    $results = @()
    $totalPackages = $Packages.Count
    $completedPackages = 0

    # Check if we need to retrieve authentic MSA token for user Store packages
    $needMSATokenRetrieval = $false
    if ($hasUserStorePackages -and $machinePackages.Count -eq 0) {
        # Only user Store packages, no machine packages
        # Check if we have an authentic token or just the fallback
        if (-not (Test-AuthenticMSAToken)) {
            Write-Host "Preparing Microsoft Store authentication (requires elevation)..." -ForegroundColor Yellow
            $needMSATokenRetrieval = $true
        }
    }

    # For machine packages OR MSA token retrieval, create a single elevated script
    if ($machinePackages.Count -gt 0 -or $needMSATokenRetrieval) {
        
        # UAC approach - create progress file for communication
        $tempProgressFile = [System.IO.Path]::GetTempFileName() + ".progress"
        $tempLogFile = [System.IO.Path]::GetTempFileName() + ".log"

        # Update initial progress
        if ($LoadingWindow) {
            Update-LoadingWindow -Window $LoadingWindow `
                                -Message "Launching elevated installation (UAC)..." `
                                -Progress ([int](($completedPackages / $totalPackages) * 100))
        }
        
        Write-Host "Preparing elevated installation for $($machinePackages.Count) package(s)..." -ForegroundColor Yellow
        
        $tempScript = [System.IO.Path]::GetTempFileName() + ".ps1"
        $tempResultFile = [System.IO.Path]::GetTempFileName() + ".json"

        # Get function code BEFORE building the here-string
        $updateProgressCode = Get-FunctionCode -FunctionName "Update-Progress"

        # Get ODT functions if needed
        $hasOdtPackages = $machinePackages | Where-Object { $_.Source -eq "odt" }
        $odtFunctionsCode = ""
        if ($hasOdtPackages) {
            $odtPathCode = Get-FunctionCode -FunctionName "Get-OfficeDeploymentToolPath"
            $odtConfigCode = Get-FunctionCode -FunctionName "New-OfficeDeploymentConfiguration"
            $odtInstallCode = Get-FunctionCode -FunctionName "Install-OfficeWithODT"
            $odtFunctionsCode = @"

# ODT Helper Functions
$odtPathCode

$odtConfigCode

$odtInstallCode

"@
        }

        # Check if there are user Store packages that will need MSA token
        $hasUserStorePackages = @($userPackages | Where-Object { $_.Source -eq "msstore" }).Count -gt 0

        # Get MSA token functions if needed
        $msaTokenFunctionsCode = ""
        if ($hasUserStorePackages) {
            $getDeviceTokenCode = Get-FunctionCode -FunctionName "Get-DeviceMSAToken"
            $updateTokenCode = Get-FunctionCode -FunctionName "Update-MSAToken"
            $invokeAsSystemCode = Get-FunctionCode -FunctionName "Invoke-AsSystem"
            $msaTokenFunctionsCode = @"

# MSA Token Helper Functions (for Store apps)
$getDeviceTokenCode

$updateTokenCode

$invokeAsSystemCode

"@
        }

        # Get portable package functions code
        $portableFunctionsCode = ""
        try {
            $portableFunctionsCode += Get-FunctionCode -FunctionName "Add-PathToEnvironment"
            $portableFunctionsCode += "`n"
            $portableFunctionsCode += Get-FunctionCode -FunctionName "Get-WingetPublisherId"
            $portableFunctionsCode += "`n"
            $portableFunctionsCode += Get-FunctionCode -FunctionName "Get-WingetPortablePackageFolderName"
            $portableFunctionsCode += "`n"
            $portableFunctionsCode += Get-FunctionCode -FunctionName "Install-PortablePackage"
        }
        catch {
            Write-Warning "Could not get portable functions code: $_"
        }

        # Get MS Store functions code for machine-scope store packages
        $hasMachineStorePackages = @($machinePackages | Where-Object { $_.Source -eq "msstore" }).Count -gt 0
        $msStoreFunctionsCode = ""
        if ($hasMachineStorePackages) {
            try {
                Write-Verbose "Injecting MS Store functions for elevated installation..."
                $storeFunctions = @(
                    "New-CorrelationVectorObject",
                    "New-LocaleObject",
                    "Get-DCatEndpointUrl",
                    "Get-CurrentMSAToken",
                    "Get-DefaultMSAToken",
                    "Get-DeviceMSAToken",
                    "Get-FE3Cookie",
                    "Get-FE3UpdateIDs",
                    "Get-FE3FileUrls",
                    "Invoke-MSHttpRequest",
                    "Invoke-DisplayCatalogQuery",
                    "Invoke-DisplayCatalogQueryOverload",
                    "Invoke-FE3SyncUpdates",
                    "Invoke-PackageManifestQuery",
                    "Invoke-SystemTokenExtraction",
                    "Update-MSAToken",
                    "Filter-PackagesByArchitecture",
                    "Get-StoreAppManifest",
                    "Get-StoreAppInfo",
                    "Get-MsStoreWin32Installer",
                    "Get-UnifiedStoreAppInfo",
                    "Get-InstalledPrograms",
                    "ConvertTo-Guid",
                    "Install-MSStoreAppx",
                    "Install-MSStoreWin32",
                    "Get-LocaleFormats",
                    "Install-MSStoreApp",
                    "Get-SystemArchitecture"
                )
                $msStoreFunctionsCode = "`n# MS Store Installation Functions`n"
                foreach ($funcName in $storeFunctions) {
                    $msStoreFunctionsCode += Get-FunctionCode -FunctionName $funcName
                    $msStoreFunctionsCode += "`n"
                }
            }
            catch {
                Write-Warning "Could not get MS Store functions code: $_"
            }
        }

        # Build installation script with progress updates
        $scriptContent = @"
# Elevated installation script
`$logFile = '$tempLogFile'
`$results = @()
`$ErrorActionPreference = 'Continue'
`$progressFile = '$tempProgressFile'
`$totalMachinePackages = $($machinePackages.Count)
`$currentMachinePackage = 0

# Write-Log: writes to log file with timestamp and to console
function Write-Log {
    param([string]`$Message, [string]`$ForegroundColor = 'White')
    `$timestamp = Get-Date -Format "HH:mm:ss"
    `$logMessage = "[`$timestamp] `$Message"
    Add-Content -Path `$logFile -Value `$logMessage -Encoding UTF8
    Write-Host `$Message -ForegroundColor `$ForegroundColor
}

# Write-HostAndLog: writes to log file (without timestamp) and to console
function Write-HostAndLog {
    param([string]`$Message, [string]`$ForegroundColor = 'White')
    Add-Content -Path `$logFile -Value `$Message -Encoding UTF8
    Write-Host `$Message -ForegroundColor `$ForegroundColor
}

# Initialize log file
Set-Content -Path `$logFile -Value "" -Encoding UTF8

$updateProgressCode
$odtFunctionsCode
$msaTokenFunctionsCode
$portableFunctionsCode
$msStoreFunctionsCode
"@
        
        $packageIndex = 0
        foreach ($package in $machinePackages) {
            $packageIndex++
            $installerUrl = $package.Installer.URL
            $silentArgs = $package.Installer.Silent
            $packageId = $package.Id
            $packageName = $package.Name
            $packageSource = $package.Source
            $packageScope = $package.Scope

            # Detect if this is a portable package
            $installerType = $package.Installer.InstallerType
            $nestedInstallerType = $package.Installer.NestedInstallerType
            if (-not $nestedInstallerType -and $package.Installer.Installers) {
                $nestedInstallerType = $package.Installer.Installers.NestedInstallerType
            }
            $isPortablePackage = ($installerType -eq "portable") -or ($nestedInstallerType -eq "portable")

            # Get architecture for portable packages
            $pkgArchitecture = $package.Installer.Architecture
            if (-not $pkgArchitecture -and $package.Installer.Installers) {
                $pkgArchitecture = $package.Installer.Installers.Architecture
            }

            # Get NestedInstallerFiles for portable packages
            $nestedInstallerFiles = $package.Installer.NestedInstallerFiles
            if (-not $nestedInstallerFiles -and $package.Installer.Installers) {
                $nestedInstallerFiles = $package.Installer.Installers.NestedInstallerFiles
            }
            $relativePath = ""
            if ($nestedInstallerFiles) {
                if ($nestedInstallerFiles -is [hashtable]) {
                    $relativePath = $nestedInstallerFiles["RelativeFilePath"]
                } elseif ($nestedInstallerFiles -is [array] -and $nestedInstallerFiles.Count -gt 0) {
                    $relativePath = $nestedInstallerFiles[0].RelativeFilePath
                    if (-not $relativePath) { $relativePath = $nestedInstallerFiles[0]["RelativeFilePath"] }
                } elseif ($nestedInstallerFiles.RelativeFilePath) {
                    $relativePath = $nestedInstallerFiles.RelativeFilePath
                }
            }

            if ($packageSource -eq "winget") {
                $escapedUrl = $installerUrl -replace "'", "''"
                $escapedSilent = $silentArgs -replace "'", "''"
                $escapedPackageName = $packageName -replace "'", "''"
                $escapedRelativePath = $relativePath -replace "'", "''"

                $scriptContent += @"

`$currentMachinePackage++
Update-Progress -PackageName '$escapedPackageName' -Current `$currentMachinePackage -Total `$totalMachinePackages

Write-HostAndLog "Installing: $packageName" -ForegroundColor Cyan

try {
    # Download installer
    `$tempPath = "`$env:TEMP\PackageInstall"
    if (-not (Test-Path `$tempPath)) {
        New-Item -Path `$tempPath -ItemType Directory -Force | Out-Null
    }

    `$installerUrl = '$escapedUrl'
    Write-HostAndLog "  URL: `$installerUrl" -ForegroundColor Gray
    `$uri = [System.Uri]`$installerUrl
    `$fileName = [System.IO.Path]::GetFileName(`$uri.LocalPath)

    # Handle URLs where filename cannot be determined from path
    if (`$fileName -eq 'download' -or [string]::IsNullOrWhiteSpace(`$fileName)) {
        # Try to extract filename from anywhere in the URL (path or query parameters)
        if (`$installerUrl -match '([^/?&=]+\.(exe|msi|msix|appx|zip))') {
            `$fileName = `$matches[1]
            Write-HostAndLog "  Extracted filename from URL: `$fileName" -ForegroundColor Gray
        }
        else {
            # Default to download.exe if we can't determine
            `$fileName = 'download.exe'
            Write-HostAndLog "  Warning: Could not determine filename, using default: `$fileName" -ForegroundColor Yellow
        }
    }

    `$installerPath = Join-Path `$tempPath `$fileName

    Write-HostAndLog "  Downloading..." -ForegroundColor Gray
    try {
        # Enable all TLS versions for maximum compatibility
        [System.Net.ServicePointManager]::SecurityProtocol = [System.Net.SecurityProtocolType]::Tls -bor [System.Net.SecurityProtocolType]::Tls11 -bor [System.Net.SecurityProtocolType]::Tls12 -bor [System.Net.SecurityProtocolType]::Tls13

        # Determine appropriate User-Agent based on URL
        `$userAgent = if (`$installerUrl -match 'sourceforge\.net') {
            # SourceForge requires wget-like User-Agent for direct download
            "Wget"
        } else {
            # Other sources use Winget User-Agent
            "Winget-CLI/1.7.10514"
        }

        `$ProgressPreference = 'SilentlyContinue'
        `$response = Invoke-WebRequest -Uri `$installerUrl -OutFile `$installerPath -UserAgent `$userAgent -UseBasicParsing -PassThru
        `$ProgressPreference = 'Continue'

        # Try to get filename from Content-Disposition header (most reliable)
        if (`$response.Headers['Content-Disposition']) {
            `$contentDisposition = `$response.Headers['Content-Disposition']
            if (`$contentDisposition -match 'filename\s*=\s*"?([^";]+)"?') {
                `$headerFileName = `$matches[1]
                if (`$headerFileName -and `$headerFileName -ne `$fileName) {
                    Write-HostAndLog "  Found filename in Content-Disposition header: `$headerFileName" -ForegroundColor Gray
                    # Rename the file to match the header
                    `$newInstallerPath = Join-Path `$tempPath `$headerFileName
                    if (Test-Path `$installerPath) {
                        Move-Item -Path `$installerPath -Destination `$newInstallerPath -Force
                        `$installerPath = `$newInstallerPath
                        `$fileName = `$headerFileName
                    }
                }
            }
        }
    }
    catch {
        Write-HostAndLog "  Download failed: `$(`$_.Exception.Message)" -ForegroundColor Red
        throw
    }

    # Verify file was downloaded
    if (-not (Test-Path `$installerPath)) {
        throw "Downloaded file does not exist: `$installerPath"
    }

    `$fileSize = (Get-Item `$installerPath).Length
    Write-HostAndLog "  File downloaded: `$fileName (`$fileSize bytes)" -ForegroundColor Gray

    `$extension = [System.IO.Path]::GetExtension(`$fileName).ToLower()
    Write-HostAndLog "  Extension detected: `$extension" -ForegroundColor Gray

    # Handle ZIP extraction if needed
    `$portableInstalled = `$false
    if (`$extension -eq '.zip') {
        # Check if this is a portable package
        `$isPortablePackage = $(if ($isPortablePackage) { '$true' } else { '$false' })
        if (`$isPortablePackage) {
            Write-HostAndLog "  Installing portable package..." -ForegroundColor Gray
            # Build Package hashtable for Install-PortablePackage
            `$pkg = @{
                Id = '$packageId'
                Name = '$escapedPackageName'
                Installer = @{
                    Scope = 'machine'
                    Architecture = '$pkgArchitecture'
                    NestedInstallerFiles = @{
                        RelativeFilePath = '$escapedRelativePath'
                    }
                }
            }
            `$success = Install-PortablePackage -Package `$pkg -InstallerPath `$installerPath -TempPath `$tempPath
            if (`$success) {
                Write-HostAndLog "  Success!" -ForegroundColor Green
                `$results += [pscustomobject]@{Name='$packageName'; Id='$packageId'; Source='$packageSource'; Success=`$true; RebootRequired=`$false}
            } else {
                Write-HostAndLog "  Failed to install portable package" -ForegroundColor Red
                `$results += [pscustomobject]@{Name='$packageName'; Id='$packageId'; Source='$packageSource'; Success=`$false; RebootRequired=`$false}
            }
            Remove-Item -Path `$installerPath -Force -ErrorAction SilentlyContinue
            `$portableInstalled = `$true
        }
        else {
            Write-HostAndLog "  Extracting ZIP..." -ForegroundColor Gray
            `$extractPath = Join-Path `$tempPath "$packageId`_extracted"
            if (Test-Path `$extractPath) { Remove-Item -Path `$extractPath -Recurse -Force }
            Expand-Archive -Path `$installerPath -DestinationPath `$extractPath -Force
            `$installerExtensions = @('.exe', '.msi', '.msix', '.appx')
            `$foundInstaller = Get-ChildItem -Path `$extractPath -Recurse -File | Where-Object { `$installerExtensions -contains `$_.Extension.ToLower() } | Select-Object -First 1
            if (`$foundInstaller) {
                `$installerPath = `$foundInstaller.FullName
                `$extension = `$foundInstaller.Extension.ToLower()
                Write-HostAndLog "  Installer extracted: `$(`$foundInstaller.Name)" -ForegroundColor Gray
            }
            else {
                throw "No installer found in ZIP archive"
            }
        }
    }

    # Skip installer execution if portable package was already installed
    if (-not `$portableInstalled) {

    # Handle MSIX/APPX separately (provision for all users in machine scope)
    if (`$extension -eq '.msix' -or `$extension -eq '.appx') {
        Write-HostAndLog "  Provisioning MSIX/APPX package for all users..." -ForegroundColor Gray
        Add-AppxProvisionedPackage -Online -PackagePath `$installerPath -SkipLicense -ErrorAction Stop
        Write-HostAndLog "  Success!" -ForegroundColor Green
        `$results += [pscustomobject]@{Name='$packageName'; Id='$packageId'; Source='$packageSource'; Success=`$true; RebootRequired=`$false}
        Remove-Item -Path `$installerPath -Force -ErrorAction SilentlyContinue
    }
    else {
        # Prepare command for MSI/EXE
        `$installCommand = `$null
        `$installArgs = @()

        if (`$extension -eq '.msi') {
            `$installCommand = 'msiexec.exe'
            `$installArgs = @('/i')
            `$installArgs += `$installerPath
            if ('$escapedSilent') {
                `$silentArgsList = '$escapedSilent' -split ' ' | Where-Object { `$_ -ne '' }
                `$installArgs += `$silentArgsList
            }
        }
        elseif (`$extension -eq '.exe') {
            `$installCommand = `$installerPath
            if ('$escapedSilent') {
                `$installArgs = '$escapedSilent' -split ' ' | Where-Object { `$_ -ne '' }
            }
        }
        else {
            throw "Unsupported extension: `$extension"
        }

        # Verify install command was set
        if ([string]::IsNullOrWhiteSpace(`$installCommand)) {
            throw "Install command was not set for extension `$extension"
        }

        # Run installer
        Write-HostAndLog "  Running installer..." -ForegroundColor Gray
        Write-HostAndLog "  Command: `$installCommand" -ForegroundColor Gray
        Write-HostAndLog "  Args: `$(`$installArgs -join ' ')" -ForegroundColor Gray
        `$process = Start-Process -FilePath `$installCommand -ArgumentList `$installArgs -PassThru -WindowStyle Hidden

        # Wait for the installer process itself to exit (not child processes)
        while (-not `$process.HasExited) {
            Start-Sleep -Milliseconds 500
        }

        if (`$process.ExitCode -eq 0 -or `$process.ExitCode -eq 3010) {
            `$reboot = (`$process.ExitCode -eq 3010)
            if (`$reboot) { Write-HostAndLog "  Success! (reboot required)" -ForegroundColor Yellow }
            else { Write-HostAndLog "  Success!" -ForegroundColor Green }
            `$results += [pscustomobject]@{Name='$packageName'; Id='$packageId'; Source='$packageSource'; Success=`$true; RebootRequired=`$reboot}
        }
        else {
            Write-HostAndLog "  Failed with exit code: `$(`$process.ExitCode)" -ForegroundColor Red
            `$results += [pscustomobject]@{Name='$packageName'; Id='$packageId'; Source='$packageSource'; Success=`$false; RebootRequired=`$false}
        }

        # Cleanup
        Remove-Item -Path `$installerPath -Force -ErrorAction SilentlyContinue
    }

    } # End of if (-not $portableInstalled)
}
catch {
    Write-HostAndLog "  Error: `$(`$_.Exception.Message)" -ForegroundColor Red
    Write-HostAndLog "  Error type: `$(`$_.Exception.GetType().FullName)" -ForegroundColor Red
    if (`$_.Exception.InnerException) {
        Write-HostAndLog "  Inner error: `$(`$_.Exception.InnerException.Message)" -ForegroundColor Red
    }
    `$results += [pscustomobject]@{Name='$packageName'; Id='$packageId'; Source='$packageSource'; Success=`$false; RebootRequired=`$false}
}

"@
            }
            elseif ($packageSource -eq "msstore") {
                $scriptContent += @"

`$currentMachinePackage++
Update-Progress -PackageName '$packageName' -Current `$currentMachinePackage -Total `$totalMachinePackages

Write-HostAndLog "Installing from Microsoft Store: $packageName" -ForegroundColor Cyan
Write-HostAndLog "  Package ID: $packageId" -ForegroundColor Gray
Write-HostAndLog "  Scope: $packageScope" -ForegroundColor Gray

try {
    # Build package object for Install-MSStoreApp
    `$storePackage = @{
        Id = '$packageId'
        Name = '$packageName'
        Scope = '$packageScope'
    }

    # Set default architecture for Store API calls (required by Install-MSStoreApp)
    `$Architecture = 'Autodetect'
    #`$VerbosePreference = 'Continue'
    `$success = Install-MSStoreApp -Package `$storePackage

    if (`$success) {
        Write-HostAndLog "  Success!" -ForegroundColor Green
        `$results += [pscustomobject]@{Name='$packageName'; Id='$packageId'; Source='$packageSource'; Success=`$true; RebootRequired=`$false}
    }
    else {
        Write-HostAndLog "  Installation returned false" -ForegroundColor Red
        `$results += [pscustomobject]@{Name='$packageName'; Id='$packageId'; Source='$packageSource'; Success=`$false; RebootRequired=`$false}
    }
}
catch {
    Write-HostAndLog "  Error: `$(`$_.Exception.Message)" -ForegroundColor Red
    Write-HostAndLog "  Error type: `$(`$_.Exception.GetType().FullName)" -ForegroundColor Red
    Write-HostAndLog "  Stack trace: `$(`$_.ScriptStackTrace)" -ForegroundColor DarkRed
    if (`$_.Exception.InnerException) {
        Write-HostAndLog "  Inner error: `$(`$_.Exception.InnerException.Message)" -ForegroundColor Red
    }
    `$results += [pscustomobject]@{Name='$packageName'; Id='$packageId'; Source='$packageSource'; Success=`$false; RebootRequired=`$false}
}

"@
            }
            elseif ($packageSource -eq "odt") {
                # Convert ODT configuration to JSON for embedding
                $odtConfigJson = $package.ODT | ConvertTo-Json -Compress -Depth 10
                $odtConfigJson = $odtConfigJson -replace "'", "''"

                $scriptContent += @"

`$currentMachinePackage++
Update-Progress -PackageName '$packageName' -Current `$currentMachinePackage -Total `$totalMachinePackages

Write-HostAndLog "Installing with Office Deployment Tool: $packageName" -ForegroundColor Cyan

try {
    # Reconstruct package object with ODT configuration
    `$odtConfig = '$odtConfigJson' | ConvertFrom-Json
    `$package = @{
        Id = '$packageId'
        Name = '$packageName'
        Source = 'odt'
        ODT = @{
            Products = @(`$odtConfig.Products)
            Language = `$odtConfig.Language
            OfficeClientEdition = `$odtConfig.OfficeClientEdition
            Channel = `$odtConfig.Channel
            DisplayLevel = `$odtConfig.DisplayLevel
            AcceptEULA = `$odtConfig.AcceptEULA
            PinIconsToTaskbar = `$odtConfig.PinIconsToTaskbar
            AutoActivate = `$odtConfig.AutoActivate
            ExcludeApps = if (`$odtConfig.ExcludeApps) { @(`$odtConfig.ExcludeApps) } else { `$null }
        }
    }

    Install-OfficeWithODT -Package `$package
    Write-HostAndLog "  Success!" -ForegroundColor Green
    `$results += [pscustomobject]@{Name='$packageName'; Id='$packageId'; Source='$packageSource'; Success=`$true; RebootRequired=`$false}
}
catch {
    Write-HostAndLog "  Error: `$(`$_.Exception.Message)" -ForegroundColor Red
    Write-HostAndLog "  Error type: `$(`$_.Exception.GetType().FullName)" -ForegroundColor Red
    if (`$_.Exception.InnerException) {
        Write-HostAndLog "  Inner error: `$(`$_.Exception.InnerException.Message)" -ForegroundColor Red
    }
    `$results += [pscustomobject]@{Name='$packageName'; Id='$packageId'; Source='$packageSource'; Success=`$false; RebootRequired=`$false}
}

"@
            }
            elseif ($packageSource -eq "windowscapability") {
                $scriptContent += @"

`$currentMachinePackage++
Update-Progress -PackageName '$packageName' -Current `$currentMachinePackage -Total `$totalMachinePackages

Write-HostAndLog "Installing Windows Capability: $packageName" -ForegroundColor Cyan
Write-HostAndLog "  Capability: $packageId" -ForegroundColor Gray

try {
    # Find the full capability name (version suffix varies)
    `$capability = Get-WindowsCapability -Online | Where-Object { `$_.Name -like '$packageId*' } | Select-Object -First 1

    if (-not `$capability) {
        throw "Windows Capability '$packageId' not found on this system"
    }

    if (`$capability.State -eq 'Installed') {
        Write-HostAndLog "  Already installed" -ForegroundColor Yellow
        `$results += [pscustomobject]@{Name='$packageName'; Id='$packageId'; Source='$packageSource'; Success=`$true; RebootRequired=`$false}
    }
    else {
        Write-HostAndLog "  Adding capability: `$(`$capability.Name)" -ForegroundColor Gray
        `$capResult = Add-WindowsCapability -Online -Name `$capability.Name -ErrorAction Stop
        `$reboot = `$capResult.RestartNeeded
        if (`$reboot) { Write-HostAndLog "  Success! (reboot required)" -ForegroundColor Yellow }
        else { Write-HostAndLog "  Success!" -ForegroundColor Green }
        `$results += [pscustomobject]@{Name='$packageName'; Id='$packageId'; Source='$packageSource'; Success=`$true; RebootRequired=`$reboot}
    }
}
catch {
    Write-HostAndLog "  Error: `$(`$_.Exception.Message)" -ForegroundColor Red
    Write-HostAndLog "  Error type: `$(`$_.Exception.GetType().FullName)" -ForegroundColor Red
    if (`$_.Exception.InnerException) {
        Write-HostAndLog "  Inner error: `$(`$_.Exception.InnerException.Message)" -ForegroundColor Red
    }
    `$results += [pscustomobject]@{Name='$packageName'; Id='$packageId'; Source='$packageSource'; Success=`$false; RebootRequired=`$false}
}

"@
            }
            elseif ($packageSource -eq "windowsfeature") {
                $scriptContent += @"

`$currentMachinePackage++
Update-Progress -PackageName '$packageName' -Current `$currentMachinePackage -Total `$totalMachinePackages

Write-HostAndLog "Enabling Windows Optional Feature: $packageName" -ForegroundColor Cyan
Write-HostAndLog "  Feature: $packageId" -ForegroundColor Gray

try {
    `$feature = Get-WindowsOptionalFeature -Online -FeatureName '$packageId' -ErrorAction Stop

    if (-not `$feature) {
        throw "Windows Optional Feature '$packageId' not found on this system"
    }

    if (`$feature.State -eq 'Enabled') {
        Write-HostAndLog "  Already enabled" -ForegroundColor Yellow
        `$results += [pscustomobject]@{Name='$packageName'; Id='$packageId'; Source='$packageSource'; Success=`$true; RebootRequired=`$false}
    }
    else {
        Write-HostAndLog "  Enabling feature..." -ForegroundColor Gray
        `$featureResult = Enable-WindowsOptionalFeature -Online -FeatureName '$packageId' -All -NoRestart -ErrorAction Stop
        `$reboot = `$featureResult.RestartNeeded
        if (`$reboot) { Write-HostAndLog "  Success! (reboot required)" -ForegroundColor Yellow }
        else { Write-HostAndLog "  Success!" -ForegroundColor Green }
        `$results += [pscustomobject]@{Name='$packageName'; Id='$packageId'; Source='$packageSource'; Success=`$true; RebootRequired=`$reboot}
    }
}
catch {
    Write-HostAndLog "  Error: `$(`$_.Exception.Message)" -ForegroundColor Red
    Write-HostAndLog "  Error type: `$(`$_.Exception.GetType().FullName)" -ForegroundColor Red
    if (`$_.Exception.InnerException) {
        Write-HostAndLog "  Inner error: `$(`$_.Exception.InnerException.Message)" -ForegroundColor Red
    }
    `$results += [pscustomobject]@{Name='$packageName'; Id='$packageId'; Source='$packageSource'; Success=`$false; RebootRequired=`$false}
}

"@
            }
        }

        # Add MSA token preparation if needed (convert boolean to string)
        if ($hasUserStorePackages) {
            $scriptContent += @"

# Prepare MSA token for user Store packages
Write-HostAndLog "Preparing Microsoft Store token for user packages..." -ForegroundColor Yellow
try {
    `$token = Get-DeviceMSAToken -Force
    if (`$token) {
        Write-HostAndLog "  MSA token retrieved and cached successfully" -ForegroundColor Green
    } else {
        Write-HostAndLog "  Warning: Failed to retrieve MSA token. User Store packages may fail to install." -ForegroundColor Yellow
    }
} catch {
    Write-HostAndLog "  Warning: Error retrieving MSA token: `$_" -ForegroundColor Yellow
}

"@
        }

        $scriptContent += @"
# Save results
try {
    `$jsonContent = ConvertTo-Json -InputObject @(`$results) -Depth 10
    `$jsonContent | Set-Content -Path '$tempResultFile' -Encoding UTF8 -Force
} catch {
    Write-Error "Failed to save results: `$_"
}

# Clean up progress file
Remove-Item -Path `$progressFile -Force -ErrorAction SilentlyContinue

Write-HostAndLog "Installation completed" -ForegroundColor Green
Start-Sleep -Seconds 2

"@
        
        Set-Content -Path $tempScript -Value $scriptContent -Encoding UTF8
        
        Write-Host "Launching elevated installation (UAC prompt)..." -ForegroundColor Yellow

        # Start elevated process
        $processInfo = New-Object System.Diagnostics.ProcessStartInfo
        $processInfo.FileName = "powershell.exe"
        $processInfo.Arguments = "-NoProfile -ExecutionPolicy Bypass -File `"$tempScript`""
        $processInfo.Verb = "RunAs"
        $processInfo.UseShellExecute = $true
        $processInfo.WindowStyle = [System.Diagnostics.ProcessWindowStyle]::Hidden

        try {
            $elevatedProcess = [System.Diagnostics.Process]::Start($processInfo)
        }
        catch {
            # UAC was cancelled or elevation failed
            $errorMessage = $_.Exception.Message

            # Clean up temp files
            Remove-Item -Path $tempScript -Force -ErrorAction SilentlyContinue
            #Remove-Item -Path $tempResultFile -Force -ErrorAction SilentlyContinue
            #Remove-Item -Path $tempProgressFile -Force -ErrorAction SilentlyContinue
            #Remove-Item -Path $tempLogFile -Force -ErrorAction SilentlyContinue

            # Return special result indicating UAC cancellation
            # The calling function will handle closing LoadingWindow and showing the dialog
            return @(@{
                Name = "UAC_CANCELLED"
                Success = $false
                ErrorMessage = $errorMessage
                UACCancelled = $true
            })
        }

        if (-not $elevatedProcess) {
            throw "Failed to start elevated process"
        }
        
        Write-Host "Elevated process started (PID: $($elevatedProcess.Id))" -ForegroundColor Green
        Write-Host "Surveillance de la progression..." -ForegroundColor Gray
        
        # Monitor progress file
        $machinePackagesProcessed = 0
        $lastUpdate = Get-Date
        $noProgressTimeout = 300  # 5 minutes
        
        while (-not $elevatedProcess.HasExited) {
            # Read transcript file for details display
            $transcriptContent = Read-FileNonBlocking -Path $tempLogFile

            if (Test-Path $tempProgressFile) {
                try {
                    $progressContent = Get-Content -Path $tempProgressFile -Raw -ErrorAction SilentlyContinue

                    if ($progressContent) {
                        $progressData = $progressContent | ConvertFrom-Json

                        $totalCurrentPackage = $completedPackages + $progressData.Current
                        $percentComplete = [int](($totalCurrentPackage / $totalPackages) * 100)

                        if ($LoadingWindow) {
                            Update-LoadingWindow -Window $LoadingWindow `
                                                -Message "Installation : $($progressData.PackageName) ($totalCurrentPackage/$totalPackages)..." `
                                                -Progress $percentComplete `
                                                -Details $transcriptContent
                        }

                        if ($progressData.Current -ne $machinePackagesProcessed) {
                            Write-Host "  Package en cours: $($progressData.PackageName) ($($progressData.Current)/$($progressData.Total))" -ForegroundColor Cyan
                            $machinePackagesProcessed = $progressData.Current
                            $lastUpdate = Get-Date
                        }
                    }
                }
                catch {
                    Write-Verbose "Error reading progression file: $_"
                }
            }
            elseif ($transcriptContent -and $LoadingWindow) {
                # Even without progress file, update details if transcript exists
                Update-LoadingWindow -Window $LoadingWindow -Details $transcriptContent
            }

            # Check for timeout
            $elapsed = (Get-Date) - $lastUpdate
            if ($elapsed.TotalSeconds -gt $noProgressTimeout) {
                Write-Warning "Timeout: aucune progression depuis $noProgressTimeout secondes"
                break
            }

            Start-Sleep -Milliseconds 500
        }
        
        Write-Host "Elevated process completed" -ForegroundColor Green

        # Wait for file to be written
        Start-Sleep -Seconds 2

        # Display elevated process log
        if (Test-Path $tempLogFile) {
            Write-Host "`n" -NoNewline
            Write-Host ("=" * 60) -ForegroundColor Yellow
            Write-Host "Elevated process output:" -ForegroundColor Yellow
            Write-Host ("=" * 60) -ForegroundColor Yellow
            Get-Content -Path $tempLogFile | ForEach-Object { Write-Host $_ }
            Write-Host ("=" * 60) -ForegroundColor Yellow
            Write-Host "`n" -NoNewline
        }

        # Read final results
        if (Test-Path $tempResultFile) {
            Write-Host "Reading results..." -ForegroundColor Gray

            try {
                $resultContent = Get-Content -Path $tempResultFile -Raw -ErrorAction Stop

                # Use -InputObject instead of piping to fix PowerShell 5.1 array parsing bug
                $parsed = ConvertFrom-Json -InputObject $resultContent
                # Force array enumeration for proper element handling
                $elevatedResults = @()
                foreach ($item in $parsed) {
                    $elevatedResults += $item
                }

                foreach ($result in $elevatedResults) {
                    $completedPackages++
                    
                    $oNewResult = [PSCustomObject]@{
                        Name = $result.Name
                        Id = $result.Id
                        Source = $result.Source
                        Success = $result.Success
                        RebootRequired = [bool]$result.RebootRequired
                    }
                    $results += $oNewResult

                    if ($result.Success) {
                        Write-Host "  ✓ $($result.Name) installed successfully" -ForegroundColor Green
                    } else {
                        Write-Host "  ✗ Failed to install $($result.Name)" -ForegroundColor Red
                    }
                }
            }
            catch {
                Write-Error "Error reading results: $_"
            }
        }
        else {
            Write-Warning "Results file not found: $tempResultFile"
        }
        
        # Cleanup
        Remove-Item -Path $tempScript -Force -ErrorAction SilentlyContinue
        Remove-Item -Path $tempProgressFile -Force -ErrorAction SilentlyContinue
        Remove-Item -Path $tempResultFile -Force -ErrorAction SilentlyContinue
        Remove-Item -Path $tempLogFile -Force -ErrorAction SilentlyContinue
    }

    # Install user packages (no elevation needed)
    foreach ($package in $userPackages) {
        $completedPackages++
        $percentComplete = [int](($completedPackages / $totalPackages) * 100)

        if ($LoadingWindow) {
            Update-LoadingWindow -Window $LoadingWindow `
                                -Message "Installation : $($package.Name) ($completedPackages/$totalPackages)..." `
                                -Progress $percentComplete
        }

        Write-Host "[$completedPackages/$totalPackages] Installing $($package.Name)..." -ForegroundColor Cyan

        try {
            $installResult = Install-Package -Package $package

            if ($installResult.Success) {
                Write-Host "  ✓ $($package.Name) installed successfully" -ForegroundColor Green
            } else {
                Write-Host "  ✗ Failed to install $($package.Name)" -ForegroundColor Red
            }

            $oNewResult = [PSCustomObject]@{
                Name = $package.Name
                Id = $package.Id
                Source = $package.Source
                Success = $installResult.Success
                RebootRequired = $installResult.RebootRequired
            }
            $results += $oNewResult
        }
        catch {
            Write-Host "  ✗ Erreur : $_" -ForegroundColor Red
            $oNewResult = [PSCustomObject]@{
                Name = $package.Name
                Id = $package.Id
                Source = $package.Source
                Success = $false
                RebootRequired = $false
            }
            $results += $oNewResult
        }
    }

    return $results
}

function Install-WithCredentials {
    <#
    .SYNOPSIS
        Installs a program with specified credentials using RunAs
    
    .PARAMETER Command
        Command to execute
    
    .PARAMETER Arguments
        Arguments for the command
    
    .PARAMETER Credential
        Credentials to use
    #>
    
    [CmdletBinding()]
    Param(
        [Parameter(Mandatory=$true)]
        [string]$Command,
        
        [Parameter(Mandatory=$true)]
        [array]$Arguments,
        
        [Parameter(Mandatory=$true)]
        [System.Management.Automation.PSCredential]$Credential
    )
    
    try {
        # Validate credential
        if ([string]::IsNullOrWhiteSpace($Credential.UserName)) {
            throw "Credential UserName is null or empty"
        }
        
        Write-Verbose "Running with credentials: $($Credential.UserName)"
        Write-Verbose "Command: $Command"
        Write-Verbose "Arguments: $($Arguments -join ' ')"
        
        # Create a temporary PowerShell script that will run the installer
        $tempScript = [System.IO.Path]::GetTempFileName() + ".ps1"
        
        $argumentString = if ($Arguments.Count -gt 0) { 
            ($Arguments | ForEach-Object { "`"$_`"" }) -join ', '
        } else { 
            "" 
        }
        
        $scriptContent = @"
# Temporary install script
`$command = '$Command'
`$arguments = @($argumentString)

Write-Host "Installing with: `$command"
Write-Host "Arguments: `$(`$arguments -join ' ')"

try {
    `$process = Start-Process -FilePath `$command ``
                             -ArgumentList `$arguments ``
                             -Wait ``
                             -PassThru ``
                             -WindowStyle Hidden
    
    Write-Host "Process exited with code: `$(`$process.ExitCode)"
    exit `$process.ExitCode
}
catch {
    Write-Error "Installation failed: `$_"
    exit 1
}
"@
        
        Set-Content -Path $tempScript -Value $scriptContent -Encoding UTF8
        Write-Verbose "Temporary script created: $tempScript"
        
        # Execute the script with credentials
        $process = Start-Process -FilePath "powershell.exe" `
                                 -ArgumentList "-NoProfile", "-ExecutionPolicy", "Bypass", "-File", "`"$tempScript`"" `
                                 -Credential $Credential `
                                 -Wait `
                                 -PassThru `
                                 -WindowStyle Hidden `
                                 -LoadUserProfile
        
        Write-Verbose "Script execution completed with exit code: $($process.ExitCode)"
        
        # Clean up
        Remove-Item -Path $tempScript -Force -ErrorAction SilentlyContinue
        
        # Check exit code
        if ($process.ExitCode -ne 0 -and $process.ExitCode -ne 3010) {
            throw "Installation failed with exit code: $($process.ExitCode)"
        }
        
    }
    catch {
        # Clean up on error
        if (Test-Path $tempScript) {
            Remove-Item -Path $tempScript -Force -ErrorAction SilentlyContinue
        }
        throw "Install-WithCredentials failed: $_"
    }
}

function Install-MSStorePackage {
    <#
    .SYNOPSIS
    Installs a Microsoft Store package

    .DESCRIPTION
    Wrapper function that calls Install-MSStoreApp to install Microsoft Store applications.
    Handles both MSIX/APPX (modern) and Win32 Store apps with automatic dependency management.
    Falls back to WinGet installation if native installation fails.

    .PARAMETER Package
    Package hashtable from config.json with Source="msstore"
    Required properties: Id (ProductId), Scope (user/machine), Name

    .EXAMPLE
    $package = @{ Id = "9NKSQGP7F2NH"; Scope = "user"; Name = "WhatsApp" }
    Install-MSStorePackage -Package $package
    #>
    [CmdletBinding()]
    Param(
        [Parameter(Mandatory=$true)]
        [hashtable]$Package
    )

    Write-Verbose "Installing from Microsoft Store: $($Package.Id)"

    try {
        # Call the new Install-MSStoreApp function
        $success = Install-MSStoreApp -Package $Package

        if (-not $success) {
            throw "Install-MSStoreApp returned false"
        }
    }
    catch {
        Write-Warning "Failed to install via native Store method: $($_.Exception.Message)"
        Write-Verbose "Falling back to WinGet installation..."

        # Fallback to WinGet method
        $wingetExe = Get-WinGetExe

        $arguments = @(
            "install"
            "--id", $Package.Id
            "--source", "msstore"
            "--exact"
            "--silent"
            "--accept-package-agreements"
            "--accept-source-agreements"
        )

        # Add scope argument if specified
        if ($Package.Scope) {
            $arguments += "--scope", $Package.Scope
        }

        Write-Verbose "Running: $wingetExe $($arguments -join ' ')"

        $process = Start-Process -FilePath $wingetExe `
                                 -ArgumentList $arguments `
                                 -Wait `
                                 -PassThru `
                                 -WindowStyle Hidden

        if ($process.ExitCode -ne 0 -and $process.ExitCode -ne 3010) {
            throw "Microsoft Store install failed with exit code: $($process.ExitCode)"
        }
        $script:LastInstallExitCode = $process.ExitCode
    }
}

function Get-WingetPortablePackagePath {
    <#
    .SYNOPSIS
        Gets the installation path for a portable WinGet package

    .DESCRIPTION
        Determines the base directory where WinGet installs portable packages
        based on the installation scope (machine or user)

    .PARAMETER PackageId
        WinGet package identifier

    .PARAMETER Scope
        Installation scope: "machine" or "user"

    .OUTPUTS
        Returns the full path to the package directory if found, otherwise $null

    .EXAMPLE
        Get-WingetPortablePackagePath "tannerhelland.PhotoDemon" "machine"
        # Returns: C:\Program Files\WinGet\Packages\tannerhelland.PhotoDemon_Microsoft.Winget.Source_8wekyb3d8bbwe
    #>

    [CmdletBinding()]
    Param(
        [Parameter(Mandatory=$true)]
        [string]$PackageId,

        [Parameter(Mandatory=$false)]
        [ValidateSet("machine", "user")]
        [string]$Scope = "machine"
    )

    # Determine base path based on scope
    if ($Scope -eq "machine") {
        $basePath = "$env:ProgramFiles\WinGet\Packages"
    } else {
        $basePath = "$env:LOCALAPPDATA\Microsoft\WinGet\Packages"
    }

    Write-Verbose "Checking portable package path: $basePath"

    if (-not (Test-Path $basePath)) {
        Write-Verbose "WinGet packages directory not found: $basePath"
        return $null
    }

    # Search for directory matching pattern: PackageId_Microsoft.Winget.Source_*
    $pattern = "$PackageId`_Microsoft.Winget.Source_*"
    $packageDirs = Get-ChildItem -Path $basePath -Directory -Filter $pattern -ErrorAction SilentlyContinue

    if ($packageDirs -and $packageDirs.Count -gt 0) {
        # Return the first match (should only be one)
        $foundPath = $packageDirs[0].FullName
        Write-Verbose "Found portable package at: $foundPath"
        return $foundPath
    }

    Write-Verbose "Portable package not found for: $PackageId"
    return $null
}

function Test-PortablePackageInstalled {
    <#
    .SYNOPSIS
        Tests if a portable WinGet package is installed

    .DESCRIPTION
        Checks if a portable package is installed by looking for its directory
        in the WinGet packages folder. For portable packages installed by WinGet,
        the presence of the directory is sufficient to confirm installation.

    .PARAMETER Package
        Package hashtable with Id and Scope properties

    .OUTPUTS
        Returns $true if the portable package directory exists, $false otherwise

    .EXAMPLE
        Test-PortablePackageInstalled -Package @{Id="tannerhelland.PhotoDemon"; Scope="machine"}
    #>

    [CmdletBinding()]
    Param(
        [Parameter(Mandatory=$true)]
        [hashtable]$Package
    )

    $packagePath = Get-WingetPortablePackagePath -PackageId $Package.Id -Scope $Package.Scope

    if (-not $packagePath) {
        Write-Verbose "Portable package path not found for: $($Package.Id)"
        return $false
    }

    # Check if the package directory exists
    if (-not (Test-Path $packagePath)) {
        Write-Verbose "Portable package directory does not exist: $packagePath"
        return $false
    }

    Write-Verbose "Portable package installed at: $packagePath"
    return $true
}

function Get-PortablePackagePathToAdd {
    <#
    .SYNOPSIS
        Determines the path that should be added to PATH environment variable for a portable package

    .DESCRIPTION
        Retrieves the NestedInstallerFiles from the package manifest and determines
        which subdirectory should be added to the PATH variable based on the
        RelativeFilePath of the main executable.

    .PARAMETER Package
        Package hashtable with Id, Scope, and Installer properties

    .PARAMETER PackagePath
        Base installation path of the portable package

    .OUTPUTS
        Returns the full path to add to PATH variable, or $null if cannot be determined

    .EXAMPLE
        Get-PortablePackagePathToAdd -Package $pkg -PackagePath "C:\Program Files\WinGet\Packages\tannerhelland.PhotoDemon_..."
        # Returns: C:\Program Files\WinGet\Packages\tannerhelland.PhotoDemon_...\PhotoDemon
    #>

    [CmdletBinding()]
    Param(
        [Parameter(Mandatory=$true)]
        [hashtable]$Package,

        [Parameter(Mandatory=$true)]
        [string]$PackagePath
    )

    try {
        # Check if Installer info is already in Package (from Get-WingetPackageInstaller)
        if ($Package.Installer -and $Package.Installer.NestedInstallerFiles) {
            $nestedInstallerFiles = $Package.Installer.NestedInstallerFiles
        }
        else {
            # Try to get installer info
            Write-Verbose "Retrieving installer information for $($Package.Id)..."
            $installer = Get-WingetPackageInstaller -PackageId $Package.Id -Scope $Package.Scope

            if (-not $installer) {
                Write-Verbose "Could not retrieve installer information"
                return $null
            }

            $nestedInstallerFiles = $installer.NestedInstallerFiles
        }

        if (-not $nestedInstallerFiles) {
            Write-Verbose "No NestedInstallerFiles found in package manifest"
            return $null
        }

        # NestedInstallerFiles is typically an array of objects with RelativeFilePath
        # Example: @{RelativeFilePath="PhotoDemon/PhotoDemon.exe"}
        $relativeFilePath = $nestedInstallerFiles[0].RelativeFilePath

        if (-not $relativeFilePath) {
            Write-Verbose "RelativeFilePath not found in NestedInstallerFiles"
            return $null
        }

        # Extract the directory part from the RelativeFilePath
        # Example: "PhotoDemon/PhotoDemon.exe" -> "PhotoDemon"
        $relativeDir = Split-Path $relativeFilePath -Parent

        if ([string]::IsNullOrEmpty($relativeDir)) {
            # If no parent directory, use the base package path
            Write-Verbose "No subdirectory found, using base package path"
            return $PackagePath
        }

        # Construct the full path to add to PATH
        $pathToAdd = Join-Path $PackagePath $relativeDir
        Write-Verbose "Determined PATH entry: $pathToAdd"

        return $pathToAdd

    } catch {
        Write-Warning "Error determining PATH for portable package $($Package.Id): $_"
        return $null
    }
}

function Get-WingetPortablePackageFolderName {
    <#
    .SYNOPSIS
        Generates the WinGet portable package folder name with hash

    .DESCRIPTION
        Creates the folder name format used by WinGet for portable packages:
        {PackageId}_Microsoft.Winget.Source_{hash}

    .PARAMETER PackageId
        The WinGet package identifier

    .OUTPUTS
        Returns the folder name with hash suffix

    .EXAMPLE
        Get-WingetPortablePackageFolderName "tannerhelland.PhotoDemon"
        # Returns: tannerhelland.PhotoDemon_Microsoft.Winget.Source_8wekyb3d8bbwe
    #>

    [CmdletBinding()]
    Param(
        [Parameter(Mandatory=$true)]
        [string]$PackageId
    )

    # Get the PublisherId dynamically from the WinGet package
    $publisherId = Get-WingetPublisherId

    $folderName = "${PackageId}_Microsoft.Winget.Source_${publisherId}"

    Write-Verbose "Generated folder name: $folderName"

    return $folderName
}

function Install-PortablePackage {
    <#
    .SYNOPSIS
        Installs a portable WinGet package

    .DESCRIPTION
        Handles the complete installation of a portable package:
        - Extracts ZIP or copies EXE to WinGet Packages directory (based on architecture and scope)
        - Creates a shortcut in the Start Menu
        - Adds the appropriate path to PATH environment variable

    .PARAMETER Package
        Package hashtable with Id, Name, Installer, Scope properties

    .PARAMETER InstallerPath
        Path to the downloaded installer file

    .PARAMETER TempPath
        Temporary directory for extraction

    .OUTPUTS
        Returns $true if successful, $false otherwise

    .EXAMPLE
        Install-PortablePackage -Package $pkg -InstallerPath "C:\temp\app.zip" -TempPath "C:\temp"
    #>

    [CmdletBinding()]
    Param(
        [Parameter(Mandatory=$true)]
        [hashtable]$Package,

        [Parameter(Mandatory=$true)]
        [string]$InstallerPath,

        [Parameter(Mandatory=$false)]
        [string]$TempPath = "$env:TEMP\PackageInstall"
    )

    try {
        $scope = $Package.Installer.Scope
        $packageId = $Package.Id
        $packageName = if ($Package.Name) { $Package.Name } else { $packageId }

        # Get architecture from Installer or from Installers sub-object
        $architecture = $Package.Installer.Architecture
        if (-not $architecture -and $Package.Installer.Installers) {
            $architecture = $Package.Installer.Installers.Architecture
        }

        Write-Verbose "Installing portable package: $packageName (Arch: $architecture, Scope: $scope)"

        # Determine base WinGet Packages directory based on scope and architecture
        if ($scope -eq "machine") {
            # For machine scope, use Program Files based on architecture
            if ($architecture -eq "x86") {
                $programFilesBase = ${env:ProgramFiles(x86)}
            } else {
                # x64, arm64, or neutral -> Program Files
                $programFilesBase = $env:ProgramFiles
            }
            $basePackagesPath = "$programFilesBase\WinGet\Packages"
        } else {
            $basePackagesPath = "$env:LOCALAPPDATA\Microsoft\WinGet\Packages"
        }

        # Create base directory if it doesn't exist
        if (-not (Test-Path $basePackagesPath)) {
            Write-Verbose "Creating WinGet Packages directory: $basePackagesPath"
            New-Item -Path $basePackagesPath -ItemType Directory -Force | Out-Null
        }

        # Generate the package-specific folder name
        $packageFolderName = Get-WingetPortablePackageFolderName -PackageId $packageId
        $packageInstallPath = Join-Path $basePackagesPath $packageFolderName

        # Create package directory
        if (Test-Path $packageInstallPath) {
            Write-Verbose "Package directory already exists, removing: $packageInstallPath"
            Remove-Item -Path $packageInstallPath -Recurse -Force -ErrorAction Stop
        }

        Write-Host "  Creating package directory..." -ForegroundColor Gray
        New-Item -Path $packageInstallPath -ItemType Directory -Force | Out-Null

        # Get file extension
        $extension = [System.IO.Path]::GetExtension($InstallerPath).ToLower()

        # Handle based on file type
        if ($extension -eq ".zip") {
            Write-Host "  Extracting portable archive..." -ForegroundColor Gray

            # Extract ZIP to package directory
            Expand-Archive -Path $InstallerPath -DestinationPath $packageInstallPath -Force

            Write-Verbose "Extracted to: $packageInstallPath"
        }
        elseif ($extension -eq ".exe") {
            Write-Host "  Copying portable executable..." -ForegroundColor Gray

            # For portable EXE, copy to package directory
            $targetFileName = [System.IO.Path]::GetFileName($InstallerPath)
            $targetPath = Join-Path $packageInstallPath $targetFileName

            Copy-Item -Path $InstallerPath -Destination $targetPath -Force

            Write-Verbose "Copied to: $targetPath"
        }
        else {
            throw "Unsupported portable package format: $extension"
        }

        Write-Host "  Portable package installed to: $packageInstallPath" -ForegroundColor Green

        # Determine the executable path from NestedInstallerFiles
        $executablePath = $null
        $nestedInstallerFiles = $Package.Installer.NestedInstallerFiles

        # Also check on Installers sub-object if not found
        if (-not $nestedInstallerFiles -and $Package.Installer.Installers) {
            $nestedInstallerFiles = $Package.Installer.Installers.NestedInstallerFiles
        }

        if ($nestedInstallerFiles) {
            # NestedInstallerFiles can be a hashtable, array, or PSCustomObject
            $relativePath = $null
            if ($nestedInstallerFiles -is [hashtable]) {
                $relativePath = $nestedInstallerFiles["RelativeFilePath"]
            } elseif ($nestedInstallerFiles -is [array] -and $nestedInstallerFiles.Count -gt 0) {
                $relativePath = $nestedInstallerFiles[0].RelativeFilePath
                if (-not $relativePath) {
                    $relativePath = $nestedInstallerFiles[0]["RelativeFilePath"]
                }
            } elseif ($nestedInstallerFiles.RelativeFilePath) {
                $relativePath = $nestedInstallerFiles.RelativeFilePath
            }

            if ($relativePath) {
                $executablePath = Join-Path $packageInstallPath $relativePath
                Write-Verbose "Executable path from NestedInstallerFiles: $executablePath"
            }
        }

        # If no NestedInstallerFiles, try to find the main executable
        if (-not $executablePath -or -not (Test-Path $executablePath)) {
            $foundExe = Get-ChildItem -Path $packageInstallPath -Recurse -Filter "*.exe" |
                        Where-Object { $_.Name -notmatch 'unins|setup|install' } |
                        Select-Object -First 1

            if ($foundExe) {
                $executablePath = $foundExe.FullName
                Write-Verbose "Found executable: $executablePath"
            }
        }

        # Create Start Menu shortcut
        if ($executablePath -and (Test-Path $executablePath)) {
            Write-Host "  Creating Start Menu shortcut..." -ForegroundColor Gray

            # Determine Start Menu folder based on scope
            if ($scope -eq "machine") {
                $startMenuPath = "$env:ProgramData\Microsoft\Windows\Start Menu\Programs"
            } else {
                $startMenuPath = "$env:APPDATA\Microsoft\Windows\Start Menu\Programs"
            }

            # Create shortcut
            $shortcutPath = Join-Path $startMenuPath "$packageName.lnk"
            $executableDir = [System.IO.Path]::GetDirectoryName($executablePath)

            try {
                $wshShell = New-Object -ComObject WScript.Shell
                $shortcut = $wshShell.CreateShortcut($shortcutPath)
                $shortcut.TargetPath = $executablePath
                $shortcut.WorkingDirectory = $executableDir
                $shortcut.Description = $packageName
                $shortcut.Save()

                # Release COM object
                [System.Runtime.InteropServices.Marshal]::ReleaseComObject($wshShell) | Out-Null

                Write-Host "  Shortcut created: $shortcutPath" -ForegroundColor Green
            }
            catch {
                Write-Warning "Failed to create Start Menu shortcut: $_"
            }

            # Add executable directory to PATH
            Write-Host "  Configuring PATH for portable application..." -ForegroundColor Gray
            $envScope = if ($scope -eq "machine") { "Machine" } else { "User" }
            $addResult = Add-PathToEnvironment -Path $executableDir -Scope $envScope

            if (-not $addResult) {
                Write-Warning "Failed to add portable package to PATH, but installation succeeded"
            }
        } else {
            Write-Warning "Could not find executable for shortcut creation"

            # Fallback: use Get-PortablePackagePathToAdd for PATH
            $pathToAdd = Get-PortablePackagePathToAdd -Package $Package -PackagePath $packageInstallPath

            if ($pathToAdd) {
                $envScope = if ($scope -eq "machine") { "Machine" } else { "User" }
                Add-PathToEnvironment -Path $pathToAdd -Scope $envScope | Out-Null
            } else {
                $envScope = if ($scope -eq "machine") { "Machine" } else { "User" }
                Add-PathToEnvironment -Path $packageInstallPath -Scope $envScope | Out-Null
            }
        }

        return $true

    } catch {
        Write-Error "Failed to install portable package: $_"
        return $false
    }
}

function Get-SortedPackagesByDependencies {
    <#
    .SYNOPSIS
        Sorts packages by scope (machine first) and respects dependencies (Requires attribute)

    .PARAMETER Packages
        Array of packages to sort

    .PARAMETER AllAvailablePackages
        All available packages (used to resolve missing dependencies)

    .OUTPUTS
        Hashtable with 'SortedPackages' (sorted array) and 'MissingDependencies' (array of missing dependency info)

    .EXAMPLE
        $result = Get-SortedPackagesByDependencies -Packages $selected -AllAvailablePackages $packages
    #>

    [CmdletBinding()]
    Param(
        [Parameter(Mandatory=$true)]
        [array]$Packages,

        [Parameter(Mandatory=$false)]
        [array]$AllAvailablePackages = @()
    )

    # Track missing dependencies
    $missingDependencies = @()

    # Create lookup of all selected package IDs
    $selectedIds = @{}
    foreach ($pkg in $Packages) {
        if (-not $selectedIds.ContainsKey($pkg.Id)) {
            $selectedIds[$pkg.Id] = $true
        }
    }

    # Create lookup of all available packages
    $allPackagesLookup = @{}
    foreach ($pkg in $AllAvailablePackages) {
        if (-not $allPackagesLookup.ContainsKey($pkg.Id)) {
            $allPackagesLookup[$pkg.Id] = $pkg
        }
    }

    # Check for missing dependencies
    foreach ($pkg in $Packages) {
        if ($pkg.Requires -and $pkg.Requires.Count -gt 0) {
            foreach ($requiredId in $pkg.Requires) {
                # Check if dependency is not in selected packages
                if (-not $selectedIds.ContainsKey($requiredId)) {
                    # Find the required package in all available packages
                    $requiredPackage = $allPackagesLookup[$requiredId]

                    # Check if the required package is already installed
                    $isAlreadyInstalled = $false
                    if ($requiredPackage) {
                        $isAlreadyInstalled = $requiredPackage.IsInstalled -eq $true
                    }

                    # Only add to missing dependencies if not already installed
                    if (-not $isAlreadyInstalled) {
                        $missingDependencies += [PSCustomObject]@{
                            Package = $pkg.Name
                            PackageId = $pkg.Id
                            RequiredId = $requiredId
                            RequiredPackage = if ($requiredPackage) { $requiredPackage } else { $null }
                        }
                    }
                    else {
                        Write-Verbose "Dependency $requiredId is already installed, skipping"
                    }
                }
            }
        }
    }

    # Separate packages by scope
    $machinePackages = @($Packages | Where-Object { $_.Scope -eq "machine" })
    $userPackages = @($Packages | Where-Object { $_.Scope -ne "machine" })

    # Function to perform topological sort (handles dependencies)
    function Get-TopologicalSort {
        param([array]$Items)

        if ($Items.Count -eq 0) {
            return @()
        }

        $sorted = New-Object System.Collections.ArrayList
        $visited = @{}
        $visiting = @{}

        # Create a lookup by Id for quick access
        $itemLookup = @{}
        foreach ($item in $Items) {
            # Use -Force to overwrite if key already exists (shouldn't happen, but just in case)
            if (-not $itemLookup.ContainsKey($item.Id)) {
                $itemLookup[$item.Id] = $item
            }
        }

        function Invoke-NodeVisit {
            param($item)

            $itemId = $item.Id

            # Check for circular dependencies
            if ($visiting.ContainsKey($itemId)) {
                Write-Warning "Circular dependency detected for package: $($item.Name)"
                return
            }

            # Skip if already visited
            if ($visited.ContainsKey($itemId)) {
                return
            }

            $visiting[$itemId] = $true

            # Visit dependencies first
            if ($item.Requires -and $item.Requires.Count -gt 0) {
                foreach ($requiredId in $item.Requires) {
                    # Only process dependency if it's in the same batch (same scope)
                    if ($itemLookup.ContainsKey($requiredId)) {
                        Invoke-NodeVisit $itemLookup[$requiredId]
                    }
                }
            }

            $visiting.Remove($itemId)
            $visited[$itemId] = $true
            [void]$sorted.Add($item)
        }

        # Visit all nodes
        foreach ($item in $Items) {
            Invoke-NodeVisit $item
        }

        return $sorted.ToArray()
    }

    # Sort each scope group by dependencies
    Write-Verbose "Sorting machine packages ($($machinePackages.Count) packages)..."
    $sortedMachine = Get-TopologicalSort -Items $machinePackages
    Write-Verbose "Machine packages sorted: $($sortedMachine.Count)"

    Write-Verbose "Sorting user packages ($($userPackages.Count) packages)..."
    $sortedUser = Get-TopologicalSort -Items $userPackages
    Write-Verbose "User packages sorted: $($sortedUser.Count)"

    # Combine: machine packages first, then user packages
    $sortedPackages = @()
    $sortedPackages += @($sortedMachine)
    $sortedPackages += @($sortedUser)

    Write-Verbose "Total sorted packages: $($sortedPackages.Count)"

    return @{
        SortedPackages = $sortedPackages
        MissingDependencies = $missingDependencies
    }
}

function Get-PackageRowIds {
    <#
    .SYNOPSIS
        Retrieves the complete row from the SQLite packages table for each package

    .PARAMETER Packages
        Array of package hashtables with Id property

    .OUTPUTS
        The input packages array with an added WingetPackage property for each package

    .EXAMPLE
        $packages = Get-PackageRowIds -Packages $packages
    #>

    [CmdletBinding()]
    Param(
        [Parameter(Mandatory=$true)]
        [array]$Packages
    )

    try {
        # Get the catalog database path
        if (-not $Global:WingetCatalog -or -not (Test-Path $Global:WingetCatalog.DatabasePath)) {
            Write-Warning "Winget catalog not loaded. Cannot retrieve package data."
            return $Packages
        }

        $dbPath = $Global:WingetCatalog.DatabasePath
        Write-Verbose "Using database: $dbPath"

        # Build a query to get all package rows in one shot
        $packageIds = $Packages | Where-Object { $_.Source -eq "winget" } | ForEach-Object { "'$($_.Id)'" }

        if ($packageIds.Count -eq 0) {
            Write-Verbose "No winget packages to look up"
            return $Packages
        }

        $packageIdsString = $packageIds -join ","

        # Query the database - select all columns including rowid
        $query = "SELECT rowid, * FROM packages WHERE id IN ($packageIdsString)"
        Write-Verbose "Query: $query"

        $results = Invoke-SqliteQuery -DataSource $dbPath -Query $query -ErrorAction Stop

        # Create a hashtable for fast lookup
        $packageLookup = @{}
        foreach ($result in $results) {
            $packageLookup[$result.id] = $result
        }

        # Add WingetPackage to each package
        foreach ($package in $Packages) {
            if ($package.Source -eq "winget" -and $packageLookup.ContainsKey($package.Id)) {
                $package.WingetPackage = $packageLookup[$package.Id]
                Write-Verbose "Package $($package.Id) -> WingetPackage added"
            }
            else {
                $package.WingetPackage = $null
            }
        }

        Write-Host "Retrieved winget package data for $($packageLookup.Count) packages" -ForegroundColor Green

        return $Packages
    }
    catch {
        Write-Warning "Failed to retrieve package data: $_"
        return $Packages
    }
}

function Get-PackageProductCodes {
    <#
    .SYNOPSIS
        Retrieves all product codes from the SQLite productcodes2 table for each package

    .PARAMETER Packages
        Array of package hashtables with WingetPackage property

    .OUTPUTS
        The input packages array with an added ProductCodes property (array) for each package

    .EXAMPLE
        $packages = Get-PackageProductCodes -Packages $packages
    #>

    [CmdletBinding()]
    Param(
        [Parameter(Mandatory=$true)]
        [array]$Packages
    )

    try {
        # Get the catalog database path
        if (-not $Global:WingetCatalog -or -not (Test-Path $Global:WingetCatalog.DatabasePath)) {
            Write-Warning "Winget catalog not loaded. Cannot retrieve product codes."
            return $Packages
        }

        $dbPath = $Global:WingetCatalog.DatabasePath
        Write-Verbose "Using database: $dbPath"

        # Get all packages with WingetPackage
        $packagesWithWingetData = $Packages | Where-Object { $null -ne $_.WingetPackage }

        if ($packagesWithWingetData.Count -eq 0) {
            Write-Verbose "No packages with WingetPackage to look up product codes"
            return $Packages
        }

        # Build list of rowids
        $rowIds = $packagesWithWingetData | ForEach-Object { $_.WingetPackage.rowid }
        $rowIdsString = ($rowIds -join ",")

        # Query the database for all product codes
        $query = "SELECT package, productcode FROM productcodes2 WHERE package IN ($rowIdsString)"
        Write-Verbose "Query: $query"

        $results = Invoke-SqliteQuery -DataSource $dbPath -Query $query -ErrorAction Stop

        # Create a hashtable to store product codes by package rowid
        $productCodesLookup = @{}
        foreach ($result in $results) {
            $packageRowId = $result.package
            $productCode = $result.productcode

            if (-not $productCodesLookup.ContainsKey($packageRowId)) {
                $productCodesLookup[$packageRowId] = @()
            }

            $productCodesLookup[$packageRowId] += $productCode
        }

        # Add ProductCodes to each package
        foreach ($package in $Packages) {
            if ($null -ne $package.WingetPackage -and $productCodesLookup.ContainsKey($package.WingetPackage.rowid)) {
                $package.ProductCodes = $productCodesLookup[$package.WingetPackage.rowid]
                Write-Verbose "Package $($package.Id) -> ProductCodes: $($package.ProductCodes.Count)"
            }
            else {
                $package.ProductCodes = @()
            }
        }

        $totalProductCodes = ($productCodesLookup.Values | ForEach-Object { $_.Count } | Measure-Object -Sum).Sum
        Write-Host "Retrieved $totalProductCodes product codes for $($productCodesLookup.Count) packages" -ForegroundColor Green

        return $Packages
    }
    catch {
        Write-Warning "Failed to retrieve product codes: $_"
        return $Packages
    }
}

function Test-PackageCompatibility {
    <#
    .SYNOPSIS
        Tests if a package is compatible with the current system

    .DESCRIPTION
        Evaluates the Requirements array from a package configuration.
        Uses language-neutral detection methods (registry EditionID, WMI ProductType).

        Supported requirement types:
        - WindowsEdition: EditionID values (Professional, Enterprise, Education, Core, etc.)
        - WindowsType: Workstation, Server, DomainController
        - MinBuild: Minimum Windows build number
        - Architecture: AMD64, ARM64, x86
    #>

    [CmdletBinding()]
    Param(
        [Parameter(Mandatory=$true)]
        [hashtable]$Package
    )

    if (-not $Package.Requirements -or $Package.Requirements.Count -eq 0) {
        return $true
    }

    # Cache system info on first call
    if (-not $script:SystemCompatInfo) {
        $sysInfo = Get-SystemInfo
        $script:SystemCompatInfo = @{
            EditionID    = $sysInfo.EditionID
            ProductType  = $sysInfo.ProductType
            Build        = $sysInfo.Build
            Architecture = $sysInfo.Architecture
        }
        Write-Verbose "System compat info: Edition=$($script:SystemCompatInfo.EditionID), Type=$($script:SystemCompatInfo.ProductType), Build=$($script:SystemCompatInfo.Build), Arch=$($script:SystemCompatInfo.Architecture)"
    }

    foreach ($req in $Package.Requirements) {
        switch ($req.Type) {
            "WindowsEdition" {
                if ($script:SystemCompatInfo.EditionID -notin $req.Values) {
                    Write-Verbose "Incompatible: EditionID '$($script:SystemCompatInfo.EditionID)' not in [$($req.Values -join ', ')]"
                    return $false
                }
            }
            "WindowsType" {
                if ($script:SystemCompatInfo.ProductType -notin $req.Values) {
                    Write-Verbose "Incompatible: WindowsType '$($script:SystemCompatInfo.ProductType)' not in [$($req.Values -join ', ')]"
                    return $false
                }
            }
            "MinBuild" {
                if ($script:SystemCompatInfo.Build -lt $req.Value) {
                    Write-Verbose "Incompatible: Build $($script:SystemCompatInfo.Build) < required $($req.Value)"
                    return $false
                }
            }
            "Architecture" {
                if ($script:SystemCompatInfo.Architecture -notin $req.Values) {
                    Write-Verbose "Incompatible: Architecture '$($script:SystemCompatInfo.Architecture)' not in [$($req.Values -join ', ')]"
                    return $false
                }
            }
            default {
                Write-Warning "Unknown requirement type: $($req.Type)"
            }
        }
    }

    return $true
}

function Test-WindowsCapabilityInstalled {
    <#
    .SYNOPSIS
        Tests if a Windows Capability is installed by checking registry

    .PARAMETER CapabilityName
        The capability name (e.g., "Rsat.ActiveDirectory.DS-LDS.Tools")

    .OUTPUTS
        Returns $true if capability is installed, $false otherwise
    #>

    [CmdletBinding()]
    Param(
        [Parameter(Mandatory=$true)]
        [string]$CapabilityName
    )

    $regPath = "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Component Based Servicing\CapabilityIndex\$CapabilityName"
    return Test-Path $regPath
}

function Test-PackageInstalled {
    <#
    .SYNOPSIS
        Tests if a package is installed by checking product codes, name, and comments

    .PARAMETER Package
        Package hashtable with ProductCodes, Name properties

    .OUTPUTS
        Returns $true if package is found installed, $false otherwise

    .EXAMPLE
        Test-PackageInstalled -Package $package
    #>

    [CmdletBinding()]
    Param(
        [Parameter(Mandatory=$true)]
        [hashtable]$Package
    )

    # Ensure installed programs are loaded
    if (-not $Global:InstalledPrograms) {
        Write-Verbose "Loading installed programs..."
        $Global:InstalledPrograms = Get-InstalledPrograms -ProgramAndFeatures -AsHashtable -IncludeAppx
    }

    # Strategy 0: Check for portable WinGet packages (before all other strategies)
    # This strategy checks if the package is a portable package installed by WinGet
    if ($Package.Source -eq "winget" -and $Package.Id) {
        $isPortable = Test-PortablePackageInstalled -Package $Package
        if ($isPortable) {
            Write-Verbose "Package $($Package.Name) found as portable WinGet package"
            return $true
        }
    }

    # Strategy 0b: Check for Windows Capabilities (RSAT, etc.)
    if ($Package.Source -eq "windowscapability" -and $Package.Id) {
        $isInstalled = Test-WindowsCapabilityInstalled -CapabilityName $Package.Id
        Write-Verbose "Windows Capability $($Package.Id) installed: $isInstalled"
        return $isInstalled
    }

    # Strategy 0c: Check for Windows Optional Features (Hyper-V, WSL, etc.)
    if ($Package.Source -eq "windowsfeature" -and $Package.Id) {
        $isInstalled = Test-WindowsFeatureInstalled -FeatureName $Package.Id
        Write-Verbose "Windows Feature $($Package.Id) installed: $isInstalled"
        return $isInstalled
    }

    # Strategy 1: Check by DetectionScript (custom PowerShell script)
    # If DetectionScript attribute exists, ONLY use this method and skip all other strategies
    if ($Package.DetectionScript) {
        Write-Verbose "Checking package $($Package.Name) using DetectionScript (exclusive mode)"

        try {
            # Execute the detection script as a script block
            $scriptBlock = [ScriptBlock]::Create($Package.DetectionScript)
            $result = & $scriptBlock

            # Script should return $true or $false
            if ($result -eq $true) {
                Write-Verbose "Package $($Package.Name) found by DetectionScript"
                return $true
            }
            else {
                Write-Verbose "Package $($Package.Name) not found by DetectionScript"
                return $false
            }
        }
        catch {
            Write-Warning "DetectionScript failed for package $($Package.Name): $_"
            return $false
        }
    }

    # Strategy 2: Check by Detection attribute (custom detection rules)
    # If Detection attribute exists, ONLY use this method and skip all other strategies
    if ($Package.Detection) {
        Write-Verbose "Checking package $($Package.Name) using Detection attribute (exclusive mode)"

        # Find programs that match ALL properties in Detection (AND logic)
        $found = $Global:InstalledPrograms | Where-Object {
            $program = $_
            $allMatch = $true

            # Check each property in Detection hashtable
            foreach ($key in $Package.Detection.Keys) {
                $detectionValue = $Package.Detection[$key]
                $programValue = $program[$key]

                # If program doesn't have this property, it's not a match
                if ($null -eq $programValue) {
                    $allMatch = $false
                    break
                }

                # Use -like operator to support wildcards
                if ($programValue -notlike $detectionValue) {
                    $allMatch = $false
                    break
                }
            }

            $allMatch
        }

        if ($found) {
            Write-Verbose "Package $($Package.Name) found by Detection attribute"
            return $true
        }
        else {
            Write-Verbose "Package $($Package.Name) not found by Detection attribute"
            return $false
        }
    }

    # Strategy 3: Check by ProductCode (exact match)
    if ($Package.ProductCodes -and $Package.ProductCodes.Count -gt 0) {
        foreach ($productCode in $Package.ProductCodes) {
            $found = $Global:InstalledPrograms | Where-Object { $_.ProductCode -eq $productCode }
            if ($found) {
                Write-Verbose "Package $($Package.Name) found by ProductCode exact match: $productCode"
                return $true
            }
        }
    }

    # Strategy 4: Check by ProductCodes matching Comments field
    if ($Package.ProductCodes -and $Package.ProductCodes.Count -gt 0) {
        foreach ($productCode in $Package.ProductCodes) {
            $found = $Global:InstalledPrograms | Where-Object { $_.Comments -eq $productCode }
            if ($found) {
                Write-Verbose "Package $($Package.Name) found by ProductCode matching Comments: $productCode"
                return $true
            }
        }
    }

    # Strategy 5: Check by Name (exact match, case-insensitive)
    if ($Package.Name) {
        $found = $Global:InstalledPrograms | Where-Object { $_.Name -eq $Package.Name }
        if ($found) {
            Write-Verbose "Package $($Package.Name) found by Name"
            return $true
        }
    }

    # Strategy 6: Check by Comments field matching package Name
    if ($Package.Name) {
        $found = $Global:InstalledPrograms | Where-Object { $_.Comments -eq $Package.Name }
        if ($found) {
            Write-Verbose "Package $($Package.Name) found by Comments matching Name"
            return $true
        }
    }

    # Strategy 7: Check by PackageName field matching package PackageName
    if ($Package.PackageName) {
        $found = $Global:InstalledPrograms | Where-Object { $_.PackageName -eq $Package.PackageName }
        if ($found) {
            Write-Verbose "Package $($Package.Name) found by PackageName matching PackageName"
            return $true
        }
    }

    # Strategy 8: Check by WingetPackage name + latest_version (for packages without ProductCodes)
    if ($Package.WingetPackage -and $Package.WingetPackage.name -and $Package.WingetPackage.latest_version) {
        $wingetNameVersionPattern = "$($Package.WingetPackage.name)*$($Package.WingetPackage.latest_version)*"

        # Check in Name field
        $found = $Global:InstalledPrograms | Where-Object { $_.Name -like $wingetNameVersionPattern }
        if ($found) {
            Write-Verbose "Package $($Package.Name) found by WingetPackage name+version in Name: $wingetNameVersionPattern"
            return $true
        }

        # Check in Comments field
        $found = $Global:InstalledPrograms | Where-Object { $_.Comments -like $wingetNameVersionPattern }
        if ($found) {
            Write-Verbose "Package $($Package.Name) found by WingetPackage name+version in Comments: $wingetNameVersionPattern"
            return $true
        }
    }

    Write-Verbose "Package $($Package.Name) not found"
    return $false
}

function Show-PackageManagerUI {
    <#
    .SYNOPSIS
        Displays a WPF GUI for managing package installation

    .PARAMETER Packages
        Array of package hashtables with Category, Name, Id, Scope, Source

    .PARAMETER IconFolder
        Path to folder containing package icons

    .EXAMPLE
        Show-PackageManagerUI -Packages $packages -IconFolder $iconFolder
    #>

    [CmdletBinding()]
    Param(
        [Parameter(Mandatory=$true)]
        [array]$Packages,

        [Parameter(Mandatory=$false)]
        [string]$IconFolder = "",

        [Parameter(Mandatory=$false)]
        [string]$IconFile = ""
    )

    Add-Type -AssemblyName PresentationFramework
    Add-Type -AssemblyName PresentationCore
    Add-Type -AssemblyName WindowsBase
    Add-Type -AssemblyName System.Drawing

    # Reset selection state at startup (before the language change loop)
    # This ensures a clean state when the script is launched, but preserves selection during language changes
    $script:selectedPackagesDict = @{}
    $script:userClickedInstall = $false
    $script:selectedOriginalCategory = "AllSoftware"

    # Loop to handle language changes
    do {
        $reloadUI = $false

        # Clear translation cache at start of loop to ensure fresh translations
        # This forces reload of translations with the current locale
        $script:TranslationCache = @{}

        # Get theme colors
        $colors = Get-ThemedColors

    # XAML Definition with themed colors
    [xml]$xaml = @"
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="$(tr 'UI.WindowTitle')"
        Height="700" Width="885"
        WindowStartupLocation="CenterScreen"
        Background="$($colors.WindowBackground)">
    
    <Window.Resources>
        $(Get-WPFScrollBarStyle -Colors $colors)
        $(Get-WPFCheckBoxStyle -Colors $colors)
        $(Get-WPFSidebarButtonStyle -Colors $colors)
        $(Get-WPFPrimaryButtonStyle -Colors $colors)
        $(Get-WPFComboBoxStyle -Colors $colors)
    </Window.Resources>
    
    <Grid>
        <Grid.ColumnDefinitions>
            <ColumnDefinition Width="220"/>
            <ColumnDefinition Width="*"/>
        </Grid.ColumnDefinitions>
        
        <!-- Left sidebar with categories -->
        <Border Grid.Column="0" Background="$($colors.SidebarBackground)" BorderBrush="$($colors.BorderColor)" BorderThickness="0,0,1,0">
            <DockPanel Margin="0,20,0,20" LastChildFill="True">
                <!-- Language selector at bottom -->
                <StackPanel DockPanel.Dock="Bottom" Margin="20,20,20,0">
                    <TextBlock Text="$(tr 'UI.Language')"
                              FontSize="12"
                              Foreground="$($colors.TextSecondary)"
                              Margin="0,0,0,5"/>
                    <ComboBox x:Name="LanguageSelector"
                             FontSize="14"/>
                </StackPanel>

                <!-- Categories section (fills remaining space) -->
                <TextBlock DockPanel.Dock="Top"
                           Text="$(tr 'UI.Categories')"
                           FontSize="18"
                           FontWeight="Bold"
                           Foreground="$($colors.TextPrimary)"
                           Margin="20,0,0,20"/>

                <ScrollViewer VerticalScrollBarVisibility="Auto"
                              HorizontalScrollBarVisibility="Disabled">
                    <StackPanel x:Name="CategoriesPanel"/>
                </ScrollViewer>
            </DockPanel>
        </Border>
        
        <!-- Main content area -->
        <Grid Grid.Column="1">
            <Grid.RowDefinitions>
                <RowDefinition Height="Auto"/>
                <RowDefinition Height="*"/>
                <RowDefinition Height="Auto"/>
            </Grid.RowDefinitions>

            <!-- Search bar -->
            <Border Grid.Row="0" Margin="20,20,20,0"
                    Background="$($colors.SidebarBackground)"
                    BorderBrush="$($colors.BorderColor)"
                    BorderThickness="1"
                    CornerRadius="6">
                <Grid>
                    <TextBox x:Name="SearchBox"
                             Padding="8,6,28,6"
                             FontSize="14"
                             Background="Transparent"
                             Foreground="$($colors.TextPrimary)"
                             CaretBrush="$($colors.TextPrimary)"
                             BorderThickness="0"
                             VerticalContentAlignment="Center"/>
                    <TextBlock x:Name="SearchPlaceholder"
                               Text="$(tr 'UI.SearchPlaceholder')"
                               Padding="10,7,0,0"
                               FontSize="14"
                               Foreground="$($colors.TextSecondary)"
                               IsHitTestVisible="False"/>
                </Grid>
            </Border>

            <!-- Package grid -->
            <ScrollViewer Grid.Row="1" VerticalScrollBarVisibility="Auto" Margin="10,20,20,10">
                <WrapPanel x:Name="PackagesPanel" Orientation="Horizontal"/>
            </ScrollViewer>
            
            <!-- Bottom bar with install button -->
            <Border Grid.Row="2" Background="$($colors.SidebarBackground)" BorderBrush="$($colors.BorderColor)" BorderThickness="0,1,0,0" Padding="20">
                <Grid>
                    <Grid.ColumnDefinitions>
                        <ColumnDefinition Width="*"/>
                        <ColumnDefinition Width="Auto"/>
                    </Grid.ColumnDefinitions>

                    <TextBlock x:Name="SelectionInfo"
                              Grid.Column="0"
                              VerticalAlignment="Center"
                              FontSize="14"
                              Foreground="$($colors.TextSecondary)"
                              Text="$(tr 'UI.NoSoftwareSelected')"/>
                    
                    <Button x:Name="InstallButton"
                           Grid.Column="1"
                           Style="{StaticResource PrimaryButtonStyle}"
                           Content="$(tr 'UI.InstallButton' -Parameters @(0))"
                           IsEnabled="False"/>
                </Grid>
            </Border>
        </Grid>
    </Grid>
</Window>
"@
    
    # Load XAML
    $reader = New-Object System.Xml.XmlNodeReader $xaml
    $window = [Windows.Markup.XamlReader]::Load($reader)

    # Set window icon if provided (title bar + taskbar)
    if ($IconFile -and (Test-Path $IconFile)) {
        Set-WindowIcon -Window $window -IconPath $IconFile
    }

    # Get controls
    $categoriesPanel = $window.FindName("CategoriesPanel")
    $packagesPanel = $window.FindName("PackagesPanel")
    $installButton = $window.FindName("InstallButton")
    $selectionInfo = $window.FindName("SelectionInfo")
    $languageSelector = $window.FindName("LanguageSelector")
    $searchBox = $window.FindName("SearchBox")
    $searchPlaceholder = $window.FindName("SearchPlaceholder")

    # Get unique categories and translate them
    $uniqueCategories = $Packages.Category | Select-Object -Unique
    $translatedCategories = @()
    $script:categoryMapping = @{}  # Map translated category to original category
    $script:reverseCategoryMapping = @{}  # Map original category to translated category

    $allSoftwareTranslated = tr 'UI.AllSoftware'
    $script:categoryMapping[$allSoftwareTranslated] = "AllSoftware"
    $script:reverseCategoryMapping["AllSoftware"] = $allSoftwareTranslated

    foreach ($cat in $uniqueCategories) {
        $translatedCat = tr "Categories.$cat"
        $translatedCategories += $translatedCat
        $script:categoryMapping[$translatedCat] = $cat
        $script:reverseCategoryMapping[$cat] = $translatedCat
    }
    # Sort translated categories alphabetically, then prepend "All Software"
    $translatedCategories = $translatedCategories | Sort-Object
    $categories = @($allSoftwareTranslated) + $translatedCategories

    # Get the currently selected category in the new language
    $script:currentCategory = $script:reverseCategoryMapping[$script:selectedOriginalCategory]
    
    # Function to create composite key
    $getPackageKey = {
        param($package)
        return "$($package.Id)-$($package.Source)"
    }
    
    # Function to update install button
    $updateInstallButton = {
        $count = $script:selectedPackagesDict.Count

        if ($count -eq 0) {
            $installButton.Content = tr 'UI.NoSoftwareSelected'
            $selectionInfo.Text = tr 'UI.NoSoftwareSelected'
        } else {
            $installButton.Content = tr 'UI.InstallButton' -Parameters @($count)
            $selectionInfo.Text = tr 'UI.SoftwareSelected' -Parameters @($count)
        }

        $installButton.IsEnabled = ($count -gt 0)
    }

    # Initialize button with correct text
    & $updateInstallButton

    # Initialize language selector - dynamically load from lang folder
    $availableLanguages = @()
    $langFolder = Join-Path (Get-ScriptDir -InputDir -FullPath) "lang"

    if (Test-Path $langFolder) {
        $langFiles = Get-ChildItem -Path $langFolder -Filter "*.json"

        foreach ($file in $langFiles) {
            $localeCode = [System.IO.Path]::GetFileNameWithoutExtension($file.Name)

            try {
                $langContent = Get-Content -Path $file.FullName -Raw -Encoding UTF8 | ConvertFrom-Json
                $displayName = if ($langContent.LanguageName) {
                    $langContent.LanguageName
                } else {
                    $localeCode
                }

                $availableLanguages += @{ Code = $localeCode; Display = $displayName }
            }
            catch {
                Write-Warning "Failed to load language file: $($file.Name)"
            }
        }
    }

    # Sort by display name
    $availableLanguages = $availableLanguages | Sort-Object -Property Display

    foreach ($lang in $availableLanguages) {
        $item = New-Object System.Windows.Controls.ComboBoxItem
        $item.Content = $lang.Display
        $item.Tag = $lang.Code
        $languageSelector.Items.Add($item) | Out-Null

        # Select current language
        if ($lang.Code -eq (Get-CurrentLocale)) {
            $languageSelector.SelectedItem = $item
        }
    }

    # Handle language change
    $languageSelector.Add_SelectionChanged({
        param($sender, $e)

        if ($sender.SelectedItem -and $sender.SelectedItem.Tag) {
            $newLocale = $sender.SelectedItem.Tag
            $currentLocale = Get-CurrentLocale

            # Only reload if language actually changed
            if ($newLocale -ne $currentLocale) {
                # Get translated messages BEFORE changing locale
                $message = (tr 'UI.LanguageChangeMessage') -replace '\\n', "`n"
                $title = tr 'UI.LanguageChangeTitle'
                $yesText = tr 'UI.Yes'
                $noText = tr 'UI.No'

                # Show themed dialog
                $buttons = @(
                    @{ text = $yesText; value = "yes" }
                    @{ text = $noText; value = "no" }
                )

                $result = Show-WPFButtonDialog -Title $title -Message $message -Buttons $buttons -Icon Question

                if ($result -eq "yes") {
                    # Set global locale override for the new language
                    $global:OverrideLocale = $newLocale

                    # Mark window to reload with new language
                    $window.Tag = @{ Action = "ReloadLanguage"; Locale = $newLocale }
                    $window.Close()
                }
                else {
                    # User clicked No, reset selector to current locale
                    foreach ($item in $languageSelector.Items) {
                        if ($item.Tag -eq $currentLocale) {
                            $languageSelector.SelectedItem = $item
                            break
                        }
                    }
                }
            }
        }
    })

    # Function to create package tile
    $newPackageTile = {
        param($package, $iconFolderPath, $themeColors)

        try {
            # Create composite key
            $packageKey = & $getPackageKey $package

            # Create tile programmatically
            $border = New-Object System.Windows.Controls.Border
            $border.Background = [System.Windows.Media.BrushConverter]::new().ConvertFrom($themeColors.TileBackground)
            $border.BorderBrush = [System.Windows.Media.BrushConverter]::new().ConvertFrom($themeColors.BorderColor)
            $border.BorderThickness = [System.Windows.Thickness]::new(1)
            $border.CornerRadius = [System.Windows.CornerRadius]::new(6)
            $border.Margin = [System.Windows.Thickness]::new(10)
            $border.Width = 180
            $border.Height = 200
            $border.Cursor = [System.Windows.Input.Cursors]::Hand

            $grid = New-Object System.Windows.Controls.Grid

            # Create rows
            $row1 = New-Object System.Windows.Controls.RowDefinition
            $row1.Height = [System.Windows.GridLength]::new(120)
            $row2 = New-Object System.Windows.Controls.RowDefinition
            $row2.Height = [System.Windows.GridLength]::new(1, [System.Windows.GridUnitType]::Star)
            $row3 = New-Object System.Windows.Controls.RowDefinition
            $row3.Height = [System.Windows.GridLength]::Auto

            $grid.RowDefinitions.Add($row1) | Out-Null
            $grid.RowDefinitions.Add($row2) | Out-Null
            $grid.RowDefinitions.Add($row3) | Out-Null

            # Icon - Try to load from iconFolder, otherwise use default
            $iconControl = $null
            $iconExtensions = @('.png', '.ico', '.jpg', '.jpeg')
            $iconFound = $false

            # Search for icon file with package Id (only if iconFolderPath is provided)
            if ($iconFolderPath -and $iconFolderPath -ne "") {
                foreach ($ext in $iconExtensions) {
                    $iconPath = Join-Path $iconFolderPath "$($package.Id)$ext"

                    if (Test-Path $iconPath) {
                        try {
                            # Load bitmap image (PNG, ICO, JPG, JPEG)
                            $image = New-Object System.Windows.Controls.Image
                            $image.Width = 80
                            $image.Height = 80
                            $image.Margin = [System.Windows.Thickness]::new(0,20,0,0)
                            $image.Stretch = [System.Windows.Media.Stretch]::Uniform
                            # High quality scaling to avoid aliasing on large images
                            [System.Windows.Media.RenderOptions]::SetBitmapScalingMode($image, [System.Windows.Media.BitmapScalingMode]::HighQuality)

                            # Load bitmap with decode size for better quality
                            $bitmap = New-Object System.Windows.Media.Imaging.BitmapImage
                            $bitmap.BeginInit()
                            $bitmap.UriSource = New-Object System.Uri $iconPath
                            $bitmap.CacheOption = [System.Windows.Media.Imaging.BitmapCacheOption]::OnLoad
                            $bitmap.DecodePixelWidth = 160  # 2x display size for crisp rendering
                            $bitmap.EndInit()
                            $bitmap.Freeze()

                            $image.Source = $bitmap
                            $iconControl = $image
                            $iconFound = $true
                            break
                        }
                        catch {
                            # Icon loading failed, continue to next extension
                        }
                    }
                }
            }

            # If no custom icon found, use default SVG icon
            if (-not $iconFound) {
                $viewbox = New-Object System.Windows.Controls.Viewbox
                $viewbox.Width = 80
                $viewbox.Height = 80
                $viewbox.Margin = [System.Windows.Thickness]::new(0,20,0,0)

                $canvas = New-Object System.Windows.Controls.Canvas
                $canvas.Width = 24
                $canvas.Height = 24

                $path = New-Object System.Windows.Shapes.Path
                $path.Fill = [System.Windows.Media.BrushConverter]::new().ConvertFrom("#0078D4")
                $path.Data = [System.Windows.Media.Geometry]::Parse("M19,4H5A2,2 0 0,0 3,6V18A2,2 0 0,0 5,20H19A2,2 0 0,0 21,18V6A2,2 0 0,0 19,4M19,18H5V8H19V18M13.5,12.67L11,14.5L9.5,13.25L7.5,15H16.5L13.5,12.67Z")

                $canvas.Children.Add($path) | Out-Null
                $viewbox.Child = $canvas
                $iconControl = $viewbox
            }

            [System.Windows.Controls.Grid]::SetRow($iconControl, 0)
            $grid.Children.Add($iconControl) | Out-Null
            
            # Name
            $textBlock = New-Object System.Windows.Controls.TextBlock
            $textBlock.Text = $package.Name
            $textBlock.FontSize = 13
            $textBlock.FontWeight = [System.Windows.FontWeights]::Medium
            $textBlock.Foreground = [System.Windows.Media.BrushConverter]::new().ConvertFrom($themeColors.TextPrimary)
            $textBlock.TextAlignment = [System.Windows.TextAlignment]::Center
            $textBlock.TextWrapping = [System.Windows.TextWrapping]::Wrap
            $textBlock.Margin = [System.Windows.Thickness]::new(10,5,10,5)
            $textBlock.VerticalAlignment = [System.Windows.VerticalAlignment]::Center
            [System.Windows.Controls.Grid]::SetRow($textBlock, 1)
            $grid.Children.Add($textBlock) | Out-Null
            
            # Checkbox, "Already installed" or "Incompatible" text
            if ($package.IsCompatible -eq $false) {
                # Show "Incompatible" text in italics
                $incompatibleText = New-Object System.Windows.Controls.TextBlock
                $incompatibleText.Text = tr 'UI.Incompatible'
                $incompatibleText.FontSize = 12
                $incompatibleText.FontStyle = [System.Windows.FontStyles]::Italic
                $incompatibleText.Foreground = [System.Windows.Media.BrushConverter]::new().ConvertFrom($themeColors.TextSecondary)
                $incompatibleText.HorizontalAlignment = [System.Windows.HorizontalAlignment]::Center
                $incompatibleText.Margin = [System.Windows.Thickness]::new(0,0,0,15)
                [System.Windows.Controls.Grid]::SetRow($incompatibleText, 2)
                $grid.Children.Add($incompatibleText) | Out-Null
            }
            elseif ($package.IsInstalled -eq $true) {
                # Show "Already installed" text in italics
                $installedText = New-Object System.Windows.Controls.TextBlock
                $installedText.Text = tr 'UI.AlreadyInstalled'
                $installedText.FontSize = 12
                $installedText.FontStyle = [System.Windows.FontStyles]::Italic
                $installedText.Foreground = [System.Windows.Media.BrushConverter]::new().ConvertFrom($themeColors.TextSecondary)
                $installedText.HorizontalAlignment = [System.Windows.HorizontalAlignment]::Center
                $installedText.Margin = [System.Windows.Thickness]::new(0,0,0,15)
                [System.Windows.Controls.Grid]::SetRow($installedText, 2)
                $grid.Children.Add($installedText) | Out-Null
            }
            else {
                # Show checkbox
                $checkbox = New-Object System.Windows.Controls.CheckBox
                $checkbox.HorizontalAlignment = [System.Windows.HorizontalAlignment]::Center
                $checkbox.Margin = [System.Windows.Thickness]::new(0,0,0,15)
                $checkbox.Tag = $packageKey
                [System.Windows.Controls.Grid]::SetRow($checkbox, 2)

                # Check if already selected
                $checkbox.IsChecked = $script:selectedPackagesDict.ContainsKey($packageKey)

                # Add event handlers
                $checkbox.Add_Checked({
                    param($s, $e)
                    $key = $s.Tag
                    $pkg = $Packages | Where-Object { (& $getPackageKey $_) -eq $key } | Select-Object -First 1

                    if ($pkg -and -not $script:selectedPackagesDict.ContainsKey($key)) {
                        $script:selectedPackagesDict[$key] = $pkg
                        Write-Verbose "Added: $($pkg.Name) (Key: $key)"
                    }

                    & $updateInstallButton
                })

                $checkbox.Add_Unchecked({
                    param($s, $e)
                    $key = $s.Tag

                    if ($script:selectedPackagesDict.ContainsKey($key)) {
                        $pkgName = $script:selectedPackagesDict[$key].Name
                        $script:selectedPackagesDict.Remove($key)
                        Write-Verbose "Removed: $pkgName (Key: $key)"
                    }

                    & $updateInstallButton
                })

                $grid.Children.Add($checkbox) | Out-Null
            }

            # Set grid background to transparent to ensure it captures mouse events
            $grid.Background = [System.Windows.Media.Brushes]::Transparent

            $border.Child = $grid

            # Only add interaction handlers if package is compatible and not installed
            if ($package.IsCompatible -ne $false -and $package.IsInstalled -ne $true) {
                # Click handler to toggle checkbox - use PreviewMouseLeftButtonDown to capture before children
                $border.Add_PreviewMouseLeftButtonDown({
                    param($s, $e)
                    $borderGrid = $s.Child
                    foreach ($child in $borderGrid.Children) {
                        if ($child -is [System.Windows.Controls.CheckBox]) {
                            $child.IsChecked = -not $child.IsChecked
                            $e.Handled = $true
                            break
                        }
                    }
                })

                # Visual feedback on hover - capture colors locally for this tile
                $localHoverColor = $themeColors.TileHover
                $localNormalColor = $themeColors.TileBackground

                $border.Add_MouseEnter({
                    param($s, $e)
                    $s.Background = [System.Windows.Media.BrushConverter]::new().ConvertFrom($localHoverColor)
                }.GetNewClosure())

                $border.Add_MouseLeave({
                    param($s, $e)
                    $s.Background = [System.Windows.Media.BrushConverter]::new().ConvertFrom($localNormalColor)
                }.GetNewClosure())
            }
            else {
                # For installed packages, disable cursor change
                $border.Cursor = [System.Windows.Input.Cursors]::Arrow
            }
            
            return $border
        }
        catch {
            Write-Error "Error creating tile for $($package.Name): $_"
            return $null
        }
    }
    
    # Function to display packages
    $showPackages = {
        param($translatedCategory)

        $packagesPanel.Children.Clear() | Out-Null

        # Get original category from mapping
        $originalCategory = $script:categoryMapping[$translatedCategory]

        # If mapping failed, log it
        if (-not $originalCategory) {
            Write-Warning "Category mapping not found for '$translatedCategory'. Available mappings: $($script:categoryMapping.Keys -join ', ')"
            return
        }

        # Get packages for the selected category and sort alphabetically by Name
        $packagesToShow = if ($originalCategory -eq "AllSoftware") {
            $Packages | Sort-Object -Property @{Expression = {$_.Name.ToString()}; Ascending = $true}
        } else {
            $Packages | Where-Object { $_.Category -eq $originalCategory } | Sort-Object -Property @{Expression = {$_.Name.ToString()}; Ascending = $true}
        }

        # Apply search filter
        $searchText = $searchBox.Text.Trim()
        if ($searchText.Length -gt 0) {
            $packagesToShow = @($packagesToShow | Where-Object { $_.Name -like "*$searchText*" })
        }

        foreach ($package in $packagesToShow) {
            try {
                $tile = & $newPackageTile $package $IconFolder $colors

                if ($null -ne $tile) {
                    $packagesPanel.Children.Add($tile) | Out-Null
                }
            }
            catch {
                Write-Warning "Failed to add tile for $($package.Name): $_"
            }
        }
    }
    
    # Create category buttons
    foreach ($category in $categories) {
        $button = New-Object System.Windows.Controls.RadioButton
        $button.Style = $window.FindResource("SidebarButtonStyle")
        $button.Content = $category
        $button.GroupName = "Categories"

        # Attach event handler BEFORE checking the button
        $button.Add_Checked({
            param($s, $e)
            $translatedCat = $s.Content
            $script:currentCategory = $translatedCat
            # Save the original category for language changes
            $script:selectedOriginalCategory = $script:categoryMapping[$translatedCat]
            & $showPackages $translatedCat
        })

        # Check if this is the currently selected category
        # This will trigger the Add_Checked event above
        if ($category -eq $script:currentCategory) {
            $button.IsChecked = $true
        }

        $categoriesPanel.Children.Add($button) | Out-Null
    }

    # Note: showPackages is automatically called by the Add_Checked event
    # when we set IsChecked = $true above, so no need to call it manually here

    # Search box: placeholder visibility toggle
    $searchBox.Add_TextChanged({
        if ($searchBox.Text.Length -gt 0) {
            $searchPlaceholder.Visibility = [System.Windows.Visibility]::Collapsed
        } else {
            $searchPlaceholder.Visibility = [System.Windows.Visibility]::Visible
        }
        # Re-filter packages with current category
        if ($script:currentCategory) {
            & $showPackages $script:currentCategory
        }
    })

    # Install button click handler - CLOSE WINDOW AND RETURN SELECTION
    $installButton.Add_Click({
        if ($script:selectedPackagesDict.Count -eq 0) {
            return
        }
        
        # Mark that user clicked Install
        $script:userClickedInstall = $true
        
        # Close the window
        $window.Close()
    })
    
    # Show initial packages (commented out - packages are shown by the checked category button)
    # & $showPackages $script:currentCategory

    # Show window
    $window.Add_Loaded({
        param($win, $e)
        # Activate window and bring to foreground
        Set-WPFWindowForeground -Window $win

        # Apply dark title bar if in dark mode
        if (Get-SystemTheme) {
            Set-DarkTitleBar -Window $win
        }
    })
    $window.ShowDialog() | Out-Null

    # Check if language change was requested
    if ($window.Tag -and $window.Tag.Action -eq "ReloadLanguage") {
        $reloadUI = $true
        Write-Verbose "Reloading UI with locale: $($window.Tag.Locale)"
    }

    } while ($reloadUI)

    # Return selected packages only if Install was clicked
    if ($script:userClickedInstall -and $script:selectedPackagesDict.Count -gt 0) {
        # Convert hashtable values to array
        $resultArray = @()
        $script:selectedPackagesDict.Values | ForEach-Object {
            $resultArray += $_
        }

        Write-Verbose "Returning $($resultArray.Count) packages after Install button click"

        return ,$resultArray
    } else {
        # User closed window without clicking Install
        Write-Verbose "Window closed without clicking Install button"
        return @()
    }
}

function Install-SelectedPackagesWithUI {
    <#
    .SYNOPSIS
        Installs selected packages with UI for progress

    .PARAMETER Packages
        Array of packages to install

    .PARAMETER IconFile
        Path to the icon file for the loading window

    .EXAMPLE
        Install-SelectedPackagesWithUI -Packages $selected -IconFile $iconFile
    #>

    [CmdletBinding()]
    Param(
        [Parameter(Mandatory=$true)]
        [array]$Packages,

        [Parameter(Mandatory=$false)]
        [string]$IconFile = ""
    )

    if ($Packages.Count -eq 0) {
        Write-Host "No software to install." -ForegroundColor Yellow
        return
    }

    # Show loading window for installation
    $loadingWindow = Show-LoadingWindow -Title (tr 'UI.InstallationInProgress') `
                                        -Message (tr 'UI.PreparingInstallation') `
                                        -IconFile $IconFile `
                                        -ShowDetailsText (tr 'UI.ShowDetails') `
                                        -HideDetailsText (tr 'UI.HideDetails') `
                                        -ShowDetailsSection
    
    try {
        $totalPackages = $Packages.Count
        
        Write-Host "`nStarting installation of $totalPackages package(s)..." -ForegroundColor Green
        Write-Host ("=" * 60) -ForegroundColor Gray
        
        # Use single UAC approach (no credential needed - UAC handles everything)
        $results = Install-AllPackagesWithSingleUAC -Packages $Packages -LoadingWindow $loadingWindow

        # Check if UAC was cancelled
        $uacCancelled = $results | Where-Object { $_.UACCancelled -eq $true } | Select-Object -First 1
        if ($uacCancelled) {
            # Close loading window first
            Close-LoadingWindow -Window $loadingWindow

            # Now show error dialog (after LoadingWindow is closed)
            Show-WPFButtonDialog -Title "Installation annulée" `
                -Message "L'installation a été annulée.`n`nDétails: $($uacCancelled.ErrorMessage)" `
                -Buttons @(@{text="OK"; value="ok"}) `
                -Icon "Error" | Out-Null

            return @()
        }

        # Count successes/failures/reboots (use @() to ensure array for .Count to work with single results)
        $successCount = @($results | Where-Object { $_.Success }).Count
        $failureCount = @($results | Where-Object { -not $_.Success }).Count
        $rebootRequired = @($results | Where-Object { $_.RebootRequired }).Count -gt 0

        # Close loading window
        Close-LoadingWindow -Window $loadingWindow

        # Show summary
        Write-Host "`n" -NoNewline
        Write-Host ("=" * 60) -ForegroundColor Gray
        Write-Host "`nInstallation summary:" -ForegroundColor Cyan
        Write-Host "  Total     : $totalPackages" -ForegroundColor White
        Write-Host "  Succeeded : $successCount" -ForegroundColor Green
        Write-Host "  Failed    : $failureCount" -ForegroundColor $(if ($failureCount -gt 0) { "Red" } else { "Gray" })

        if ($failureCount -gt 0) {
            Write-Host "`nFailed packages:" -ForegroundColor Red
            $results | Where-Object { -not $_.Success } | ForEach-Object {
                Write-Host "  • $($_.Name)" -ForegroundColor Red
            }
        }

        if ($rebootRequired) {
            Write-Host "`nReboot required for:" -ForegroundColor Yellow
            $results | Where-Object { $_.RebootRequired } | ForEach-Object {
                Write-Host "  • $($_.Name)" -ForegroundColor Yellow
            }
        }

        Write-Host "`n" -NoNewline
        Write-Host ("=" * 60) -ForegroundColor Gray

        # Show final message box with theme
        $message = if ($failureCount -eq 0) {
            tr 'UI.AllInstalledSuccessfully' -Parameters @($successCount, $totalPackages)
        } else {
            tr 'UI.InstallationCompletedWithFailures' -Parameters @($failureCount, $successCount, $totalPackages)
        }

        if ($rebootRequired) {
            $rebootPackages = @($results | Where-Object { $_.RebootRequired } | ForEach-Object { $_.Name }) -join ", "
            $message += "`n`n" + (tr 'UI.RebootRequired' -Parameters @($rebootPackages))
        }

        $icon = if ($rebootRequired) { "Warning" } elseif ($failureCount -eq 0) { "Information" } else { "Warning" }
        $buttons = if ($rebootRequired) {
            @(@{ text = (tr 'UI.RebootNow'); value = "reboot" }, @{ text = (tr 'UI.RebootLater'); value = "later" })
        } else {
            @(@{ text = "OK"; value = "ok" })
        }

        $dialogResult = Show-WPFButtonDialog -Title (tr 'UI.InstallationCompleted') -Message $message -Buttons $buttons -Icon $icon

        if ($dialogResult -eq "reboot") {
            Restart-Computer -Force
        }
        
        return $results
    }
    catch {
        Close-LoadingWindow -Window $loadingWindow
        Write-Error "Installation error: $_"
        Show-WPFButtonDialog -Title (tr 'UI.Error') `
            -Message (tr 'UI.InstallationError' -Parameters @($_)) `
            -Buttons @(@{text="OK"; value="ok"}) `
            -Icon "Error" | Out-Null
        return $null
    }
}

$InstalledPrograms = $null
# Load default packages
$packages = Get-Content -Path $(Get-RootScriptConfigFile -configFileName "apps.json") | ConvertFrom-Json | ConvertTo-Hashtable

# Load custom overrides (hide packages / add custom packages)
$customConfigFile = Get-RootScriptConfigFile -configFileName "apps_custom.json"
if ($customConfigFile) {
    $customConfig = Get-Content -Path $customConfigFile | ConvertFrom-Json | ConvertTo-Hashtable

    # Hide packages by ID
    if ($customConfig.Hide -and $customConfig.Hide.Count -gt 0) {
        $packages = $packages | Where-Object { $_.Id -notin $customConfig.Hide }
    }

    # Add custom packages
    if ($customConfig.Packages -and $customConfig.Packages.Count -gt 0) {
        $packages += $customConfig.Packages
    }
}

# Normalize scope: sources requiring elevation are always machine-scoped
$elevatedSources = @("windowscapability", "windowsfeature")
foreach ($pkg in $packages) {
    if ($pkg.Source -in $elevatedSources -and $pkg.Scope -ne "machine") {
        $pkg.Scope = "machine"
    }
}

$iconFolder = (Get-ScriptDir -InputDir) + "\icons\"
$iconFile = (Get-ScriptDir -InputDir -FullPath) + "\" + (Get-RootScriptName) + ".ico"
# Show loading window pour l'import des modules uniquement
$loadingWindow = Show-LoadingWindow -Title (tr 'UI.Initialization') `
                                    -Message (tr 'UI.ImportingModules') `
                                    -IconFile $iconFile

try {
    # Load catalog
    Update-LoadingWindow -Window $loadingWindow -Message (tr 'UI.LoadingWingetCatalog') -Progress 40
    Get-WingetPackageCatalog | Out-Null

    # Get package rowids from database
    Update-LoadingWindow -Window $loadingWindow -Message (tr 'UI.RetrievingPackageIds') -Progress 60
    $packages = Get-PackageRowIds -Packages $packages

    # Get product codes for packages
    Update-LoadingWindow -Window $loadingWindow -Message (tr 'UI.RetrievingProductCodes') -Progress 80
    $packages = Get-PackageProductCodes -Packages $packages

    # Check compatibility and installation status for each package
    Update-LoadingWindow -Window $loadingWindow -Message (tr 'UI.CheckingInstalledPackages') -Progress 90
    foreach ($package in $packages) {
        $package.IsCompatible = Test-PackageCompatibility -Package $package
        if ($package.IsCompatible) {
            $package.IsInstalled = Test-PackageInstalled -Package $package
        } else {
            $package.IsInstalled = $false
        }
    }

    # Close loading window
    Update-LoadingWindow -Window $loadingWindow -Message (tr 'UI.Finalizing') -Progress 100
    Close-LoadingWindow -Window $loadingWindow

} catch {
    Close-LoadingWindow -Window $loadingWindow
    Write-Error "Initialization error: $_"
    exit 1
}

# Show selection window
Write-Host "Opening selection interface..." -ForegroundColor Green

# Convert icon folder to absolute path for WPF
if ($iconFolder -and (Test-Path $iconFolder)) {
    $iconFolderAbsolute = (Resolve-Path $iconFolder).Path
} else {
    $iconFolderAbsolute = ""
}

$selected = Show-PackageManagerUI -Packages $packages -IconFolder $iconFolderAbsolute -IconFile $iconFile

# Si des packages sont sélectionnés, charger les informations d'installation
if ($selected -and $selected.Count -gt 0) {
    Write-Host "`n$($selected.Count) package(s) selected" -ForegroundColor Green

    # Sort packages: machine scope first, then respect dependencies
    Write-Host "Sorting packages by scope and dependencies..." -ForegroundColor Cyan
    $sortResult = Get-SortedPackagesByDependencies -Packages $selected -AllAvailablePackages $packages

    # Check for missing dependencies
    if ($sortResult.MissingDependencies -and $sortResult.MissingDependencies.Count -gt 0) {
        # Build message with missing dependencies
        $message = (tr 'UI.MissingDependenciesHeader') + "`n`n"

        $packagesToAdd = @()
        $missingNotFound = @()

        foreach ($missing in $sortResult.MissingDependencies) {
            if ($missing.RequiredPackage) {
                $message += "• $($missing.Package) → $($missing.RequiredPackage.Name)`n"
                if ($packagesToAdd -notcontains $missing.RequiredPackage) {
                    $packagesToAdd += $missing.RequiredPackage
                }
            }
            else {
                $message += "• $($missing.Package) → $($missing.RequiredId) (" + (tr 'UI.NotFound') + ")`n"
                $missingNotFound += $missing
            }
        }

        if ($packagesToAdd.Count -gt 0) {
            $message += "`n" + (tr 'UI.MissingPackagesFound') + "`n"
            foreach ($pkg in $packagesToAdd) {
                $message += "  + $($pkg.Name)`n"
            }

            $buttons = @(
                @{ text = (tr 'UI.AddAutomatically'); value = "add" },
                @{ text = (tr 'UI.ContinueWithoutAdding'); value = "continue" },
                @{ text = (tr 'UI.CancelInstallation'); value = "cancel" }
            )

            $choice = Show-WPFButtonDialog -Title (tr 'UI.MissingDependencies') -Message $message -Buttons $buttons -Icon Warning

            if ($choice -eq "add") {
                # Add missing packages to selection (avoid duplicates)
                $addedCount = 0
                foreach ($pkgToAdd in $packagesToAdd) {
                    $alreadyExists = $selected | Where-Object { $_.Id -eq $pkgToAdd.Id }
                    if (-not $alreadyExists) {
                        $selected += $pkgToAdd
                        $addedCount++
                        Write-Verbose "Ajout de $($pkgToAdd.Name) ($($pkgToAdd.Id))"
                    }
                }
                Write-Host "Dependencies added ($addedCount package(s)). Total: $($selected.Count). Re-sorting..." -ForegroundColor Green

                # Re-sort with complete package list
                try {
                    $sortResult = Get-SortedPackagesByDependencies -Packages $selected -AllAvailablePackages $packages #-Verbose
                    Write-Host "Sorting completed. Sorted packages: $($sortResult.SortedPackages.Count)" -ForegroundColor Green

                    if ($sortResult.SortedPackages.Count -eq 0) {
                        Write-Warning "Sorting returned 0 packages!"
                        Write-Host "Packages before sorting: $($selected.Count)" -ForegroundColor Yellow
                        Write-Host "Package details:" -ForegroundColor Yellow
                        foreach ($pkg in $selected) {
                            Write-Host "  - $($pkg.Name) (Id: $($pkg.Id), Scope: $($pkg.Scope))" -ForegroundColor Yellow
                        }
                    }
                }
                catch {
                    Write-Error "Sorting error: $_"
                    exit 1
                }
            }
            elseif ($choice -eq "cancel" -or $null -eq $choice) {
                Write-Host "Installation cancelled by user." -ForegroundColor Yellow
                exit 0
            }
            # If "continue", proceed without adding
        }
        elseif ($missingNotFound.Count -gt 0) {
            # Only missing packages that cannot be found
            $buttons = @(
                @{ text = (tr 'UI.ContinueAnyway'); value = "continue" },
                @{ text = (tr 'UI.CancelInstallation'); value = "cancel" }
            )

            $choice = Show-WPFButtonDialog -Title (tr 'UI.DependenciesNotFound') -Message $message -Buttons $buttons -Icon Error

            if ($choice -eq "cancel" -or $null -eq $choice) {
                Write-Host "Installation cancelled by user." -ForegroundColor Yellow
                exit 0
            }
        }
    }

    $selected = $sortResult.SortedPackages

    # Verify sorted packages
    if (-not $selected -or $selected.Count -eq 0) {
        Write-Error "Error sorting packages. The list is empty."
        exit 1
    }

    # Display installation order
    Write-Host "`nInstallation order ($($selected.Count) package(s)):" -ForegroundColor Yellow
    $index = 1
    foreach ($pkg in $selected) {
        if (-not $pkg) {
            Write-Warning "Null package detected at index $index"
            continue
        }
        $scopeLabel = if ($pkg.Scope -eq "machine") { "[Machine]" } else { "[User]" }
        $requiresLabel = if ($pkg.Requires -and $pkg.Requires.Count -gt 0) { " (Dependencies: $($pkg.Requires -join ', '))" } else { "" }
        Write-Host "  $index. $scopeLabel $($pkg.Name)$requiresLabel" -ForegroundColor Gray
        $index++
    }
    Write-Host ""

    # Show loading window for installers
    $loadingWindow = Show-LoadingWindow -Title (tr 'UI.Preparation') `
                                        -Message (tr 'UI.RetrievingInstallationInfo') `
                                        -IconFile $iconFile
    
    try {
        $wingetPackages = @($selected | Where-Object { $_.Source -eq "winget" })
        $totalPackages = $wingetPackages.Count
        $currentPackage = 0
        
        foreach ($package in $wingetPackages) {
            $currentPackage++
            $percentComplete = [int](($currentPackage / $totalPackages) * 100)
            
            Update-LoadingWindow -Window $loadingWindow `
                                -Message ((tr 'UI.RetrievingPackage' -Parameters @($package.Name, $currentPackage, $totalPackages))) `
                                -Progress $percentComplete
            
            try {
                $package.Installer = Get-WingetPackageInstaller -PackageId $package.Id `
                                                                -Architecture "x64" `
                                                                -BackupArchitecture "x86" `
                                                                -Scope $package.Scope
                
                Write-Host "  ✓ $($package.Name)" -ForegroundColor Cyan
            }
            catch {
                Write-Warning "Unable to retrieve installer for $($package.Name): $_"
                $package.Installer = $null
            }
        }
        
        Close-LoadingWindow -Window $loadingWindow

        # Lancer l'installation avec UI (credentials + progression) - TOUS LES PACKAGES
        $results = Install-SelectedPackagesWithUI -Packages $selected -IconFile $iconFile
        
        # Afficher les résultats détaillés
        if ($results) {
            Write-Host "`nInstallation details:" -ForegroundColor Cyan
            $results | Format-Table Name, Source, Success -AutoSize
        }

        exit 0
    } catch {
        Close-LoadingWindow -Window $loadingWindow
        Write-Error "Error retrieving installers: $_"
        exit 1
    }

} else {
    Write-Host "`nNo software selected." -ForegroundColor Yellow
    exit 0
}
