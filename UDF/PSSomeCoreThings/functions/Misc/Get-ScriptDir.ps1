function Get-ScriptDir {
    <#
    .SYNOPSIS
        Gets application directories (input, output, working, or tools)

    .DESCRIPTION
        Returns the path to a standard application subfolder relative to the root script.
        Supports dev folder structure detection for organized project layouts.

    .PARAMETER InputDir
        Return the input directory path.

    .PARAMETER OutputDir
        Return the output directory path.

    .PARAMETER WorkingDir
        Return the working directory path.

    .PARAMETER ToolsDir
        Return the tools directory path (requires ToolName).

    .PARAMETER ToolName
        Name of the tool subfolder under tools.

    .PARAMETER CustomModules
        Return the custom (client-specific) modules directory path, named
        "Custom_Modules" under the root script directory. Can be redirected by a
        root script argument (e.g. to point at a client's shared modules folder),
        and is not nested under the project name when a .devfolder marker is present.

    .OUTPUTS
        [String]. Directory path.

    .EXAMPLE
        $inputDir = Get-ScriptDir -InputDir

    .EXAMPLE
        $toolsDir = Get-ScriptDir -ToolsDir -ToolName "7zip"

    .EXAMPLE
        $customModulesDir = Get-ScriptDir -CustomModules

    .NOTES
        Author  : Loïc Ade
        Version : 1.5.0

        1.0.0 - First version

        1.1.0 (2026-03-05)
            - Corrected bugs of Get-RootScriptPath
            - Removes -FullPath parameter (always returns full path)

        1.2.0 (2026-03-08)
            - InputDir, OutputDir and WorkingDir can be overridden by root script parameters
            - ParameterSetNames renamed to match parameter names
            - Folder name derived from ParameterSetName

        1.3.0 (2026-03-10)
            - Uses Get-RootScriptInfo instead of Get-RootScriptPath, Get-RootScriptName and Get-RootScriptArguments

        1.4.0 (2026-04-23)
            - Emits a Write-Warning when a redirected directory argument is provided but does not exist,
              instead of silently falling back to the default path

        1.5.0 (2026-06-14)
            - Added -CustomModules parameter set returning the "Custom_Modules" folder
              (client-specific modules), redirectable by a root script argument and not
              nested under the project name under .devfolder

    #>

    Param(
        [Parameter(ParameterSetName = "InputDir", Mandatory)]
        [switch]$InputDir,
        [Parameter(ParameterSetName = "OutputDir", Mandatory)]
        [switch]$OutputDir,
        [Parameter(ParameterSetName = "WorkingDir", Mandatory)]
        [switch]$WorkingDir,
        [Parameter(ParameterSetName = "ToolsDir", Mandatory)]
        [switch]$ToolsDir,
        [Parameter(ParameterSetName = "ToolsDir", Mandatory)]
        [string]$ToolName,
        [Parameter(ParameterSetName = "CustomModules", Mandatory)]
        [switch]$CustomModules
    )
    Begin {
        $rootInfo = Get-RootScriptInfo
    }
    Process {
        if ($InputDir -or $OutputDir -or $WorkingDir -or $CustomModules) {
            $sRootArgValue = $rootInfo.Arguments[$PSCmdlet.ParameterSetName]
            if (-not [string]::IsNullOrEmpty($sRootArgValue)) {
                if (Test-Path $sRootArgValue -PathType Container) {
                    return $sRootArgValue
                }
                Write-Warning "Redirected $($PSCmdlet.ParameterSetName) path does not exist: '$sRootArgValue'. Falling back to default."
            }
        }

        $sFolderName = if ($PSCmdlet.ParameterSetName -eq "CustomModules") {
            # Fixed name (not derived): the custom client modules folder.
            "Custom_Modules"
        } else {
            $sTmp = $PSCmdlet.ParameterSetName -replace 'Dir$', ''
            $sTmp.Substring(0, 1).ToLower() + $sTmp.Substring(1)
        }
        $sResult = $rootInfo.Directory + "\" + $sFolderName
        if ($PSCmdlet.ParameterSetName -eq "ToolsDir") {
            $sResult += "\" + $ToolName
        }
        if (Test-Path ($rootInfo.Directory + "\.devfolder")) {
            $sResult = switch ($PSCmdlet.ParameterSetName) {
                # Tools and custom modules are not nested under the project name.
                "ToolsDir" { $sResult }
                "CustomModules" { $sResult }
                default { $sResult + "\" + $rootInfo.Name }
            }
        }
        return $sResult
    }
    End {}
}