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

    .PARAMETER FullPath
        Return the full absolute path instead of relative.

    .OUTPUTS
        [String]. Directory path.

    .EXAMPLE
        $inputDir = Get-ScriptDir -InputDir

    .EXAMPLE
        $toolsDir = Get-ScriptDir -ToolsDir -ToolName "7zip"

    .NOTES
        Author  : Loïc Ade
        Version : 1.0.0

        1.0.0 - First version
        1.1.0 (2026-03-05)
            - Corrected bugs of Get-RootScriptPath
            - Removes -FullPath parameter (always returns full path)
    #>

    Param(
        [Parameter(ParameterSetName = "input", Mandatory)]
        [switch]$InputDir,
        [Parameter(ParameterSetName = "output", Mandatory)]
        [switch]$OutputDir,
        [Parameter(ParameterSetName = "working_dir", Mandatory)]
        [switch]$WorkingDir,
        [Parameter(ParameterSetName = "tools", Mandatory)]
        [switch]$ToolsDir,
        [Parameter(ParameterSetName = "tools", Mandatory)]
        [string]$ToolName
    )
    Begin {
        function Resolve-RelativePath {
            Param(
                [string]$From,
                [string]$To
            )
            $oLocationBefore = Get-Location
            Set-Location $From 
            Resolve-Path -Path $To -Relative
            Set-Location $oLocationBefore
        }
    }
    Process {
        $sRootPath = Get-RootScriptPath
        $sResult = $sRootPath + "\" + $PSCmdlet.ParameterSetName 
        if ($PSCmdlet.ParameterSetName -eq "tools") {
            $sResult += "\" + $ToolName
        }
        if (Test-Path ($sRootPath + "\.devfolder")) {
            $sResult = switch ($PSCmdlet.ParameterSetName) {
                "tools" { $sResult } 
                default {$sResult + "\" + (Get-RootScriptName)}
            }
        }
        return $sResult
    }
    End {}
}