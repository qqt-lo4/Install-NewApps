function Get-RootScriptPath {
    <#
    .SYNOPSIS
        Gets the root script's directory path

    .DESCRIPTION
        Walks the call stack to determine the directory of the top-level calling script.

    .PARAMETER FullPath
        Resolve and return the full absolute path.

    .OUTPUTS
        [String]. The root script directory path.

    .EXAMPLE
        $scriptDir = Get-RootScriptPath -FullPath

    .NOTES
        Author  : Loïc Ade
        Version : 1.0.0
    #>

    Param(
        [switch]$FullPath
    )
    $scriptCallStack = Get-PSCallStack | Where-Object { $_.Command -ne '<ScriptBlock>' } 
    $rootScriptFullPath = $scriptCallStack[-1].InvocationInfo.InvocationName
    $rootScriptName = $scriptCallStack[-1].InvocationInfo.MyCommand.Name
    $sResult = if (($rootScriptFullPath.Length - $rootScriptName.Length) -lt 0) {
        ""
    } else {
        $rootScriptFullPath.Remove($rootScriptFullPath.Length - $rootScriptName.Length)
    }
    if ($FullPath.IsPresent) {
        if ($sResult -eq "") {
            (Resolve-Path -Path ".").Path
        } else {
            (Resolve-Path -Path $sResult).Path
        }
    } else {
        return $sResult
    }
}
