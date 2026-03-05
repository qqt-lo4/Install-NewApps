function Get-RootScriptPath {
    <#
    .SYNOPSIS
        Gets the root script's directory path

    .DESCRIPTION
        Walks the call stack to determine the directory of the top-level calling script.

    .OUTPUTS
        [String]. The root script directory path.

    .EXAMPLE
        $scriptDir = Get-RootScriptPath

    .NOTES
        Author  : Loïc Ade
        Version : 1.0.0
    #>
    $scriptCallStack = Get-PSCallStack | Where-Object { $_.Command -ne '<ScriptBlock>' } 
    $rootScriptFullPath = $scriptCallStack[-1].ScriptName
    $rootScriptName = $scriptCallStack[-1].Command.ToString()
    return $rootScriptFullPath.Substring(0, ($rootScriptFullPath.Length - $rootScriptName.Length - 1))
}
