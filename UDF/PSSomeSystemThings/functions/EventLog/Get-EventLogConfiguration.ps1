function Get-EventLogConfiguration {
    <#
    .SYNOPSIS
        Reports event log configuration of a Windows host, with remote execution support.

    .DESCRIPTION
        Wraps Get-WinEvent -ListLog and returns the configuration (enabled state, retention
        mode, maximum size, current record count, backing file) of the selected logs.

        Log selection mirrors the Event Viewer tree an auditor expects:
            - the "Windows Logs" (Application, Security, Setup, System, ForwardedEvents);
            - optionally the classic "server role" logs that sit directly under
              "Applications and Services Logs" (DNS Server, Directory Service,
              DFS Replication, File Replication Service, ...). These are the classic
              (IsClassicLog) logs that actually hold records, which keeps the noisy
              Microsoft-Windows-* operational channels out.

        Supports remote execution via -ComputerName / -Credential or an existing -Session.
        When neither is supplied the local machine is queried.

    .PARAMETER WindowsLogs
        Names of the "Windows Logs" always included regardless of record count.
        Default: Application, Security, Setup, System, ForwardedEvents.

    .PARAMETER IncludeClassicRoleLogs
        Also include the populated classic server-role logs. Default: $true.

    .PARAMETER ComputerName
        Remote computer name(s) to query.

    .PARAMETER Credential
        Credentials for remote execution.

    .PARAMETER Session
        Existing PSSession(s) for remote execution. Mutually exclusive with -ComputerName.

    .OUTPUTS
        [PSCustomObject[]] One object per log: LogName, Enabled, LogMode, MaximumSizeMB,
        RecordCount, IsClassic, LogFilePath.

    .EXAMPLE
        Get-EventLogConfiguration -ComputerName SRV01 -Credential $cred

    .NOTES
        Author  : Loic Ade
        Version : 1.0.0

        Version History:
        1.0.0 - Initial version
    #>
    [CmdletBinding()]
    Param(
        [string[]]$WindowsLogs = @('Application', 'Security', 'Setup', 'System', 'ForwardedEvents'),

        [bool]$IncludeClassicRoleLogs = $true,

        [Alias("Cn")]
        [string[]]$ComputerName,

        [pscredential]$Credential,

        [System.Management.Automation.Runspaces.PSSession[]]$Session
    )

    Begin {
        if ($Session -and $ComputerName) {
            throw "Incompatible arguments : you can't use Session and ComputerName at the same time"
        }
        $oScriptBlock = {
            Param([object]$Params)
            $aWinLogs    = $Params.WindowsLogs
            $bIncludeRole = $Params.IncludeClassicRoleLogs
            Get-WinEvent -ListLog * -ErrorAction SilentlyContinue |
                Where-Object { ($_.LogName -in $aWinLogs) -or ($bIncludeRole -and $_.IsClassicLog -and $_.RecordCount -gt 0) } |
                ForEach-Object {
                    [PSCustomObject][ordered]@{
                        LogName       = $_.LogName
                        Enabled       = $_.IsEnabled
                        LogMode       = "$($_.LogMode)"
                        MaximumSizeMB = [math]::Round($_.MaximumSizeInBytes / 1MB, 1)
                        RecordCount   = $_.RecordCount
                        IsClassic     = $_.IsClassicLog
                        LogFilePath   = $_.LogFilePath
                    }
                } | Sort-Object @{ Expression = { $_.LogName -notin $aWinLogs } }, LogName
        }
    }
    Process {
        $hRemote = @{}
        if ($Session) {
            $hRemote.Session = $Session
        } elseif ($ComputerName) {
            $hRemote.ComputerName = $ComputerName
            if ($Credential) { $hRemote.Credential = $Credential }
        }
        $oArgs = @{ WindowsLogs = $WindowsLogs; IncludeClassicRoleLogs = $IncludeClassicRoleLogs }
        Invoke-Command @hRemote -ScriptBlock $oScriptBlock -ArgumentList $oArgs |
            Select-Object -Property * -ExcludeProperty RunspaceId, PSComputerName, PSShowComputerName
    }
}
