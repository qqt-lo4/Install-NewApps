function Get-EventLogEntry {
    <#
    .SYNOPSIS
        Retrieves the most recent events of selected logs of a Windows host, with remote support.

    .DESCRIPTION
        Returns up to -MaxEventsPerLog of the newest events from each selected log. Logs are
        selected with the same rule as Get-EventLogConfiguration: the "Windows Logs" plus,
        optionally, the populated classic server-role logs. Every returned event carries its
        source LogName so the caller can group events per log.

        Supports remote execution via -ComputerName / -Credential or an existing -Session.
        When neither is supplied the local machine is queried.

    .PARAMETER MaxEventsPerLog
        Maximum number of events retrieved per log, most recent first. Default: 100.

    .PARAMETER WindowsLogs
        Names of the "Windows Logs" considered. Default: Application, Security, Setup, System,
        ForwardedEvents.

    .PARAMETER IncludeClassicRoleLogs
        Also include the populated classic server-role logs. Default: $true.

    .PARAMETER ComputerName
        Remote computer name(s) to query.

    .PARAMETER Credential
        Credentials for remote execution.

    .PARAMETER Session
        Existing PSSession(s) for remote execution. Mutually exclusive with -ComputerName.

    .OUTPUTS
        [PSCustomObject[]] One object per event: LogName, TimeCreated, Level, EventId,
        Provider, Message (first line).

    .EXAMPLE
        Get-EventLogEntry -ComputerName SRV01 -Credential $cred -MaxEventsPerLog 100

    .NOTES
        Author  : Loic Ade
        Version : 1.0.0

        Version History:
        1.0.0 - Initial version
    #>
    [CmdletBinding()]
    Param(
        [int]$MaxEventsPerLog = 100,

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
            $aWinLogs     = $Params.WindowsLogs
            $bIncludeRole = $Params.IncludeClassicRoleLogs
            $iMax         = $Params.MaxEventsPerLog

            $hLevel = @{ 0 = 'Information'; 1 = 'Critical'; 2 = 'Error'; 3 = 'Warning'; 4 = 'Information'; 5 = 'Verbose' }

            $aLogs = @(Get-WinEvent -ListLog * -ErrorAction SilentlyContinue |
                Where-Object { (($_.LogName -in $aWinLogs) -or ($bIncludeRole -and $_.IsClassicLog)) -and $_.RecordCount -gt 0 })

            foreach ($oLog in $aLogs) {
                try {
                    Get-WinEvent -LogName $oLog.LogName -MaxEvents $iMax -ErrorAction Stop |
                        ForEach-Object {
                            [PSCustomObject][ordered]@{
                                LogName     = $oLog.LogName
                                TimeCreated = $_.TimeCreated
                                Level       = if ($hLevel.ContainsKey([int]$_.Level)) { $hLevel[[int]$_.Level] } else { "$($_.LevelDisplayName)" }
                                EventId     = $_.Id
                                Provider    = $_.ProviderName
                                Message     = if ($_.Message) { ($_.Message -split "`r?`n")[0] } else { '' }
                            }
                        }
                } catch {
                    # An empty match between ListLog and the read is not an error worth surfacing.
                    if ($_.FullyQualifiedErrorId -notlike 'NoMatchingEventsFound*') { throw }
                }
            }
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
        $oArgs = @{
            WindowsLogs            = $WindowsLogs
            IncludeClassicRoleLogs = $IncludeClassicRoleLogs
            MaxEventsPerLog        = $MaxEventsPerLog
        }
        Invoke-Command @hRemote -ScriptBlock $oScriptBlock -ArgumentList $oArgs |
            Select-Object -Property * -ExcludeProperty RunspaceId, PSComputerName, PSShowComputerName
    }
}
