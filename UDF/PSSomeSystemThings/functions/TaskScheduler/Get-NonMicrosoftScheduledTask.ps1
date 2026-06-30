function Get-NonMicrosoftScheduledTask {
    <#
    .SYNOPSIS
        Lists non-Microsoft scheduled tasks of a host, with remote execution support.

    .DESCRIPTION
        Wraps Get-ScheduledTask and returns the tasks that do not live under the built-in
        \Microsoft\ task path, with the account they run as, their triggers and the action
        executed. Surfaces third-party / custom scheduled jobs (persistence / integrity).

        Supports remote execution via -ComputerName / -Credential or an existing -Session.
        When neither is supplied the local machine is queried.

    .PARAMETER ComputerName
        Remote computer name(s) to query.

    .PARAMETER Credential
        Credentials for remote execution.

    .PARAMETER Session
        Existing PSSession(s) for remote execution. Mutually exclusive with -ComputerName.

    .OUTPUTS
        [PSCustomObject[]] One object per task: TaskName, TaskPath, State, RunAs, Author,
        Triggers, Action.

    .EXAMPLE
        Get-NonMicrosoftScheduledTask -ComputerName SRV01 -Credential $cred

    .NOTES
        Author  : Loic Ade
        Version : 1.0.0

        Version History:
        1.0.0 - Initial version
    #>
    [CmdletBinding()]
    Param(
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
            if (-not (Get-Command Get-ScheduledTask -ErrorAction SilentlyContinue)) { return }
            Get-ScheduledTask -ErrorAction SilentlyContinue |
                Where-Object { $_.TaskPath -notlike '\Microsoft\*' } |
                ForEach-Object {
                    $t = $_
                    $aTriggers = @(foreach ($g in @($t.Triggers)) {
                        $sType = ($g.CimClass.CimClassName -replace '^MSFT_Task', '' -replace 'Trigger$', '')
                        if ($g.StartBoundary) { "$sType @ $($g.StartBoundary)" } else { $sType }
                    })
                    $aActions = @(foreach ($a in @($t.Actions)) {
                        (@($a.Execute, $a.Arguments) | Where-Object { $_ }) -join ' '
                    })
                    [PSCustomObject][ordered]@{
                        TaskName = $t.TaskName
                        TaskPath = $t.TaskPath
                        State    = "$($t.State)"
                        RunAs    = $t.Principal.UserId
                        Author   = $t.Author
                        Triggers = ($aTriggers -join '; ')
                        Action   = ($aActions -join ' | ')
                    }
                } | Sort-Object TaskPath, TaskName
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
        Invoke-Command @hRemote -ScriptBlock $oScriptBlock |
            Select-Object -Property * -ExcludeProperty RunspaceId, PSComputerName, PSShowComputerName
    }
}
