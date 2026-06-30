function Get-FirewallProfileStatus {
    <#
    .SYNOPSIS
        Reports the Windows Firewall profile status of a host, with remote execution support.

    .DESCRIPTION
        Wraps Get-NetFirewallProfile and returns, per profile (Domain / Private / Public),
        whether the firewall is enabled and the default inbound/outbound actions plus the
        logging configuration. Evidence that host-based filtering is in place.

        Supports remote execution via -ComputerName / -Credential or an existing -Session.
        When neither is supplied the local machine is queried.

    .PARAMETER ComputerName
        Remote computer name(s) to query.

    .PARAMETER Credential
        Credentials for remote execution.

    .PARAMETER Session
        Existing PSSession(s) for remote execution. Mutually exclusive with -ComputerName.

    .OUTPUTS
        [PSCustomObject[]] One object per profile: Profile, Enabled, DefaultInboundAction,
        DefaultOutboundAction, AllowInboundRules, LogBlocked, LogAllowed, LogFileName.

    .EXAMPLE
        Get-FirewallProfileStatus -ComputerName SRV01 -Credential $cred

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
            if (-not (Get-Command Get-NetFirewallProfile -ErrorAction SilentlyContinue)) { return }
            Get-NetFirewallProfile -ErrorAction SilentlyContinue | ForEach-Object {
                [PSCustomObject][ordered]@{
                    Profile               = $_.Name
                    Enabled               = [bool]$_.Enabled
                    DefaultInboundAction  = "$($_.DefaultInboundAction)"
                    DefaultOutboundAction = "$($_.DefaultOutboundAction)"
                    AllowInboundRules     = "$($_.AllowInboundRules)"
                    LogBlocked            = "$($_.LogBlocked)"
                    LogAllowed            = "$($_.LogAllowed)"
                    LogFileName           = $_.LogFileName
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
        Invoke-Command @hRemote -ScriptBlock $oScriptBlock |
            Select-Object -Property * -ExcludeProperty RunspaceId, PSComputerName, PSShowComputerName
    }
}
