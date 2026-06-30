function Get-SmbServerSecurityConfiguration {
    <#
    .SYNOPSIS
        Reports the SMB server security configuration of a host, with remote execution support.

    .DESCRIPTION
        Wraps Get-SmbServerConfiguration and returns the security-relevant settings: whether
        the legacy SMBv1 protocol is enabled (a common audit finding), whether SMB signing is
        required/enabled, and whether SMB encryption is enforced.

        Supports remote execution via -ComputerName / -Credential or an existing -Session.
        When neither is supplied the local machine is queried.

    .PARAMETER ComputerName
        Remote computer name(s) to query.

    .PARAMETER Credential
        Credentials for remote execution.

    .PARAMETER Session
        Existing PSSession(s) for remote execution. Mutually exclusive with -ComputerName.

    .OUTPUTS
        [PSCustomObject] SMB server security fields.

    .EXAMPLE
        Get-SmbServerSecurityConfiguration -ComputerName SRV01 -Credential $cred

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
            if (-not (Get-Command Get-SmbServerConfiguration -ErrorAction SilentlyContinue)) { return }
            $c = Get-SmbServerConfiguration -ErrorAction SilentlyContinue
            if (-not $c) { return }
            [PSCustomObject][ordered]@{
                EnableSMB1Protocol      = $c.EnableSMB1Protocol
                EnableSMB2Protocol      = $c.EnableSMB2Protocol
                RequireSecuritySignature = $c.RequireSecuritySignature
                EnableSecuritySignature = $c.EnableSecuritySignature
                EncryptData             = $c.EncryptData
                RejectUnencryptedAccess = $c.RejectUnencryptedAccess
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
