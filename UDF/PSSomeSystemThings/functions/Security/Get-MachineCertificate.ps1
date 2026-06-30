function Get-MachineCertificate {
    <#
    .SYNOPSIS
        Lists machine certificates of a host, with remote execution support.

    .DESCRIPTION
        Enumerates the certificates of the LocalMachine certificate store (default: the
        personal "My" store) and returns subject, issuer, validity, days remaining, private-key
        presence and enhanced key usages. Useful to spot expired / soon-to-expire certificates.

        Supports remote execution via -ComputerName / -Credential or an existing -Session.
        When neither is supplied the local machine is queried.

    .PARAMETER StoreName
        LocalMachine store to read. Default: My.

    .PARAMETER ComputerName
        Remote computer name(s) to query.

    .PARAMETER Credential
        Credentials for remote execution.

    .PARAMETER Session
        Existing PSSession(s) for remote execution. Mutually exclusive with -ComputerName.

    .OUTPUTS
        [PSCustomObject[]] One object per certificate: Subject, Issuer, NotBefore, NotAfter,
        DaysRemaining, HasPrivateKey, Thumbprint, FriendlyName, EnhancedKeyUsage.

    .EXAMPLE
        Get-MachineCertificate -ComputerName SRV01 -Credential $cred

    .NOTES
        Author  : Loic Ade
        Version : 1.0.0

        Version History:
        1.0.0 - Initial version
    #>
    [CmdletBinding()]
    Param(
        [string]$StoreName = 'My',

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
            $sStore = $Params.StoreName
            $sPath  = "Cert:\LocalMachine\$sStore"
            if (-not (Test-Path $sPath)) { return }
            $dtNow = Get-Date
            Get-ChildItem -Path $sPath -ErrorAction SilentlyContinue | ForEach-Object {
                $c = $_
                [PSCustomObject][ordered]@{
                    Subject         = $c.Subject
                    Issuer          = $c.Issuer
                    NotBefore       = $c.NotBefore
                    NotAfter        = $c.NotAfter
                    DaysRemaining   = [int]([math]::Floor(($c.NotAfter - $dtNow).TotalDays))
                    HasPrivateKey   = $c.HasPrivateKey
                    Thumbprint      = $c.Thumbprint
                    FriendlyName    = $c.FriendlyName
                    EnhancedKeyUsage = (@($c.EnhancedKeyUsageList | ForEach-Object { $_.FriendlyName }) -join ', ')
                }
            } | Sort-Object NotAfter
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
        Invoke-Command @hRemote -ScriptBlock $oScriptBlock -ArgumentList @{ StoreName = $StoreName } |
            Select-Object -Property * -ExcludeProperty RunspaceId, PSComputerName, PSShowComputerName
    }
}
