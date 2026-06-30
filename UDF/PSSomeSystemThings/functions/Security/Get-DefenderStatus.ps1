function Get-DefenderStatus {
    <#
    .SYNOPSIS
        Reports Microsoft Defender Antivirus status of a host, with remote execution support.

    .DESCRIPTION
        Wraps Get-MpComputerStatus and returns the key antivirus posture fields: service and
        real-time protection state, signature age/last update, last scan and tamper protection.
        Evidence of malware-protection coverage. Returns nothing when Defender is not present
        on the target (cmdlet unavailable).

        Supports remote execution via -ComputerName / -Credential or an existing -Session.
        When neither is supplied the local machine is queried.

    .PARAMETER ComputerName
        Remote computer name(s) to query.

    .PARAMETER Credential
        Credentials for remote execution.

    .PARAMETER Session
        Existing PSSession(s) for remote execution. Mutually exclusive with -ComputerName.

    .OUTPUTS
        [PSCustomObject] Antivirus status fields.

    .EXAMPLE
        Get-DefenderStatus -ComputerName SRV01 -Credential $cred

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
            if (-not (Get-Command Get-MpComputerStatus -ErrorAction SilentlyContinue)) { return }
            $s = Get-MpComputerStatus -ErrorAction SilentlyContinue
            if (-not $s) { return }
            [PSCustomObject][ordered]@{
                AMServiceEnabled          = $s.AMServiceEnabled
                RealTimeProtectionEnabled = $s.RealTimeProtectionEnabled
                AntivirusEnabled          = $s.AntivirusEnabled
                IsTamperProtected         = $s.IsTamperProtected
                AntivirusSignatureAgeDays = $s.AntivirusSignatureAge
                AntivirusSignatureUpdated = $s.AntivirusSignatureLastUpdated
                LastQuickScan             = $s.QuickScanEndTime
                LastFullScan              = $s.FullScanEndTime
                AMProductVersion          = $s.AMProductVersion
                AMEngineVersion           = $s.AMEngineVersion
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
