function Get-BitLockerStatus {
    <#
    .SYNOPSIS
        Reports BitLocker volume encryption status of a host, with remote execution support.

    .DESCRIPTION
        Wraps Get-BitLockerVolume and returns, per volume, the protection and encryption
        status, method, percentage and configured key protectors. Evidence of data-at-rest
        protection. Returns nothing when the BitLocker module is not present on the target
        (e.g. the feature is not installed).

        Supports remote execution via -ComputerName / -Credential or an existing -Session.
        When neither is supplied the local machine is queried.

    .PARAMETER ComputerName
        Remote computer name(s) to query.

    .PARAMETER Credential
        Credentials for remote execution.

    .PARAMETER Session
        Existing PSSession(s) for remote execution. Mutually exclusive with -ComputerName.

    .OUTPUTS
        [PSCustomObject[]] One object per volume: MountPoint, VolumeType, VolumeStatus,
        ProtectionStatus, EncryptionPercentage, EncryptionMethod, KeyProtectors.

    .EXAMPLE
        Get-BitLockerStatus -ComputerName SRV01 -Credential $cred

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
            if (-not (Get-Command Get-BitLockerVolume -ErrorAction SilentlyContinue)) { return }
            Get-BitLockerVolume -ErrorAction SilentlyContinue | ForEach-Object {
                [PSCustomObject][ordered]@{
                    MountPoint           = $_.MountPoint
                    VolumeType           = "$($_.VolumeType)"
                    VolumeStatus         = "$($_.VolumeStatus)"
                    ProtectionStatus     = "$($_.ProtectionStatus)"
                    EncryptionPercentage = $_.EncryptionPercentage
                    EncryptionMethod     = "$($_.EncryptionMethod)"
                    KeyProtectors        = (@($_.KeyProtector | ForEach-Object { "$($_.KeyProtectorType)" }) -join ', ')
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
