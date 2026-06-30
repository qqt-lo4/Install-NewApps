function Get-SecurityCenterProduct {
    <#
    .SYNOPSIS
        Lists the security products registered with the Windows Security Center, with remote support.

    .DESCRIPTION
        Queries the WMI namespace root\SecurityCenter2 (AntiVirusProduct, AntiSpywareProduct,
        FirewallProduct) and returns, per registered product, its name and decoded state
        (enabled / up-to-date) together with the raw productState and the backing executable.
        This surfaces the third-party and built-in security tools as Windows itself sees them.

        IMPORTANT: root\SecurityCenter2 exists only on CLIENT editions of Windows
        (Windows 10/11). Windows Server SKUs do not expose the Security Center, so this
        returns nothing on servers (the namespace lookup fails and is swallowed).

        The enabled / up-to-date flags are decoded from the productState DWORD with the
        well-known (community) heuristic: byte 2 (0x10 bit) = enabled, byte 3 (0x10 bit) =
        out-of-date. The raw value is also returned so the decode can be audited.

        Supports remote execution via -ComputerName / -Credential or an existing -Session.
        When neither is supplied the local machine is queried.

    .PARAMETER ComputerName
        Remote computer name(s) to query.

    .PARAMETER Credential
        Credentials for remote execution.

    .PARAMETER Session
        Existing PSSession(s) for remote execution. Mutually exclusive with -ComputerName.

    .OUTPUTS
        [PSCustomObject[]] One object per product: Category, ProductName, Enabled, UpToDate,
        ProductState, Timestamp, ExePath.

    .EXAMPLE
        Get-SecurityCenterProduct -ComputerName PC01 -Credential $cred

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
            $sNamespace = 'root/SecurityCenter2'
            # Probe the namespace first: on Windows Server it does not exist and Get-CimInstance
            # throws - treat that as "no Security Center on this host".
            try { $null = Get-CimInstance -Namespace $sNamespace -ClassName AntiVirusProduct -ErrorAction Stop }
            catch { return }

            function ConvertFrom-ProductState {
                Param([int]$State)
                $sHex = '{0:X6}' -f $State
                $iEnabled = [Convert]::ToInt32($sHex.Substring(2, 2), 16)
                $iDefs    = [Convert]::ToInt32($sHex.Substring(4, 2), 16)
                [PSCustomObject]@{
                    Enabled  = (($iEnabled -band 0x10) -ne 0)
                    UpToDate = (($iDefs -band 0x10) -eq 0)
                    Hex      = "0x$sHex"
                }
            }

            foreach ($oPair in @(
                @('Antivirus',   'AntiVirusProduct'),
                @('AntiSpyware', 'AntiSpywareProduct'),
                @('Firewall',    'FirewallProduct')
            )) {
                $sCategory = $oPair[0]
                $sClass    = $oPair[1]
                try {
                    Get-CimInstance -Namespace $sNamespace -ClassName $sClass -ErrorAction Stop | ForEach-Object {
                        $oDecoded = ConvertFrom-ProductState ([int]$_.productState)
                        [PSCustomObject][ordered]@{
                            Category     = $sCategory
                            ProductName  = $_.displayName
                            Enabled      = $oDecoded.Enabled
                            UpToDate     = if ($sClass -eq 'FirewallProduct') { $null } else { $oDecoded.UpToDate }
                            ProductState = $oDecoded.Hex
                            Timestamp    = $_.timestamp
                            ExePath      = $_.pathToSignedProductExe
                        }
                    }
                } catch {}
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
