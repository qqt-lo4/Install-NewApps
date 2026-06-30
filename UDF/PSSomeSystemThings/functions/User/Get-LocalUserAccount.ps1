function Get-LocalUserAccount {
    <#
    .SYNOPSIS
        Lists local user accounts of a Windows host, with remote execution support.

    .DESCRIPTION
        Enumerates local user accounts through the WinNT ADSI provider rather than
        Get-LocalUser, so the collection works on legacy hosts (Server 2012 R2 / PowerShell
        5.1 without the Microsoft.PowerShell.LocalAccounts module) as well as current ones.

        The remote side only gathers raw values (the userFlags bitmask, the account SID);
        the userFlags bits are decoded locally against the standard ADS_USER_FLAG layout
        (WinNT userFlags share it with AD userAccountControl). PSSomeActiveDirectoryThings'
        Convert-ADUACBit is intentionally NOT reused here: its [ADS_USER_FLAG_ENUM] parameter
        type is module-private (the enum is dot-sourced) and is not resolvable when the
        function is called from another module, so a direct bitwise test is used instead.

        Supports remote execution via -ComputerName / -Credential or an existing -Session.
        When neither is supplied the local machine is queried.

    .PARAMETER ComputerName
        Remote computer name(s) to query.

    .PARAMETER Credential
        Credentials for remote execution.

    .PARAMETER Session
        Existing PSSession(s) for remote execution. Mutually exclusive with -ComputerName.

    .OUTPUTS
        [PSCustomObject[]] One object per local user: Name, FullName, Description, Disabled,
        Locked, PasswordNeverExpires, PasswordCannotChange, LastLogin, SID.

    .EXAMPLE
        Get-LocalUserAccount

    .EXAMPLE
        Get-LocalUserAccount -ComputerName SRV01 -Credential $cred

    .NOTES
        Author  : Loic Ade
        Version : 1.0.0

        Version History:
        1.0.0 - Initial version (WinNT ADSI collection, userFlags decoded via Convert-ADUACBit)
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

        # Raw collection - runs on the target, no module dependency. Returns the userFlags
        # bitmask untouched so the host can decode it with the shared AD helper.
        $oScriptBlock = {
            $oComputer = [ADSI]"WinNT://$env:COMPUTERNAME,computer"
            $oComputer.psbase.Children |
                Where-Object { $_.SchemaClassName -eq 'User' } |
                ForEach-Object {
                    $u = $_
                    $sSid = $null
                    try { $sSid = (New-Object System.Security.Principal.SecurityIdentifier(($u.objectSID.Value), 0)).Value } catch {}
                    $oLastLogin = $null
                    try { $oLastLogin = $u.LastLogin.Value } catch {}
                    [PSCustomObject][ordered]@{
                        Name        = $u.Name.Value
                        FullName    = $u.FullName.Value
                        Description = $u.Description.Value
                        UserFlags   = [int]($u.UserFlags.Value)
                        LastLogin   = $oLastLogin
                        SID         = $sSid
                    }
                }
        }

        # ADS_USER_FLAG bit constants (same layout as AD userAccountControl).
        $ADS_UF_ACCOUNTDISABLE    = 0x0002
        $ADS_UF_LOCKOUT           = 0x0010
        $ADS_UF_PASSWD_CANT_CHANGE = 0x0040
        $ADS_UF_DONT_EXPIRE_PASSWD = 0x10000
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
            ForEach-Object {
                $flags = [int]$_.UserFlags
                [PSCustomObject][ordered]@{
                    Name                 = $_.Name
                    FullName             = $_.FullName
                    Description          = $_.Description
                    Disabled             = (($flags -band $ADS_UF_ACCOUNTDISABLE)    -ne 0)
                    Locked               = (($flags -band $ADS_UF_LOCKOUT)            -ne 0)
                    PasswordNeverExpires = (($flags -band $ADS_UF_DONT_EXPIRE_PASSWD) -ne 0)
                    PasswordCannotChange = (($flags -band $ADS_UF_PASSWD_CANT_CHANGE) -ne 0)
                    LastLogin            = $_.LastLogin
                    SID                  = $_.SID
                }
            } | Sort-Object Name
    }
}
