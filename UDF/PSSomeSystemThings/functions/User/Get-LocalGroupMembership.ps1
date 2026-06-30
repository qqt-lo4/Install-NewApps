function Get-LocalGroupMembership {
    <#
    .SYNOPSIS
        Lists local groups and their members of a Windows host, with remote execution support.

    .DESCRIPTION
        Enumerates local groups and the members of each group through the WinNT ADSI
        provider rather than Get-LocalGroup / Get-LocalGroupMember, so the collection works
        on legacy hosts (Server 2012 R2 / PowerShell 5.1 without the
        Microsoft.PowerShell.LocalAccounts module) as well as current ones.

        The remote side gathers each member's account name (from its WinNT ADsPath, which
        the target's LSA resolves for both local and domain members) and its SID. Any member
        whose ADsPath did not resolve to a name (orphaned / cross-forest SID) is resolved on
        the host with Resolve-ADSidName from PSSomeActiveDirectoryThings when that module is
        loaded; otherwise the raw SID is kept. Members are returned as DOMAIN\Name strings,
        directly answering "who has administrator / user access" to the server.

        Supports remote execution via -ComputerName / -Credential or an existing -Session.
        When neither is supplied the local machine is queried.

    .PARAMETER ComputerName
        Remote computer name(s) to query.

    .PARAMETER Credential
        Credentials for remote execution.

    .PARAMETER Session
        Existing PSSession(s) for remote execution. Mutually exclusive with -ComputerName.

    .OUTPUTS
        [PSCustomObject[]] One object per local group: Name, Description, MemberCount, Members.

    .EXAMPLE
        Get-LocalGroupMembership

    .EXAMPLE
        Get-LocalGroupMembership -ComputerName SRV01 -Credential $cred |
            Where-Object Name -eq 'Administrators'

    .NOTES
        Author  : Loic Ade
        Version : 1.0.0

        Version History:
        1.0.0 - Initial version (WinNT ADSI collection, SID fallback via Resolve-ADSidName)
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

        # Raw collection - runs on the target, no module dependency. Per group it returns
        # the member list as {Account, SID, Class} so the host can resolve any leftover SID.
        $oScriptBlock = {
            function Get-AdsiMemberInfo {
                Param($Member)
                $sPath = try { $Member.GetType().InvokeMember('ADsPath', 'GetProperty', $null, $Member, $null) } catch { '' }
                $sClass = try { $Member.GetType().InvokeMember('Class', 'GetProperty', $null, $Member, $null) } catch { '' }
                $sAccount = (($sPath -replace '^WinNT://', '') -replace '/', '\')
                $sSid = $null
                try {
                    $oBytes = $Member.GetType().InvokeMember('objectSID', 'GetProperty', $null, $Member, $null)
                    if ($oBytes) { $sSid = (New-Object System.Security.Principal.SecurityIdentifier($oBytes, 0)).Value }
                } catch {}
                [PSCustomObject]@{ Account = $sAccount; SID = $sSid; Class = $sClass }
            }

            $oComputer = [ADSI]"WinNT://$env:COMPUTERNAME,computer"
            $oComputer.psbase.Children |
                Where-Object { $_.SchemaClassName -eq 'Group' } |
                ForEach-Object {
                    $g = $_
                    $aMembers = @()
                    try { $aMembers = @($g.psbase.Invoke('Members') | ForEach-Object { Get-AdsiMemberInfo $_ }) } catch {}
                    [PSCustomObject][ordered]@{
                        Name        = $g.Name.Value
                        Description = $g.Description.Value
                        Members     = $aMembers
                    }
                }
        }

        $bHasSidResolver = [bool](Get-Command Resolve-ADSidName -ErrorAction SilentlyContinue)
        $script:bHasSidResolver = $bHasSidResolver
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
                $aNames = @(foreach ($m in $_.Members) {
                    $sName = $m.Account
                    # ADsPath could not resolve to a name (orphaned / cross-forest SID):
                    # fall back to the shared resolver against the host's LSA.
                    if ((-not $sName -or $sName -match '^S-1-') -and $m.SID -and $script:bHasSidResolver) {
                        $sResolved = Resolve-ADSidName -Sid $m.SID
                        if ($sResolved) { $sName = $sResolved }
                    }
                    if (-not $sName) { $sName = $m.SID }
                    $sName
                })
                [PSCustomObject][ordered]@{
                    Name        = $_.Name
                    Description = $_.Description
                    MemberCount = $aNames.Count
                    Members     = ($aNames | Sort-Object) -join '; '
                }
            } | Sort-Object Name
    }
}
