function Get-AuditPolicySetting {
    <#
    .SYNOPSIS
        Reports the effective advanced audit policy of a Windows host, with remote support.

    .DESCRIPTION
        Runs auditpol.exe and returns the effective setting (No Auditing / Success /
        Failure / Success and Failure) of every audit subcategory. This is the evidence
        that security logging is configured, complementing the event-log content that shows
        it is actually working.

        Supports remote execution via -ComputerName / -Credential or an existing -Session.
        When neither is supplied the local machine is queried. Note: auditpol reports the
        policy of the host it runs on, so remote execution reflects the target's policy.

    .PARAMETER ComputerName
        Remote computer name(s) to query.

    .PARAMETER Credential
        Credentials for remote execution.

    .PARAMETER Session
        Existing PSSession(s) for remote execution. Mutually exclusive with -ComputerName.

    .OUTPUTS
        [PSCustomObject[]] One object per subcategory: Subcategory, Setting.

    .EXAMPLE
        Get-AuditPolicySetting -ComputerName SRV01 -Credential $cred

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
            # auditpol /r emits CSV whose HEADER NAMES are localized (Subcategory vs
            # Sous-categorie, ...) but whose COLUMN ORDER is fixed:
            #   0 Machine Name, 1 Policy Target, 2 Subcategory, 3 Subcategory GUID,
            #   4 Inclusion Setting, 5 Exclusion Setting
            # Parsing by position keeps this language-agnostic. Requires elevation;
            # an unprivileged caller gets no CSV and the function returns nothing.
            $aCsv = @(& auditpol.exe /get /category:* /r 2>$null | ConvertFrom-Csv)
            if ($aCsv.Count -eq 0) { return }
            $aCols = @($aCsv[0].PSObject.Properties.Name)
            if ($aCols.Count -lt 5) { return }
            $sSubcatCol = $aCols[2]
            $sGuidCol   = $aCols[3]
            $sSettingCol = $aCols[4]
            $aCsv |
                Where-Object { $_.$sGuidCol -match '^\{?[0-9A-Fa-f-]{36}\}?$' } |
                ForEach-Object {
                    [PSCustomObject][ordered]@{
                        Subcategory = $_.$sSubcatCol
                        Setting     = $_.$sSettingCol
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
