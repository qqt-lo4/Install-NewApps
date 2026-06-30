function Get-LocalPasswordPolicy {
    <#
    .SYNOPSIS
        Reports the local password & lockout policy of a host, with remote execution support.

    .DESCRIPTION
        Exports the local Account Policies with secedit and returns the password and lockout
        settings (minimum length, complexity, ages, history, lockout threshold/duration). The
        [System Access] keys secedit emits are language-neutral, so the result is independent
        of the host UI language. On a domain member the effective policy may be overridden by
        the domain policy; this reports the host's local database. Requires elevation.

        Supports remote execution via -ComputerName / -Credential or an existing -Session.
        When neither is supplied the local machine is queried.

    .PARAMETER ComputerName
        Remote computer name(s) to query.

    .PARAMETER Credential
        Credentials for remote execution.

    .PARAMETER Session
        Existing PSSession(s) for remote execution. Mutually exclusive with -ComputerName.

    .OUTPUTS
        [PSCustomObject] Password/lockout policy fields.

    .EXAMPLE
        Get-LocalPasswordPolicy -ComputerName SRV01 -Credential $cred

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
            $sTmp = Join-Path $env:TEMP ("secpol_{0}.inf" -f [System.Guid]::NewGuid().ToString('N'))
            try {
                $null = & secedit.exe /export /areas SECURITYPOLICY /cfg $sTmp 2>$null
                if (-not (Test-Path $sTmp)) { return }
                # secedit writes the INF as Unicode; read [System Access] key=value pairs.
                $hKv = @{}
                foreach ($sLine in (Get-Content -LiteralPath $sTmp -ErrorAction Stop)) {
                    if ($sLine -match '^\s*([A-Za-z]\w+)\s*=\s*(.+?)\s*$') { $hKv[$Matches[1]] = $Matches[2] }
                }
                # Lockout/age durations are in minutes (lockout) and seconds (password age):
                # secedit reports password ages in SECONDS divided by 86400 already? No - it
                # reports MaximumPasswordAge in DAYS and lockout durations in MINUTES. Keep raw
                # values with explicit units to avoid mis-scaling across OS versions.
                [PSCustomObject][ordered]@{
                    MinimumPasswordLength    = $hKv['MinimumPasswordLength']
                    PasswordComplexity       = if ($hKv.ContainsKey('PasswordComplexity')) { [int]$hKv['PasswordComplexity'] -eq 1 } else { $null }
                    MinimumPasswordAgeDays   = $hKv['MinimumPasswordAge']
                    MaximumPasswordAgeDays   = $hKv['MaximumPasswordAge']
                    PasswordHistorySize      = $hKv['PasswordHistorySize']
                    ClearTextPasswordStored  = if ($hKv.ContainsKey('ClearTextPassword')) { [int]$hKv['ClearTextPassword'] -eq 1 } else { $null }
                    LockoutThreshold         = $hKv['LockoutBadCount']
                    LockoutDurationMinutes   = $hKv['LockoutDuration']
                    ResetLockoutCounterMinutes = $hKv['ResetLockoutCount']
                }
            } finally {
                Remove-Item -LiteralPath $sTmp -Force -ErrorAction SilentlyContinue
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
