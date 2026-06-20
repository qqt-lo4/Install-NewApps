function Invoke-AsCredential {
    <#
    .SYNOPSIS
        Runs a script block impersonated as the given credential (in-process).

    .DESCRIPTION
        Logs on with the supplied credential (LogonUser), impersonates that token on
        the current thread, runs the script block, then reverts. The script block (and
        anything it calls, including native callbacks) executes under the credential's
        security context. Returns whatever the script block returns.

        The default logon type is LOGON32_LOGON_NEW_CREDENTIALS (9) with
        LOGON32_PROVIDER_DEFAULT (0) - equivalent to "runas /netonly": the local
        identity is unchanged but network access (e.g. an LDAP/DPAPI-NG call to a DC)
        uses the supplied credential. This avoids the WinRM/CredSSP double-hop and does
        not require the "log on locally" right. Pass -LogonType to change it (e.g. 2 =
        INTERACTIVE, 8 = NETWORK_CLEARTEXT) when a full local identity is needed.

        Unlike Invoke-AsSystem (scheduled task), this is fully in-process: no temp
        files, no child process. It only changes the calling thread's token for the
        duration of the script block.

    .PARAMETER Credential
        The credential to impersonate.

    .PARAMETER ScriptBlock
        The script block to run under the impersonated context.

    .PARAMETER ArgumentList
        Arguments passed to the script block. Wrap an array/byte[] argument so it is
        passed as a single argument, e.g. -ArgumentList (,$bytes).

    .PARAMETER LogonType
        Win32 logon type. Default 9 (LOGON32_LOGON_NEW_CREDENTIALS).

    .PARAMETER LogonProvider
        Win32 logon provider. Default 0 (LOGON32_PROVIDER_DEFAULT).

    .OUTPUTS
        Whatever the script block returns.

    .EXAMPLE
        Invoke-AsCredential -Credential $cred -ScriptBlock { whoami }

    .EXAMPLE
        $plain = Invoke-AsCredential -Credential $cred -ScriptBlock {
            param([byte[]]$Blob) Unprotect-Thing $Blob
        } -ArgumentList (,$encryptedBytes)

    .NOTES
        Author  : Loïc Ade
        Version : 1.0.0

        CHANGELOG:

        Version 1.0.0 - 2026-06-15 - Loïc Ade
            - Initial release. In-process impersonation runner (LogonUser +
              ImpersonateLoggedOnUser + RevertToSelf), default logon type
              NEW_CREDENTIALS (9). Factored out so any caller (e.g. LAPS
              decryption) can run code under a credential without a child process.
    #>
    [CmdletBinding()]
    Param(
        [Parameter(Mandatory)]
        [pscredential]$Credential,
        [Parameter(Mandatory)]
        [scriptblock]$ScriptBlock,
        [object[]]$ArgumentList = @(),
        [int]$LogonType = 9,
        [int]$LogonProvider = 0
    )

    if (-not ([System.Management.Automation.PSTypeName]'PSSome_Impersonation').Type) {
        Add-Type -TypeDefinition @"
        using System;
        using System.Runtime.InteropServices;

        public static class PSSome_Impersonation
        {
            [DllImport("advapi32.dll", SetLastError = true, CharSet = CharSet.Unicode)]
            public static extern bool LogonUser(string lpszUsername, string lpszDomain, string lpszPassword, int dwLogonType, int dwLogonProvider, out IntPtr phToken);

            [DllImport("advapi32.dll", SetLastError = true)]
            public static extern bool ImpersonateLoggedOnUser(IntPtr hToken);

            [DllImport("advapi32.dll", SetLastError = true)]
            public static extern bool RevertToSelf();

            [DllImport("kernel32.dll", SetLastError = true)]
            public static extern bool CloseHandle(IntPtr handle);
        }
"@
    }

    $oNetCred = $Credential.GetNetworkCredential()
    $sUser    = $oNetCred.UserName
    $sDomain  = if ($oNetCred.Domain) { $oNetCred.Domain } else { "." }

    $tokenHandle = [IntPtr]::Zero
    if (-not [PSSome_Impersonation]::LogonUser($sUser, $sDomain, $oNetCred.Password, $LogonType, $LogonProvider, [ref]$tokenHandle)) {
        throw (New-Object System.ComponentModel.Win32Exception([System.Runtime.InteropServices.Marshal]::GetLastWin32Error()))
    }

    $bImpersonated = $false
    try {
        $bImpersonated = [PSSome_Impersonation]::ImpersonateLoggedOnUser($tokenHandle)
        if (-not $bImpersonated) {
            throw (New-Object System.ComponentModel.Win32Exception([System.Runtime.InteropServices.Marshal]::GetLastWin32Error()))
        }
        return (& $ScriptBlock @ArgumentList)
    } finally {
        if ($bImpersonated) { [void][PSSome_Impersonation]::RevertToSelf() }
        if ($tokenHandle -ne [IntPtr]::Zero) { [void][PSSome_Impersonation]::CloseHandle($tokenHandle) }
    }
}
