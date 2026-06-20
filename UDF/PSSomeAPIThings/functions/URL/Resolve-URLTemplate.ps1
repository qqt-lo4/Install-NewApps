function Resolve-URLTemplate {
    <#
    .SYNOPSIS
        Expands {placeholders} in a URL template using values from a hashtable
        and returns the remaining arguments for use as a query string.

    .DESCRIPTION
        Many REST APIs embed parameters directly in the URL path (e.g.
        "/users/{userId}/groups/{groupId}"). This helper walks the URL template,
        finds each "{name}" placeholder, replaces it with the matching value
        from the provided arguments hashtable, and returns both the resolved URL
        and a copy of the arguments hashtable with the consumed keys removed
        (those remaining arguments are typically passed as a query string or body).

        Placeholder names must match exactly one argument key. A missing key
        causes a terminating error.

    .PARAMETER Endpoint
        URL template containing zero or more "{name}" placeholders.
        Placeholder names must match [a-zA-Z_]+.

    .PARAMETER Arguments
        Hashtable whose keys are used to replace placeholders in the URL.
        Any key matching a placeholder is consumed (removed from the returned
        Arguments hashtable).

    .OUTPUTS
        [PSCustomObject] with properties:
        - Endpoint  : URL with placeholders replaced
        - Arguments : copy of the input hashtable without the consumed keys

    .EXAMPLE
        $r = Resolve-URLTemplate -Endpoint "/users/{userId}/groups/{groupId}" `
            -Arguments @{ userId = 42; groupId = 7; expand = 'members' }
        $r.Endpoint  # "/users/42/groups/7"
        $r.Arguments # @{ expand = 'members' }

    .EXAMPLE
        # No placeholders: endpoint is returned unchanged, arguments are passed through.
        $r = Resolve-URLTemplate -Endpoint "/users" -Arguments @{ limit = 50 }
        $r.Endpoint  # "/users"
        $r.Arguments # @{ limit = 50 }

    .NOTES
        Author  : Loic Ade
        Version : 1.0.0
        Dependencies: Copy-Hashtable (PSSomeDataThings)

        1.0.0 (2026-04-20) - Initial version, extracted from Invoke-BeyondTrustPAMAPI
    #>
    [CmdletBinding()]
    Param(
        [Parameter(Mandatory, Position = 0)]
        [string]$Endpoint,

        [Parameter(Mandatory, Position = 1)]
        [AllowNull()]
        [hashtable]$Arguments
    )

    $ss = Select-String -InputObject $Endpoint -Pattern "{[a-zA-Z_]+}" -AllMatches
    if (-not $ss) {
        return [PSCustomObject]@{
            Endpoint  = $Endpoint
            Arguments = $Arguments
        }
    }

    $sEndpoint = $Endpoint
    $aURLKeys = @()
    foreach ($sValue in $ss.Matches.Groups.Value) {
        # Strip the enclosing braces to get the bare key name
        $sArgumentsKey = $sValue[1..($sValue.Length - 2)] -join ""
        $aURLKeys += $sArgumentsKey
        if ($Arguments -and ($sArgumentsKey -in $Arguments.Keys)) {
            $sEndpoint = $sEndpoint -replace $sValue, $Arguments[$sArgumentsKey]
        } else {
            throw "Argument '$sArgumentsKey' not found in arguments hashtable"
        }
    }

    $hArguments = Copy-Hashtable -InputObject $Arguments -Not -Properties $aURLKeys

    return [PSCustomObject]@{
        Endpoint  = $sEndpoint
        Arguments = $hArguments
    }
}
