function Remove-SensitiveProperties {
    <#
    .SYNOPSIS
        Removes or masks properties that look like passwords or secrets from an object.

    .DESCRIPTION
        Recursively walks through an object (PSCustomObject, Hashtable, or ordered dictionary)
        and removes or masks properties whose names match sensitive patterns AND whose values
        look like actual secrets (not booleans, not small numbers).

    .PARAMETER InputObject
        The object to sanitize. Accepts pipeline input.

    .PARAMETER Patterns
        Array of wildcard patterns to match against property names.
        Default: @('*password*', '*secret*', '*passphrase*', '*pre-shared*', '*preshared*',
                   '*api-key*', '*apikey*', '*token*', '*credential*', '*private-key*', '*privatekey*')

    .PARAMETER Action
        What to do with matched properties:
        - "Mask"   : replace the value with MaskValue (default)
        - "Remove" : delete the property entirely

    .PARAMETER MaskValue
        The replacement text when Action is "Mask". Default: "[REDACTED]"

    .PARAMETER MaxDepth
        Maximum depth to recurse. Default: 10.

    .OUTPUTS
        The sanitized object (new copy, original is not modified).

    .EXAMPLE
        $gp.'remote-access' | Remove-SensitiveProperties
        # Masks l2tp-pre-shared-key and similar properties with "[REDACTED]"

    .EXAMPLE
        $obj | Remove-SensitiveProperties -Action Remove
        # Removes sensitive properties entirely

    .EXAMPLE
        $obj | Remove-SensitiveProperties -MaskValue "***"
        # Custom mask value

    .NOTES
        Author  : Loic Ade
        Version : 1.0.0

        1.0.0 (2026-04-08) - Initial version
    #>
    [CmdletBinding()]
    Param(
        [Parameter(Mandatory, ValueFromPipeline)]
        [object]$InputObject,

        [string[]]$Patterns = @(
            '*password*', '*passwd*', '*secret*', '*passphrase*',
            '*pre-shared*', '*preshared*',
            '*api-key*', '*apikey*', '*api_key*',
            '*token*', '*credential*',
            '*private-key*', '*privatekey*', '*private_key*'
        ),

        [ValidateSet('Mask', 'Remove')]
        [string]$Action = 'Mask',

        [string]$MaskValue = '[REDACTED]',

        [int]$MaxDepth = 10
    )

    Process {
        if ($null -eq $InputObject -or $MaxDepth -le 0) { return $InputObject }

        # Helper: check if a property name matches any sensitive pattern
        function Test-SensitiveName {
            Param([string]$Name)
            foreach ($sPattern in $Patterns) {
                if ($Name -like $sPattern) { return $true }
            }
            return $false
        }

        # Helper: check if a value looks like an actual secret (not a boolean/flag/small number)
        function Test-SensitiveValue {
            Param($Value)
            if ($null -eq $Value) { return $false }
            if ($Value -is [bool]) { return $false }
            if ($Value -is [int] -or $Value -is [long] -or $Value -is [double]) {
                return [Math]::Abs($Value) -ge 1000
            }
            if ($Value -is [string]) {
                $sLower = $Value.ToLower()
                if ($sLower -in @('true', 'false', 'yes', 'no', '0', '1')) { return $false }
                return $Value.Length -gt 0
            }
            return $false
        }

        # Recursive params to pass down
        $hRecurse = @{
            Patterns  = $Patterns
            Action    = $Action
            MaskValue = $MaskValue
            MaxDepth  = $MaxDepth - 1
        }

        if ($InputObject -is [System.Collections.IDictionary]) {
            $hResult = [ordered]@{}
            foreach ($sKey in @($InputObject.Keys)) {
                if ((Test-SensitiveName $sKey) -and (Test-SensitiveValue $InputObject[$sKey])) {
                    if ($Action -eq 'Mask') { $hResult[$sKey] = $MaskValue }
                } else {
                    $hResult[$sKey] = Remove-SensitiveProperties -InputObject $InputObject[$sKey] @hRecurse
                }
            }
            return $hResult

        } elseif ($InputObject -is [PSCustomObject]) {
            $hResult = [ordered]@{}
            foreach ($oProp in $InputObject.PSObject.Properties) {
                if ((Test-SensitiveName $oProp.Name) -and (Test-SensitiveValue $oProp.Value)) {
                    if ($Action -eq 'Mask') { $hResult[$oProp.Name] = $MaskValue }
                } else {
                    $hResult[$oProp.Name] = Remove-SensitiveProperties -InputObject $oProp.Value @hRecurse
                }
            }
            return [PSCustomObject]$hResult

        } elseif ($InputObject -is [array]) {
            return @($InputObject | ForEach-Object {
                Remove-SensitiveProperties -InputObject $_ @hRecurse
            })

        } else {
            return $InputObject
        }
    }
}
