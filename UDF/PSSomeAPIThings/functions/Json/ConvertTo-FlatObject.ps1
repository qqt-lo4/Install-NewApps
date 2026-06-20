function ConvertTo-FlatObject {
    <#
    .SYNOPSIS
        Flattens a nested object into a single-level hashtable with dot-separated property names.

    .DESCRIPTION
        Recursively walks through an object's properties and sub-objects, producing a flat
        ordered hashtable where nested properties are represented with dot-separated keys.

        Arrays are indexed with [n] notation. Supports PSCustomObject, Hashtable, and
        OrderedDictionary inputs.

    .PARAMETER InputObject
        The object to flatten. Accepts pipeline input.

    .PARAMETER Prefix
        Internal parameter for recursion. Do not use directly.

    .PARAMETER MaxDepth
        Maximum depth to recurse. Default: 10.

    .PARAMETER Separator
        Separator between property levels. Default: "."

    .OUTPUTS
        [ordered] hashtable with dot-separated keys and scalar values.

    .EXAMPLE
        $gp.'stateful-inspection' | ConvertTo-FlatObject
        # Returns: @{ "tcp-start-timeout" = 40; "tcp-session-timeout" = 3600; ... }

    .EXAMPLE
        $gp.'remote-access' | ConvertTo-FlatObject
        # Returns: @{ "encrypt-dns-traffic" = True; "vpn-authentication-and-encryption.encryption-method" = "ike_v1_only"; ... }

    .EXAMPLE
        $obj | ConvertTo-FlatObject -Separator "/"
        # Uses "/" instead of "." as separator

    .NOTES
        Author  : Loic Ade
        Version : 1.0.0

        1.0.0 (2026-04-08) - Initial version
    #>
    [CmdletBinding()]
    Param(
        [Parameter(Mandatory, ValueFromPipeline)]
        [object]$InputObject,

        [string]$Prefix = "",

        [int]$MaxDepth = 10,

        [string]$Separator = "."
    )

    Process {
        $hResult = [ordered]@{}

        if ($MaxDepth -le 0) {
            $sKey = if ($Prefix) { $Prefix } else { '_value' }
            $hResult[$sKey] = "$InputObject"
            return $hResult
        }

        # Get properties based on object type
        $aProps = @()
        if ($InputObject -is [System.Collections.IDictionary]) {
            $aProps = @($InputObject.Keys | ForEach-Object { @{ Name = $_; Value = $InputObject[$_] } })
        } elseif ($InputObject -is [PSCustomObject]) {
            $aProps = @($InputObject.PSObject.Properties | ForEach-Object { @{ Name = $_.Name; Value = $_.Value } })
        } else {
            # Scalar value
            $sKey = if ($Prefix) { $Prefix } else { '_value' }
            $hResult[$sKey] = $InputObject
            return $hResult
        }

        foreach ($oProp in $aProps) {
            $sName = $oProp.Name
            $oVal = $oProp.Value
            $sFullKey = if ($Prefix) { "$Prefix$Separator$sName" } else { $sName }

            if ($null -eq $oVal) {
                $hResult[$sFullKey] = $null
            } elseif ($oVal -is [string] -or $oVal -is [bool] -or $oVal -is [int] -or $oVal -is [long] -or $oVal -is [double] -or $oVal -is [datetime]) {
                # Scalar value
                $hResult[$sFullKey] = $oVal
            } elseif ($oVal -is [array] -or ($oVal -is [System.Collections.IEnumerable] -and $oVal -isnot [string] -and $oVal -isnot [System.Collections.IDictionary])) {
                # Array: check if all elements are scalars
                $aItems = @($oVal)
                $bAllScalar = $true
                foreach ($oItem in $aItems) {
                    if ($null -ne $oItem -and $oItem -isnot [string] -and $oItem -isnot [bool] -and $oItem -isnot [int] -and $oItem -isnot [long] -and $oItem -isnot [double]) {
                        $bAllScalar = $false
                        break
                    }
                }

                if ($bAllScalar) {
                    # Array of scalars: join into a single value
                    $hResult[$sFullKey] = ($aItems | ForEach-Object { "$_" }) -join ', '
                } else {
                    # Mixed array: index each element
                    $iIdx = 0
                    foreach ($oItem in $aItems) {
                        $sArrayKey = "$sFullKey[$iIdx]"
                        if ($oItem -is [string] -or $oItem -is [bool] -or $oItem -is [int] -or $oItem -is [long] -or $oItem -is [double]) {
                            $hResult[$sArrayKey] = $oItem
                        } elseif ($oItem -is [PSCustomObject] -or $oItem -is [System.Collections.IDictionary]) {
                            $hSub = ConvertTo-FlatObject -InputObject $oItem -Prefix $sArrayKey -MaxDepth ($MaxDepth - 1) -Separator $Separator
                            foreach ($sSubKey in $hSub.Keys) { $hResult[$sSubKey] = $hSub[$sSubKey] }
                        } else {
                            $hResult[$sArrayKey] = "$oItem"
                        }
                        $iIdx++
                    }
                }
            } elseif ($oVal -is [PSCustomObject] -or $oVal -is [System.Collections.IDictionary]) {
                # Nested object: recurse
                $hSub = ConvertTo-FlatObject -InputObject $oVal -Prefix $sFullKey -MaxDepth ($MaxDepth - 1) -Separator $Separator
                foreach ($sSubKey in $hSub.Keys) { $hResult[$sSubKey] = $hSub[$sSubKey] }
            } else {
                $hResult[$sFullKey] = "$oVal"
            }
        }

        return $hResult
    }
}
