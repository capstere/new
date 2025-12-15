#requires -Version 5.1
Set-StrictMode -Version Latest

<#
.SYNOPSIS
    Returns the loaded RuleBank root object.
.DESCRIPTION
    Central guard that checks if RuleBank.ps1 has been dot-sourced and exposes the
    global `$RuleBank` hashtable in a safe manner.
.OUTPUTS
    Hashtable with RuleBank data or $null when unavailable.
.EXAMPLE
    $rb = Get-RuleBankRoot
    if ($rb) { $rb.Version }
#>
function Get-RuleBankRoot {
    [CmdletBinding()]
    param()

    $rbVar = Get-Variable -Name RuleBank -Scope Global -ErrorAction SilentlyContinue
    if (-not $rbVar) {
        Write-Error "RuleBank.ps1 is not loaded. Dot-source Modules/RuleBank.ps1 before using the rule engine." -ErrorAction Continue
        return $null
    }

    $rb = $rbVar.Value
    if (-not $rb -or ($rb -isnot [hashtable])) {
        Write-Error "RuleBank was found but is not a hashtable (expected a data-only root object)." -ErrorAction Continue
        return $null
    }

    return $rb
}

function Normalize-RuleBankFieldKey {
    param([string]$Name)
    if ($null -eq $Name) { return '' }
    return ([regex]::Replace(($Name + ''), '[^a-z0-9]', '')).ToLowerInvariant()
}

function Get-RuleBankRowValue {
    param(
        $Row,
        [string]$TargetField
    )
    $normTarget = Normalize-RuleBankFieldKey $TargetField
    if (-not $Row) { return $null }

    if ($Row -is [System.Data.DataRow]) {
        foreach ($col in $Row.Table.Columns) {
            if ((Normalize-RuleBankFieldKey $col.ColumnName) -eq $normTarget) {
                return $Row[$col.ColumnName]
            }
        }
    }
    elseif ($Row -is [System.Collections.IDictionary]) {
        foreach ($key in $Row.Keys) {
            if ((Normalize-RuleBankFieldKey $key) -eq $normTarget) {
                return $Row[$key]
            }
        }
    }
    else {
        foreach ($prop in $Row.PSObject.Properties) {
            if ((Normalize-RuleBankFieldKey $prop.Name) -eq $normTarget) {
                return $prop.Value
            }
        }
    }

    return $null
}

function Normalize-RuleBankValue {
    param($Value)
    if ($null -eq $Value) { return $null }
    $text = ($Value + '').Trim()
    if ([string]::IsNullOrWhiteSpace($text)) { return $null }
    return $text
}

<#
.SYNOPSIS
    Selects an assay profile from RuleBank based on assay identity.
.DESCRIPTION
    Performs a case-insensitive lookup against RuleBank.AssayProfiles. Falls back to
    the 'default' profile when no match is found and records the reason in the
    IdentityFindings list. Also evaluates IdentityHints (AssayVersionExpected) when present.
.PARAMETER AssayName
    Assay value from the dataset (CSV).
.PARAMETER AssayVersion
    Assay version extracted from the dataset.
.PARAMETER DatasetInfo
    Optional hashtable with context (distinct assays/versions, row count).
.OUTPUTS
    PSCustomObject with ProfileKey, DisplayName, Mode, Profile, IdentityFindings, VersionMatch, AssayVersionExpected, Reason.
.EXAMPLE
    Get-AssayRuleProfile -AssayName 'Xpert MRSA NxG' -AssayVersion '3'
#>
function Get-AssayRuleProfile {
    [CmdletBinding()]
    param(
        [string]$AssayName,
        [string]$AssayVersion,
        [hashtable]$DatasetInfo
    )

    $identityFindings = New-Object 'System.Collections.Generic.List[string]'
    $rb = Get-RuleBankRoot
    if (-not $rb) { return $null }

    $profiles = $rb['AssayProfiles']
    if (-not $profiles) {
        $identityFindings.Add('RuleBank.AssayProfiles saknas.')
        return [pscustomobject]@{
            ProfileKey            = 'default'
            DisplayName           = 'default'
            Mode                  = 'GlobalOnly'
            Profile               = $null
            IdentityFindings      = $identityFindings
            VersionMatch          = $null
            AssayVersionExpected  = $null
            Reason                = 'AssayProfiles missing'
            AssayKey              = 'default'
        }
    }

    $profileKey = $null
    $normalizedInput = ($AssayName + '').Trim().ToLowerInvariant()
    if ($normalizedInput) {
        foreach ($key in $profiles.Keys) {
            if ((($key + '').Trim().ToLowerInvariant()) -eq $normalizedInput) {
                $profileKey = $key
                break
            }
        }
    }

    $reason = $null
    if (-not $profileKey) {
        $profileKey = 'default'
        $reason = 'Unknown assay'
    }

    $profile = $profiles[$profileKey]
    if (-not $profile) {
        $profile = [ordered]@{ Mode = 'GlobalOnly' }
        $reason = "Profile '$profileKey' saknas i RuleBank."
    }

    if ($reason) { $identityFindings.Add($reason) }

    $assayKeyOut = if ($profile.PSObject.Properties['AssayKey']) { $profile.AssayKey } else { $profileKey }
    $displayName = if ($profile.PSObject.Properties['DisplayName'] -and $profile.DisplayName) { $profile.DisplayName } else { $profileKey }
    $mode = if ($profile.PSObject.Properties['Mode'] -and $profile.Mode) { $profile.Mode } else { 'GlobalOnly' }

    $versionExpected = $null
    $versionMatch = $null
    if ($profile.PSObject.Properties['IdentityHints'] -and $profile.IdentityHints -and $profile.IdentityHints.AssayVersionExpected) {
        $versionExpected = $profile.IdentityHints.AssayVersionExpected
        if ($AssayVersion -ne $null -and $AssayVersion -ne '') {
            $versionMatch = ([string]$AssayVersion -eq [string]$versionExpected)
            if (-not $versionMatch) {
                $identityFindings.Add(("AssayVersion expected {0} men hittade {1}." -f $versionExpected, $AssayVersion))
            }
        } else {
            $identityFindings.Add(("AssayVersion saknas; förväntade {0}." -f $versionExpected))
        }
    }

    return [pscustomobject]@{
        AssayKey             = $assayKeyOut
        ProfileKey           = $profileKey
        DisplayName          = $displayName
        Mode                 = $mode
        Profile              = $profile
        IdentityFindings     = $identityFindings
        VersionMatch         = $versionMatch
        AssayVersionExpected = $versionExpected
        Reason               = $reason
    }
}

<#
.SYNOPSIS
    Builds a lightweight assay context for the current dataset.
.DESCRIPTION
    Validates identity fields (Assay, Assay Version, Reagent Lot ID) for consistency,
    selects an assay profile from RuleBank, and returns a structured context without
    throwing for normal QC deviations. Findings are collected in IdentityFindings.
.PARAMETER Rows
    Dataset rows (from CSV) as objects or DataRow instances.
.PARAMETER FallbackAssayName
    Optional assay name used when the dataset is empty or missing the field.
.PARAMETER FallbackAssayVersion
    Optional version used when no version is present in the dataset.
.OUTPUTS
    PSCustomObject with assay identity, selected profile, mode, findings, and dataset info.
.EXAMPLE
    $ctx = New-AssayRuleContext -Rows $csvRows -FallbackAssayName $assay
#>
function New-AssayRuleContext {
    [CmdletBinding()]
    param(
        [object[]]$Rows,
        [string]$FallbackAssayName,
        [string]$FallbackAssayVersion
    )

    $identityFindings = New-Object 'System.Collections.Generic.List[string]'
    $rb = Get-RuleBankRoot
    $identityFields = @('Assay','Assay Version','Reagent Lot ID')
    if ($rb -and $rb.Global -and $rb.Global.Identity -and $rb.Global.Identity.Fields) {
        $identityFields = @($rb.Global.Identity.Fields)
    }

    $rowsArr = @()
    if ($Rows) { $rowsArr = @($Rows | Where-Object { $_ }) }
    $rowCount = $rowsArr.Count
    if ($rowCount -eq 0) {
        $identityFindings.Add('Inga CSV-rader tillgängliga för RuleBank-analys.')
    }

    $distinctMap = @{}
    foreach ($field in $identityFields) {
        $distinctMap[$field] = New-Object 'System.Collections.Generic.HashSet[string]' ([StringComparer]::OrdinalIgnoreCase)
    }

    foreach ($row in $rowsArr) {
        foreach ($field in $identityFields) {
            $val = Normalize-RuleBankValue (Get-RuleBankRowValue -Row $row -TargetField $field)
            if ($val) {
                [void]$distinctMap[$field].Add($val)
            }
        }
    }

    $datasetInfo = @{
        RowCount       = $rowCount
        DistinctValues = @{}
    }
    foreach ($field in $identityFields) {
        $datasetInfo.DistinctValues[$field] = @($distinctMap[$field])
    }

    foreach ($field in $identityFields) {
        $values = @($distinctMap[$field])
        if ($values.Count -gt 1) {
            $identityFindings.Add(("Multiple values for '{0}': {1}" -f $field, ($values -join ', ')))
        } elseif ($values.Count -eq 0) {
            $identityFindings.Add(("Missing value for '{0}' i datasetet." -f $field))
        }
    }

    $assayField = ($identityFields | Where-Object { (Normalize-RuleBankFieldKey $_) -eq 'assay' } | Select-Object -First 1)
    $assayVersionField = ($identityFields | Where-Object { (Normalize-RuleBankFieldKey $_) -eq 'assayversion' } | Select-Object -First 1)
    $lotField = ($identityFields | Where-Object { (Normalize-RuleBankFieldKey $_) -eq 'reagentlotid' } | Select-Object -First 1)

    $assayName = $null
    if ($assayField -and $distinctMap[$assayField] -and $distinctMap[$assayField].Count -gt 0) {
        $assayName = @($distinctMap[$assayField])[0]
    } elseif ($FallbackAssayName) {
        $assayName = $FallbackAssayName
    }

    $assayVersion = $null
    if ($assayVersionField -and $distinctMap[$assayVersionField] -and $distinctMap[$assayVersionField].Count -gt 0) {
        $assayVersion = @($distinctMap[$assayVersionField])[0]
    } elseif ($FallbackAssayVersion) {
        $assayVersion = $FallbackAssayVersion
    }

    $reagentLotId = $null
    if ($lotField -and $distinctMap[$lotField] -and $distinctMap[$lotField].Count -gt 0) {
        $reagentLotId = @($distinctMap[$lotField])[0]
    }

    $profileContext = Get-AssayRuleProfile -AssayName $assayName -AssayVersion $assayVersion -DatasetInfo $datasetInfo
    if (-not $profileContext) {
        $profileContext = [pscustomobject]@{
            AssayKey             = 'default'
            ProfileKey           = 'default'
            DisplayName          = 'default'
            Mode                 = 'GlobalOnly'
            Profile              = $null
            IdentityFindings     = New-Object 'System.Collections.Generic.List[string]'
            VersionMatch         = $null
            AssayVersionExpected = $null
            Reason               = 'RuleBank missing'
        }
        $profileContext.IdentityFindings.Add('RuleBank saknas – defaultprofil används.')
    }

    if ($profileContext.IdentityFindings) {
        foreach ($msg in $profileContext.IdentityFindings) {
            if ($msg) { $identityFindings.Add($msg) }
        }
    }

    return [pscustomobject]@{
        AssayName            = $assayName
        AssayVersion         = $assayVersion
        ReagentLotId         = $reagentLotId
        ProfileKey           = $profileContext.ProfileKey
        Mode                 = $profileContext.Mode
        Profile              = $profileContext.Profile
        IdentityFindings     = $identityFindings
        VersionMatch         = $profileContext.VersionMatch
        AssayVersionExpected = $profileContext.AssayVersionExpected
        DatasetInfo          = $datasetInfo
    }
}
