#requires -Version 5.1
<#
  RuleEngine.ps1  (PURE LOGIC)
  PowerShell 5.1 compatible.

  Responsibilities:
  - Build a safe context from a parsed CSV bundle (no EPPlus dependencies).
  - Evaluate rules (pure logic).
  - NEVER modify RuleBank (data-only).
#>

Set-StrictMode -Off

# -----------------------------
# Map helpers (Hashtable + OrderedDictionary safe)
# -----------------------------
function Test-MapHasKey {
    param([object]$Map, [object]$Key)
    if ($null -eq $Map -or $null -eq $Key) { return $false }

    try {
        if ($Map -is [hashtable]) { return $Map.ContainsKey($Key) }
        if ($Map -is [System.Collections.Specialized.OrderedDictionary]) { return $Map.Contains($Key) }
        if ($Map -is [System.Collections.IDictionary]) { return $Map.Contains($Key) }

        if ($Map.PSObject -and $Map.PSObject.Methods['ContainsKey']) { return [bool]$Map.ContainsKey($Key) }
        if ($Map.PSObject -and $Map.PSObject.Methods['Contains'])    { return [bool]$Map.Contains($Key) }
    } catch {}
    return $false
}

function Get-MapValue {
    param([object]$Map, [object]$Key, [object]$Default = $null)
    try {
        if (Test-MapHasKey -Map $Map -Key $Key) { return $Map[$Key] }
    } catch {}
    return $Default
}

function Add-MapCount {
    param([object]$Map, [object]$Key, [int]$Delta = 1)
    if ($null -eq $Map -or $null -eq $Key) { return }
    try {
        if (-not (Test-MapHasKey $Map $Key)) { $Map[$Key] = 0 }
        $Map[$Key] = [int]$Map[$Key] + [int]$Delta
    } catch {}
}

# -----------------------------
# Severity
# -----------------------------
$script:SeverityOrder = @{
    'Error' = 3
    'Warn'  = 2
    'Info'  = 1
}

function Get-SeverityRank {
    param([string]$Severity)
    $k = ($Severity + '').Trim()
    if ($script:SeverityOrder.ContainsKey($k)) { return $script:SeverityOrder[$k] }
    return 0
}

# -----------------------------
# Normalization
# -----------------------------
if (-not (Get-Command Normalize-HeaderName -ErrorAction SilentlyContinue)) {
    function Normalize-HeaderName {
        param([string]$s)
        if ([string]::IsNullOrWhiteSpace($s)) { return '' }
        $t = $s.Trim().ToLowerInvariant()
        $t = $t.Replace('_',' ')
        $t = ($t -replace '[\(\)]',' ')
        $t = ($t -replace '\s+',' ').Trim()
        return $t
    }
}

function Normalize-AssayName {
    param([string]$Name)
    if ([string]::IsNullOrWhiteSpace($Name)) { return '' }
    $x = $Name.Trim().ToLowerInvariant()
    $x = $x -replace '[_]+',' '
    $x = $x -replace '[^a-z0-9]+',' '
    $x = ($x -replace '\s+',' ').Trim()
    return $x
}

function Get-HeaderIndexValue {
    param(
        [object]$HeaderIndex,
        [string[]]$Keys
    )
    if (-not $HeaderIndex) { return -1 }

    foreach ($k in ($Keys | Where-Object { $_ })) {
        $n = Normalize-HeaderName $k

        if (Test-MapHasKey $HeaderIndex $n) { return [int](Get-MapValue $HeaderIndex $n -1) }
        if (Test-MapHasKey $HeaderIndex $k) { return [int](Get-MapValue $HeaderIndex $k -1) }
    }
    return -1
}

# -----------------------------
# Assay canonical resolution (Profiles > RuleBank aliases > Config map > Fallback)
# -----------------------------
function Resolve-AssayCanonicalName {
    param(
        [string]$RawAssay,
        $AssayMap,
        [pscustomobject]$RuleBank
    )

    $norm = Normalize-AssayName $RawAssay
    $canon = $null
    $source = 'Fallback'
    $matched = $null

    # 1) Exact profile key match (normalized compare)
    try {
        if (-not [string]::IsNullOrWhiteSpace($norm) -and $RuleBank -and $RuleBank.AssayProfiles) {
            foreach ($k in $RuleBank.AssayProfiles.Keys) {
                if ($norm -eq (Normalize-AssayName ($k + ''))) {
                    $canon = $k; $source = 'Profiles'; $matched = $k
                    break
                }
            }
        }
    } catch {}

    # 2) RuleBank aliases
    try {
        if (-not $canon -and $RuleBank -and $RuleBank.AssayAliases) {
            foreach ($k in $RuleBank.AssayAliases.Keys) {
                if ($norm -eq (Normalize-AssayName ($k + ''))) {
                    $canon = $RuleBank.AssayAliases[$k]
                    $source = 'RuleBankAlias'
                    $matched = $k
                    break
                }
            }
        }
    } catch {}

    # 3) Config map (hashtable or objects with Aliases/Tab/Assay/Canonical)
    if (-not $canon -and $AssayMap) {
        try {
            if ($AssayMap -is [hashtable]) {
                foreach ($k in $AssayMap.Keys) {
                    if ($norm -eq (Normalize-AssayName ($k + ''))) {
                        $canon = $AssayMap[$k]
                        $source = 'ConfigAssayMap'
                        $matched = $k
                        break
                    }
                }
            } else {
                foreach ($item in $AssayMap) {
                    if (-not $item) { continue }

                    $target = $null
                    try {
                        if ($item.PSObject.Properties.Match('Assay').Count -gt 0)       { $target = $item.Assay }
                        elseif ($item.PSObject.Properties.Match('Canonical').Count -gt 0) { $target = $item.Canonical }
                        elseif ($item.PSObject.Properties.Match('Tab').Count -gt 0)     { $target = $item.Tab }
                    } catch {}

                    $aliases = $null
                    try { if ($item.PSObject.Properties.Match('Aliases').Count -gt 0) { $aliases = $item.Aliases } } catch {}

                    foreach ($al in ($aliases | Where-Object { $_ })) {
                        if ($norm -eq (Normalize-AssayName ($al + ''))) {
                            $canon = $target
                            $source = 'ConfigAssayMap'
                            $matched = $al
                            break
                        }
                    }
                    if ($canon) { break }
                }
            }
        } catch {}
    }

    # 4) Fallback
    if (-not $canon) {
        $canon = if ($RawAssay) { $RawAssay } else { '_DEFAULT' }
        $source = if ($RawAssay) { 'Fallback' } else { 'Default' }
        $matched = $RawAssay
    }

    return [pscustomobject]@{
        Raw         = $RawAssay
        Normalized  = $norm
        Canonical   = $canon
        MatchSource = $source
        MatchedKey  = $matched
    }
}

function Get-AssayRuleProfile {
    param(
        [string]$Canonical,
        [pscustomobject]$RuleBank
    )

    $profile = $null
    $keyUsed = $null

    try {
        if ($RuleBank -and $RuleBank.AssayProfiles) {
            if ($Canonical -and (Test-MapHasKey $RuleBank.AssayProfiles $Canonical)) {
                $profile = $RuleBank.AssayProfiles[$Canonical]
                $keyUsed = $Canonical
            } elseif (Test-MapHasKey $RuleBank.AssayProfiles '_DEFAULT') {
                $profile = $RuleBank.AssayProfiles['_DEFAULT']
                $keyUsed = '_DEFAULT'
            } elseif (Test-MapHasKey $RuleBank.AssayProfiles 'default') {
                $profile = $RuleBank.AssayProfiles['default']
                $keyUsed = 'default'
            }
        }
    } catch {}

    if (-not $profile) {
        $profile = [ordered]@{ Mode='GlobalOnly'; Description='Fallback profile'; DisplayName='_DEFAULT' }
        $keyUsed = '_DEFAULT'
    }

    return [pscustomobject]@{
        Profile = $profile
        Key     = $keyUsed
        Mode    = if ($profile -and $profile.Mode) { $profile.Mode } else { '' }
    }
}

# -----------------------------
# Parsing helpers
# -----------------------------
function Parse-ErrorCode {
    param([string]$Text, [pscustomobject]$RuleBank)

    if ([string]::IsNullOrWhiteSpace($Text)) { return $null }
    $rxTxt = $null
    try { if ($RuleBank -and $RuleBank.ErrorBank -and $RuleBank.ErrorBank.ExtractRegex) { $rxTxt = $RuleBank.ErrorBank.ExtractRegex } } catch {}
    if (-not $rxTxt) { $rxTxt = '(?i)\b(?:error|err)?\s*(?:code)?\s*[:#]?\s*(?<Code>\d{3,})\b' }

    try {
        $m = [regex]::Match($Text, $rxTxt)
        if ($m.Success -and $m.Groups['Code'] -and $m.Groups['Code'].Value) {
            return $m.Groups['Code'].Value.ToUpperInvariant()
        }
    } catch {}
    return $null
}

function Convert-ToDoubleOrNull {
    param([string]$Text)
    if ([string]::IsNullOrWhiteSpace($Text)) { return $null }
    $t = ($Text + '').Trim()
    $t = $t -replace ',', '.'
    $v = $null
    if ([double]::TryParse($t, [System.Globalization.NumberStyles]::Any, [System.Globalization.CultureInfo]::InvariantCulture, [ref]$v)) {
        return [double]$v
    }
    return $null
}

function Parse-SampleIdParts {
    param(
        [string]$SampleId,
        [pscustomobject]$RuleBank
    )

    $defaultResult = [pscustomobject]@{ Role=''; Id=''; Idx=''; Success=$false }
    if ([string]::IsNullOrWhiteSpace($SampleId)) { return $defaultResult }

    $pattern = '^(?<role>[A-Za-z0-9_]+)[-_](?<id>\d{1,4})[-_](?<idx>\d{1,3})$'
    try {
        if ($RuleBank -and $RuleBank.Global -and $RuleBank.Global.SampleId -and $RuleBank.Global.SampleId.Parse -and $RuleBank.Global.SampleId.Parse.Pattern) {
            $pattern = $RuleBank.Global.SampleId.Parse.Pattern
        }
    } catch {}

    try {
        $m = [regex]::Match($SampleId.Trim(), $pattern)
        if ($m.Success) {
            return [pscustomobject]@{
                Role    = if ($m.Groups['role']) { $m.Groups['role'].Value } else { '' }
                Id      = if ($m.Groups['id'])   { $m.Groups['id'].Value }   else { '' }
                Idx     = if ($m.Groups['idx'])  { $m.Groups['idx'].Value }  else { '' }
                Success = $true
            }
        }
    } catch {}
    return $defaultResult
}

function Add-FailureTag {
    param(
        [pscustomobject]$Row,
        [string]$RuleId,
        [string]$Severity
    )
    if (-not $Row) { return }

    if (-not $Row.FailureTags) { $Row.FailureTags = New-Object System.Collections.Generic.List[string] }
    if (-not $Row.FailureTags.Contains($RuleId)) { [void]$Row.FailureTags.Add($RuleId) }
    if (-not $Row.PrimaryRule) { $Row.PrimaryRule = $RuleId }

    $cur  = Get-SeverityRank $Row.Severity
    $next = Get-SeverityRank $Severity
    if ($next -gt $cur) { $Row.Severity = $Severity }
}


# -----------------------------
# Display / Human-friendly naming
# -----------------------------
function Resolve-AssayDisplayName {
    param(
        [string]$Canonical,
        [pscustomobject]$RuleBank
    )

    $c = ($Canonical + '').Trim()
    if ([string]::IsNullOrWhiteSpace($c)) { return '' }

    # 1) Explicit display map
    try {
        if ($RuleBank -and $RuleBank.AssayDisplayNames -and (Test-MapHasKey $RuleBank.AssayDisplayNames $c)) {
            $v = $RuleBank.AssayDisplayNames[$c]
            if ($v) { return ($v + '') }
        }
    } catch {}

    # 2) Profile DisplayName
    try {
        if ($RuleBank -and $RuleBank.AssayProfiles -and (Test-MapHasKey $RuleBank.AssayProfiles $c)) {
            $p = $RuleBank.AssayProfiles[$c]
            if ($p -and $p.DisplayName) { return ($p.DisplayName + '') }
        }
    } catch {}

    return $c
}

# -----------------------------
# Context builder
# -----------------------------
function New-AssayRuleContext {
    param(
        [pscustomobject]$Bundle,
        $AssayMap,
        [pscustomobject]$RuleBank
    )

    $ctx = [ordered]@{
        CsvPath              = if ($Bundle) { $Bundle.Path } else { '' }
        Bundle               = $Bundle
        Delimiter            = if ($Bundle -and $Bundle.Delimiter) { $Bundle.Delimiter } else { ',' }
        HeaderRowIndex       = if ($Bundle) { [int]$Bundle.HeaderRowIndex } else { 0 }
        DataStartRowIndex    = if ($Bundle) { [int]$Bundle.DataStartRowIndex } else { 0 }
        Headers              = if ($Bundle) { $Bundle.Headers } else { @() }
        HeaderIndex          = if ($Bundle) { $Bundle.HeaderIndex } else { @{} }
        Indices              = [ordered]@{}
        AssayRaw             = $null
        AssayVersion         = $null
        WorkCenters          = @()
        ReagentLotIds        = @()
        WorkCentersDisplay   = @()
        ReagentLotDisplay     = ''
        AssayNormalized      = $null
        AssayCanonical       = $null
        AssayDisplayName     = $null
        MatchSource          = $null
        ProfileMatched       = $null
        ProfileMode          = $null
        Profile              = $null
        StatusCounts         = @{}
        ErrorCodeCounts      = @{}         # Hashtable
        DataRows             = @()
        TotalTests           = 0
        MaxPressureMax       = $null
        MaxPressureCount     = 0
        ColumnMissing        = @()
        Debug                = @()
        ParseErrors          = New-Object System.Collections.Generic.List[string]
        DuplicateCounts      = [ordered]@{}
        UniqueSampleIds      = 0
        UniqueCartridgeSN    = 0
        AssayCanonicalSource = $null
        TestTypeCounts       = @{}
    }

    if (-not $Bundle -or -not $Bundle.Lines) {
        $ctx.Debug = @([pscustomobject]@{ Name='Context'; Value='Bundle or Lines missing' })
        return [pscustomobject]$ctx
    }

    # Index lookup
    $idxAssay       = Get-HeaderIndexValue -HeaderIndex $ctx.HeaderIndex -Keys @('assay')
    $idxAssayVer    = Get-HeaderIndexValue -HeaderIndex $ctx.HeaderIndex -Keys @('assay version')
    $idxReagentLot  = Get-HeaderIndexValue -HeaderIndex $ctx.HeaderIndex -Keys @('reagent lot id','reagent lot')
    $idxSample      = Get-HeaderIndexValue -HeaderIndex $ctx.HeaderIndex -Keys @('sample id','sampleid')
    $idxCartridge   = Get-HeaderIndexValue -HeaderIndex $ctx.HeaderIndex -Keys @('cartridge s/n','cartridge sn','cartridge serial')
    $idxModule      = Get-HeaderIndexValue -HeaderIndex $ctx.HeaderIndex -Keys @('module sn','module s/n','module serial','module')
    $idxTestType    = Get-HeaderIndexValue -HeaderIndex $ctx.HeaderIndex -Keys @('test type','testtype')
    $idxStatus      = Get-HeaderIndexValue -HeaderIndex $ctx.HeaderIndex -Keys @('status')
    $idxTestResult  = Get-HeaderIndexValue -HeaderIndex $ctx.HeaderIndex -Keys @('test result','result')
    $idxError       = Get-HeaderIndexValue -HeaderIndex $ctx.HeaderIndex -Keys @('error','error code','error codes')
    $idxMaxPressure = Get-HeaderIndexValue -HeaderIndex $ctx.HeaderIndex -Keys @('max pressure (psi)','max pressure psi','max pressure')
    $idxWorkCenter  = Get-HeaderIndexValue -HeaderIndex $ctx.HeaderIndex -Keys @('work center','workcenter')

    $ctx.Indices = [ordered]@{
        Assay       = $idxAssay
        AssayVer    = $idxAssayVer
        ReagentLot  = $idxReagentLot
        SampleId    = $idxSample
        CartridgeSn = $idxCartridge
        ModuleSn    = $idxModule
        TestType    = $idxTestType
        Status      = $idxStatus
        TestResult  = $idxTestResult
        Error       = $idxError
        MaxPressure = $idxMaxPressure
        WorkCenter  = $idxWorkCenter
    }

    foreach ($kv in $ctx.Indices.GetEnumerator()) {
        if ($kv.Value -lt 0) { $ctx.ColumnMissing += $kv.Name }
    }

    # Parser selection
    $parse = $null
    if (Get-Command Parse-CsvLine -ErrorAction SilentlyContinue) {
        $parse = { param($ln,$delim) (Parse-CsvLine -line $ln -delim $delim) }
    } elseif (Get-Command ConvertTo-CsvFields -ErrorAction SilentlyContinue) {
        $parse = { param($ln,$delim) (ConvertTo-CsvFields $ln) }
    } else {
        $parse = { param($ln,$delim) ([regex]::Split($ln, [regex]::Escape($delim) + '(?=(?:[^"]*"[^"]*")*[^"]*$)')) }
    }

    $lines    = $Bundle.Lines
    $startRow = $ctx.DataStartRowIndex
    if ($startRow -lt 0) { $startRow = [int]$ctx.HeaderRowIndex + 1 }
    if ($startRow -lt 0) { $startRow = 0 }

    $reagentSet       = New-Object 'System.Collections.Generic.HashSet[string]'
    $workCenterSet    = New-Object 'System.Collections.Generic.HashSet[string]'
    $sampleCounts     = @{}
    $cartridgeCounts  = @{}
    $testTypeCounts   = @{}

    $statusCounts = @{
        'Done'        = 0
        'Error'       = 0
        'Invalid'     = 0
        'Aborted'     = 0
        'Incomplete'  = 0
        'In Progress' = 0
        'Other'       = 0
    }

    $errorCounts = @{}  # Hashtable
    $rows = New-Object 'System.Collections.Generic.List[object]'
    $maxPressureMax = $null
    $maxPressureCount = 0

    for ($r = $startRow; $r -lt $lines.Count; $r++) {
        $ln = $lines[$r]
        if (-not $ln) { continue }

        $cols = $null
        try {
            $cols = & $parse $ln $ctx.Delimiter
        } catch {
            [void]$ctx.ParseErrors.Add(("Row {0}: Parse failed: {1}" -f ($r+1), $_.Exception.Message))
            continue
        }

        if (-not $cols -or ($cols -join '').Trim().Length -eq 0) { continue }

        try {
            $rowObj = [ordered]@{
                RowIndex     = ($r + 1)
                DataIndex    = ($r - $startRow + 1)
                Assay        = if ($idxAssay -ge 0 -and $cols.Count -gt $idxAssay) { ($cols[$idxAssay] + '').Trim() } else { '' }
                AssayVersion = if ($idxAssayVer -ge 0 -and $cols.Count -gt $idxAssayVer) { ($cols[$idxAssayVer] + '').Trim() } else { '' }
                ReagentLotId = if ($idxReagentLot -ge 0 -and $cols.Count -gt $idxReagentLot) { ($cols[$idxReagentLot] + '').Trim() } else { '' }
                SampleID     = if ($idxSample -ge 0 -and $cols.Count -gt $idxSample) { ($cols[$idxSample] + '').Trim() } else { '' }
                CartridgeSN  = if ($idxCartridge -ge 0 -and $cols.Count -gt $idxCartridge) { ($cols[$idxCartridge] + '').Trim() } else { '' }
                ModuleSN     = if ($idxModule -ge 0 -and $cols.Count -gt $idxModule) { ($cols[$idxModule] + '').Trim() } else { '' }
                TestType     = if ($idxTestType -ge 0 -and $cols.Count -gt $idxTestType) { ($cols[$idxTestType] + '').Trim() } else { '' }
                Status       = if ($idxStatus -ge 0 -and $cols.Count -gt $idxStatus) { ($cols[$idxStatus] + '').Trim() } else { '' }
                TestResult   = if ($idxTestResult -ge 0 -and $cols.Count -gt $idxTestResult) { ($cols[$idxTestResult] + '').Trim() } else { '' }
                ErrorRaw     = if ($idxError -ge 0 -and $cols.Count -gt $idxError) { ($cols[$idxError] + '').Trim() } else { '' }
                ErrorCode    = $null
                MaxPressure  = $null
                WorkCenter   = if ($idxWorkCenter -ge 0 -and $cols.Count -gt $idxWorkCenter) { ($cols[$idxWorkCenter] + '').Trim() } else { '' }
                FailureTags  = New-Object System.Collections.Generic.List[string]
                Notes        = ''
                Severity     = 'Info'
                ParsedRole   = ''
                ParsedId     = ''
                ParsedIdx    = ''
                PrimaryRule  = ''
            }

            if ($rowObj.ReagentLotId) {
                $m = [regex]::Match($rowObj.ReagentLotId, '(?<!\d)(\d{5})(?!\d)')
                if ($m.Success) { $rowObj.ReagentLotId = $m.Groups[1].Value }
                [void]$reagentSet.Add($rowObj.ReagentLotId)
            }
            if ($rowObj.WorkCenter) { [void]$workCenterSet.Add($rowObj.WorkCenter) }

            if ($idxError -ge 0) {
                $code = Parse-ErrorCode -Text $rowObj.ErrorRaw -RuleBank $RuleBank
                if ($code) {
                    $rowObj.ErrorCode = $code
                    Add-MapCount -Map $errorCounts -Key $code -Delta 1
                }
            }

            if ($idxMaxPressure -ge 0 -and $cols.Count -gt $idxMaxPressure) {
                $mp = Convert-ToDoubleOrNull (($cols[$idxMaxPressure] + '').Trim())
                if ($mp -ne $null) {
                    $rowObj.MaxPressure = [double]$mp
                    if ($maxPressureMax -eq $null -or $mp -gt $maxPressureMax) { $maxPressureMax = [double]$mp }
                    $maxPressureCount++
                }
            }

            try {
                $parsed = Parse-SampleIdParts -SampleId $rowObj.SampleID -RuleBank $RuleBank
                if ($parsed -and $parsed.Success) {
                    $rowObj.ParsedRole = $parsed.Role
                    $rowObj.ParsedId   = $parsed.Id
                    $rowObj.ParsedIdx  = $parsed.Idx
                }
            } catch {}

            if ($rowObj.SampleID) {
                if (-not $sampleCounts.ContainsKey($rowObj.SampleID)) { $sampleCounts[$rowObj.SampleID] = 0 }
                $sampleCounts[$rowObj.SampleID]++
            }
            if ($rowObj.CartridgeSN) {
                if (-not $cartridgeCounts.ContainsKey($rowObj.CartridgeSN)) { $cartridgeCounts[$rowObj.CartridgeSN] = 0 }
                $cartridgeCounts[$rowObj.CartridgeSN]++
            }
            if ($rowObj.TestType) {
                if (-not $testTypeCounts.ContainsKey($rowObj.TestType)) { $testTypeCounts[$rowObj.TestType] = 0 }
                $testTypeCounts[$rowObj.TestType]++
            }

            $st = ($rowObj.Status + '').Trim()
            if ($st) {
                switch -Regex ($st) {
                    '(?i)^done$'          { $statusCounts['Done']++ }
                    '(?i)^error$'         { $statusCounts['Error']++ }
                    '(?i)^invalid$'       { $statusCounts['Invalid']++ }
                    '(?i)^aborted$'       { $statusCounts['Aborted']++ }
                    '(?i)^incomplete$'    { $statusCounts['Incomplete']++ }
                    '(?i)^in\s*progress$' { $statusCounts['In Progress']++ }
                    default               { $statusCounts['Other']++ }
                }
            } else {
                $statusCounts['Other']++
            }

            [void]$rows.Add([pscustomobject]$rowObj)
        } catch {
            [void]$ctx.ParseErrors.Add(("Row {0}: Parse/convert failed: {1}" -f ($r+1), $_.Exception.Message))
        }
    }

    $ctx.DataRows        = $rows.ToArray()
    $ctx.TotalTests      = $ctx.DataRows.Count
    $ctx.StatusCounts    = $statusCounts
    $ctx.ErrorCodeCounts = $errorCounts

    # HashSet -> string[]
    try { $ctx.ReagentLotIds = @($reagentSet | Sort-Object) } catch { $ctx.ReagentLotIds = @() }
    try { $ctx.WorkCenters   = @($workCenterSet | Sort-Object) } catch { $ctx.WorkCenters = @() }

    $ctx.UniqueSampleIds   = ($sampleCounts.Keys | Where-Object { $_ }).Count
    $ctx.UniqueCartridgeSN = ($cartridgeCounts.Keys | Where-Object { $_ }).Count

    $dupSample = @($sampleCounts.Keys | Where-Object { $sampleCounts[$_] -gt 1 })
    $dupCart   = @($cartridgeCounts.Keys | Where-Object { $cartridgeCounts[$_] -gt 1 })
    $ctx.DuplicateCounts = [ordered]@{
        SampleId    = $dupSample.Count
        CartridgeSN = $dupCart.Count
    }

    $ctx.MaxPressureMax   = $maxPressureMax
    $ctx.MaxPressureCount = $maxPressureCount
    $ctx.TestTypeCounts   = $testTypeCounts

    # Set AssayRaw / AssayVersion to most common value
    try {
        $assayGroups = $ctx.DataRows | Where-Object { $_.Assay } | Group-Object -Property Assay | Sort-Object -Property Count -Descending
        if ($assayGroups -and $assayGroups.Count -gt 0) { $ctx.AssayRaw = ($assayGroups[0].Name + '') }

        $verGroups = $ctx.DataRows | Where-Object { $_.AssayVersion } | Group-Object -Property AssayVersion | Sort-Object -Property Count -Descending
        if ($verGroups -and $verGroups.Count -gt 0) { $ctx.AssayVersion = ($verGroups[0].Name + '') }
    } catch {}

    $resolved = Resolve-AssayCanonicalName -RawAssay $ctx.AssayRaw -AssayMap $AssayMap -RuleBank $RuleBank
    $ctx.AssayNormalized      = $resolved.Normalized
    $ctx.AssayCanonical       = $resolved.Canonical
    $ctx.MatchSource          = $resolved.MatchSource
    $ctx.AssayCanonicalSource = $resolved.MatchSource

    $profileResolved = Get-AssayRuleProfile -Canonical $ctx.AssayCanonical -RuleBank $RuleBank
    if ($profileResolved) {
        $ctx.Profile        = $profileResolved.Profile
        $ctx.ProfileMatched = $profileResolved.Key
        $ctx.ProfileMode    = $profileResolved.Mode
    }

    # Human-friendly display values (how we talk)
    try { $ctx.AssayDisplayName = Resolve-AssayDisplayName -Canonical $ctx.AssayCanonical -RuleBank $RuleBank } catch { $ctx.AssayDisplayName = $ctx.AssayCanonical }
    try {
        if ($ctx.ReagentLotIds -and $ctx.ReagentLotIds.Count -gt 0) {
            $ctx.ReagentLotDisplay = 'LSP ' + ($ctx.ReagentLotIds -join ', LSP ')
        } else {
            $ctx.ReagentLotDisplay = ''
        }
    } catch {
        try { $ctx.ReagentLotDisplay = ($ctx.ReagentLotIds -join ', ') } catch { $ctx.ReagentLotDisplay = '' }
    }
    try { $ctx.WorkCentersDisplay = $ctx.WorkCenters } catch { $ctx.WorkCentersDisplay = @() }

    $ctx.Debug = @(
        [pscustomobject]@{ Name='CsvPath';        Value=$ctx.CsvPath },
        [pscustomobject]@{ Name='Delimiter';     Value=$ctx.Delimiter },
        [pscustomobject]@{ Name='HeaderRow';     Value=$ctx.HeaderRowIndex },
        [pscustomobject]@{ Name='DataStartRow';  Value=$ctx.DataStartRowIndex },
        [pscustomobject]@{ Name='RowCount';      Value=$ctx.TotalTests },
        [pscustomobject]@{ Name='ParseErrors';   Value=$ctx.ParseErrors.Count },
        [pscustomobject]@{ Name='AssayRaw';      Value=$ctx.AssayRaw },
        [pscustomobject]@{ Name='AssayCanonical';Value=$ctx.AssayCanonical },
        [pscustomobject]@{ Name='MatchSource';   Value=$ctx.MatchSource },
        [pscustomobject]@{ Name='AssayVersion';  Value=$ctx.AssayVersion },
        [pscustomobject]@{ Name='ReagentLotIds'; Value=($ctx.ReagentLotIds -join ', ') },
        [pscustomobject]@{ Name='WorkCenters';   Value=($ctx.WorkCenters -join ', ') },
        [pscustomobject]@{ Name='MaxPressureMax';Value=$ctx.MaxPressureMax }
    )

    return [pscustomobject]$ctx
}

# -----------------------------
# Engine (pure logic)
# -----------------------------
function Invoke-AssayRuleEngine {
    param(
        [pscustomobject]$Context,
        [pscustomobject]$RuleBank
    )

    $result = [pscustomobject]@{
        Findings               = @()
        AffectedTests          = @()
        AffectedTestsTruncated = 0
        ErrorSummary           = @()
        PressureStats          = [pscustomobject]@{ Max=$null; OverWarn=0; OverFail=0; WarnThreshold=$null; FailThreshold=$null }
        DuplicatesTable        = @()
        Debug                  = @()
        OverallSeverity        = 'Info'
        OverallStatus          = 'PASS'
        StatusCounts           = @{}
        TotalTests             = 0
        UniqueErrorCodes       = 0
        IdentityFlags          = @{}
        SeverityCounts         = @{ Error=0; Warn=0; Info=0 }  # Hashtable
        Exception              = $null
        BaselineExpected       = 0
        BaselineDelta          = 0
        ErrorRowCount          = 0
    }

    $findings = New-Object System.Collections.Generic.List[object]

    function New-Finding {
        param(
            [string]$Severity,
            [string]$RuleId,
            [string]$Message,
            [object[]]$Rows,
            [string]$Evidence = $null
        )
        $affCount = if ($Rows) { [int]$Rows.Count } else { 0 }
        $example = ''
        if ($Rows -and $Rows.Count -gt 0) {
            $sample = ($Rows | Where-Object { $_.SampleID } | Select-Object -First 1 -ExpandProperty SampleID)
            if ($sample) { $example = $sample }
            elseif ($Rows[0].RowIndex) { $example = "Row $($Rows[0].RowIndex)" }
        }
        return [pscustomobject]@{
            Severity  = $Severity
            RuleId    = $RuleId
            Title     = $Message
            Message   = $Message
            Count     = $affCount
            Example   = $example
            Evidence  = ($Evidence + '')
        }
    }

    if (-not $Context) {
        [void]$findings.Add((New-Finding -Severity 'Error' -RuleId 'ENGINE_EXCEPTION' -Message 'Context missing' -Rows $null -Evidence 'Context object is null'))
        $result.Findings = $findings.ToArray()
        $result.OverallSeverity = 'Error'
        $result.OverallStatus = 'FAIL'
        return [pscustomobject]$result
    }

    try {
        $rows = $Context.DataRows
        $result.TotalTests   = [int]$Context.TotalTests
        $result.StatusCounts = $Context.StatusCounts
        $result.Debug        = $Context.Debug

        # Parse warnings as Info (visible but non-blocking)
        try {
            if ($Context.ParseErrors -and $Context.ParseErrors.Count -gt 0) {
                $msg = ("CSV parse warnings: {0}" -f [int]$Context.ParseErrors.Count)
                $ev  = ($Context.ParseErrors | Select-Object -First 3) -join ' | '
                [void]$findings.Add((New-Finding -Severity 'Info' -RuleId 'CSV_PARSE_WARNINGS' -Message $msg -Rows $null -Evidence $ev))
            }
        } catch {}

        # Pressure thresholds
        $warnThreshold = 85
        $failThreshold = 90
        try {
            if ($RuleBank -and $RuleBank.Global -and $RuleBank.Global.MaxPressure) {
                if ($RuleBank.Global.MaxPressure.WarnThreshold -ne $null) { $warnThreshold = [double]$RuleBank.Global.MaxPressure.WarnThreshold }
                if ($RuleBank.Global.MaxPressure.FailThreshold -ne $null) { $failThreshold = [double]$RuleBank.Global.MaxPressure.FailThreshold }
            }
        } catch {}

        $result.PressureStats.Max = $Context.MaxPressureMax
        $result.PressureStats.WarnThreshold = $warnThreshold
        $result.PressureStats.FailThreshold = $failThreshold

        try {
            $overWarn = @($rows | Where-Object { $_.MaxPressure -ne $null -and [double]$_.MaxPressure -ge [double]$warnThreshold }).Count
            $overFail = @($rows | Where-Object { $_.MaxPressure -ne $null -and [double]$_.MaxPressure -ge [double]$failThreshold }).Count
            $result.PressureStats.OverWarn = [int]$overWarn
            $result.PressureStats.OverFail = [int]$overFail
        } catch {}

        # Baseline policy
        $expected = 0
        $allowedMissing = 0
        $allowedExtra   = 0
        $baselineSeverity = 'Error'
        $baselineRuleId   = 'BASELINE_SAMPLE_SIZE'
        try {
            if ($RuleBank -and $RuleBank.Baseline -and $RuleBank.Baseline.SampleSizePolicy) {
                $policy = $RuleBank.Baseline.SampleSizePolicy
                if ($policy.Expected -ne $null)       { $expected = [int]$policy.Expected }
                if ($policy.AllowedMissing -ne $null) { $allowedMissing = [int]$policy.AllowedMissing }
                if ($policy.AllowedExtra -ne $null)   { $allowedExtra = [int]$policy.AllowedExtra }
                if ($policy.Severity)                 { $baselineSeverity = $policy.Severity }
                if ($policy.RuleId)                   { $baselineRuleId = $policy.RuleId }
            }
        } catch {}

        $diff = $result.TotalTests - $expected
        $result.BaselineExpected = $expected
        $result.BaselineDelta    = $diff

        if ($expected -gt 0) {
            if ($diff -lt 0 -and [Math]::Abs($diff) -gt $allowedMissing) {
                [void]$findings.Add((New-Finding -Severity $baselineSeverity -RuleId $baselineRuleId -Message ("{0} tests missing (expected {1}, got {2})" -f ([Math]::Abs($diff)), $expected, $result.TotalTests) -Rows $null -Evidence ("AllowedMissing=$allowedMissing; AllowedExtra=$allowedExtra")))
            } elseif ($diff -gt 0 -and $diff -gt $allowedExtra) {
                [void]$findings.Add((New-Finding -Severity $baselineSeverity -RuleId $baselineRuleId -Message ("{0} extra tests (expected {1}, got {2})" -f $diff, $expected, $result.TotalTests) -Rows $null -Evidence ("AllowedMissing=$allowedMissing; AllowedExtra=$allowedExtra")))
            }
        }

        foreach ($col in ($Context.ColumnMissing | Sort-Object)) {
            [void]$findings.Add((New-Finding -Severity 'Info' -RuleId ('COLUMN_MISSING_' + $col.ToUpper()) -Message ("Column missing: {0}" -f $col) -Rows $null))
        }

        # Identity checks (constant fields)
        $identityDefs = @()
        try {
            if ($RuleBank -and $RuleBank.Global -and $RuleBank.Global.Identity -and $RuleBank.Global.Identity.Fields) {
                $identityDefs = @($RuleBank.Global.Identity.Fields.GetEnumerator() | ForEach-Object {
                    [pscustomobject]@{ Name=$_.Key; Config=$_.Value }
                })
            }
        } catch {}
        if ($identityDefs.Count -eq 0) {
            $identityDefs = @(
                [pscustomobject]@{ Name='Assay';        Config=[ordered]@{ RuleId='IDENTITY_ASSAY'; Severity='Error' } },
                [pscustomobject]@{ Name='AssayVersion'; Config=[ordered]@{ RuleId='IDENTITY_ASSAY_VERSION'; Severity='Error' } },
                [pscustomobject]@{ Name='ReagentLotId'; Config=[ordered]@{ RuleId='IDENTITY_REAGENT_LOT_ID'; Severity='Error' } }
            )
        }

        $result.IdentityFlags = [ordered]@{}
        foreach ($idDef in $identityDefs) {
            $propName = $idDef.Name
            $ruleId   = if ($idDef.Config.RuleId) { $idDef.Config.RuleId } else { 'IDENTITY_' + ($propName -replace '\W','_').ToUpper() }
            $sevId    = if ($idDef.Config.Severity) { $idDef.Config.Severity } else { 'Error' }

            $vals = @($rows | ForEach-Object { $_.$propName } | Where-Object { $_ }) | Select-Object -Unique
            if ($vals.Count -gt 1) {
                $baselineVal = $vals[0]
                $affected = @($rows | Where-Object { $_.$propName -and $_.$propName -ne $baselineVal })
                foreach ($r in $affected) { Add-FailureTag -Row $r -RuleId $ruleId -Severity $sevId }
                $result.IdentityFlags[$propName] = 'Mismatch'
                [void]$findings.Add((New-Finding -Severity $sevId -RuleId $ruleId -Message ("{0} mismatch (expected '{1}')" -f $propName, $baselineVal) -Rows $affected))
            } elseif ($vals.Count -eq 1) {
                $result.IdentityFlags[$propName] = 'Ok'
            } else {
                $result.IdentityFlags[$propName] = 'Missing'
            }
        }

        # Duplicate counts (summary only)
        if ($Context.DuplicateCounts) {
            foreach ($k in $Context.DuplicateCounts.Keys) {
                if ([int]$Context.DuplicateCounts[$k] -gt 0) {
                    [void]$findings.Add((New-Finding -Severity 'Warn' -RuleId ("DUPLICATE_" + $k.ToUpper()) -Message ("Duplicate {0} detected" -f $k) -Rows $null -Evidence ("Count={0}" -f $Context.DuplicateCounts[$k])))
                }
            }
        }

        # Status checks (Error/Invalid/Aborted/Incomplete -> Error)
        $failStatuses = @('error','invalid','aborted','incomplete')
        $passStatuses = @('done')
        try { if ($RuleBank -and $RuleBank.Global -and $RuleBank.Global.Status -and $RuleBank.Global.Status.FailStatuses) { $failStatuses = @($RuleBank.Global.Status.FailStatuses | ForEach-Object { $_.ToString().ToLower() }) } } catch {}
        try { if ($RuleBank -and $RuleBank.Global -and $RuleBank.Global.Status -and $RuleBank.Global.Status.PassStatuses) { $passStatuses = @($RuleBank.Global.Status.PassStatuses | ForEach-Object { $_.ToString().ToLower() }) } } catch {}

        $rowsFailStatus = @($rows | Where-Object { $_.Status -and ($failStatuses -contains ($_.Status.ToLower())) })
        if ($rowsFailStatus.Count -gt 0) {
            foreach ($r in $rowsFailStatus) { Add-FailureTag -Row $r -RuleId 'STATUS_FAILED' -Severity 'Error' }
            [void]$findings.Add((New-Finding -Severity 'Error' -RuleId 'STATUS_FAILED' -Message 'Error/Invalid/Failed status present.' -Rows $rowsFailStatus))
        }

        # Apply RuleBank custom rules (optional)
        if ($RuleBank -and $RuleBank.Rules) {
            foreach ($rule in $RuleBank.Rules) {
                if (-not $rule) { continue }
                try {
                    $rid = $rule.RuleId
                    $sev = if ($rule.Severity) { $rule.Severity } else { 'Warn' }
                    $msg = $rule.Message
                    $affRows = @()
                    if ($rule.Filter) {
                        $affRows = @($rows | Where-Object $rule.Filter)
                    } elseif ($rule.FilterScript) {
                        $affRows = @($rows | Where-Object $rule.FilterScript)
                    }

                    if ($affRows.Count -gt 0) {
                        foreach ($r in $affRows) { Add-FailureTag -Row $r -RuleId $rid -Severity $sev }
                        [void]$findings.Add((New-Finding -Severity $sev -RuleId $rid -Message $msg -Rows $affRows))
                    }
                } catch {}
            }
        }

        # Error code breakdown (table)
        $errorCodeCounts = $Context.ErrorCodeCounts
        $knownCodes = @()
        try { if ($RuleBank -and $RuleBank.ErrorBank -and $RuleBank.ErrorBank.Codes) { $knownCodes = @($RuleBank.ErrorBank.Codes.Keys | ForEach-Object { $_.ToString() }) } } catch {}

        if ($errorCodeCounts -and $errorCodeCounts.Count -gt 0) {
            $errDetails = New-Object System.Collections.Generic.List[object]
            foreach ($code in ($errorCodeCounts.Keys | Sort-Object)) {
                $def = $null
                try { if ($RuleBank -and $RuleBank.ErrorBank -and $RuleBank.ErrorBank.Codes) { $def = $RuleBank.ErrorBank.Codes[$code] } } catch {}
                $exampleRow = ($rows | Where-Object { $_.ErrorCode -eq $code } | Select-Object -First 1)
                [void]$errDetails.Add([pscustomobject]@{
                    ErrorCode       = $code
                    Name            = if ($def -and $def.Name) { $def.Name } else { '' }
                    Classification  = if ($def -and $def.Classification) { $def.Classification } else { '' }
                    Count           = [int]$errorCodeCounts[$code]
                    ExampleSampleID = if ($exampleRow) { $exampleRow.SampleID } else { '' }
                })
            }
            $result.ErrorSummary = @($errDetails | Sort-Object -Property Count -Descending, ErrorCode)
            $result.UniqueErrorCodes = [int]$result.ErrorSummary.Count

            if ($knownCodes.Count -gt 0) {
                $rowsUnknownErr = @($rows | Where-Object { $_.ErrorCode -and ($knownCodes -notcontains $_.ErrorCode) })
                if ($rowsUnknownErr.Count -gt 0) {
                    foreach ($r in $rowsUnknownErr) { Add-FailureTag -Row $r -RuleId 'UNKNOWN_ERROR_CODE' -Severity 'Warn' }
                    $unknownCodesTxt = (($rowsUnknownErr | Select-Object -ExpandProperty ErrorCode) | Select-Object -Unique) -join ', '
                    [void]$findings.Add((New-Finding -Severity 'Warn' -RuleId 'UNKNOWN_ERROR_CODE' -Message 'Error code not in known list.' -Rows $rowsUnknownErr -Evidence ("Unknown codes: {0}" -f $unknownCodesTxt)))
                }
            }
        } else {
            $result.UniqueErrorCodes = 0
        }

        # Affected tests (Warn/Error)
        $affected = @($rows | Where-Object { $_.FailureTags -and $_.FailureTags.Count -gt 0 -and (Get-SeverityRank $_.Severity) -ge (Get-SeverityRank 'Warn') })
        $affected = @(
            $affected | Sort-Object `
                @{ Expression = { Get-SeverityRank $_.Severity }; Descending = $true }, `
                @{ Expression = { $_.ErrorCode } }, `
                @{ Expression = { $_.SampleID } }, `
                @{ Expression = { $_.RowIndex } }
        )
        $result.AffectedTests = $affected
        $result.ErrorRowCount = $affected.Count

        if ($findings.Count -eq 0) {
            [void]$findings.Add((New-Finding -Severity 'Info' -RuleId 'NO_FINDINGS' -Message 'No rule findings detected.' -Rows $null))
        }

        $sortedFindings = @(
            $findings | Sort-Object `
                @{ Expression = { Get-SeverityRank $_.Severity }; Descending = $true }, `
                @{ Expression = { $_.Count }; Descending = $true }
        )
        $result.Findings = $sortedFindings

        foreach ($f in $sortedFindings) {
            if ($result.SeverityCounts.ContainsKey($f.Severity)) { $result.SeverityCounts[$f.Severity] = [int]$result.SeverityCounts[$f.Severity] + 1 }
        }

        $worst = 'Info'
        foreach ($f in $sortedFindings) {
            if ((Get-SeverityRank $f.Severity) -gt (Get-SeverityRank $worst)) { $worst = $f.Severity }
        }
        $result.OverallSeverity = $worst
        switch ($worst) {
            'Error' { $result.OverallStatus = 'FAIL' }
            'Warn'  { $result.OverallStatus = 'WARN' }
            default { $result.OverallStatus = 'PASS' }
        }

    } catch {
        $ev = $null
        try { $ev = $_.Exception.ToString() } catch { $ev = $_.Exception.Message }

        [void]$findings.Add([pscustomobject]@{
            Severity = 'Error'
            RuleId   = 'ENGINE_EXCEPTION'
            Title    = 'RuleEngine exception'
            Message  = 'RuleEngine exception'
            Count    = 0
            Evidence = $ev
            Example  = ''
        })
        $result.Findings = $findings.ToArray()
        $result.OverallSeverity = 'Error'
        $result.OverallStatus = 'FAIL'
        $result.Exception = $_.Exception
    }

    return [pscustomobject]$result
}
