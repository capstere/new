#requires -Version 5.1

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

        # Fallback via methods if present
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

function Set-MapValue {
    param([object]$Map, [object]$Key, [object]$Value)
    if ($null -eq $Map -or $null -eq $Key) { return $false }
    try { $Map[$Key] = $Value; return $true } catch { return $false }
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
# Normalization helpers
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
                        if ($item.PSObject.Properties.Match('Assay').Count -gt 0)      { $target = $item.Assay }
                        elseif ($item.PSObject.Properties.Match('Canonical').Count -gt 0){ $target = $item.Canonical }
                        elseif ($item.PSObject.Properties.Match('Tab').Count -gt 0)    { $target = $item.Tab }
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
# Context builder
# -----------------------------
function New-AssayRuleContext {
    param(
        [pscustomobject]$Bundle,
        $AssayMap,
        [pscustomobject]$RuleBank
    )

    $ctx = [ordered]@{
        CsvPath             = if ($Bundle) { $Bundle.Path } else { '' }
        Bundle              = $Bundle
        Delimiter           = if ($Bundle -and $Bundle.Delimiter) { $Bundle.Delimiter } else { ',' }
        HeaderRowIndex      = if ($Bundle) { [int]$Bundle.HeaderRowIndex } else { 0 }
        DataStartRowIndex   = if ($Bundle) { [int]$Bundle.DataStartRowIndex } else { 0 }
        Headers             = if ($Bundle) { $Bundle.Headers } else { @() }
        HeaderIndex         = if ($Bundle) { $Bundle.HeaderIndex } else { @{} }
        Indices             = [ordered]@{}
        AssayRaw            = $null
        AssayVersion        = $null
        WorkCenters         = @()
        ReagentLotIds        = @()
        AssayNormalized      = $null
        AssayCanonical       = $null
        MatchSource          = $null
        ProfileMatched       = $null
        ProfileMode          = $null
        Profile              = $null
        StatusCounts         = @{}
        ErrorCodeCounts      = @{}         # hashtable (safe for ContainsKey)
        DataRows             = @()
        TotalTests           = 0
        MaxPressureMax       = $null
        MaxPressureAvg       = $null
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

    $reagentSet    = New-Object 'System.Collections.Generic.HashSet[string]'
    $workCenterSet = New-Object 'System.Collections.Generic.HashSet[string]'
    $sampleCounts  = @{}
    $cartridgeCounts = @{}
    $testTypeCounts = @{}

    $statusCounts = @{
        'Done'        = 0
        'Error'       = 0
        'Invalid'     = 0
        'Aborted'     = 0
        'Incomplete'  = 0
        'In Progress' = 0
        'Other'       = 0
    }

    $errorCounts = @{}  # <-- IMPORTANT: Hashtable (no OrderedDictionary ContainsKey crash)
    $rows = New-Object System.Collections.Generic.List[object]
    $maxPressureMax = $null
    $maxPressureSum = 0.0
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
                    if ($maxPressureMax -eq $null -or $mp -gt $maxPressureMax) { $maxPressureMax = $mp }
                    $maxPressureSum += [double]$mp
                    $maxPressureCount++
                }
            }

            $parsed = Parse-SampleIdParts -SampleId $rowObj.SampleID -RuleBank $RuleBank
            if ($parsed) {
                $rowObj.ParsedRole = $parsed.Role
                $rowObj.ParsedId   = $parsed.Id
                $rowObj.ParsedIdx  = $parsed.Idx
            }

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

            $st = ($rowObj.Status + '')
            if ($st) {
                switch -Regex ($st) {
                    '(?i)^done$'          { $statusCounts['Done']++ }
                    '(?i)^error$'         { $statusCounts['Error']++ }
                    '(?i)^invalid$'        { $statusCounts['Invalid']++ }
                    '(?i)^aborted$'        { $statusCounts['Aborted']++ }
                    '(?i)^incomplete$'     { $statusCounts['Incomplete']++ }
                    '(?i)^in\s*progress$'  { $statusCounts['In Progress']++ }
                    default                { $statusCounts['Other']++ }
                }
            } else {
                $statusCounts['Other']++
            }

            [void]$rows.Add([pscustomobject]$rowObj)

            if (-not $ctx.AssayRaw -and $rowObj.Assay) { $ctx.AssayRaw = $rowObj.Assay }
            if (-not $ctx.AssayVersion -and $rowObj.AssayVersion) { $ctx.AssayVersion = $rowObj.AssayVersion }

            $ctx.TotalTests++
        } catch {
            [void]$ctx.ParseErrors.Add(("Row {0}: Build row failed: {1}" -f ($r+1), $_.Exception.Message))
            continue
        }
    }

    $ctx.ReagentLotIds      = @($reagentSet | Sort-Object)
    $ctx.WorkCenters        = @($workCenterSet | Sort-Object)
    $ctx.StatusCounts       = $statusCounts
    $ctx.ErrorCodeCounts    = $errorCounts
    $ctx.DataRows           = $rows
    $ctx.MaxPressureMax     = $maxPressureMax
    $ctx.MaxPressureCount   = $maxPressureCount
    $ctx.MaxPressureAvg     = if ($maxPressureCount -gt 0) { [double]($maxPressureSum / $maxPressureCount) } else { $null }
    $ctx.UniqueSampleIds    = ($sampleCounts.Keys | Where-Object { $_ }).Count
    $ctx.UniqueCartridgeSN  = ($cartridgeCounts.Keys | Where-Object { $_ }).Count
    $ctx.DuplicateCounts    = [ordered]@{
        SampleId    = ($sampleCounts.GetEnumerator() | Where-Object { $_.Value -gt 1 }).Count
        CartridgeSN = ($cartridgeCounts.GetEnumerator() | Where-Object { $_.Value -gt 1 }).Count
    }
    $ctx.TestTypeCounts     = $testTypeCounts

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

    # Debug info
    $ctx.Debug = @(
        [pscustomobject]@{ Name='CsvPath';          Value=$ctx.CsvPath },
        [pscustomobject]@{ Name='Delimiter';       Value=$ctx.Delimiter },
        [pscustomobject]@{ Name='HeaderRowIndex';  Value=$ctx.HeaderRowIndex },
        [pscustomobject]@{ Name='DataStartRow';    Value=$ctx.DataStartRowIndex },
        [pscustomobject]@{ Name='RowCount';        Value=$ctx.TotalTests },
        [pscustomobject]@{ Name='ParseErrors';     Value=$ctx.ParseErrors.Count },
        [pscustomobject]@{ Name='AssayRaw';        Value=$ctx.AssayRaw },
        [pscustomobject]@{ Name='AssayCanonical';  Value=$ctx.AssayCanonical },
        [pscustomobject]@{ Name='MatchSource';     Value=$ctx.MatchSource },
        [pscustomobject]@{ Name='AssayVersion';    Value=$ctx.AssayVersion },
        [pscustomobject]@{ Name='ReagentLotIds';   Value=($ctx.ReagentLotIds -join ', ') },
        [pscustomobject]@{ Name='WorkCenters';     Value=($ctx.WorkCenters -join ', ') },
        [pscustomobject]@{ Name='MaxPressureMax';  Value=$ctx.MaxPressureMax },
        [pscustomobject]@{ Name='MaxPressureAvg';  Value=$(if ($ctx.MaxPressureAvg -ne $null) { [math]::Round([double]$ctx.MaxPressureAvg,2) } else { '—' }) }
    )

    return [pscustomobject]$ctx
}

# -----------------------------
# Engine
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
        ErrorCodesTable        = @()
        PressureStats          = [pscustomobject]@{ Max=$null; Avg=$null; OverWarn=0; OverFail=0; WarnThreshold=$null; FailThreshold=$null }
        DuplicatesTable        = @()
        Debug                  = @()
        OverallSeverity        = 'Info'
        OverallStatus          = 'PASS'
        StatusCounts           = @{}
        TotalTests             = 0
        UniqueErrorCodes       = 0
        IdentityFlags          = @{}
        SeverityCounts         = @{ Error=0; Warn=0; Info=0 }  # <-- Hashtable, safe
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
            [string]$Classification = '',
            [bool]$GeneratesRetest = $false,
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
            Severity        = $Severity
            RuleId          = $RuleId
            Message         = $Message
            Count           = $affCount
            AffectedCount   = $affCount
            Title           = $Message
            Example         = $example
            Evidence        = ($Evidence + '')
            Classification  = $Classification
            GeneratesRetest = [bool]$GeneratesRetest
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
        $result.TotalTests    = [int]$Context.TotalTests
        $result.StatusCounts  = $Context.StatusCounts
        $result.Debug         = $Context.Debug

        # Surface parse errors as INFO (so you see them, but it doesn't kill the run)
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
        $result.PressureStats.Avg = $Context.MaxPressureAvg
        $result.PressureStats.WarnThreshold = $warnThreshold
        $result.PressureStats.FailThreshold = $failThreshold

        # Baseline policy (respects RuleBank, but never crashes)
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

        # Identity checks
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
                [void]$findings.Add((New-Finding -Severity $sevId -RuleId $ruleId -Message ("{0} is not constant" -f $propName) -Rows $affected -Evidence ("Values: {0}" -f ($vals -join '; '))))
                $result.IdentityFlags[$propName] = $true
            } else {
                $result.IdentityFlags[$propName] = $false
            }
        }

        # Duplicates
        $dupSample   = @($rows | Group-Object SampleID   | Where-Object { $_.Name -and $_.Count -gt 1 })
        $dupCartridge= @($rows | Group-Object CartridgeSN| Where-Object { $_.Name -and $_.Count -gt 1 })

        $sevDupSample = 'Error'
        $sevDupCart   = 'Error'
        try { if ($RuleBank -and $RuleBank.Global -and $RuleBank.Global.Uniqueness -and $RuleBank.Global.Uniqueness.SampleId -and $RuleBank.Global.Uniqueness.SampleId.Severity) { $sevDupSample = $RuleBank.Global.Uniqueness.SampleId.Severity } } catch {}
        try { if ($RuleBank -and $RuleBank.Global -and $RuleBank.Global.Uniqueness -and $RuleBank.Global.Uniqueness.CartridgeSN -and $RuleBank.Global.Uniqueness.CartridgeSN.Severity) { $sevDupCart = $RuleBank.Global.Uniqueness.CartridgeSN.Severity } } catch {}

        $dupSampleRows = New-Object System.Collections.Generic.List[object]
        foreach ($grp in $dupSample) {
            foreach ($r in $grp.Group) { Add-FailureTag -Row $r -RuleId 'DUP_SAMPLE_ID' -Severity $sevDupSample; [void]$dupSampleRows.Add($r) }
        }
        if ($dupSampleRows.Count -gt 0) {
            [void]$findings.Add((New-Finding -Severity $sevDupSample -RuleId 'DUP_SAMPLE_ID' -Message 'Duplicate Sample ID detected.' -Rows $dupSampleRows.ToArray() -Evidence (($dupSample | ForEach-Object { "$($_.Name) x$($_.Count)" }) -join '; ')))
        }

        $dupCartRows = New-Object System.Collections.Generic.List[object]
        foreach ($grp in $dupCartridge) {
            foreach ($r in $grp.Group) { Add-FailureTag -Row $r -RuleId 'DUP_CARTRIDGE_SN' -Severity $sevDupCart; [void]$dupCartRows.Add($r) }
        }
        if ($dupCartRows.Count -gt 0) {
            [void]$findings.Add((New-Finding -Severity $sevDupCart -RuleId 'DUP_CARTRIDGE_SN' -Message 'Duplicate Cartridge S/N detected.' -Rows $dupCartRows.ToArray() -Evidence (($dupCartridge | ForEach-Object { "$($_.Name) x$($_.Count)" }) -join '; ')))
        }

        $dupTable = New-Object System.Collections.Generic.List[object]
        foreach ($grp in $dupSample) {
            [void]$dupTable.Add([pscustomobject]@{
                Type = 'SampleID'; Key=$grp.Name; Count=$grp.Count
                ExampleRowIndexes = (($grp.Group | Select-Object -First 3 -ExpandProperty RowIndex) -join ', ')
            })
        }
        foreach ($grp in $dupCartridge) {
            [void]$dupTable.Add([pscustomobject]@{
                Type = 'CartridgeSN'; Key=$grp.Name; Count=$grp.Count
                ExampleRowIndexes = (($grp.Group | Select-Object -First 3 -ExpandProperty RowIndex) -join ', ')
            })
        }
        $result.DuplicatesTable = $dupTable.ToArray()

        # Status checks
        $failStatuses = @('error','invalid','aborted','incomplete')
        $passStatuses = @('done')
        try { if ($RuleBank -and $RuleBank.Global -and $RuleBank.Global.Status -and $RuleBank.Global.Status.FailStatuses) { $failStatuses = @($RuleBank.Global.Status.FailStatuses | ForEach-Object { $_.ToString().ToLower() }) } } catch {}
        try { if ($RuleBank -and $RuleBank.Global -and $RuleBank.Global.Status -and $RuleBank.Global.Status.PassStatuses) { $passStatuses = @($RuleBank.Global.Status.PassStatuses | ForEach-Object { $_.ToString().ToLower() }) } } catch {}

        $rowsFailStatus = @($rows | Where-Object { $_.Status -and ($failStatuses -contains ($_.Status.ToLower())) })
        if ($rowsFailStatus.Count -gt 0) {
            foreach ($r in $rowsFailStatus) { Add-FailureTag -Row $r -RuleId 'STATUS_FAILED' -Severity 'Error' }
            [void]$findings.Add((New-Finding -Severity 'Error' -RuleId 'STATUS_FAILED' -Message 'Error/Invalid/Failed status present.' -Rows $rowsFailStatus))
        }

        $rowsUnknownStatus = @($rows | Where-Object {
            $_.Status -and ($passStatuses -notcontains ($_.Status.ToLower())) -and ($failStatuses -notcontains ($_.Status.ToLower())) -and ($_.Status -notmatch '(?i)in\s*progress')
        })
        if ($rowsUnknownStatus.Count -gt 0) {
            foreach ($r in $rowsUnknownStatus) { Add-FailureTag -Row $r -RuleId 'STATUS_UNEXPECTED' -Severity 'Warn' }
            [void]$findings.Add((New-Finding -Severity 'Warn' -RuleId 'STATUS_UNEXPECTED' -Message 'Unexpected Status value.' -Rows $rowsUnknownStatus))
        }

        # Error codes
        $knownCodes = @()
        try {
            if ($RuleBank -and $RuleBank.ErrorBank -and $RuleBank.ErrorBank.Codes -and $RuleBank.ErrorBank.Codes.Keys.Count -gt 0) {
                $knownCodes = @($RuleBank.ErrorBank.Codes.Keys | ForEach-Object { $_.ToString() })
            }
        } catch {}

        $errorCodeCounts = @{}
        $rowsWithError = @($rows | Where-Object { $_.ErrorCode })
        foreach ($r in $rowsWithError) {
            Add-FailureTag -Row $r -RuleId 'ERROR_CODE_PRESENT' -Severity 'Warn'
            if (-not $errorCodeCounts.ContainsKey($r.ErrorCode)) { $errorCodeCounts[$r.ErrorCode] = 0 }
            $errorCodeCounts[$r.ErrorCode]++
            if ($knownCodes.Count -gt 0 -and ($knownCodes -notcontains $r.ErrorCode)) {
                Add-FailureTag -Row $r -RuleId 'UNKNOWN_ERROR_CODE' -Severity 'Warn'
            }
        }
        $result.UniqueErrorCodes = $errorCodeCounts.Count

        $rowsErrorTextOnly = @($rows | Where-Object { $_.ErrorRaw -and (-not $_.ErrorCode) })
        if ($rowsErrorTextOnly.Count -gt 0) {
            foreach ($r in $rowsErrorTextOnly) { Add-FailureTag -Row $r -RuleId 'ERROR_TEXT_PRESENT' -Severity 'Warn' }
        }

        $rowsFailNoCode = @($rows | Where-Object { ($_.Status -match '(?i)(error|invalid)') -and (-not $_.ErrorCode) })
        if ($rowsFailNoCode.Count -gt 0) {
            foreach ($r in $rowsFailNoCode) { Add-FailureTag -Row $r -RuleId 'STATUS_WITHOUT_ERRORCODE' -Severity 'Warn' }
            [void]$findings.Add((New-Finding -Severity 'Warn' -RuleId 'STATUS_WITHOUT_ERRORCODE' -Message 'Error/Invalid status without error code.' -Rows $rowsFailNoCode))
        }

        # Max pressure thresholds
        $pressureWarnRows = New-Object System.Collections.Generic.List[object]
        $pressureFailRows = New-Object System.Collections.Generic.List[object]
        foreach ($r in ($rows | Where-Object { $_.MaxPressure -ne $null })) {
            if ([double]$r.MaxPressure -ge $failThreshold) {
                $sev = 'Error'
                try { if ($RuleBank -and $RuleBank.Global -and $RuleBank.Global.MaxPressure -and $RuleBank.Global.MaxPressure.Severity) { $sev = $RuleBank.Global.MaxPressure.Severity } } catch {}
                Add-FailureTag -Row $r -RuleId 'MAX_PRESSURE_FAIL' -Severity $sev
                [void]$pressureFailRows.Add($r)
                $result.PressureStats.OverFail++
                $result.PressureStats.OverWarn++
            } elseif ([double]$r.MaxPressure -ge $warnThreshold) {
                Add-FailureTag -Row $r -RuleId 'MAX_PRESSURE_WARN' -Severity 'Warn'
                [void]$pressureWarnRows.Add($r)
                $result.PressureStats.OverWarn++
            }
        }
        if ($pressureFailRows.Count -gt 0) {
            [void]$findings.Add((New-Finding -Severity 'Error' -RuleId 'MAX_PRESSURE_FAIL' -Message ("Max Pressure >= {0}" -f $failThreshold) -Rows $pressureFailRows.ToArray()))
        }
        if ($pressureWarnRows.Count -gt 0) {
            [void]$findings.Add((New-Finding -Severity 'Warn' -RuleId 'MAX_PRESSURE_WARN' -Message ("Max Pressure >= {0}" -f $warnThreshold) -Rows $pressureWarnRows.ToArray()))
        }

        # Error summary table
        $errDetails = New-Object System.Collections.Generic.List[object]
        foreach ($code in $errorCodeCounts.Keys) {
            $defKey = if ($knownCodes -contains $code) { $code } elseif ($knownCodes -contains 'default') { 'default' } else { $null }
            $def = $null
            if ($defKey) {
                try { $def = $RuleBank.ErrorBank.Codes[$defKey] } catch {}
            }
            $exampleRow = ($rows | Where-Object { $_.ErrorCode -eq $code } | Select-Object -First 1)
            [void]$errDetails.Add([pscustomobject]@{
                ErrorCode       = $code
                Name            = if ($def -and $def.Name) { $def.Name } else { '' }
                Group           = if ($def -and $def.Group) { $def.Group } else { '' }
                Classification  = if ($def -and $def.Classification) { $def.Classification } else { '' }
                GeneratesRetest = if ($def -and $def.GeneratesRetest -ne $null) { [bool]$def.GeneratesRetest } else { $false }
                Count           = [int]$errorCodeCounts[$code]
                ExampleSampleID = if ($exampleRow) { $exampleRow.SampleID } else { '' }
            })
        }
        if ($errDetails.Count -gt 0) {
            $result.ErrorSummary    = @($errDetails | Sort-Object -Property Count -Descending, ErrorCode)
            $result.ErrorCodesTable = $result.ErrorSummary
        }

        # Unknown error codes warning
        if ($knownCodes.Count -gt 0) {
            $rowsUnknownErr = @($rows | Where-Object { $_.ErrorCode -and ($knownCodes -notcontains $_.ErrorCode) })
            if ($rowsUnknownErr.Count -gt 0) {
                foreach ($r in $rowsUnknownErr) { Add-FailureTag -Row $r -RuleId 'UNKNOWN_ERROR_CODE' -Severity 'Warn' }
                $unknownCodesTxt = (($rowsUnknownErr | Select-Object -ExpandProperty ErrorCode) | Select-Object -Unique) -join ', '
                [void]$findings.Add((New-Finding -Severity 'Warn' -RuleId 'UNKNOWN_ERROR_CODE' -Message 'Error code not in known list.' -Rows $rowsUnknownErr -Evidence ("Unknown codes: {0}" -f $unknownCodesTxt)))
            }
        }

        # Affected tests (Warn/Error)
        $affected = @($rows | Where-Object { $_.FailureTags -and $_.FailureTags.Count -gt 0 -and (Get-SeverityRank $_.Severity) -ge (Get-SeverityRank 'Warn') })
        $affected = @($affected | Sort-Object @{Expression={ -1*(Get-SeverityRank $_.Severity) }}, @{Expression={$_.ErrorCode}}, @{Expression={$_.SampleID}}, @{Expression={$_.RowIndex}})
        $result.AffectedTests = $affected
        $result.ErrorRowCount = $affected.Count

        if ($findings.Count -eq 0) {
            [void]$findings.Add((New-Finding -Severity 'Info' -RuleId 'NO_FINDINGS' -Message 'No rule findings detected.' -Rows $null))
        }

        $sortedFindings = @($findings | Sort-Object @{Expression={ -1 * (Get-SeverityRank $_.Severity) }}, @{Expression={ $_.Count }; Descending=$true})
        $result.Findings = $sortedFindings

        # SeverityCounts (number of findings per severity)
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
            Severity        = 'Error'
            RuleId          = 'ENGINE_EXCEPTION'
            Message         = 'RuleEngine exception'
            Count           = 0
            Evidence        = $ev
            Classification  = 'Engine'
            GeneratesRetest = $false
            Example         = ''
        })
        $result.Findings = $findings.ToArray()
        $result.OverallSeverity = 'Error'
        $result.OverallStatus = 'FAIL'
        $result.Exception = $_.Exception
    }

    return [pscustomobject]$result
}

# -----------------------------
# Writer (Information2)
# -----------------------------
function Write-Information2Sheet {
    param(
        [object]$Worksheet,
        [pscustomobject]$Context,
        [pscustomobject]$Evaluation,
        [string]$CsvPath,
        [string]$ScriptVersion
    )

    if (-not $Worksheet) { return }

    try {
        # Default guards
        if (-not $Evaluation) {
            $Evaluation = [pscustomobject]@{
                OverallStatus          = 'FAIL'
                OverallSeverity        = 'Error'
                Findings               = @()
                AffectedTests          = @()
                AffectedTestsTruncated = 0
                ErrorSummary           = @()
                PressureStats          = [ordered]@{ Max=$null; Avg=$null; OverWarn=0; OverFail=0; WarnThreshold=$null; FailThreshold=$null }
                DuplicatesTable        = @()
                Debug                  = @()
                UniqueErrorCodes       = 0
                IdentityFlags          = [ordered]@{}
                SeverityCounts         = @{ Error=0; Warn=0; Info=0 }
                BaselineExpected       = 0
                BaselineDelta          = 0
                ErrorRowCount          = 0
            }
        }
        if (-not $Context) {
            $Context = [pscustomobject]@{
                TotalTests           = 0
                StatusCounts         = @{}
                AssayRaw             = ''
                AssayCanonical       = ''
                AssayVersion         = ''
                ReagentLotIds        = @()
                WorkCenters          = @()
                MaxPressureMax       = $null
                MaxPressureAvg       = $null
                UniqueSampleIds      = 0
                UniqueCartridgeSN    = 0
                DuplicateCounts      = [ordered]@{ SampleId=0; CartridgeSN=0 }
                AssayCanonicalSource = ''
                ParseErrors          = @()
            }
        }

        # Hard reset (removes spök-CF + dropdown validations)
        try { $Worksheet.Cells.Clear() } catch {}
        try { $Worksheet.ConditionalFormatting.Clear() } catch {}
        try { $Worksheet.DataValidations.Clear() } catch {}
        try { $Worksheet.Tables.Clear() } catch {}

        $ps = if ($Evaluation.PressureStats) { $Evaluation.PressureStats } else { [ordered]@{ Max=$null; Avg=$null; OverWarn=0; OverFail=0; WarnThreshold=$null; FailThreshold=$null } }
        $sevCounts = if ($Evaluation.SeverityCounts) { $Evaluation.SeverityCounts } else { @{ Error=0; Warn=0; Info=0 } }

        $statusCounts = @{ Done=0; Error=0; Invalid=0; Aborted=0; Incomplete=0; Other=0 }
        foreach ($k in @($statusCounts.Keys)) {
            try { if ($Context.StatusCounts -and (Test-MapHasKey $Context.StatusCounts $k)) { $statusCounts[$k] = [int]$Context.StatusCounts[$k] } } catch {}
        }

        $r = 1
        $Worksheet.Cells["A$r"].Value = "Information2 – QC Rule Summary"
        $Worksheet.Cells["A$r:O$r"].Merge = $true
        $Worksheet.Cells["A$r"].Style.Font.Bold = $true
        $Worksheet.Cells["A$r"].Style.Font.Size = 16
        $r++

        $Worksheet.Cells["A$r"].Value = "CSV"
        $Worksheet.Cells["A$r"].Style.Font.Bold = $true
        $Worksheet.Cells["B$r"].Value = $(if ($CsvPath) { Split-Path $CsvPath -Leaf } else { '—' })
        $Worksheet.Cells["C$r"].Value = "Generated"
        $Worksheet.Cells["D$r"].Value = (Get-Date).ToString('yyyy-MM-dd HH:mm')
        if ($ScriptVersion) { $Worksheet.Cells["E$r"].Value = "Version: $ScriptVersion" }
        $r++

        $Worksheet.Cells["A$r"].Value = "Assay (canonical)"
        $Worksheet.Cells["A$r"].Style.Font.Bold = $true
        $Worksheet.Cells["B$r"].Value = $Context.AssayCanonical
        $Worksheet.Cells["C$r"].Value = "Assay Version"
        $Worksheet.Cells["C$r"].Style.Font.Bold = $true
        $Worksheet.Cells["D$r"].Value = $Context.AssayVersion
        $Worksheet.Cells["E$r"].Value = "Lot"
        $Worksheet.Cells["E$r"].Style.Font.Bold = $true
        $Worksheet.Cells["F$r"].Value = ($Context.ReagentLotIds -join ', ')
        $Worksheet.Cells["G$r"].Value = "WorkCenter"
        $Worksheet.Cells["G$r"].Style.Font.Bold = $true
        $Worksheet.Cells["H$r"].Value = ($Context.WorkCenters -join ', ')
        $Worksheet.Cells["I$r"].Value = "Match source"
        $Worksheet.Cells["I$r"].Style.Font.Bold = $true
        $Worksheet.Cells["J$r"].Value = $Context.AssayCanonicalSource
        $r++

        $Worksheet.Cells["A$r"].Value = "Row count"
        $Worksheet.Cells["A$r"].Style.Font.Bold = $true
        $Worksheet.Cells["B$r"].Value = [int]$Context.TotalTests
        $Worksheet.Cells["C$r"].Value = "Unique SampleID"
        $Worksheet.Cells["C$r"].Style.Font.Bold = $true
        $Worksheet.Cells["D$r"].Value = [int]$Context.UniqueSampleIds
        $Worksheet.Cells["E$r"].Value = "Unique CartridgeSN"
        $Worksheet.Cells["E$r"].Style.Font.Bold = $true
        $Worksheet.Cells["F$r"].Value = [int]$Context.UniqueCartridgeSN
        $r += 2

        $Worksheet.Cells["A$r"].Value = "KPI / Summary"
        $Worksheet.Cells["A$r"].Style.Font.Bold = $true
        $Worksheet.Cells["A$r"].Style.Font.Size = 12
        $r++

        $Worksheet.Cells["A$r"].Value = "Findings: ERROR / WARN / INFO"
        $Worksheet.Cells["B$r"].Value = ("{0} / {1} / {2}" -f $sevCounts.Error, $sevCounts.Warn, $sevCounts.Info)
        $r++
        $Worksheet.Cells["A$r"].Value = "Error rows (affected tests)"
        $Worksheet.Cells["B$r"].Value = [int]$Evaluation.ErrorRowCount
        $r++
        $Worksheet.Cells["A$r"].Value = "Distinct error codes"
        $Worksheet.Cells["B$r"].Value = [int]$Evaluation.UniqueErrorCodes
        $r++
        $Worksheet.Cells["A$r"].Value = "Baseline (expected/actual/delta)"
        $Worksheet.Cells["B$r"].Value = $Evaluation.BaselineExpected
        $Worksheet.Cells["C$r"].Value = $Context.TotalTests
        $Worksheet.Cells["D$r"].Value = $Evaluation.BaselineDelta
        $r++
        $Worksheet.Cells["A$r"].Value = "Max Pressure (max / avg)"
        $Worksheet.Cells["B$r"].Value = if ($ps.Max -ne $null) { [Math]::Round([double]$ps.Max,2) } else { '—' }
        $Worksheet.Cells["C$r"].Value = if ($ps.Avg -ne $null) { [Math]::Round([double]$ps.Avg,2) } else { '—' }
        $r++
        $Worksheet.Cells["A$r"].Value = "Max Pressure >= Warn / Fail"
        $Worksheet.Cells["B$r"].Value = [int]$ps.OverWarn
        $Worksheet.Cells["C$r"].Value = [int]$ps.OverFail
        $r++
        $Worksheet.Cells["A$r"].Value = "Status counts (Done/Error/Invalid/Aborted/Incomplete)"
        $Worksheet.Cells["B$r"].Value = ("{0}/{1}/{2}/{3}/{4}" -f $statusCounts.Done, $statusCounts.Error, $statusCounts.Invalid, $statusCounts.Aborted, $statusCounts.Incomplete)
        $r += 2

        # Findings
        $Worksheet.Cells["A$r"].Value = "Findings"
        $Worksheet.Cells["A$r"].Style.Font.Bold = $true
        $Worksheet.Cells["A$r"].Style.Font.Size = 12
        $r++

        $Worksheet.Cells["A$r"].Value = "Severity"
        $Worksheet.Cells["B$r"].Value = "RuleId"
        $Worksheet.Cells["C$r"].Value = "Title"
        $Worksheet.Cells["D$r"].Value = "Message"
        $Worksheet.Cells["E$r"].Value = "Count"
        $Worksheet.Cells["F$r"].Value = "Example"
        $Worksheet.Cells["G$r"].Value = "Evidence"
        $Worksheet.Cells["A$r:G$r"].Style.Font.Bold = $true
        $Worksheet.Cells["A$r:G$r"].Style.WrapText = $true
        $r++

        $findStart = $r
        foreach ($f in $Evaluation.Findings) {
            $Worksheet.Cells["A$r"].Value = $f.Severity
            $Worksheet.Cells["B$r"].Value = $f.RuleId
            $Worksheet.Cells["C$r"].Value = $f.Title
            $Worksheet.Cells["D$r"].Value = $f.Message
            $Worksheet.Cells["E$r"].Value = [int]$f.Count
            $Worksheet.Cells["F$r"].Value = $f.Example
            $Worksheet.Cells["G$r"].Value = $f.Evidence
            $r++
        }
        $findEnd = $r - 1
        if ($findEnd -ge $findStart) {
            $Worksheet.Cells["A$findStart:G$findEnd"].AutoFilter = $true
            $Worksheet.Cells["C$findStart:G$findEnd"].Style.WrapText = $true

            # Conditional formatting (relative formulas)
            $addr = "A$findStart:A$findEnd"
            $cfErr = $Worksheet.ConditionalFormatting.AddExpression($addr)
            $cfErr.Formula = 'ISNUMBER(SEARCH("Error",A1))'
            $cfErr.Style.Fill.PatternType = 'Solid'
            $cfErr.Style.Fill.BackgroundColor.Color = [System.Drawing.ColorTranslator]::FromHtml('#ffc7ce')

            $cfWarn = $Worksheet.ConditionalFormatting.AddExpression($addr)
            $cfWarn.Formula = 'ISNUMBER(SEARCH("Warn",A1))'
            $cfWarn.Style.Fill.PatternType = 'Solid'
            $cfWarn.Style.Fill.BackgroundColor.Color = [System.Drawing.ColorTranslator]::FromHtml('#ffe699')
        }
        $r += 2

        # Error Summary
        $Worksheet.Cells["A$r"].Value = "Error Summary"
        $Worksheet.Cells["A$r"].Style.Font.Bold = $true
        $Worksheet.Cells["A$r"].Style.Font.Size = 12
        $r++
        $Worksheet.Cells["A$r"].Value = "Code"
        $Worksheet.Cells["B$r"].Value = "Name"
        $Worksheet.Cells["C$r"].Value = "Classification"
        $Worksheet.Cells["D$r"].Value = "Count"
        $Worksheet.Cells["E$r"].Value = "Example SampleID"
        $Worksheet.Cells["A$r:E$r"].Style.Font.Bold = $true
        $Worksheet.Cells["A$r:E$r"].Style.WrapText = $true
        $r++

        $errStart = $r
        foreach ($err in $Evaluation.ErrorSummary) {
            $Worksheet.Cells["A$r"].Value = $err.ErrorCode
            $Worksheet.Cells["B$r"].Value = $err.Name
            $Worksheet.Cells["C$r"].Value = if ($err.Classification) { $err.Classification } else { $err.Group }
            $Worksheet.Cells["D$r"].Value = [int]$err.Count
            $Worksheet.Cells["E$r"].Value = $err.ExampleSampleID
            $r++
        }
        $errEnd = $r - 1
        if ($errEnd -ge $errStart) {
            $Worksheet.Cells["A$errStart:E$errEnd"].AutoFilter = $true
            $Worksheet.Cells["A$errStart:E$errEnd"].Style.WrapText = $true
        }
        $r += 2

        # Affected Tests
        $Worksheet.Cells["A$r"].Value = "Affected Tests (Warn/Error)"
        $Worksheet.Cells["A$r"].Style.Font.Bold = $true
        $Worksheet.Cells["A$r"].Style.Font.Size = 12
        $r++
        $Worksheet.Cells["A$r"].Value = "Severity"
        $Worksheet.Cells["B$r"].Value = "RuleId"
        $Worksheet.Cells["C$r"].Value = "Sample ID"
        $Worksheet.Cells["D$r"].Value = "Cartridge S/N"
        $Worksheet.Cells["E$r"].Value = "Module S/N"
        $Worksheet.Cells["F$r"].Value = "Test Type"
        $Worksheet.Cells["G$r"].Value = "Status"
        $Worksheet.Cells["H$r"].Value = "Test Result"
        $Worksheet.Cells["I$r"].Value = "Max Pressure (PSI)"
        $Worksheet.Cells["J$r"].Value = "Error"
        $Worksheet.Cells["K$r"].Value = "ErrorCode"
        $Worksheet.Cells["L$r"].Value = "WorkCenter"
        $Worksheet.Cells["M$r"].Value = "Row#"
        $Worksheet.Cells["A$r:M$r"].Style.Font.Bold = $true
        $Worksheet.Cells["A$r:M$r"].Style.WrapText = $true

        $affStart = ++$r
        foreach ($row in $Evaluation.AffectedTests) {
            $Worksheet.Cells["A$r"].Value = $row.Severity
            $Worksheet.Cells["B$r"].Value = $row.PrimaryRule
            $Worksheet.Cells["C$r"].Value = $row.SampleID
            $Worksheet.Cells["D$r"].Value = $row.CartridgeSN
            $Worksheet.Cells["E$r"].Value = $row.ModuleSN
            $Worksheet.Cells["F$r"].Value = $row.TestType
            $Worksheet.Cells["G$r"].Value = $row.Status
            $Worksheet.Cells["H$r"].Value = $row.TestResult
            $Worksheet.Cells["I$r"].Value = if ($row.MaxPressure -ne $null) { [Math]::Round([double]$row.MaxPressure,2) } else { $null }
            $Worksheet.Cells["J$r"].Value = $row.ErrorRaw
            $Worksheet.Cells["K$r"].Value = $row.ErrorCode
            $Worksheet.Cells["L$r"].Value = $row.WorkCenter
            $Worksheet.Cells["M$r"].Value = [int]$row.RowIndex
            $r++
        }
        $affEnd = $r - 1
        if ($affEnd -ge $affStart) {
            $Worksheet.Cells["A$affStart:M$affEnd"].AutoFilter = $true
            $Worksheet.View.FreezePanes($affStart, 1)

            try {
                $addrA = "A$affStart:A$affEnd"
                $addrG = "G$affStart:G$affEnd"

                $cfAffErr = $Worksheet.ConditionalFormatting.AddExpression($addrA)
                $cfAffErr.Formula = 'ISNUMBER(SEARCH("Error",A1))'
                $cfAffErr.Style.Fill.PatternType = 'Solid'
                $cfAffErr.Style.Fill.BackgroundColor.Color = [System.Drawing.ColorTranslator]::FromHtml('#ffc7ce')

                $cfAffWarn = $Worksheet.ConditionalFormatting.AddExpression($addrA)
                $cfAffWarn.Formula = 'ISNUMBER(SEARCH("Warn",A1))'
                $cfAffWarn.Style.Fill.PatternType = 'Solid'
                $cfAffWarn.Style.Fill.BackgroundColor.Color = [System.Drawing.ColorTranslator]::FromHtml('#ffe699')

                $cfDone = $Worksheet.ConditionalFormatting.AddExpression($addrG)
                $cfDone.Formula = 'ISNUMBER(SEARCH("Done",G1))'
                $cfDone.Style.Fill.PatternType = 'Solid'
                $cfDone.Style.Fill.BackgroundColor.Color = [System.Drawing.ColorTranslator]::FromHtml('#c6efce')
            } catch {}

            if ($ps.FailThreshold -ne $null) {
                $cfMpFail = $Worksheet.ConditionalFormatting.AddGreaterThan("I$affStart:I$affEnd", [double]$ps.FailThreshold)
                $cfMpFail.Style.Fill.PatternType = 'Solid'
                $cfMpFail.Style.Fill.BackgroundColor.Color = [System.Drawing.ColorTranslator]::FromHtml('#ffc7ce')
            }
            if ($ps.WarnThreshold -ne $null) {
                $cfMpWarn = $Worksheet.ConditionalFormatting.AddGreaterThan("I$affStart:I$affEnd", [double]$ps.WarnThreshold)
                $cfMpWarn.Style.Fill.PatternType = 'Solid'
                $cfMpWarn.Style.Fill.BackgroundColor.Color = [System.Drawing.ColorTranslator]::FromHtml('#ffe699')
            }
        }

        $Worksheet.Cells.Style.Font.Name = 'Arial'
        $Worksheet.Cells.Style.Font.Size = 10
        try { if ($Worksheet.Dimension) { $Worksheet.Cells[$Worksheet.Dimension.Address].AutoFitColumns() } } catch {}

    } catch {
        try {
            $Worksheet.Cells.Clear()
            $Worksheet.Cells["A1"].Value = "Information2 – RuleEngine error"
            $Worksheet.Cells["A1"].Style.Font.Bold = $true
            $Worksheet.Cells["A2"].Value = $_.Exception.Message
            $Worksheet.Cells["A2"].Style.WrapText = $true
        } catch {}
    }
}
