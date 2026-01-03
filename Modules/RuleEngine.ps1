#requires -Version 5.1


#region REGION: CSV
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

# -----------------------------
# Sample ID (base + underscore-bag derived fields)
# -----------------------------
function Get-SampleIdBase {
    param([string]$SampleId)
    if ([string]::IsNullOrWhiteSpace($SampleId)) { return '' }
    # Some exports append extra info after comma; keep the first token as canonical SampleID
    $first = ($SampleId -split '\s*,\s*' | Select-Object -First 1)
    return ($first + '').Trim()
}

function Parse-SampleIdDerived {
    [CmdletBinding()]
    param([string]$SampleId)

    $d = [ordered]@{
        IsParsed          = $false
        Prefix            = $null
        Bag               = $null
        ControlIndex      = $null
        BagSampleRaw      = $null
        BagSampleDigits   = $null
        SampleNo2         = $null   # always 2-digit ("01")
        ReplacementLevel  = 0       # number of "A" (A/AA/AAA)
        HasDelaminationD  = $false
        HasX              = $false
        HasPlus           = $false
    }

    if ([string]::IsNullOrWhiteSpace($SampleId)) { return [pscustomobject]$d }

    $parts = ($SampleId + '').Split('_')
    if ($parts.Count -lt 4) { return [pscustomobject]$d }

    $d.Prefix       = $parts[0]
    $bagTxt         = $parts[1]
    $ctrlTxt        = $parts[2]
    $bagSample      = $parts[3]
    $d.BagSampleRaw = $bagSample

    $tmp = 0
    if ([int]::TryParse(($bagTxt + ''),  [ref]$tmp)) { $d.Bag = $tmp }
    if ([int]::TryParse(($ctrlTxt + ''), [ref]$tmp)) { $d.ControlIndex = $tmp }

    $digits = [regex]::Replace([string]$bagSample, '\D+', '')
    if ($digits) {
        $d.BagSampleDigits = $digits
        $d.SampleNo2       = $digits.PadLeft(2,'0')
    }

    # Replacement tokens: A / AA / AAA (count occurrences)
    try {
        $d.ReplacementLevel = ([regex]::Matches(($bagSample + ''), 'A')).Count
    } catch { $d.ReplacementLevel = 0 }

    if (($bagSample + '') -match 'D')   { $d.HasDelaminationD = $true }
    if (($bagSample + '') -match 'X')   { $d.HasX             = $true }
    if (($bagSample + '') -match '\+')  { $d.HasPlus          = $true }

    $d.IsParsed = (-not [string]::IsNullOrWhiteSpace($d.Prefix)) -and
                  ($d.Bag -ne $null) -and
                  ($d.ControlIndex -ne $null) -and
                  (-not [string]::IsNullOrWhiteSpace($d.SampleNo2))

    return [pscustomobject]$d
}

function Normalize-TestResult {
    param([string]$Text)
    if ([string]::IsNullOrWhiteSpace($Text)) { return '' }
    $t = ($Text + '').Trim()
    # Normalize semicolon spacing, collapse whitespace
    $t = [regex]::Replace($t, '\s*;\s*', '; ')
    $t = [regex]::Replace($t, '\s+', ' ').Trim()
    return $t
}

# -----------------------------
# RunIndex (single pass) – duplicates, bag-samples, error clustering
# -----------------------------
function New-RunIndex {
    [pscustomobject]([ordered]@{
        AssayCounts               = @{}
        AssayVersionCounts        = @{}
        ReagentLotCounts          = @{}

        ControlMaterialCounts     = @{}   # Prefix => count
        BagSamples                = @{}   # Bag(int) => HashSet[string] ("01","02"...)

        SampleIdRows              = @{}   # SampleID => List[int] rowIndex
        CartridgeSnRows           = @{}   # CartridgeSN => List[int] rowIndex

        ReplacementRows           = New-Object 'System.Collections.Generic.List[int]'
        DelaminationRows          = New-Object 'System.Collections.Generic.List[int]'

        ErrorByInstrumentModule   = @{}   # "Instrument|Module" => List[int] rowIndex (any error)
    })
}

function Add-RunIndexCount {
    param([hashtable]$Map, [string]$Key)
    if ($null -eq $Map) { return }
    if ([string]::IsNullOrWhiteSpace($Key)) { $Key = '<BLANK>' }
    if ($Map.ContainsKey($Key)) { $Map[$Key] = [int]$Map[$Key] + 1 }
    else { $Map[$Key] = 1 }
}

function Add-RunIndexRowRef {
    param([hashtable]$Map, [string]$Key, [int]$RowIndex)
    if ($null -eq $Map) { return }
    if ([string]::IsNullOrWhiteSpace($Key)) { return }
    if (-not $Map.ContainsKey($Key)) { $Map[$Key] = New-Object 'System.Collections.Generic.List[int]' }
    $null = $Map[$Key].Add($RowIndex)
}

function Update-RunIndexFromRow {
    param([pscustomobject]$Idx, [pscustomobject]$Row)

    if (-not $Idx -or -not $Row) { return }

    Add-RunIndexCount  $Idx.AssayCounts        $Row.Assay
    Add-RunIndexCount  $Idx.AssayVersionCounts $Row.AssayVersion
    Add-RunIndexCount  $Idx.ReagentLotCounts   $Row.ReagentLotId

    Add-RunIndexRowRef $Idx.SampleIdRows       $Row.SampleID     $Row.RowIndex
    Add-RunIndexRowRef $Idx.CartridgeSnRows    $Row.CartridgeSN  $Row.RowIndex

    $d = $Row.SampleIdDerived
    if ($d -and $d.IsParsed) {

        Add-RunIndexCount $Idx.ControlMaterialCounts ($d.Prefix + '')

        if ($d.Bag -ne $null -and $d.SampleNo2) {
            if (-not $Idx.BagSamples.ContainsKey($d.Bag)) {
                $Idx.BagSamples[$d.Bag] = New-Object 'System.Collections.Generic.HashSet[string]'
            }
            $null = $Idx.BagSamples[$d.Bag].Add(($d.SampleNo2 + ''))
        }

        if ($d.ReplacementLevel -gt 0)  { $null = $Idx.ReplacementRows.Add($Row.RowIndex) }
        if ($d.HasDelaminationD)        { $null = $Idx.DelaminationRows.Add($Row.RowIndex) }
    }

    # Any termination/error clustering per Instrument+Module
    $hasAnyError = $false
    try { if ($Row.ErrorCode) { $hasAnyError = $true } } catch {}
    if (-not $hasAnyError) {
        try {
            if ($Row.Status -and ($Row.Status -notmatch '^(?i)Done$')) { $hasAnyError = $true }
            elseif ($Row.TestResult -and ($Row.TestResult -match '^(?i)(ERROR|INVALID|NO RESULT)$')) { $hasAnyError = $true }
        } catch {}
    }

    if ($hasAnyError -and $Row.InstrumentSN -and $Row.ModuleSN) {
        $key = ("{0}|{1}" -f $Row.InstrumentSN, $Row.ModuleSN)
        Add-RunIndexRowRef $Idx.ErrorByInstrumentModule $key $Row.RowIndex
    }
}

function Get-DuplicateEvidence {
    param([hashtable]$Map, [int]$MaxItems = 20)

    $dupKeys = @($Map.Keys | Where-Object { $Map[$_].Count -gt 1 })
    $lines = New-Object 'System.Collections.Generic.List[string]'

    $shown = 0
    foreach ($k in $dupKeys) {
        $rows = @($Map[$k] | Sort-Object)
        $lines.Add(("Lines {0} for {1}" -f ($rows -join ', '), $k))
        $shown++
        if ($shown -ge $MaxItems) { break }
    }

    $extra = $dupKeys.Count - $shown
    if ($extra -gt 0) { $lines.Add(("… +{0} more duplicate values" -f $extra)) }

    [pscustomobject]@{
        DuplicateValueCount = $dupKeys.Count
        Evidence            = ($lines -join "`r`n")
    }
}

function Get-MissingBagSamples {
    param(
        [pscustomobject]$Idx,
        [int]$BagMin = 0,
        [int]$BagMax = 10,
        [int]$Bag0SampleMax = 10,
        [int]$OtherBagSampleMax = 20
    )

    $missing = New-Object 'System.Collections.Generic.List[string]'

    if (-not $Idx -or -not $Idx.BagSamples) { return $missing }

    for ($bag = $BagMin; $bag -le $BagMax; $bag++) {
        $max = if ($bag -eq 0) { $Bag0SampleMax } else { $OtherBagSampleMax }

        for ($j = 1; $j -le $max; $j++) {
            $s = $j.ToString().PadLeft(2,'0')
            if (-not $Idx.BagSamples.ContainsKey($bag) -or -not $Idx.BagSamples[$bag].Contains($s)) {
                $missing.Add(("Bag {0} is missing sample {1}" -f $bag, $s))
            }
        }
    }

    return $missing
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
        ReagentLotIds       = @()
        AssayNormalized     = $null
        AssayCanonical      = $null
        MatchSource         = $null
        ProfileMatched      = $null
        ProfileMode         = $null
        Profile             = $null
        StatusCounts        = @{}
        ErrorCodeCounts     = @{}         # Hashtable
        DataRows            = @()
        TotalTests          = 0
        MaxPressureMax      = $null
        MaxPressureCount    = 0
        ColumnMissing       = @()
        Debug               = @()
        ParseErrors         = New-Object System.Collections.Generic.List[string]
        DuplicateCounts     = [ordered]@{}
        DuplicateEvidence   = [ordered]@{}
        RunIndex            = $null
        UniqueSampleIds     = 0
        UniqueCartridgeSN   = 0
        AssayCanonicalSource= $null
        TestTypeCounts      = @{}
    }

    if (-not $Bundle -or -not $Bundle.Lines) {
        $ctx.Debug = @([pscustomobject]@{ Name='Context'; Value='Bundle or Lines missing' })
        return [pscustomobject]$ctx
    }

    # Index lookup

# Index lookup
$idxAssay       = Get-HeaderIndexValue -HeaderIndex $ctx.HeaderIndex -Keys @('assay')
$idxAssayVer    = Get-HeaderIndexValue -HeaderIndex $ctx.HeaderIndex -Keys @('assay version')
$idxSample      = Get-HeaderIndexValue -HeaderIndex $ctx.HeaderIndex -Keys @('sample id','sampleid')
$idxCartridge   = Get-HeaderIndexValue -HeaderIndex $ctx.HeaderIndex -Keys @('cartridge s/n','cartridge sn','cartridge serial')
$idxReagentLot  = Get-HeaderIndexValue -HeaderIndex $ctx.HeaderIndex -Keys @('reagent lot id','reagent lot')
$idxTestType    = Get-HeaderIndexValue -HeaderIndex $ctx.HeaderIndex -Keys @('test type','testtype')
$idxInstrument  = Get-HeaderIndexValue -HeaderIndex $ctx.HeaderIndex -Keys @('instrument s/n','instrument sn')
$idxModule      = Get-HeaderIndexValue -HeaderIndex $ctx.HeaderIndex -Keys @('module s/n','module sn','module serial','module')
$idxSwVersion   = Get-HeaderIndexValue -HeaderIndex $ctx.HeaderIndex -Keys @('s/w version','sw version','software version')
$idxStartTime   = Get-HeaderIndexValue -HeaderIndex $ctx.HeaderIndex -Keys @('start time','starttime')
$idxStatus      = Get-HeaderIndexValue -HeaderIndex $ctx.HeaderIndex -Keys @('status')
$idxTestResult  = Get-HeaderIndexValue -HeaderIndex $ctx.HeaderIndex -Keys @('test result','result')
$idxError       = Get-HeaderIndexValue -HeaderIndex $ctx.HeaderIndex -Keys @('error','error code','error codes')
$idxMaxPressure = Get-HeaderIndexValue -HeaderIndex $ctx.HeaderIndex -Keys @('max pressure (psi)','max pressure psi','max pressure')
$idxWorkCenter  = Get-HeaderIndexValue -HeaderIndex $ctx.HeaderIndex -Keys @('work center','workcenter')

$ctx.Indices = [ordered]@{
    Assay       = $idxAssay
    AssayVer    = $idxAssayVer
    SampleId    = $idxSample
    CartridgeSn = $idxCartridge
    ReagentLot  = $idxReagentLot
    TestType    = $idxTestType
    InstrumentSn= $idxInstrument
    ModuleSn    = $idxModule
    SwVersion   = $idxSwVersion
    StartTime   = $idxStartTime
    Status      = $idxStatus
    TestResult  = $idxTestResult
    Error       = $idxError
    MaxPressure = $idxMaxPressure
    WorkCenter  = $idxWorkCenter
}

    # Required vs optional columns:
    # - Required: core identity + status/result/error/pressure
    # - Optional: instrument/module SW/start time/work center (engine will degrade gracefully)
    $requiredCols = @('Assay','AssayVer','SampleId','CartridgeSn','ReagentLot','TestType','ModuleSn','Status','TestResult','Error','MaxPressure')
    foreach ($name in $requiredCols) {
        try {
            if (-not $ctx.Indices.Contains($name) -or $ctx.Indices[$name] -lt 0) { $ctx.ColumnMissing += $name }
        } catch {
            $ctx.ColumnMissing += $name
        }
    }
    if ($ctx.ColumnMissing.Count -gt 0) {
        [void]$ctx.ParseErrors.Add("Missing required column(s): " + ($ctx.ColumnMissing -join ', '))
        return [pscustomobject]$ctx
    }

        # Parser selection (single source of truth: CsvBundle.Parse-CsvLine)
        if (-not (Get-Command Parse-CsvLine -ErrorAction SilentlyContinue)) {
            [void]$ctx.ParseErrors.Add("Parse-CsvLine not available. CsvBundle.ps1 must be loaded before RuleEngine.")
            $ctx.Debug = @([pscustomobject]@{ Name='CsvParser'; Value='Missing Parse-CsvLine (CsvBundle not loaded)' })
            return [pscustomobject]$ctx
        }
        $parse = { param($ln,$delim) (Parse-CsvLine -line $ln -delim $delim) }


    $lines    = $Bundle.Lines
    $startRow = $ctx.DataStartRowIndex
    if ($startRow -lt 0) { $startRow = [int]$ctx.HeaderRowIndex + 1 }
    if ($startRow -lt 0) { $startRow = 0 }


$reagentSet     = New-Object 'System.Collections.Generic.HashSet[string]'
$workCenterSet  = New-Object 'System.Collections.Generic.HashSet[string]'
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

$errorCounts = @{}  # Hashtable
$rows        = New-Object 'System.Collections.Generic.List[object]'
$runIdx      = New-RunIndex

$maxPressureMax   = $null
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
        $sampleRaw = if ($idxSample -ge 0 -and $cols.Count -gt $idxSample) { ($cols[$idxSample] + '').Trim() } else { '' }
        $sampleBase = Get-SampleIdBase $sampleRaw

        $testResultRaw = if ($idxTestResult -ge 0 -and $cols.Count -gt $idxTestResult) { ($cols[$idxTestResult] + '').Trim() } else { '' }

        $rowObj = [ordered]@{
            RowIndex       = ($r + 1)
            DataIndex      = ($r - $startRow + 1)

            Assay          = if ($idxAssay -ge 0 -and $cols.Count -gt $idxAssay) { ($cols[$idxAssay] + '').Trim() } else { '' }
            AssayVersion   = if ($idxAssayVer -ge 0 -and $cols.Count -gt $idxAssayVer) { ($cols[$idxAssayVer] + '').Trim() } else { '' }
            ReagentLotId   = if ($idxReagentLot -ge 0 -and $cols.Count -gt $idxReagentLot) { ($cols[$idxReagentLot] + '').Trim() } else { '' }

            SampleIDRaw    = $sampleRaw
            SampleID       = $sampleBase

            CartridgeSN    = if ($idxCartridge -ge 0 -and $cols.Count -gt $idxCartridge) { ($cols[$idxCartridge] + '').Trim() } else { '' }

            TestType       = if ($idxTestType -ge 0 -and $cols.Count -gt $idxTestType) { ($cols[$idxTestType] + '').Trim() } else { '' }

            InstrumentSN   = if ($idxInstrument -ge 0 -and $cols.Count -gt $idxInstrument) { ($cols[$idxInstrument] + '').Trim() } else { '' }
            ModuleSN       = if ($idxModule -ge 0 -and $cols.Count -gt $idxModule) { ($cols[$idxModule] + '').Trim() } else { '' }

            SwVersion      = if ($idxSwVersion -ge 0 -and $cols.Count -gt $idxSwVersion) { ($cols[$idxSwVersion] + '').Trim() } else { '' }
            StartTime      = if ($idxStartTime -ge 0 -and $cols.Count -gt $idxStartTime) { ($cols[$idxStartTime] + '').Trim() } else { '' }

            Status         = if ($idxStatus -ge 0 -and $cols.Count -gt $idxStatus) { ($cols[$idxStatus] + '').Trim() } else { '' }

            TestResultRaw  = $testResultRaw
            TestResult     = Normalize-TestResult $testResultRaw

            ErrorRaw       = if ($idxError -ge 0 -and $cols.Count -gt $idxError) { ($cols[$idxError] + '').Trim() } else { '' }
            ErrorCode      = $null

            MaxPressure    = $null
            WorkCenter     = if ($idxWorkCenter -ge 0 -and $cols.Count -gt $idxWorkCenter) { ($cols[$idxWorkCenter] + '').Trim() } else { '' }

            SampleIdDerived= $null

            FailureTags    = New-Object System.Collections.Generic.List[string]
            Notes          = ''
            Severity       = 'Info'

            # Legacy parsed fields (global pattern)
            ParsedRole     = ''
            ParsedId       = ''
            ParsedIdx      = ''
            PrimaryRule    = ''
        }

        # Normalize ReagentLotId: extract 5 digits
        if ($rowObj.ReagentLotId) {
            $m = [regex]::Match($rowObj.ReagentLotId, '(?<!\d)(\d{5})(?!\d)')
            if ($m.Success) { $rowObj.ReagentLotId = $m.Groups[1].Value }
            [void]$reagentSet.Add($rowObj.ReagentLotId)
        }
        if ($rowObj.WorkCenter) { [void]$workCenterSet.Add($rowObj.WorkCenter) }

        # Error code (numeric or alphanumeric)
        if ($idxError -ge 0) {
            $code = Parse-ErrorCode -Text $rowObj.ErrorRaw -RuleBank $RuleBank
            if ($code) {
                $rowObj.ErrorCode = $code
                Add-MapCount -Map $errorCounts -Key $code -Delta 1
            }
        }

        # Max pressure
        if ($idxMaxPressure -ge 0 -and $cols.Count -gt $idxMaxPressure) {
            $mp = Convert-ToDoubleOrNull (($cols[$idxMaxPressure] + '').Trim())
            if ($mp -ne $null) {
                $rowObj.MaxPressure = [double]$mp
                if ($maxPressureMax -eq $null -or $mp -gt $maxPressureMax) { $maxPressureMax = [double]$mp }
                $maxPressureCount++
            }
        }

        # Derived: underscore-bag format (assay gated later; parsing itself is safe)
        try { $rowObj.SampleIdDerived = Parse-SampleIdDerived -SampleId $rowObj.SampleID } catch {}

        # Legacy parsed fields (if any rules still use them)
        try {
            $parsed = Parse-SampleIdParts -SampleId $rowObj.SampleID -RuleBank $RuleBank
            if ($parsed -and $parsed.Success) {
                $rowObj.ParsedRole = $parsed.Role
                $rowObj.ParsedId   = $parsed.Id
                $rowObj.ParsedIdx  = $parsed.Idx
            }
        } catch {}

        # Counting
        if ($rowObj.TestType) {
            if (-not $testTypeCounts.ContainsKey($rowObj.TestType)) { $testTypeCounts[$rowObj.TestType] = 0 }
            $testTypeCounts[$rowObj.TestType]++
        }

        # Status counts
        $statusKey = if ($rowObj.Status) { $rowObj.Status } else { 'Other' }
        if ($statusCounts.ContainsKey($statusKey)) { $statusCounts[$statusKey]++ } else { $statusCounts['Other']++ }

        $rowPsc = [pscustomobject]$rowObj
        Update-RunIndexFromRow -Idx $runIdx -Row $rowPsc

        [void]$rows.Add($rowPsc)
    } catch {
        [void]$ctx.ParseErrors.Add(("Row {0}: Parse/convert failed: {1}" -f ($r+1), $_.Exception.Message))
    }
}

    $ctx.DataRows        = $rows.ToArray()
    $ctx.TotalTests      = $ctx.DataRows.Count
    $ctx.StatusCounts    = $statusCounts
    $ctx.ErrorCodeCounts = $errorCounts

    # HashSet -> string[]
    if ($reagentSet -is [System.Collections.IEnumerable]) {
        $reagentArray = New-Object string[] $reagentSet.Count
        try { $reagentSet.CopyTo($reagentArray) } catch { $reagentArray = foreach ($id in $reagentSet) { [string]$id } }
        $ctx.ReagentLotIds = $reagentArray
    } else { $ctx.ReagentLotIds = @() }

    if ($workCenterSet -is [System.Collections.IEnumerable]) {
        $workCenterArray = New-Object string[] $workCenterSet.Count
        try { $workCenterSet.CopyTo($workCenterArray) } catch { $workCenterArray = foreach ($wc in $workCenterSet) { [string]$wc } }
        $ctx.WorkCenters = $workCenterArray
    } else { $ctx.WorkCenters = @() }

    
# RunIndex + duplicates (SampleID/CartridgeSN)
$ctx.RunIndex = $runIdx

$ctx.UniqueSampleIds   = ($runIdx.SampleIdRows.Keys | Measure-Object).Count
$ctx.UniqueCartridgeSN = ($runIdx.CartridgeSnRows.Keys | Measure-Object).Count

$dupSampleEv = Get-DuplicateEvidence -Map $runIdx.SampleIdRows
$dupCartEv   = Get-DuplicateEvidence -Map $runIdx.CartridgeSnRows

$ctx.DuplicateCounts = [ordered]@{
    SampleId    = [int]$dupSampleEv.DuplicateValueCount
    CartridgeSN = [int]$dupCartEv.DuplicateValueCount
}
$ctx.DuplicateEvidence = [ordered]@{
    SampleId    = $dupSampleEv.Evidence
    CartridgeSN = $dupCartEv.Evidence
}

    $ctx.MaxPressureMax   = $maxPressureMax
    $ctx.MaxPressureCount = $maxPressureCount
    $ctx.TestTypeCounts   = $testTypeCounts

    # Sätt AssayRaw/AssayVersion till vanligaste förekomsten
    try {
        $assayGroups = $ctx.DataRows | Where-Object { $_.Assay } | Group-Object -Property Assay | Sort-Object -Property Count -Descending
        if ($assayGroups -and $assayGroups.Count -gt 0) { $ctx.AssayRaw = ($assayGroups[0].Name + '') }

        $verGroups = $ctx.DataRows | Where-Object { $_.AssayVersion } | Group-Object -Property AssayVersion | Sort-Object -Property Count -Descending
        if ($verGroups -and $verGroups.Count -gt 0) { $ctx.AssayVersion = ($verGroups[0].Name + '') }
    } catch {}

    # Canonical name resolution
    $resolved = Resolve-AssayCanonicalName -RawAssay $ctx.AssayRaw -AssayMap $AssayMap -RuleBank $RuleBank
    $ctx.AssayNormalized      = $resolved.Normalized
    $ctx.AssayCanonical       = $resolved.Canonical
    $ctx.MatchSource          = $resolved.MatchSource
    $ctx.AssayCanonicalSource = $resolved.MatchSource

    # Rule profile
    $profileResolved = Get-AssayRuleProfile -Canonical $ctx.AssayCanonical -RuleBank $RuleBank
    if ($profileResolved) {
        $ctx.Profile        = $profileResolved.Profile
        $ctx.ProfileMatched = $profileResolved.Key
        $ctx.ProfileMode    = $profileResolved.Mode
    }

    # Debug info (ingen Avg längre)
    $ctx.Debug = @(
        [pscustomobject]@{ Name='CsvPath';          Value=$ctx.CsvPath },
        [pscustomobject]@{ Name='Delimiter';        Value=$ctx.Delimiter },
        [pscustomobject]@{ Name='HeaderRowIndex';   Value=$ctx.HeaderRowIndex },
        [pscustomobject]@{ Name='DataStartRow';     Value=$ctx.DataStartRowIndex },
        [pscustomobject]@{ Name='RowCount';         Value=$ctx.TotalTests },
        [pscustomobject]@{ Name='ParseErrors';      Value=$ctx.ParseErrors.Count },
        [pscustomobject]@{ Name='AssayRaw';         Value=$ctx.AssayRaw },
        [pscustomobject]@{ Name='AssayCanonical';   Value=$ctx.AssayCanonical },
        [pscustomobject]@{ Name='MatchSource';      Value=$ctx.MatchSource },
        [pscustomobject]@{ Name='AssayVersion';     Value=$ctx.AssayVersion },
        [pscustomobject]@{ Name='ReagentLotIds';    Value=($ctx.ReagentLotIds -join ', ') },
        [pscustomobject]@{ Name='WorkCenters';      Value=($ctx.WorkCenters -join ', ') },
        [pscustomobject]@{ Name='MaxPressureMax';   Value=$ctx.MaxPressureMax }
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
        BucketCounts           = [ordered]@{}
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

        # Parse warnings som INFO
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

        # Sätt Max/Trösklar
        $result.PressureStats.Max = $Context.MaxPressureMax
        $result.PressureStats.WarnThreshold = $warnThreshold
        $result.PressureStats.FailThreshold = $failThreshold

        # OverWarn/OverFail (antal rader >= tröskel)
        try {
            $overWarn = @($rows | Where-Object { $_.MaxPressure -ne $null -and $_.MaxPressure -ge [double]$warnThreshold }).Count
            $overFail = @($rows | Where-Object { $_.MaxPressure -ne $null -and $_.MaxPressure -ge [double]$failThreshold }).Count
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

        # Identity checks (oförändrad)
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


# -----------------------------
# Global & profile-gated rules (single-pass semantics)
# -----------------------------

function Get-PolicySeverity {
    param([string]$Key, [string]$Default = 'Warn')
    try {
        if ($RuleBank -and $RuleBank.Global -and $RuleBank.Global.ValidationPolicy -and $RuleBank.Global.ValidationPolicy.ContainsKey($Key)) {
            $v = ($RuleBank.Global.ValidationPolicy[$Key] + '').Trim()
            if ($v) { return $v }
        }
    } catch {}
    return $Default
}

# Row lookup by RowIndex (for Evidence -> affected rows)
$rowByIndex = @{}
foreach ($rr in $rows) {
    try {
        if ($rr.RowIndex -ne $null) { $rowByIndex[[int]$rr.RowIndex] = $rr }
    } catch {}
}

# Duplicates: SampleID / CartridgeSN (flag + evidence)
if ($Context.RunIndex) {

    # Duplicate SampleID
    try {
        $dupSampleRows = New-Object 'System.Collections.Generic.List[object]'
        foreach ($k in $Context.RunIndex.SampleIdRows.Keys) {
            $lst = $Context.RunIndex.SampleIdRows[$k]
            if ($lst -and $lst.Count -gt 1) {
                foreach ($ri in $lst) { if ($rowByIndex.ContainsKey([int]$ri)) { $null = $dupSampleRows.Add($rowByIndex[[int]$ri]) } }
            }
        }
        if ($dupSampleRows.Count -gt 0) {
            $sev = Get-PolicySeverity -Key 'DUPLICATE_SAMPLE_ID' -Default 'Warn'
            foreach ($rr in $dupSampleRows) { Add-FailureTag -Row $rr -Tag 'DUPLICATE_SAMPLE_ID' -Severity $sev -PrimaryRuleId 'DUPLICATE_SAMPLE_ID' }
            [void]$findings.Add((New-Finding -Severity $sev -RuleId 'DUPLICATE_SAMPLE_ID' -Message 'Duplicate Sample ID detected' -Rows $dupSampleRows -Classification 'Duplicates' -Evidence $Context.DuplicateEvidence.SampleId))
        }
    } catch {}

    # Duplicate CartridgeSN
    try {
        $dupCartRows = New-Object 'System.Collections.Generic.List[object]'
        foreach ($k in $Context.RunIndex.CartridgeSnRows.Keys) {
            $lst = $Context.RunIndex.CartridgeSnRows[$k]
            if ($lst -and $lst.Count -gt 1) {
                foreach ($ri in $lst) { if ($rowByIndex.ContainsKey([int]$ri)) { $null = $dupCartRows.Add($rowByIndex[[int]$ri]) } }
            }
        }
        if ($dupCartRows.Count -gt 0) {
            $sev = Get-PolicySeverity -Key 'DUPLICATE_CARTRIDGE_SN' -Default 'Warn'
            foreach ($rr in $dupCartRows) { Add-FailureTag -Row $rr -Tag 'DUPLICATE_CARTRIDGE_SN' -Severity $sev -PrimaryRuleId 'DUPLICATE_CARTRIDGE_SN' }
            [void]$findings.Add((New-Finding -Severity $sev -RuleId 'DUPLICATE_CARTRIDGE_SN' -Message 'Duplicate Cartridge S/N detected' -Rows $dupCartRows -Classification 'Duplicates' -Evidence $Context.DuplicateEvidence.CartridgeSN))
        }
    } catch {}

    # Replacements (A/AA/AAA) – informational by default
    try {
        if ($Context.RunIndex.ReplacementRows -and $Context.RunIndex.ReplacementRows.Count -gt 0) {
            $repRows = New-Object 'System.Collections.Generic.List[object]'
            foreach ($ri in $Context.RunIndex.ReplacementRows) { if ($rowByIndex.ContainsKey([int]$ri)) { $null = $repRows.Add($rowByIndex[[int]$ri]) } }
            $sev = Get-PolicySeverity -Key 'REPLACEMENTS_PRESENT' -Default 'Info'
            foreach ($rr in $repRows) { Add-FailureTag -Row $rr -Tag 'REPLACEMENT_A' -Severity $sev -PrimaryRuleId 'REPLACEMENTS_PRESENT' }
            [void]$findings.Add((New-Finding -Severity $sev -RuleId 'REPLACEMENTS_PRESENT' -Message 'Replacements detected (A/AA/AAA)' -Rows $repRows -Classification 'Replacements'))
        }
    } catch {}

    # Delaminations (D) – informational by default
    try {
        if ($Context.RunIndex.DelaminationRows -and $Context.RunIndex.DelaminationRows.Count -gt 0) {
            $dRows = New-Object 'System.Collections.Generic.List[object]'
            foreach ($ri in $Context.RunIndex.DelaminationRows) { if ($rowByIndex.ContainsKey([int]$ri)) { $null = $dRows.Add($rowByIndex[[int]$ri]) } }
            $sev = Get-PolicySeverity -Key 'DELAMINATIONS_PRESENT' -Default 'Info'
            foreach ($rr in $dRows) { Add-FailureTag -Row $rr -Tag 'DELAMINATION_D' -Severity $sev -PrimaryRuleId 'DELAMINATIONS_PRESENT' }
            [void]$findings.Add((New-Finding -Severity $sev -RuleId 'DELAMINATIONS_PRESENT' -Message 'Delaminations detected (D)' -Rows $dRows -Classification 'Observation'))
        }
    } catch {}
}

# Profile-gated: underscore-bag parsing expectations, missing samples, wrong test type
$enableUnderscoreBag = $false
$enableWrongTestType = $false
try { if ($Context.Profile -and $Context.Profile.Enable) { $enableUnderscoreBag = [bool]$Context.Profile.Enable.UnderscoreBagSampleRules } } catch {}
try { if ($Context.Profile -and $Context.Profile.Enable) { $enableWrongTestType = [bool]$Context.Profile.Enable.WrongTestTypeByControlIndex } } catch {}

if ($enableUnderscoreBag -and $Context.RunIndex) {
    # SampleID parse failures (only when profile says underscore-bag should be used)
    try {
        $badRows = @($rows | Where-Object { -not $_.SampleIdDerived -or -not $_.SampleIdDerived.IsParsed })
        if ($badRows.Count -gt 0) {
            $sev = Get-PolicySeverity -Key 'SAMPLEID_PARSE_FAILED' -Default 'Warn'
            foreach ($rr in $badRows) { Add-FailureTag -Row $rr -Tag 'SAMPLEID_PARSE_FAILED' -Severity $sev -PrimaryRuleId 'SAMPLEID_PARSE_FAILED' }
            [void]$findings.Add((New-Finding -Severity $sev -RuleId 'SAMPLEID_PARSE_FAILED' -Message 'Sample ID format deviation (underscore-bag expected)' -Rows $badRows -Classification 'Sample ID or Test Type Deviation'))
        }
    } catch {}

    # Missing bag samples (bag 0 => 10, other bags => 20 by default)
    try {
        $bagMin = 0; $bagMax = 10; $bag0max = 10; $bagOtherMax = 20
        try {
            if ($RuleBank -and $RuleBank.Global -and $RuleBank.Global.SampleId -and $RuleBank.Global.SampleId.UnderscoreBag) {
                if ($RuleBank.Global.SampleId.UnderscoreBag.BagMin -ne $null) { $bagMin = [int]$RuleBank.Global.SampleId.UnderscoreBag.BagMin }
                if ($RuleBank.Global.SampleId.UnderscoreBag.BagMax -ne $null) { $bagMax = [int]$RuleBank.Global.SampleId.UnderscoreBag.BagMax }
                if ($RuleBank.Global.SampleId.UnderscoreBag.Bag0SampleMax -ne $null) { $bag0max = [int]$RuleBank.Global.SampleId.UnderscoreBag.Bag0SampleMax }
                if ($RuleBank.Global.SampleId.UnderscoreBag.OtherBagSampleMax -ne $null) { $bagOtherMax = [int]$RuleBank.Global.SampleId.UnderscoreBag.OtherBagSampleMax }
            }
        } catch {}

        $missing = Get-MissingBagSamples -Idx $Context.RunIndex -BagMin $bagMin -BagMax $bagMax -Bag0SampleMax $bag0max -OtherBagSampleMax $bagOtherMax
        if ($missing -and $missing.Count -gt 0) {
            $sev = Get-PolicySeverity -Key 'MISSING_BAG_SAMPLES' -Default 'Info'
            $ev  = ($missing | Select-Object -First 30) -join "`r`n"
            if ($missing.Count -gt 30) { $ev += "`r`n… +$($missing.Count - 30) more" }
            [void]$findings.Add((New-Finding -Severity $sev -RuleId 'MISSING_BAG_SAMPLES' -Message 'Missing samples by bag' -Rows $null -Classification 'Observation' -Evidence $ev))
        }
    } catch {}
}

if ($enableWrongTestType -and $Context.Profile -and $Context.Profile.WrongTestType) {
    try {
        $controls = $Context.Profile.WrongTestType.Controls
        $defaultT = $Context.Profile.WrongTestType.Default
        $badT = New-Object 'System.Collections.Generic.List[object]'

        foreach ($rr in $rows) {
            if (-not $rr.SampleIdDerived -or -not $rr.SampleIdDerived.IsParsed) { continue }
            $ci = $rr.SampleIdDerived.ControlIndex
            if ($ci -eq $null) { continue }

            $expected = $null
            if ($controls -and $controls.ContainsKey([int]$ci)) { $expected = $controls[[int]$ci] }
            elseif ($defaultT) { $expected = $defaultT }

            if ($expected -and $rr.TestType -and ($rr.TestType -ne $expected)) {
                $null = $badT.Add($rr)
            }
        }

        if ($badT.Count -gt 0) {
            $sev = Get-PolicySeverity -Key 'WRONG_TEST_TYPE' -Default 'Warn'
            foreach ($rr in $badT) { Add-FailureTag -Row $rr -Tag 'WRONG_TEST_TYPE' -Severity $sev -PrimaryRuleId 'WRONG_TEST_TYPE' }
            [void]$findings.Add((New-Finding -Severity $sev -RuleId 'WRONG_TEST_TYPE' -Message 'Wrong Test Type (derived from Sample ID control index)' -Rows $badT -Classification 'Sample ID or Test Type Deviation'))
        }
    } catch {}
}

# -----------------------------
# Error classification (Global)
# -----------------------------
$minorCodes = @()
try {
    if ($RuleBank -and $RuleBank.Global -and $RuleBank.Global.ErrorClassification -and $RuleBank.Global.ErrorClassification.MinorFunctionalErrorCodes) {
        $minorCodes = @($RuleBank.Global.ErrorClassification.MinorFunctionalErrorCodes | ForEach-Object { ($_ + '').Trim() } | Where-Object { $_ })
    }
} catch {}
if (-not $minorCodes -or $minorCodes.Count -eq 0) {
    # Safe defaults
    $minorCodes = @('2008','2009','2125','2096','2097','2037','5006','5007','5008','5009','5017','5018','5019','5001','5002','5003','5004','5005','5015','5016','5011')
}

$minorRows = New-Object 'System.Collections.Generic.List[object]'
$instrRows = New-Object 'System.Collections.Generic.List[object]'

foreach ($rr in $rows) {
    $isTerminated = $false
    try {
        if ($rr.TestResult -and ($rr.TestResult -match '^(?i)(ERROR|INVALID|NO RESULT)$')) { $isTerminated = $true }
        elseif ($rr.Status -and ($rr.Status -notmatch '^(?i)Done$')) { $isTerminated = $true }
    } catch {}

    if (-not $isTerminated) { continue }

    $isMinor = $false
    try {
        if ($rr.MaxPressure -ne $null -and [double]$rr.MaxPressure -ge [double]$failThreshold) { $isMinor = $true }
        elseif ($rr.ErrorCode -and ($minorCodes -contains (($rr.ErrorCode + '').Trim()))) { $isMinor = $true }
    } catch {}

    if ($isMinor) { $null = $minorRows.Add($rr) }
    else { $null = $instrRows.Add($rr) }
}

if ($minorRows.Count -gt 0) {
    $sev = Get-PolicySeverity -Key 'MINOR_FUNCTIONAL_ERROR' -Default 'Warn'
    foreach ($rr in $minorRows) { Add-FailureTag -Row $rr -Tag 'MINOR_FUNCTIONAL_ERROR' -Severity $sev -PrimaryRuleId 'MINOR_FUNCTIONAL_ERROR' }
    [void]$findings.Add((New-Finding -Severity $sev -RuleId 'MINOR_FUNCTIONAL_ERROR' -Message 'Minor Functional Error (no re-test)' -Rows $minorRows -Classification 'Minor Functional Error' -GeneratesRetest $false))
}
if ($instrRows.Count -gt 0) {
    $sev = Get-PolicySeverity -Key 'INSTRUMENT_ERROR' -Default 'Error'
    foreach ($rr in $instrRows) { Add-FailureTag -Row $rr -Tag 'INSTRUMENT_ERROR' -Severity $sev -PrimaryRuleId 'INSTRUMENT_ERROR' }
    [void]$findings.Add((New-Finding -Severity $sev -RuleId 'INSTRUMENT_ERROR' -Message 'Instrument Error (re-test/replacement needed)' -Rows $instrRows -Classification 'Instrument Error' -GeneratesRetest $true))
}

# Observation: repeated errors on same Instrument+Module
if ($Context.RunIndex -and $Context.RunIndex.ErrorByInstrumentModule) {
    try {
        $repeatRows = New-Object 'System.Collections.Generic.List[object]'
        $lines = New-Object 'System.Collections.Generic.List[string]'

        foreach ($k in $Context.RunIndex.ErrorByInstrumentModule.Keys) {
            $lst = $Context.RunIndex.ErrorByInstrumentModule[$k]
            if ($lst -and $lst.Count -gt 1) {
                $sorted = @($lst | Sort-Object)
                $lines.Add(("Lines {0} for {1}" -f ($sorted -join ', '), $k))
                foreach ($ri in $sorted) { if ($rowByIndex.ContainsKey([int]$ri)) { $null = $repeatRows.Add($rowByIndex[[int]$ri]) } }
            }
        }

        if ($lines.Count -gt 0) {
            $sev = Get-PolicySeverity -Key 'REPEAT_ERRORS_IN_MODULE' -Default 'Info'
            $ev = ($lines | Select-Object -First 30) -join "`r`n"
            if ($lines.Count -gt 30) { $ev += "`r`n… +$($lines.Count - 30) more" }
            foreach ($rr in $repeatRows) { Add-FailureTag -Row $rr -Tag 'REPEAT_ERRORS_IN_MODULE' -Severity $sev -PrimaryRuleId 'REPEAT_ERRORS_IN_MODULE' }
            [void]$findings.Add((New-Finding -Severity $sev -RuleId 'REPEAT_ERRORS_IN_MODULE' -Message 'Observation: repeated errors on same Instrument+Module' -Rows $repeatRows -Classification 'Observation' -Evidence $ev))
        }
    } catch {}
}

# -----------------------------
# Assay-specific expected results (profile-driven)
# -----------------------------
try {
    if ($Context.Profile -and $Context.Profile.ExpectedResults) {
        $exp = $Context.Profile.ExpectedResults
        $extraAllow = $null
        try { $extraAllow = $Context.Profile.AdditionalAcceptedResults } catch {}

        $majorRows = New-Object 'System.Collections.Generic.List[object]'

        foreach ($rr in $rows) {
            # Only evaluate "Done" non-termination results
            if (-not $rr.Status -or ($rr.Status -notmatch '^(?i)Done$')) { continue }
            if (-not $rr.TestType) { continue }
            if (-not $rr.TestResult) { continue }
            if ($rr.TestResult -match '^(?i)(ERROR|INVALID|NO RESULT)$') { continue }

            if (-not $exp.ContainsKey($rr.TestType)) { continue }

            $patterns = @($exp[$rr.TestType])
            $ok = $false
            foreach ($pat in $patterns) {
                if ([string]::IsNullOrWhiteSpace($pat)) { continue }
                try { if ($rr.TestResult -match $pat) { $ok = $true; break } } catch {}
            }

            if (-not $ok -and $extraAllow -and $extraAllow.ContainsKey($rr.TestType)) {
                foreach ($pat in @($extraAllow[$rr.TestType])) {
                    if ([string]::IsNullOrWhiteSpace($pat)) { continue }
                    try { if ($rr.TestResult -match $pat) { $ok = $true; break } } catch {}
                }
            }

            if (-not $ok) { $null = $majorRows.Add($rr) }
        }

        if ($majorRows.Count -gt 0) {
            $sev = Get-PolicySeverity -Key 'MAJOR_FUNCTIONAL_ERROR' -Default 'Error'
            foreach ($rr in $majorRows) { Add-FailureTag -Row $rr -Tag 'MAJOR_FUNCTIONAL_ERROR' -Severity $sev -PrimaryRuleId 'MAJOR_FUNCTIONAL_ERROR' }
            [void]$findings.Add((New-Finding -Severity $sev -RuleId 'MAJOR_FUNCTIONAL_ERROR' -Message 'Major Functional Error (unexpected Test Result)' -Rows $majorRows -Classification 'Major Functional Error' -GeneratesRetest $true))
        }
    }
} catch {}
        # RuleBank rules (oförändrad)
        if ($RuleBank -and $RuleBank.Rules) {
            foreach ($rule in $RuleBank.Rules) {
                if (-not $rule) { continue }
                try {
                    $rid = $rule.RuleId
                    $sev = if ($rule.Severity) { $rule.Severity } else { 'Warn' }
                    $msg = $rule.Message
                    $class = if ($rule.Classification) { $rule.Classification } else { '' }
                    $generates = $false
                    try { if ($rule.GeneratesRetest -ne $null) { $generates = [bool]$rule.GeneratesRetest } } catch {}

                    $affRows = @()
                    if ($rule.Filter) {
                        $affRows = @($rows | Where-Object $rule.Filter)
                    } elseif ($rule.FilterScript) {
                        $affRows = @($rows | Where-Object $rule.FilterScript)
                    }

                    if ($affRows.Count -gt 0) {
                        foreach ($r in $affRows) { Add-FailureTag -Row $r -RuleId $rid -Severity $sev }
                        [void]$findings.Add((New-Finding -Severity $sev -RuleId $rid -Message $msg -Rows $affRows -Classification $class -GeneratesRetest $generates))
                    }
                } catch {}
            }
        }

        # Error code breakdown
        $errorCodeCounts = $Context.ErrorCodeCounts
        $knownCodes = @()
        try { if ($RuleBank -and $RuleBank.ErrorBank -and $RuleBank.ErrorBank.Codes) { $knownCodes = $RuleBank.ErrorBank.Codes.Keys } } catch {}
        if ($errorCodeCounts -and $errorCodeCounts.Count -gt 0) {
            $errDetails = New-Object System.Collections.Generic.List[object]
            foreach ($code in ($errorCodeCounts.Keys | Sort-Object)) {
                $defKey = $code
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
                $result.ErrorSummary = @(
                    $errDetails | Sort-Object `
                        @{ Expression = 'Count';     Descending = $true }, `
                        @{ Expression = 'ErrorCode'; Descending = $false }
                )
                $result.ErrorCodesTable = $result.ErrorSummary
                try { $result.UniqueErrorCodes = @($errorCodeCounts.Keys).Count } catch {}
            }
        }

        # Unknown error codes (Info) – still show to highlight new/unmapped codes, but do not elevate severity
        if ($knownCodes.Count -gt 0) {
            $rowsUnknownErr = @($rows | Where-Object { $_.ErrorCode -and ($knownCodes -notcontains $_.ErrorCode) })
            if ($rowsUnknownErr.Count -gt 0) {
                $unknownCodesTxt = (($rowsUnknownErr | Select-Object -ExpandProperty ErrorCode) | Select-Object -Unique) -join ', '
                [void]$findings.Add((New-Finding -Severity 'Info' -RuleId 'UNKNOWN_ERROR_CODE' -Message 'Observation: Error code not in RuleBank ErrorBank list.' -Rows $rowsUnknownErr -Classification 'Observation' -Evidence ("Unknown codes: {0}" -f $unknownCodesTxt)))
            }
        }

        # Affected tests – stabil sortering
        $affected = @($rows | Where-Object { $_.FailureTags -and $_.FailureTags.Count -gt 0 })
        $affected = @(
            $affected | Sort-Object `
                @{ Expression = { Get-SeverityRank $_.Severity }; Descending = $true }, `
                @{ Expression = { $_.ErrorCode } }, `
                @{ Expression = { $_.SampleID } }, `
                @{ Expression = { $_.RowIndex } }
        )
        $result.AffectedTests = $affected
        $result.ErrorRowCount = $affected.Count

# Bucket counts (unique rows per bucket)
$result.BucketCounts = [ordered]@{
    'Minor Functional Error'           = @($rows | Where-Object { $_.FailureTags -contains 'MINOR_FUNCTIONAL_ERROR' } | Select-Object -ExpandProperty RowIndex -Unique).Count
    'Major Functional Error'           = @($rows | Where-Object { $_.FailureTags -contains 'MAJOR_FUNCTIONAL_ERROR' } | Select-Object -ExpandProperty RowIndex -Unique).Count
    'Instrument Error'                 = @($rows | Where-Object { $_.FailureTags -contains 'INSTRUMENT_ERROR' } | Select-Object -ExpandProperty RowIndex -Unique).Count
    'Replacements'                     = @($rows | Where-Object { $_.FailureTags -contains 'REPLACEMENT_A' } | Select-Object -ExpandProperty RowIndex -Unique).Count
    'Sample ID or Test Type Deviation' = @($rows | Where-Object { ($_.FailureTags -contains 'SAMPLEID_PARSE_FAILED') -or ($_.FailureTags -contains 'WRONG_TEST_TYPE') } | Select-Object -ExpandProperty RowIndex -Unique).Count
    'Observation'                      = @($rows | Where-Object { ($_.FailureTags -contains 'REPEAT_ERRORS_IN_MODULE') -or ($_.FailureTags -contains 'DELAMINATION_D') } | Select-Object -ExpandProperty RowIndex -Unique).Count
    'Duplicates'                       = @($rows | Where-Object { ($_.FailureTags -contains 'DUPLICATE_SAMPLE_ID') -or ($_.FailureTags -contains 'DUPLICATE_CARTRIDGE_SN') } | Select-Object -ExpandProperty RowIndex -Unique).Count
}


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
# Writer (Information2) – NO Avg + EPPlus CF fix
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
        if (-not $Evaluation) { $Evaluation = [pscustomobject]@{ Findings=@(); AffectedTests=@(); BucketCounts=[ordered]@{}; SeverityCounts=@{Error=0;Warn=0;Info=0}; PressureStats=$null; ErrorSummary=@(); UniqueErrorCodes=0; BaselineExpected=0; BaselineDelta=0; ErrorRowCount=0 } }
        if (-not $Context)    { $Context    = [pscustomobject]@{ TotalTests=0; AssayCanonical=''; AssayVersion=''; ReagentLotIds=@(); WorkCenters=@(); DuplicateCounts=[ordered]@{SampleId=0;CartridgeSN=0}; DuplicateEvidence=[ordered]@{SampleId='';CartridgeSN=''}; RunIndex=$null } }

        $rowsAll = @($Evaluation.AffectedTests)
        $rowsMajor = @($rowsAll | Where-Object { $_.FailureTags -contains 'MAJOR_FUNCTIONAL_ERROR' } | Sort-Object -Property RowIndex)
        $rowsMinor = @($rowsAll | Where-Object { $_.FailureTags -contains 'MINOR_FUNCTIONAL_ERROR' } | Sort-Object -Property RowIndex)
        $rowsInstr = @($rowsAll | Where-Object { $_.FailureTags -contains 'INSTRUMENT_ERROR' } | Sort-Object -Property RowIndex)
        $rowsRepl  = @($rowsAll | Where-Object { $_.FailureTags -contains 'REPLACEMENT_A' } | Sort-Object -Property RowIndex)
        $rowsObs   = @($rowsAll | Where-Object { ($_.FailureTags -contains 'REPEAT_ERRORS_IN_MODULE') -or ($_.FailureTags -contains 'DELAMINATION_D') -or ($_.FailureTags -contains 'UNKNOWN_ERROR_CODE') } | Sort-Object -Property RowIndex)
        $rowsBadId = @($rowsAll | Where-Object { $_.FailureTags -contains 'SAMPLEID_PARSE_FAILED' } | Sort-Object -Property RowIndex)
        $rowsBadTT = @($rowsAll | Where-Object { $_.FailureTags -contains 'WRONG_TEST_TYPE' } | Sort-Object -Property RowIndex)
        $rowsDupId = @($rowsAll | Where-Object { $_.FailureTags -contains 'DUPLICATE_SAMPLE_ID' } | Sort-Object -Property RowIndex)
        $rowsDupSN = @($rowsAll | Where-Object { $_.FailureTags -contains 'DUPLICATE_CARTRIDGE_SN' } | Sort-Object -Property RowIndex)

        $missingRows = @()
        try {
            $missingLines = New-Object 'System.Collections.Generic.List[string]'
            foreach ($f in @($Evaluation.Findings)) {
                if ($f -and $f.RuleId -eq 'MISSING_BAG_SAMPLES' -and $f.Evidence) {
                    foreach ($ln in ($f.Evidence -split "`r?`n")) { $missingLines.Add($ln) | Out-Null }
                }
            }
            foreach ($ln in $missingLines) {
                $t = ($ln + '').Trim()
                if (-not $t) { continue }
                if ($t -match '^\s*…') { continue }
                $missingRows += [pscustomobject]@{ Item = $t; Value = '' }
            }
        } catch {}

        $controlRows = @()
        try {
            if ($Context.RunIndex -and $Context.RunIndex.ControlMaterialCounts) {
                foreach ($k in ($Context.RunIndex.ControlMaterialCounts.Keys | Sort-Object)) {
                    $controlRows += [pscustomobject]@{ Item = $k; Value = [int]$Context.RunIndex.ControlMaterialCounts[$k] }
                }
            }
        } catch {}

        $standardHeaders = @(
            'Sample ID','Cartridge S/N','Test Type','Instrument S/N','Module S/N',
            'Start Time','Status','Test Result','Max Pressure (PSI)','Error'
        )
        $sectionDefs = @(
            [pscustomobject]@{ Title = 'Major Functional';           Mode = 'Standard'; Rows = $rowsMajor },
            [pscustomobject]@{ Title = 'Minor Functional';           Mode = 'Standard'; Rows = $rowsMinor },
            [pscustomobject]@{ Title = 'Instrument Error';           Mode = 'Standard'; Rows = $rowsInstr },
            [pscustomobject]@{ Title = 'Replacements';               Mode = 'Standard'; Rows = $rowsRepl  },
            [pscustomobject]@{ Title = 'Observation';                Mode = 'Standard'; Rows = $rowsObs   },
            [pscustomobject]@{ Title = 'Incorrect Sample ID';        Mode = 'Standard'; Rows = $rowsBadId },
            [pscustomobject]@{ Title = 'Incorrect Test Type';        Mode = 'Standard'; Rows = $rowsBadTT },
            [pscustomobject]@{ Title = 'Duplicate of Sample ID';     Mode = 'Standard'; Rows = $rowsDupId },
            [pscustomobject]@{ Title = 'Duplicate of Cartridge S/N'; Mode = 'Standard'; Rows = $rowsDupSN },
            [pscustomobject]@{ Title = 'Missing Samples';            Mode = 'KeyValue'; Rows = $missingRows },
            [pscustomobject]@{ Title = 'Control Materials';          Mode = 'KeyValue'; Rows = $controlRows }
        )

        $headerLayout = @(
            @{ Title = 'Major Functional';           Headers = $standardHeaders },
            @{ Title = 'Minor Functional';           Headers = $standardHeaders },
            @{ Title = 'Instrument Error';           Headers = $standardHeaders },
            @{ Title = 'Replacements';               Headers = $standardHeaders },
            @{ Title = 'Observation';                Headers = $standardHeaders },
            @{ Title = 'Incorrect Sample ID';        Headers = $standardHeaders },
            @{ Title = 'Incorrect Test Type';        Headers = $standardHeaders },
            @{ Title = 'Duplicate of Sample ID';     Headers = $standardHeaders },
            @{ Title = 'Duplicate of Cartridge S/N'; Headers = $standardHeaders },
            @{ Title = 'Missing Samples';            Headers = @('Hint') },
            @{ Title = 'Control Materials';          Headers = @('Count') }
        )

        $titleSet = New-Object 'System.Collections.Generic.HashSet[string]'
        foreach ($s in $sectionDefs) { [void]$titleSet.Add((Normalize-HeaderName $s.Title)) }

        $getSectionRow = {
            param([string]$Title)
            if (-not $Worksheet.Dimension) { return 0 }
            $target = Normalize-HeaderName $Title
            for ($row = $Worksheet.Dimension.Start.Row; $row -le $Worksheet.Dimension.End.Row; $row++) {
                $txt = ($Worksheet.Cells[$row,1].Text + '').Trim()
                if (-not $txt) { continue }
                if ((Normalize-HeaderName $txt) -eq $target) { return $row }
            }
            return 0
        }

        foreach ($s in $sectionDefs) { $s.Row = & $getSectionRow $s.Title }
        $found = @($sectionDefs | Where-Object { $_.Row -gt 0 })

        if ($found.Count -eq 0) {
            $row = 1
            foreach ($sec in $headerLayout) {
                $Worksheet.Cells[$row,1].Value = $sec.Title
                $Worksheet.Cells[$row,1].Style.Font.Bold = $true
                for ($i = 0; $i -lt $sec.Headers.Count; $i++) {
                    $Worksheet.Cells[$row,2 + $i].Value = $sec.Headers[$i]
                    $Worksheet.Cells[$row,2 + $i].Style.Font.Bold = $true
                }
                $row += 2
            }
            foreach ($s in $sectionDefs) { $s.Row = & $getSectionRow $s.Title }
            $found = @($sectionDefs | Where-Object { $_.Row -gt 0 })
        }

        if ($found.Count -eq 0) { return }

        $sorted = @($found | Sort-Object -Property Row)
        for ($i = 0; $i -lt $sorted.Count; $i++) {
            $sorted[$i].NextRow = if ($i + 1 -lt $sorted.Count) { $sorted[$i + 1].Row } else { ($Worksheet.Dimension.End.Row + 1) }
        }

        $isHeaderRow = {
            param([int]$Row)
            if ($Row -lt 1 -or -not $Worksheet.Dimension) { return $false }
            $txt = ($Worksheet.Cells[$Row,1].Text + '').Trim()
            if (-not $txt) { return $false }
            return $titleSet.Contains((Normalize-HeaderName $txt))
        }

        $getValue = {
            param([string]$Header, [pscustomobject]$Row)
            switch -Regex ($Header) {
                '^sample id$'        { return $Row.SampleID }
                '^cartridge\s*s/?n$' { return $Row.CartridgeSN }
                '^test type$'        { return $Row.TestType }
                '^instrument\s*s/?n$' { return $Row.InstrumentSN }
                '^module\s*s/?n$'   { return $Row.ModuleSN }
                '^start time$'       { return $Row.StartTime }
                '^status$'           { return $Row.Status }
                '^test result$'      { return $Row.TestResult }
                '^max pressure'      { return if ($Row.MaxPressure -ne $null) { [Math]::Round([double]$Row.MaxPressure,2) } else { $null } }
                '^error$'            { return $Row.ErrorRaw }
                '^error code$'       { return $Row.ErrorCode }
                '^row#?$'            { return $Row.RowIndex }
                '^workcenter$'       { return $Row.WorkCenter }
                '^primary rule$'     { return $Row.PrimaryRule }
                default              { return $null }
            }
        }

        foreach ($sec in (@($sorted | Sort-Object -Property Row -Descending))) {
            $headerRow = [int]$sec.Row
            $dataStart = $headerRow + 1
            $nextRow   = [int]$sec.NextRow
            $dataEnd   = $nextRow - 1
            $existing  = if ($dataEnd -ge $dataStart) { $dataEnd - $dataStart + 1 } else { 0 }
            $rows      = @($sec.Rows)
            $desired   = if ($rows.Count -gt 0) { $rows.Count } else { 1 }

            if ($existing -gt $desired) {
                $Worksheet.DeleteRow($dataStart + $desired, $existing - $desired)
            } elseif ($existing -lt $desired) {
                $toAdd = $desired - $existing
                $copyFrom = $dataStart
                if ($existing -eq 0) {
                    $copyFrom = $dataStart - 1
                    if ((& $isHeaderRow $copyFrom) -and $copyFrom -gt 1) { $copyFrom = $copyFrom - 1 }
                }
                if ($copyFrom -gt 0) { $Worksheet.InsertRow($dataStart + $existing, $toAdd, $copyFrom) }
                else { $Worksheet.InsertRow($dataStart + $existing, $toAdd) }
            }

            $dataEnd = $dataStart + $desired - 1

            if ($sec.Mode -eq 'KeyValue') {
                $Worksheet.Cells[$dataStart,1,$dataEnd,2].Value = $null
                $matrix = New-Object 'object[,]' $desired, 2
                for ($i = 0; $i -lt $desired; $i++) {
                    if ($i -ge $rows.Count) { continue }
                    $matrix[$i,0] = $rows[$i].Item
                    $matrix[$i,1] = $rows[$i].Value
                }
                $Worksheet.Cells[$dataStart,1,$dataEnd,2].Value = $matrix
                continue
            }

            $firstCol = 2
            $lastCol = $firstCol
            $maxCol = $Worksheet.Dimension.End.Column
            for ($c = $firstCol; $c -le $maxCol; $c++) {
                $txt = ($Worksheet.Cells[$headerRow,$c].Text + '').Trim()
                if ($txt) { $lastCol = $c }
            }
            if ($lastCol -lt $firstCol) { continue }

            $Worksheet.Cells[$dataStart,1,$dataEnd,$lastCol].Value = $null

            $headerMap = @{}
            for ($c = $firstCol; $c -le $lastCol; $c++) {
                $h = ($Worksheet.Cells[$headerRow,$c].Text + '').Trim()
                if ($h) { $headerMap[(Normalize-HeaderName $h)] = $c }
            }

            $colCount = $lastCol - $firstCol + 1
            $matrix = New-Object 'object[,]' $desired, $colCount

            for ($i = 0; $i -lt $desired; $i++) {
                if ($i -ge $rows.Count) { continue }
                $row = $rows[$i]
                foreach ($pair in $headerMap.GetEnumerator()) {
                    $colIndex = [int]$pair.Value - $firstCol
                    $value = & $getValue $pair.Key $row
                    if ($value -ne $null) { $matrix[$i,$colIndex] = $value }
                }
            }
            $Worksheet.Cells[$dataStart,$firstCol,$dataEnd,$lastCol].Value = $matrix
        }

    } catch {
        if (Get-Command Log-Exception -ErrorAction SilentlyContinue) {
            Log-Exception -Message "Information2 misslyckades att byggas" -ErrorRecord $_ -Severity 'Warn'
        } else {
            Gui-Log "⚠️ Information2 misslyckades att byggas: $($_.Exception.Message)" 'Warn'
        }
    }
}
#endregion REGION: CSV
