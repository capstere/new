function Import-RuleCsv {
param([Parameter(Mandatory)][string]$Path)
if (-not (Test-Path -LiteralPath $Path)) { return @() }

$delim = ‘,’
try { $delim = Get-CsvDelimiter -Path $Path } catch {}

try {
$lines = Get-Content -LiteralPath $Path -ErrorAction Stop
if (-not $lines -or $lines.Count -lt 1) { return @() }
return @(ConvertFrom-Csv -InputObject ($lines -join “`n”) -Delimiter $delim)
} catch {
try { return @(Import-Csv -LiteralPath $Path -Delimiter $delim) } catch { return @() }
}

}

function _RuleEngine_Log {
param(
[Parameter(Mandatory)][string]$Text,
[ValidateSet(‘Info’,‘Warn’,‘Error’)][string]$Severity = ‘Info’
)
try {
$cmd = Get-Command -Name Gui-Log -ErrorAction SilentlyContinue
if ($cmd) { Gui-Log -Text $Text -Severity $Severity -Category ‘RuleEngine’ }
} catch {}
}

function Test-RuleBankIntegrity {
param(
[Parameter(Mandatory)][pscustomobject]$RuleBank,
[Parameter(Mandatory=$false)][string]$Source = ‘’
)

function _EnsureArray([string]$name) {
$v = $null
try { $v = $RuleBank.$name } catch { $v = $null }
if ($null -eq $v) {
try { $RuleBank | Add-Member -NotePropertyName $name -NotePropertyValue @() -Force } catch {}
return @()
}
return @($v)
}

function _RequireColumns([string]$tableName, [string[]]$cols) {
$rows = _EnsureArray $tableName
if (-not $rows -or $rows.Count -eq 0) { return } # empty table is allowed
$first = $rows[0]
foreach ($c in $cols) {
$ok = $false
try {
if ($first -is [hashtable]) {
$ok = $first.ContainsKey($c)
} else {
$p = $first.PSObject.Properties[$c]
$ok = ($p -ne $null)
}
} catch { $ok = $false }
if (-not $ok) {
$src = $Source
if (-not $src) { $src = ‘RuleBank’ }
throw (“RuleBank (Load-RuleBank): Tabell ‘” + $tableName + “’ saknar kolumn: “ + $c + “ (” + $src + “)”)
}
}
}

foreach ($t in @(‘ResultCallPatterns’,‘SampleExpectationRules’,‘ErrorCodes’,‘MissingSamplesConfig’,‘SampleIdMarkers’,‘ParityCheckConfig’,‘SampleNumberRules’,‘TestTypePolicy’)) {
$null = _EnsureArray $t
}

_RequireColumns ‘ResultCallPatterns’     @(‘Assay’,‘Call’,‘MatchType’,‘Pattern’,‘Enabled’,‘Priority’)
_RequireColumns ‘SampleExpectationRules’ @(‘Assay’,‘SampleIdMatchType’,‘SampleIdPattern’,‘Expected’,‘Enabled’,‘Priority’)
_RequireColumns ‘ErrorCodes’             @(‘ErrorCode’,‘Name’,‘GeneratesRetest’)
_RequireColumns ‘SampleIdMarkers’        @(‘AssayPattern’,‘MarkerType’,‘Marker’,‘Enabled’)
_RequireColumns ‘ParityCheckConfig’      @(‘AssayPattern’,‘Enabled’,‘CartridgeField’,‘SampleTokenIndex’,‘SuffixX’,‘SuffixPlus’,‘MinValidCartridgeSNPercent’,‘Priority’)
_RequireColumns ‘SampleNumberRules’      @(‘AssayPattern’,‘SampleTypeCode’,‘BagNoPattern’,‘SampleNumberTokenIndex’,‘SampleNumberRegex’,‘SampleNumberMin’,‘SampleNumberMax’,‘SampleNumberPad’,‘Enabled’,‘Priority’)
_RequireColumns ‘TestTypePolicy’         @(‘AssayPattern’,‘AllowedTestTypes’,‘Enabled’,‘Priority’)

return $true

}

function Load-RuleBank {
param([Parameter(Mandatory)][string]$RuleBankDir)
$requireCompiled = $false
try {
if (Get-Command Get-ConfigValue -ErrorAction SilentlyContinue) {
$requireCompiled = [bool](Get-ConfigValue -Name ‘RuleBankRequireCompiled’ -Default $false)
} else {
$cfg = $null
if ($global:Config) { $cfg = $global:Config }
elseif (Get-Variable -Name Config -Scope Script -ErrorAction SilentlyContinue) { $cfg = (Get-Variable -Name Config -Scope Script -ValueOnly -ErrorAction SilentlyContinue) }

```
    if ($cfg -is [System.Collections.IDictionary]) {
        if ($cfg.Contains('RuleBankRequireCompiled')) {
            $requireCompiled = [bool]$cfg['RuleBankRequireCompiled']
        }
    } elseif ($cfg -is [hashtable]) {
        if ($cfg.ContainsKey('RuleBankRequireCompiled')) {
            $requireCompiled = [bool]$cfg['RuleBankRequireCompiled']
        }
    }
}
```

} catch { $requireCompiled = $false }

if (-not (Test-Path -LiteralPath $RuleBankDir)) {
throw (“RuleBank (Load-RuleBank): Directory not found: “ + $RuleBankDir)
}

$rb = [ordered]@{
Dir = $RuleBankDir
ResultCallPatterns = @()
SampleExpectationRules = @()
ErrorCodes = @()
MissingSamplesConfig = @()
SampleIdMarkers = @()
ParityCheckConfig = @()
SampleNumberRules = @()
TestTypePolicy = @()
}

$expectedTables = @(‘ResultCallPatterns’,‘SampleExpectationRules’,‘ErrorCodes’,‘MissingSamplesConfig’,‘SampleIdMarkers’,‘ParityCheckConfig’,‘SampleNumberRules’,‘TestTypePolicy’)

function _HasKey([object]$dict, [string]$key) {
try {
if ($dict -is [hashtable]) { return $dict.ContainsKey($key) }
if ($dict -is [System.Collections.IDictionary]) { return $dict.Contains($key) }
} catch {}
return $false
}

$compiledCandidates = @(
(Join-Path $RuleBankDir ‘RuleBank.compiled.ps1’),
(Join-Path $RuleBankDir ‘build\RuleBank.compiled.ps1’),
(Join-Path $RuleBankDir ‘RuleBank.compiled.psd1’),
(Join-Path $RuleBankDir ‘build\RuleBank.compiled.psd1’)
)

foreach ($cp in $compiledCandidates) {
if (-not (Test-Path -LiteralPath $cp)) { continue }

```
try {
    $ht = $null
    if ($cp.ToLowerInvariant().EndsWith('.ps1')) {
        $ht = & $cp
    } else {
        $ht = Import-PowerShellDataFile -Path $cp
    }

    if ($null -eq $ht -or -not ($ht -is [System.Collections.IDictionary] -or $ht -is [hashtable])) {
        throw ("RuleBank (Load-RuleBank): Compiled artifact did not return a dictionary: " + $cp)
    }

    foreach ($t in $expectedTables) {
        if (-not (_HasKey $ht $t)) {
            throw ("RuleBank (Load-RuleBank): Compiled artifact missing table '{0}' ({1})" -f $t, $cp)
        }
    }

    foreach ($t in $expectedTables) {
        $rb[$t] = @($ht[$t])
    }

    try { $rb.ResultCallPatterns = @($rb.ResultCallPatterns | Sort-Object { try { [int]((Get-RowField -Row $_ -FieldName 'Priority') + '') } catch { 0 } } -Descending) } catch {}
    try { $rb.SampleExpectationRules = @($rb.SampleExpectationRules | Sort-Object { try { [int]((Get-RowField -Row $_ -FieldName 'Priority') + '') } catch { 0 } } -Descending) } catch {}
    try { $rb.ParityCheckConfig = @($rb.ParityCheckConfig | Sort-Object { try { [int]((Get-RowField -Row $_ -FieldName 'Priority') + '') } catch { 0 } } -Descending) } catch {}
    try { $rb.SampleIdMarkers = @($rb.SampleIdMarkers | Sort-Object { try { [int]((Get-RowField -Row $_ -FieldName 'Priority') + '') } catch { 0 } } -Descending) } catch {}
    try { $rb.SampleNumberRules = @($rb.SampleNumberRules | Sort-Object { try { [int]((Get-RowField -Row $_ -FieldName 'Priority') + '') } catch { 0 } } -Descending) } catch {}
    try { $rb.TestTypePolicy = @($rb.TestTypePolicy | Sort-Object { try { [int]((Get-RowField -Row $_ -FieldName 'Priority') + '') } catch { 0 } } -Descending) } catch {}

    $rbObj = [pscustomobject]$rb
    Test-RuleBankIntegrity -RuleBank $rbObj -Source ("compiled:" + $cp)

    try {
        $cnt = @()
        foreach ($t in $expectedTables) {
            $cnt += ("{0}={1}" -f $t, (@($rbObj.$t).Count))
        }
        _RuleEngine_Log -Text ("🧠 RuleBank laddad från compiled. " + ($cnt -join ', ')) -Severity 'Info'
    } catch {}

    return (Compile-RuleBank -RuleBank $rbObj)

} catch {
    if ($requireCompiled) {
        throw ("RuleBank (Load-RuleBank): Compiled artifact failed to load: {0} ({1})" -f $cp, $_.Exception.Message)
    }
}
```

}

if ($requireCompiled) {
throw (“RuleBank (Load-RuleBank): Compiled artifact missing. Expected RuleBank.compiled.ps1 in: {0}” -f $RuleBankDir)
}

# –– CSV fallback ––

$map = @(
@{ Key=‘ResultCallPatterns’;      File=‘01_ResultCallPatterns.csv’ },
@{ Key=‘SampleExpectationRules’;  File=‘02_SampleExpectationRules.csv’ },
@{ Key=‘ErrorCodes’;              File=‘03_ErrorCodes.csv’ },
@{ Key=‘MissingSamplesConfig’;    File=‘04_MissingSamplesConfig.csv’ },
@{ Key=‘SampleIdMarkers’;         File=‘05_SampleIdMarkers.csv’ },
@{ Key=‘ParityCheckConfig’;       File=‘06_ParityCheckConfig.csv’ },
@{ Key=‘SampleNumberRules’;       File=‘07_SampleNumberRules.csv’ },
@{ Key=‘TestTypePolicy’;          File=‘08_TestTypePolicy.csv’ }
)

foreach ($m in $map) {
$p = Join-Path $RuleBankDir $m.File
$rb[$m.Key] = @(Import-RuleCsv -Path $p)
}

try { $rb.ResultCallPatterns = @($rb.ResultCallPatterns | Sort-Object { [int]($_.Priority) } -Descending) } catch {}
try { $rb.SampleExpectationRules = @($rb.SampleExpectationRules | Sort-Object { [int]($_.Priority) } -Descending) } catch {}
try { $rb.ParityCheckConfig = @($rb.ParityCheckConfig | Sort-Object { [int]($_.Priority) } -Descending) } catch {}
try { $rb.SampleIdMarkers = @($rb.SampleIdMarkers | Sort-Object { [int]($_.Priority) } -Descending) } catch {}
try { $rb.SampleNumberRules = @($rb.SampleNumberRules | Sort-Object { [int]($_.Priority) } -Descending) } catch {}
try { $rb.TestTypePolicy = @($rb.TestTypePolicy | Sort-Object { [int]($_.Priority) } -Descending) } catch {}

$rbObj2 = [pscustomobject]$rb
Test-RuleBankIntegrity -RuleBank $rbObj2 -Source ‘csv’
try {
$cnt = @()
foreach ($t in $expectedTables) {
$cnt += (”{0}={1}” -f $t, (@($rbObj2.$t).Count))
}
_RuleEngine_Log -Text (“🧠 RuleBank laddad från CSV. “ + ($cnt -join ’, ’)) -Severity ‘Info’
} catch {}

return (Compile-RuleBank -RuleBank $rbObj2)

}

function Compile-RuleBank {
param([Parameter(Mandatory)][pscustomobject]$RuleBank)

$compiled = [ordered]@{
RegexCache = @{}
PatternsByAssay = @{}
ExpectRulesByAssay = @{}
MarkerByAssayType = @{}
PolicyByAssay = @{}
SampleNumRuleByAssayCode = @{}
}

try {
foreach ($r in @($RuleBank.ResultCallPatterns)) {
if (-not $r) { continue }
if (-not (Test-RuleEnabled $r)) { continue }
$mt = ((Get-RowField -Row $r -FieldName ‘MatchType’) + ‘’).Trim().ToUpperInvariant()
if ($mt -ne ‘REGEX’) { continue }
$pat = ((Get-RowField -Row $r -FieldName ‘Pattern’) + ‘’)
if (-not ($pat.Trim())) { continue }
if (-not $compiled.RegexCache.ContainsKey($pat)) {
try {
$compiled.RegexCache[$pat] = New-Object System.Text.RegularExpressions.Regex($pat, [System.Text.RegularExpressions.RegexOptions]::IgnoreCase)
} catch {
# Invalid regex: cache a regex that never matches to preserve “false” outcome deterministically
$compiled.RegexCache[$pat] = New-Object System.Text.RegularExpressions.Regex(‘a\A’, [System.Text.RegularExpressions.RegexOptions]::IgnoreCase)
}
}
}
} catch {}

try { $RuleBank | Add-Member -NotePropertyName ‘Compiled’ -NotePropertyValue ([pscustomobject]$compiled) -Force } catch { $RuleBank.Compiled = [pscustomobject]$compiled }
return $RuleBank

}

function Get-ResultCallPatternsForAssay {
param(
[Parameter(Mandatory)][pscustomobject]$RuleBank,
[Parameter(Mandatory)][string]$Assay
)
$aKey = (($Assay + ‘’).Trim())
if (-not $aKey) { $aKey = ‘(blank)’ }

$c = $RuleBank.Compiled
if (-not $c) { return @($RuleBank.ResultCallPatterns) }

if ($c.PatternsByAssay.ContainsKey($aKey)) { return @($c.PatternsByAssay[$aKey]) }

$list = New-Object System.Collections.Generic.List[object]
foreach ($r in @($RuleBank.ResultCallPatterns)) {
if (-not $r) { continue }
if (-not (Test-RuleEnabled $r)) { continue }
$ruleAssay = (Get-RowField -Row $r -FieldName ‘Assay’)
if (-not (Test-AssayMatch -RuleAssay $ruleAssay -RowAssay $Assay)) { continue }
$list.Add($r)
}
$arr = $list.ToArray()
$c.PatternsByAssay[$aKey] = $arr
return @($arr)

}

function Get-ExpectationRulesForAssay {
param(
[Parameter(Mandatory)][pscustomobject]$RuleBank,
[Parameter(Mandatory)][string]$Assay
)
$aKey = (($Assay + ‘’).Trim())
if (-not $aKey) { $aKey = ‘(blank)’ }

$c = $RuleBank.Compiled
if (-not $c) { return @($RuleBank.SampleExpectationRules) }

if ($c.ExpectRulesByAssay.ContainsKey($aKey)) { return @($c.ExpectRulesByAssay[$aKey]) }

$list = New-Object System.Collections.Generic.List[object]
foreach ($r in @($RuleBank.SampleExpectationRules)) {
if (-not $r) { continue }
if (-not (Test-RuleEnabled $r)) { continue }
$ruleAssay = (Get-RowField -Row $r -FieldName ‘Assay’)
if (-not (Test-AssayMatch -RuleAssay $ruleAssay -RowAssay $Assay)) { continue }
$list.Add($r)
}
$arr = $list.ToArray()
$c.ExpectRulesByAssay[$aKey] = $arr
return @($arr)

}

function Match-TextFast {
param(
[Parameter(Mandatory)][string]$Text,
[Parameter(Mandatory)][string]$Pattern,
[Parameter(Mandatory)][string]$MatchType,
[Parameter(Mandatory=$false)][object]$RegexCache
)

$t = ($Text + ‘’)
$p = ($Pattern + ‘’)
$m = ($MatchType + ‘’).Trim().ToUpperInvariant()
if (-not $m) { $m = ‘CONTAINS’ }

try {
switch ($m) {
‘REGEX’  {
if (($RegexCache -is [hashtable]) -and $RegexCache.ContainsKey($p)) {
return $RegexCache[$p].IsMatch($t)
}
return [regex]::IsMatch($t, $p, [System.Text.RegularExpressions.RegexOptions]::IgnoreCase)
}
‘EQUALS’ {
return (($t.Trim()).ToUpperInvariant() -eq ($p.Trim()).ToUpperInvariant())
}
‘PREFIX’ {
$tt = ($t.Trim()).ToUpperInvariant()
$pp = ($p.Trim()).ToUpperInvariant()
if (-not $pp) { return $true }
return ($tt.Length -ge $pp.Length -and $tt.Substring(0, $pp.Length) -eq $pp)
}
‘SUFFIX’ {
$tt = ($t.Trim()).ToUpperInvariant()
$pp = ($p.Trim()).ToUpperInvariant()
if (-not $pp) { return $true }
return ($tt.Length -ge $pp.Length -and $tt.Substring($tt.Length - $pp.Length) -eq $pp)
}
default {
if (-not $p) { return $true }
return ($t.IndexOf($p, [System.StringComparison]::OrdinalIgnoreCase) -ge 0)
}
}
} catch {
return $false
}

}

function Get-TestTypePolicyForAssayCached {
param(
[Parameter(Mandatory)][pscustomobject]$RuleBank,
[Parameter(Mandatory)][string]$Assay
)
$aKey = (($Assay + ‘’).Trim())
if (-not $aKey) { $aKey = ‘(blank)’ }

$c = $RuleBank.Compiled
if (-not $c) { return (Get-TestTypePolicyForAssay -Assay $Assay -Policies $RuleBank.TestTypePolicy) }

if ($c.PolicyByAssay.ContainsKey($aKey)) { return $c.PolicyByAssay[$aKey] }

$pol = $null
foreach ($p in @($RuleBank.TestTypePolicy)) {
try {
if (((Get-RowField -Row $p -FieldName ‘Enabled’) + ‘’).Trim().Length -gt 0 -and ((Get-RowField -Row $p -FieldName ‘Enabled’) + ‘’).Trim().ToUpperInvariant() -in @(‘FALSE’,‘0’,‘NO’,‘N’)) { continue }
$pat = ((Get-RowField -Row $p -FieldName ‘AssayPattern’) + ‘’)
if (Test-AssayMatch -RuleAssay $pat -RowAssay $Assay) { $pol = $p; break }
} catch {}
}
$c.PolicyByAssay[$aKey] = $pol
return $pol

}

function Get-SampleNumberRuleForRowCached {
param(
[Parameter(Mandatory)][pscustomobject]$RuleBank,
[Parameter(Mandatory)][string]$Assay,
[Parameter(Mandatory)][string]$ControlCode,
[Parameter(Mandatory=$false)][string]$BagNo = ‘’
)

$aKey = (($Assay + ‘’).Trim()); if (-not $aKey) { $aKey = ‘(blank)’ }
$ccKey = (($ControlCode + ‘’).Trim())
$bnKey = (($BagNo + ‘’).Trim())
$key = $aKey + ‘|’ + $ccKey + ‘|’ + $bnKey

$c = $RuleBank.Compiled
if (-not $c) {
return (Get-SampleNumberRuleForRow -Assay $Assay -ControlCode $ControlCode -BagNo $BagNo -Rules $RuleBank.SampleNumberRules)
}

if ($c.SampleNumRuleByAssayCode.ContainsKey($key)) { return $c.SampleNumRuleByAssayCode[$key] }

$rule = Get-SampleNumberRuleForRow -Assay $Assay -ControlCode $ControlCode -BagNo $BagNo -Rules $RuleBank.SampleNumberRules
$c.SampleNumRuleByAssayCode[$key] = $rule
return $rule

}

function Get-RowField {
param(
[Parameter(Mandatory)][object]$Row,
[Parameter(Mandatory)][string]$FieldName
)

if ($null -eq $Row) { return ‘’ }
try {
if ($Row -is [hashtable]) {
if ($Row.ContainsKey($FieldName) -and $null -ne $Row[$FieldName]) { return $Row[$FieldName] }
return ‘’
}
if ($Row -is [System.Collections.IDictionary]) {
if ($Row.Contains($FieldName) -and $null -ne $Row[$FieldName]) { return $Row[$FieldName] }
return ‘’
}
} catch {}

try {
$p = $Row.PSObject.Properties[$FieldName]
if ($p -and $null -ne $p.Value) { return $p.Value }
} catch {}

return ‘’

}

function Test-RuleEnabled {
param([Parameter(Mandatory)][object]$Rule)
$en = (Get-RowField -Row $Rule -FieldName ‘Enabled’)
if ($en -eq $null) { return $true }
$s = ($en + ‘’).Trim().ToUpperInvariant()
if (-not $s) { return $true }
return ($s -in @(‘TRUE’,‘1’,‘YES’,‘Y’))
}

function Test-AssayMatch {
param(
[Parameter(Mandatory)][string]$RuleAssay,
[Parameter(Mandatory)][string]$RowAssay
)
$ra = ($RuleAssay + ‘’).Trim()
if (-not $ra -or $ra -eq ‘*’) { return $true }

$row = ($RowAssay + ‘’).Trim()
if ($ra -like ‘*[*?]*’) {
return ($row -like $ra)
}
return ($ra -ieq $row)

}

function Get-TestTypePolicyForAssay {
param(
[Parameter(Mandatory)][string]$Assay,
[Parameter(Mandatory)][object[]]$Policies
)
if (-not $Policies) { return $null }
foreach ($p in $Policies) {
try {
if ((($p.Enabled + ‘’)).Trim().Length -gt 0 -and ($p.Enabled + ‘’).Trim().ToUpperInvariant() -in @(‘FALSE’,‘0’,‘NO’,‘N’)) { continue }
$pat = ($p.AssayPattern + ‘’)
if (Test-AssayMatch -RuleAssay $pat -RowAssay $Assay) { return $p }
} catch {}
}
return $null
}

function Match-Text {
param(
[Parameter(Mandatory)][string]$Text,
[Parameter(Mandatory)][string]$Pattern,
[Parameter(Mandatory)][string]$MatchType
)

$t = ($Text + ‘’)
$p = ($Pattern + ‘’)
$m = ($MatchType + ‘’).Trim().ToUpperInvariant()
if (-not $m) { $m = ‘CONTAINS’ }

try {
switch ($m) {
‘REGEX’  { return [regex]::IsMatch($t, $p, [System.Text.RegularExpressions.RegexOptions]::IgnoreCase) }
‘EQUALS’ { return (($t.Trim()).ToUpperInvariant() -eq ($p.Trim()).ToUpperInvariant()) }
‘PREFIX’ {
$tt = ($t.Trim()).ToUpperInvariant()
$pp = ($p.Trim()).ToUpperInvariant()
if (-not $pp) { return $true }
return ($tt.Length -ge $pp.Length -and $tt.Substring(0, $pp.Length) -eq $pp)
}
‘SUFFIX’ {
$tt = ($t.Trim()).ToUpperInvariant()
$pp = ($p.Trim()).ToUpperInvariant()
if (-not $pp) { return $true }
return ($tt.Length -ge $pp.Length -and $tt.Substring($tt.Length - $pp.Length) -eq $pp)
}
default {
return (($t.ToUpperInvariant()).Contains($p.ToUpperInvariant()))
}
}
} catch {
return $false
}

}

function Get-ObservedCallDetailed {
param(
[Parameter(Mandatory)][object]$Row,
[Parameter(Mandatory=$false)][object[]]$Patterns = @(),
[Parameter(Mandatory=$false)][object]$RegexCache = $null
)

if (-not $Patterns) { $Patterns = @() }

$status = (Get-RowField -Row $Row -FieldName ‘Status’)
$errTxt = (Get-RowField -Row $Row -FieldName ‘Error’)
$testResult = (Get-RowField -Row $Row -FieldName ‘Test Result’)
$assay = (Get-RowField -Row $Row -FieldName ‘Assay’)

if (($errTxt + ‘’).Trim()) {
return [pscustomobject]@{ Call=‘ERROR’; Reason=‘Error column populated’ }
}
$st = ($status + ‘’).Trim()
if ($st -and ($st -ine ‘Done’)) {
return [pscustomobject]@{ Call=‘ERROR’; Reason=(“Status=” + $st) }
}

$tr = ($testResult + ‘’).Trim()
if (-not $tr) { return [pscustomobject]@{ Call=‘UNKNOWN’; Reason=‘Blank Test Result’ } }

$ass = ($assay + ‘’)
if ($ass -match ‘MTB’) {
if ($tr -match ‘(?i)\bMTB\s+TRACE\s+DETECTED\b’) {
return [pscustomobject]@{ Call=‘POS’; Reason=‘MTB Trace detected (override)’ }
}
if ($tr -match ‘(?i)\bMTB\s+DETECTED\b’) {
return [pscustomobject]@{ Call=‘POS’; Reason=‘MTB detected (override)’ }
}
if ($tr -match ‘(?i)\bMTB\s+NOT\s+DETECTED\b’) {
return [pscustomobject]@{ Call=‘NEG’; Reason=‘MTB not detected (override)’ }
}
}

$hasErr = $false
$hasNeg = $false
$hasPos = $false
$isMixed = $false
if (-not ($ass -match ‘(?i)MTB’)) {
$u = ($tr + ‘’).ToUpperInvariant()
$u = [regex]::Replace($u, ‘\s+’, ’ ’).Trim()

```
$hasErr = ($u -match '\bINVALID\b') -or ($u -match 'NO\s+RESULT') -or ($u -match '\bERROR\b')
$hasNeg = ($u -match 'NOT\s+DETECTED') -or ($u -match '\bNEGATIVE\b')

$uNoNotDetected = ($u -replace 'NOT\s+DETECTED', '')
$hasPos = ($uNoNotDetected -match '\bDETECTED\b') -or ($uNoNotDetected -match '\bPOSITIVE\b')

$isMixed = ($hasPos -and $hasNeg)
```

}

foreach ($r in $Patterns) {
if (-not (Test-RuleEnabled $r)) { continue }
$ruleAssay = (Get-RowField -Row $r -FieldName ‘Assay’)
if (-not (Test-AssayMatch -RuleAssay $ruleAssay -RowAssay $assay)) { continue }

```
$pat  = (Get-RowField -Row $r -FieldName 'Pattern')
if (-not ($pat + '').Trim()) { continue }
$mt   = (Get-RowField -Row $r -FieldName 'MatchType')

if (Match-TextFast -Text $tr -Pattern $pat -MatchType $mt -RegexCache $RegexCache) {
    $call = ((Get-RowField -Row $r -FieldName 'Call') + '').Trim().ToUpperInvariant()
    if ($call) {
        $note = (Get-RowField -Row $r -FieldName 'Note')
        $why = if (($note + '').Trim()) { $note } else { ("Matched " + $mt + ": " + $pat) }
        # If Test Result contains both POS and NEG tokens (multi-target), treat as MIXED for non-MTB.
        if ($isMixed -and ($call -in @('POS','NEG'))) {
            return [pscustomobject]@{ Call='MIXED'; Reason=('Mixed POS+NEG tokens (pattern matched ' + $call + ')') }
        }
        return [pscustomobject]@{ Call=$call; Reason=$why }
    }
}
```

}

if (-not ($ass -match ‘(?i)MTB’)) {
$u = ($tr + ‘’).ToUpperInvariant()
$u = [regex]::Replace($u, ‘\s+’, ’ ’).Trim()

```
$hasErr = ($u -match '\bINVALID\b') -or ($u -match 'NO\s+RESULT') -or ($u -match '\bERROR\b')
$hasNeg = ($u -match 'NOT\s+DETECTED') -or ($u -match '\bNEGATIVE\b')

$uNoNotDetected = ($u -replace 'NOT\s+DETECTED', '')
$hasPos = ($uNoNotDetected -match '\bDETECTED\b') -or ($uNoNotDetected -match '\bPOSITIVE\b')

if ($hasErr) {
    return [pscustomobject]@{ Call='ERROR'; Reason='Generic fallback: ERROR/INVALID/NO RESULT token' }
}
if ($hasPos) {
    if ($hasNeg) {
        return [pscustomobject]@{ Call='MIXED'; Reason='Generic fallback: Mixed POS+NEG tokens' }
    }
    return [pscustomobject]@{ Call='POS'; Reason='Generic fallback: DETECTED/POSITIVE token' }
}
if ($hasNeg) {
    return [pscustomobject]@{ Call='NEG'; Reason='Generic fallback: NOT DETECTED/NEGATIVE token' }
}
```

}

return [pscustomobject]@{ Call=‘UNKNOWN’; Reason=‘No pattern matched’ }

}
function Get-ExpectedCallDetailed {
param(
[Parameter(Mandatory)][object]$Row,
[Parameter(Mandatory=$false)][object[]]$Rules = @(),
[Parameter(Mandatory=$false)][object]$RegexCache = $null
)

if (-not $Rules) { $Rules = @() }

$sampleId = (Get-RowField -Row $Row -FieldName ‘Sample ID’)
$assay    = (Get-RowField -Row $Row -FieldName ‘Assay’)
$sid = ($sampleId + ‘’).Trim()
if (-not $sid) { return [pscustomobject]@{ Call=’’; Reason=‘Blank Sample ID’ } }

foreach ($r in $Rules) {
if (-not (Test-RuleEnabled $r)) { continue }
$ruleAssay = (Get-RowField -Row $r -FieldName ‘Assay’)
if (-not (Test-AssayMatch -RuleAssay $ruleAssay -RowAssay $assay)) { continue }

```
$mtype = (Get-RowField -Row $r -FieldName 'SampleIdMatchType')
$pat   = (Get-RowField -Row $r -FieldName 'SampleIdPattern')
if (-not ($pat + '').Trim()) { continue }

if (Match-TextFast -Text $sid -Pattern $pat -MatchType $mtype -RegexCache $RegexCache) {
    $call = ((Get-RowField -Row $r -FieldName 'Expected') + '').Trim().ToUpperInvariant()
    if ($call) {
        $note = (Get-RowField -Row $r -FieldName 'Note')
        $why = if (($note + '').Trim()) { $note } else { ("Matched " + $mtype + ": " + $pat) }
        return [pscustomobject]@{ Call=$call; Reason=$why }
    }
}
```

}

return [pscustomobject]@{ Call=’’; Reason=‘No expectation rule matched’ }

}

function Get-ExpectedTestTypeDerived {
param([Parameter(Mandatory)][string]$SampleId)

$parts = $SampleId.Split(’_’)
if ($parts.Count -ge 3) {
$tc = $parts[2]
switch -Regex ($tc) {
‘^0$’ { return ‘Negative Control 1’ }
‘^1$’ { return ‘Positive Control 1’ }
‘^2$’ { return ‘Positive Control 2’ }
‘^3$’ { return ‘Positive Control 3’ }
‘^4$’ { return ‘Positive Control 4’ }
‘^5$’ { return ‘Positive Control 5’ }
default { }
}
}
return ‘Specimen’

}

function Build-ErrorCodeLookup {
param([Parameter(Mandatory)][object[]]$ErrorCodes)
$lut = @{
Codes = @{}
Defaults = New-Object System.Collections.Generic.List[object]
NamedBlanks = New-Object System.Collections.Generic.List[object]
}

foreach ($r in $ErrorCodes) {
$code = ((Get-RowField -Row $r -FieldName ‘ErrorCode’) + ‘’).Trim()
$name = (Get-RowField -Row $r -FieldName ‘Name’)
$ret  = (Get-RowField -Row $r -FieldName ‘GeneratesRetest’)

```
if ($code -eq '####') {
    $lut.Defaults.Add([pscustomobject]@{ ErrorCode='####'; Name=$name; GeneratesRetest=$ret })
    continue
}

if (-not $code) {
    if (($name + '').Trim()) { $lut.NamedBlanks.Add([pscustomobject]@{ ErrorCode=''; Name=$name; GeneratesRetest=$ret }) }
    continue
}

if ($code -match '^\d{4,5}$') {
    $lut.Codes[$code] = [pscustomobject]@{ ErrorCode=$code; Name=$name; GeneratesRetest=$ret }
}
```

}

return $lut

}

function Get-ErrorInfo {
param(
[Parameter(Mandatory)][object]$Row,
[Parameter(Mandatory)][hashtable]$ErrorLut,
[Parameter(Mandatory)][string]$DelamPattern
)

$errTxt = (Get-RowField -Row $Row -FieldName ‘Error’)
$mpTxt  = (Get-RowField -Row $Row -FieldName ‘Max Pressure (PSI)’)

$code = ‘’
$hasErr = (($errTxt + ‘’).Trim().Length -gt 0)

if ($hasErr) {
if (($errTxt + ‘’) -match ‘(\d{4,5})’) { $code = $Matches[1] }
}

$name = ‘’
$retest = ‘’
$isKnownCode = $false

if ($hasErr) {
if ($code -and $ErrorLut.Codes.ContainsKey($code)) {
$name   = $ErrorLut.Codes[$code].Name
$retest = $ErrorLut.Codes[$code].GeneratesRetest
$isKnownCode = $true  # Felkoden finns i RuleBank = Minor Functional
} else {
# Felkoden finns INTE i RuleBank = Instrument Error
$isKnownCode = $false
$picked = $null
try {
foreach ($d in $ErrorLut.Defaults) {
if (($d.Name + ‘’) -match ‘(?i)Delamination’ -and ($errTxt + ‘’) -match $DelamPattern) {
$picked = $d; break
}
}
if (-not $picked -and $ErrorLut.Defaults.Count -gt 0) { $picked = $ErrorLut.Defaults[$ErrorLut.Defaults.Count - 1] }
} catch {}
if ($picked) {
$name   = $picked.Name
$retest = $picked.GeneratesRetest
}
}
}

$pressure = $null
try {
if (($mpTxt + ‘’).Trim()) { $pressure = [double]($mpTxt + ‘’) }
} catch {}

$pressureFlag = $false
if ($pressure -ne $null -and $pressure -ge 90) { $pressureFlag = $true }

# If pressure flag and there is a named blank describing it, surface the name

if ($pressureFlag -and -not $name) {
try {
foreach ($b in $ErrorLut.NamedBlanks) {
if (($b.Name + ‘’) -match ‘(?i)Max\s+Pressure’) {
$name = $b.Name
$retest = $b.GeneratesRetest
break
}
}
} catch {}
}

return [pscustomobject]@{
ErrorCode       = $code
ErrorName       = $name
GeneratesRetest = $retest
MaxPressure     = $pressure
PressureFlag    = $pressureFlag
IsKnownCode     = $isKnownCode
}

}

function Classify-Deviation {
param(
[AllowEmptyString()][string]$Expected,
[AllowEmptyString()][string]$Observed
)
$e = ($Expected + ‘’).Trim().ToUpperInvariant()
$o = ($Observed + ‘’).Trim().ToUpperInvariant()

if (-not $e) { return ‘UNKNOWN’ }
if ($o -eq ‘ERROR’) { return ‘ERROR’ }
if ($o -eq ‘UNKNOWN’ -or -not $o) { return ‘UNKNOWN’ }
if ($e -eq $o) { return ‘OK’ }
if ($o -eq ‘MIXED’) {
if ($e -eq ‘POS’) { return ‘FN’ }
if ($e -eq ‘NEG’) { return ‘FP’ }
return ‘MISMATCH’
}
if ($e -eq ‘POS’ -and $o -eq ‘NEG’) { return ‘FN’ }
if ($e -eq ‘NEG’ -and $o -eq ‘POS’) { return ‘FP’ }
return ‘MISMATCH’

}

function Split-CsvLineQuoted {
param(
[Parameter(Mandatory)][string]$Line,
[Parameter(Mandatory)][string]$Delimiter
)
$d = [regex]::Escape($Delimiter)
$rx = $d + ‘(?=(?:(?:[^”]*”){2})*[^”]*$)’
return [regex]::Split($Line, $rx)
}

function Get-HeaderFromTestSummaryFile {
param([Parameter(Mandatory)][string]$CsvPath)

if (-not (Test-Path -LiteralPath $CsvPath)) { return @() }

$delim = ‘,’
try { $delim = Get-CsvDelimiter -Path $CsvPath } catch {}

$lines = @()
try { $lines = Get-Content -LiteralPath $CsvPath -ErrorAction Stop } catch { return @() }

# Test Summary: header is line 8 (index 7)

if (-not $lines -or $lines.Count -lt 8) { return @() }
$hdrLine = $lines[7]
if (-not $hdrLine) { return @() }

$headers = Split-CsvLineQuoted -Line $hdrLine -Delimiter $delim
$headers = @($headers | ForEach-Object { (($_ + ‘’) -replace ‘^”|”$’,’’).Trim() })
return $headers

}

function Convert-FieldRowsToObjects {
param(
[Parameter(Mandatory)][object[]]$FieldRows,
[Parameter(Mandatory)][string[]]$Headers
)

$out = New-Object System.Collections.Generic.List[object]

foreach ($r in $FieldRows) {
if ($null -eq $r) { continue }
$arr = $r
if ($arr -isnot [object[]]) { continue }

```
$o = [ordered]@{}
$max = [Math]::Min($Headers.Count, $arr.Count)
for ($i=0; $i -lt $max; $i++) {
    $h = $Headers[$i]
    if (-not $h) { continue }
    $o[$h] = $arr[$i]
}
$out.Add([pscustomobject]$o)
```

}

return $out.ToArray()

}

function Get-MarkerValue {
param(
[Parameter(Mandatory)][pscustomobject]$RuleBank,
[Parameter(Mandatory)][string]$Assay,
[Parameter(Mandatory)][string]$MarkerType
)

# Memoized marker lookup (one-shot performance)

$aKey = (($Assay + ‘’).Trim())
if (-not $aKey) { $aKey = ‘(blank)’ }
$tKey = (($MarkerType + ‘’).Trim().ToUpperInvariant())
$mKey = ($aKey + ‘|’ + $tKey)

try {
if ($RuleBank.Compiled -and $RuleBank.Compiled.MarkerByAssayType) {
$mc = $RuleBank.Compiled.MarkerByAssayType
if ($mc.ContainsKey($mKey)) {
$v = $mc[$mKey]
if ($v -eq ‘**MISS**’) { return ‘’ }
return ($v + ‘’)
}
}
} catch {}

foreach ($r in $RuleBank.SampleIdMarkers) {
if (-not (Test-RuleEnabled $r)) { continue }

```
$ap = ((Get-RowField -Row $r -FieldName 'AssayPattern') + '').Trim()
if (-not (Test-AssayMatch -RuleAssay $ap -RowAssay $Assay)) { continue }

$mt = ((Get-RowField -Row $r -FieldName 'MarkerType') + '').Trim()
if (-not $mt) { continue }
if ($mt -ine $MarkerType) { continue }

$m = ((Get-RowField -Row $r -FieldName 'Marker') + '').Trim()
try { if ($RuleBank.Compiled -and $RuleBank.Compiled.MarkerByAssayType) { $RuleBank.Compiled.MarkerByAssayType[$mKey] = $m } } catch {}
return $m
```

}
try { if ($RuleBank.Compiled -and $RuleBank.Compiled.MarkerByAssayType) { $RuleBank.Compiled.MarkerByAssayType[$mKey] = ‘**MISS**’ } } catch {}

return ‘’

}

function Get-IntMarkerValue {
param(
[Parameter(Mandatory)][pscustomobject]$RuleBank,
[Parameter(Mandatory)][string]$Assay,
[Parameter(Mandatory)][string]$MarkerType,
[Parameter(Mandatory)][int]$Default
)
$v = Get-MarkerValue -RuleBank $RuleBank -Assay $Assay -MarkerType $MarkerType
if (-not $v) { return $Default }
try { return [int]$v } catch { return $Default }
}

function Get-ParityConfigForAssay {
param(
[Parameter(Mandatory)][pscustomobject]$RuleBank,
[Parameter(Mandatory)][string]$Assay
)

$cfg = [ordered]@{
UseParity = $false
CartridgeField = ‘Cartridge S/N’
TokenIndex = 3
XChar = ‘X’
PlusChar = ‘+’
NumericRatioThreshold = 0.60
DelaminationMarkerType = ‘DelaminationCodeRegex’
DelamRegex = ‘D\d{1,2}[A-Z]?’
ValidSuffixRegex = ‘X|+’
SampleTypeCodeTokenIndex = 2
SampleNumberTokenIndex = 3
}

$delam = Get-MarkerValue -RuleBank $RuleBank -Assay $Assay -MarkerType ‘DelaminationCodeRegex’
if ($delam) { $cfg.DelamRegex = $delam }

$suffix = Get-MarkerValue -RuleBank $RuleBank -Assay $Assay -MarkerType ‘SuffixChars’
if ($suffix) {
while ($suffix -like ‘*\*’) { $suffix = $suffix.Replace(’\’,’') }
$cfg.ValidSuffixRegex = $suffix
}

$stIdx = Get-IntMarkerValue -RuleBank $RuleBank -Assay $Assay -MarkerType ‘SampleTypeCodeTokenIndex’ -Default 2
$snIdx = Get-IntMarkerValue -RuleBank $RuleBank -Assay $Assay -MarkerType ‘SampleNumberTokenIndex’ -Default 3
$cfg.SampleTypeCodeTokenIndex = $stIdx
$cfg.SampleNumberTokenIndex = $snIdx

foreach ($r in $RuleBank.ParityCheckConfig) {
if (-not (Test-RuleEnabled $r)) { continue }

```
$ap = ((Get-RowField -Row $r -FieldName 'AssayPattern') + '').Trim()
if (-not (Test-AssayMatch -RuleAssay $ap -RowAssay $Assay)) { continue }

# First match wins since RuleBank.ParityCheckConfig is priority-sorted DESC.
$cfg.UseParity = $true

$cf = ((Get-RowField -Row $r -FieldName 'CartridgeField') + '').Trim()
if ($cf) { $cfg.CartridgeField = $cf }

$ti = ((Get-RowField -Row $r -FieldName 'SampleTokenIndex') + '').Trim()
if ($ti) { try { $cfg.TokenIndex = [int]$ti } catch {} }

$sx = ((Get-RowField -Row $r -FieldName 'SuffixX') + '').Trim()
if ($sx) { $cfg.XChar = $sx.Substring(0,1).ToUpperInvariant() }

$sp = ((Get-RowField -Row $r -FieldName 'SuffixPlus') + '').Trim()
if ($sp) { $cfg.PlusChar = $sp.Substring(0,1) }

$dmt = ((Get-RowField -Row $r -FieldName 'DelaminationMarkerType') + '').Trim()
if ($dmt) { $cfg.DelaminationMarkerType = $dmt }

$minPct = ((Get-RowField -Row $r -FieldName 'MinValidCartridgeSNPercent') + '').Trim()
if ($minPct) {
    try { $cfg.NumericRatioThreshold = ([double]$minPct) / 100.0 } catch {}
}

break
```

}

if ($cfg.DelaminationMarkerType) {
$delam2 = Get-MarkerValue -RuleBank $RuleBank -Assay $Assay -MarkerType $cfg.DelaminationMarkerType
if ($delam2) { $cfg.DelamRegex = $delam2 }
}

return [pscustomobject]$cfg

}

function Get-ControlCodeFromRow {
param(
[Parameter(Mandatory)][object]$Row,
[Parameter(Mandatory)][int]$SampleTypeCodeTokenIndex
)

$sid = (Get-RowField -Row $Row -FieldName ‘Sample ID’)
if (($sid + ‘’).Trim()) {
$parts = ($sid + ‘’).Split(’_’)
if ($parts.Count -gt $SampleTypeCodeTokenIndex) {
$cc = ($parts[$SampleTypeCodeTokenIndex] + ‘’).Trim()
if ($cc -match ‘^\d+$’) { return $cc }
}
if ($parts.Count -ge 3) {
$cc2 = ($parts[2] + ‘’).Trim()
if ($cc2 -match ‘^\d+$’) { return $cc2 }
}
}

$tt = (Get-RowField -Row $Row -FieldName ‘Test Type’)
if (($tt + ‘’) -match ‘(?i)Negative\s+Control’) { return ‘0’ }
if (($tt + ‘’) -match ‘(?i)Positive\s+Control\s+(\d+)’) { return $Matches[1] }

return ‘’

}

function Get-SampleTokenAndBase {
param(
[Parameter(Mandatory)][string]$SampleId,
[Parameter(Mandatory)][int]$TokenIndex,
[Parameter(Mandatory)][string]$DelamPattern,
[Parameter(Mandatory)][string]$ValidSuffixRegex,
[Parameter(Mandatory)][string]$XChar,
[Parameter(Mandatory)][string]$PlusChar
)

$tok = ‘’
$base = ‘’

$parts = $SampleId.Split(’_’)
if ($parts.Count -gt $TokenIndex) {
$tok = ($parts[$TokenIndex] + ‘’).Trim()
}

if (-not $tok) { return [pscustomobject]@{ SampleToken=’’; BaseToken=’’; ActualSuffix=’’; SampleNum=’’; SampleNumRaw=’’; } }

# strip trailing delamination code if present INSIDE token

$rx = “([_-]?(?:” + $DelamPattern + “))$”
try {
$base = [regex]::Replace($tok, $rx, ‘’, [System.Text.RegularExpressions.RegexOptions]::IgnoreCase)
} catch {
$base = $tok
}

$base = ($base + ‘’).Trim()

$act = ‘’
if ($base.Length -ge 1) {
$last = $base.Substring($base.Length - 1, 1)
if ($last -match (’^(?:’ + $ValidSuffixRegex + ‘)$’)) {
$u = $last.ToUpperInvariant()
if ($u -eq $XChar.ToUpperInvariant()) { $act = $XChar.ToUpperInvariant() }
elseif ($last -eq $PlusChar) { $act = $PlusChar }
else { $act = $u }
}
}

$numRaw = ‘’
$num = ‘’
if ($base -match ‘^(\d{1,4})’) {
$numRaw = $Matches[1]
$num = $numRaw
}

return [pscustomobject]@{ SampleToken=$tok; BaseToken=$base; ActualSuffix=$act; SampleNum=$num; SampleNumRaw=$numRaw }

}

function Parse-SampleIdBasic {
param(
[Parameter(Mandatory)][string]$SampleId,
[Parameter(Mandatory)][string]$DelamRegex,
[Parameter(Mandatory)][string]$XChar,
[Parameter(Mandatory)][string]$PlusChar
)

$out = [ordered]@{
Prefix = ‘’
BagNo = ‘’
SampleCode = ‘’
RunToken = ‘’
RunNoRaw = ‘’
RunNo = ‘’
RunSuffix = ‘’
ReplacementLevel = 0
DelamPresent = $false
DelamToken = ‘’
DelamCodes = @()
}

$sid = ($SampleId + ‘’).Trim()
if (-not $sid) { return [pscustomobject]$out }

$parts = $sid.Split(’_’)
if ($parts.Count -ge 1) { $out.Prefix = (($parts[0] + ‘’).Trim()).ToUpperInvariant() }
if ($parts.Count -ge 2) { $out.BagNo = (($parts[1] + ‘’).Trim()) }
if ($parts.Count -ge 3) { $out.SampleCode = (($parts[2] + ‘’).Trim()) }
if ($parts.Count -ge 4) { $out.RunToken = (($parts[3] + ‘’).Trim()) }

if ($parts.Count -ge 5) {
$dt = (($parts[4] + ‘’).Trim())
if ($dt) {
$out.DelamToken = $dt
if ($dt -match ‘^(?i)D’) { $out.DelamPresent = $true }
}
}

if (-not $out.DelamPresent -and $DelamRegex) {
try {
$rx = ‘(?i)(?:^|[*-])(’ + $DelamRegex + ’)(?:$|[,;* -])’
if ([regex]::IsMatch($sid, $rx)) { $out.DelamPresent = $true }
} catch {}
}

$rt = ($out.RunToken + ‘’).Trim()
if ($rt.Length -ge 1) {
$last = $rt.Substring($rt.Length - 1, 1)
if ($last -eq $PlusChar -or $last.ToUpperInvariant() -eq $XChar.ToUpperInvariant()) {
$out.RunSuffix = $last.ToUpperInvariant()
$core = $rt.Substring(0, $rt.Length - 1)

```
    if ($core -match '(?i)(A{1,3})$') {
        $a = $Matches[1]
        $out.ReplacementLevel = $a.Length
        $core = $core.Substring(0, $core.Length - $a.Length)
    }

    if ($core -match '^(\d{1,4})') {
        $out.RunNoRaw = $Matches[1]
        $out.RunNo = $out.RunNoRaw
    }
} else {

    if ($rt -match '^(\d{1,4})') {
        $out.RunNoRaw = $Matches[1]
        $out.RunNo = $out.RunNoRaw
    }
}
```

}

# Parse delamination codes list (if present)

if ($out.DelamToken) {
$codes = @()
foreach ($c in ($out.DelamToken -split ‘,’)) {
$t = ($c + ‘’).Trim()
if ($t) { $codes += $t }
}
$out.DelamCodes = $codes
}

return [pscustomobject]$out

}

function Get-SampleNumberRuleForRow {
param(
[Parameter(Mandatory)][string]$Assay,
[Parameter(Mandatory)][string]$ControlCode,
[Parameter(Mandatory=$false)][string]$BagNo = ‘’,
[Parameter(Mandatory)][object[]]$Rules
)

foreach ($r in $Rules) {
if (-not (Test-RuleEnabled $r)) { continue }

```
$ap = ((Get-RowField -Row $r -FieldName 'AssayPattern') + '').Trim()
if (-not (Test-AssayMatch -RuleAssay $ap -RowAssay $Assay)) { continue }

$bp = ((Get-RowField -Row $r -FieldName 'BagNoPattern') + '').Trim()
if ($bp) {
    $bn = ($BagNo + '').Trim()
    if (-not $bn) { continue }
    $bagOk = $false
    try { $bagOk = [regex]::IsMatch($bn, $bp, [System.Text.RegularExpressions.RegexOptions]::IgnoreCase) } catch { $bagOk = $false }
    if (-not $bagOk) { continue }
}

$cc = ((Get-RowField -Row $r -FieldName 'SampleTypeCode') + '').Trim()
if (-not $cc -or $cc -eq '*') { return $r }
if ($ControlCode -and ($cc -eq $ControlCode)) { return $r }
```

}

return $null

}

function Invoke-RuleEngine {
param(
[Parameter(Mandatory=$true)][AllowEmptyCollection()][object[]]$CsvObjects,
[Parameter(Mandatory)][pscustomobject]$RuleBank,
[Parameter(Mandatory=$false)][string]$CsvPath
)

if (-not $CsvObjects -or $CsvObjects.Count -eq 0) {
return [pscustomobject]@{ Rows=@(); Summary=[pscustomobject]@{ Total=0; ObservedCounts=@{}; DeviationCounts=@{}; RetestYes=0 }; TopDeviations=@() }
}

if (-not $RuleBank) { throw ‘RuleEngine: RuleBank is null.’ }

Test-RuleBankIntegrity -RuleBank $RuleBank -Source ‘runtime’

try { $RuleBank = Compile-RuleBank -RuleBank $RuleBank } catch {}

$needsConvert = $false
try {
if ($CsvObjects[0] -is [object[]]) { $needsConvert = $true }
else {
$p1 = $CsvObjects[0].PSObject.Properties.Match(‘Sample ID’)
if ($p1.Count -eq 0) { $needsConvert = $true }
}
} catch { $needsConvert = $true }

if ($needsConvert) {
if (-not $CsvPath) { throw ‘RuleEngine: CsvPath is required to convert field-array rows to objects.’ }
$hdr = Get-HeaderFromTestSummaryFile -CsvPath $CsvPath
if (-not $hdr -or $hdr.Count -lt 5) { throw ‘RuleEngine: Could not read CSV header (line 8).’ }
$CsvObjects = Convert-FieldRowsToObjects -FieldRows $CsvObjects -Headers $hdr
if (-not $CsvObjects -or $CsvObjects.Count -eq 0) {
return [pscustomobject]@{ Rows=@(); Summary=[pscustomobject]@{ Total=0; ObservedCounts=@{}; DeviationCounts=@{}; RetestYes=0 }; TopDeviations=@() }
}
}

$byAssay = @{}
foreach ($row in $CsvObjects) {
$a = (Get-RowField -Row $row -FieldName ‘Assay’)
$key = (($a + ‘’).Trim())
if (-not $key) { $key = ‘(blank)’ }
if (-not $byAssay.ContainsKey($key)) { $byAssay[$key] = New-Object System.Collections.Generic.List[object] }
$byAssay[$key].Add($row)
}

$results = New-Object System.Collections.Generic.List[object]
$errLut = Build-ErrorCodeLookup -ErrorCodes $RuleBank.ErrorCodes

foreach ($assayKey in $byAssay.Keys) {
$group = $byAssay[$assayKey]
if (-not $group -or $group.Count -eq 0) { continue }

```
$parCfg = Get-ParityConfigForAssay -RuleBank $RuleBank -Assay $assayKey

$patternsForAssay = @(Get-ResultCallPatternsForAssay -RuleBank $RuleBank -Assay $assayKey)
$expectForAssay   = @(Get-ExpectationRulesForAssay -RuleBank $RuleBank -Assay $assayKey)
if (-not $patternsForAssay) { $patternsForAssay = @() }
if (-not $expectForAssay) { $expectForAssay = @() }

$regexCache = $null
try { if ($RuleBank.Compiled -and $RuleBank.Compiled.RegexCache) { $regexCache = $RuleBank.Compiled.RegexCache } } catch { $regexCache = $null }
$delamPattern = $parCfg.DelamRegex
$validSuffix = $parCfg.ValidSuffixRegex

#
# Collect suffix statistics and parity pairs for the current assay group.
#
# A numeric list of Cartridge S/N last-digit values (for parity) and
# a list of parity pairs containing the last digit and the actual suffix.
# Also maintain counts of observed suffixes (X/+), regardless of whether
# parsing the sample ID produces a suffix. If a suffix isn't found via
# Get-SampleTokenAndBase and the Sample ID contains "X" or "+", we
# treat that as the actual suffix (similar to the VBA macro logic).
#
$numeric = New-Object System.Collections.Generic.List[long]
$parityPairs = New-Object System.Collections.Generic.List[object]
$suffixCounts = @{}
$suffixCounts[$parCfg.XChar.ToUpperInvariant()] = 0
$suffixCounts[$parCfg.PlusChar] = 0

foreach ($row in $group) {
    # Determine the last character of the Cartridge S/N as a possible numeric for parity.
    $sn = (Get-RowField -Row $row -FieldName $parCfg.CartridgeField)
    $snLastChar = ''
    if (($sn + '') -ne '') {
        $snStr = ($sn + '').Trim()
        if ($snStr.Length -ge 1) { $snLastChar = $snStr.Substring($snStr.Length - 1, 1) }
    }
    $snN = $null
    if ($snLastChar -match '[0-9]') {
        try { $snN = [int]$snLastChar } catch { $snN = $null }
        if ($snN -ne $null) { try { $numeric.Add([long]$snN) } catch {} }
    }

    # Determine the actual suffix for this row.
    $sid = (Get-RowField -Row $row -FieldName 'Sample ID')
    $actSuffix = ''
    if (($sid + '').Trim()) {
        $t = Get-SampleTokenAndBase -SampleId ($sid + '') -TokenIndex $parCfg.TokenIndex -DelamPattern $delamPattern -ValidSuffixRegex $validSuffix -XChar $parCfg.XChar -PlusChar $parCfg.PlusChar
        if ($t.ActualSuffix) {
            $actSuffix = $t.ActualSuffix
        } else {
            # Fallback: when a delamination token exists ("_D" in SampleID)
            # or parsing didn't find a suffix, inspect the SampleID for X or +.
            if ((($sid + '') -match 'X')) { $actSuffix = $parCfg.XChar.ToUpperInvariant() }
            elseif ((($sid + '') -match '\+')) { $actSuffix = $parCfg.PlusChar }
        }
    }
    # Track suffix counts and parity pairs if applicable.
    if ($actSuffix -and $suffixCounts.ContainsKey($actSuffix)) {
        $suffixCounts[$actSuffix]++
        if ($snN -ne $null) {
            try { $parityPairs.Add([pscustomobject]@{ SN = $snN; Sfx = $actSuffix }) } catch {}
        }
    }
}

$numRatio = 0.0
try { $numRatio = [double]$numeric.Count / [double]$group.Count } catch {}

$useParity = ($parCfg.UseParity -and $numeric.Count -gt 0 -and $numRatio -ge $parCfg.NumericRatioThreshold)

# Determine parity mapping and majority suffix for this assay group.
$minSn      = $null
$parityForX = $null
if ($useParity) {
    # Determine the smallest SN digit (not used but kept for backward compatibility)
    try { $minSn = ($numeric | Measure-Object -Minimum).Minimum } catch { $minSn = $null }
    if ($parityPairs) {
        # Compute how well each mapping fits the observed suffixes.
        $map0Matches = 0; $map1Matches = 0; $totalPairs = 0
        foreach ($pp in $parityPairs) {
            $snTmp = $null
            try { $snTmp = [int]$pp.SN } catch { $snTmp = $null }
            if ($snTmp -eq $null) { continue }
            $sfx = (($pp.Sfx + '')).Trim()
            if (-not $sfx) { continue }
            # Mapping0: even -> X, odd -> +
            $exp0 = if (([int]($snTmp % 2)) -eq 0) { $parCfg.XChar.ToUpperInvariant() } else { $parCfg.PlusChar }
            # Mapping1: odd -> X, even -> +
            $exp1 = if (([int]($snTmp % 2)) -eq 1) { $parCfg.XChar.ToUpperInvariant() } else { $parCfg.PlusChar }
            if ($sfx -eq $exp0) { $map0Matches++ }
            if ($sfx -eq $exp1) { $map1Matches++ }
            $totalPairs++
        }
        if ($totalPairs -gt 0) {
            # Choose mapping with the most matches
            if ($map0Matches -ge $map1Matches) { $parityForX = 0 } else { $parityForX = 1 }
            # Require that the winning mapping meets the numeric ratio threshold; otherwise disable parity.
            $maxMatches = [double]([Math]::Max($map0Matches, $map1Matches))
            $ratio      = $maxMatches / [double]$totalPairs
            if ($ratio -lt $parCfg.NumericRatioThreshold) {
                $parityForX = $null
            }
        }
    }
    if ($parityForX -eq $null) { $useParity = $false }
}

# Determine majority suffix when parity isn't used or mapping was inconclusive.
$majSuffix = ''
if (-not $useParity) {
    $xCount = $suffixCounts[$parCfg.XChar.ToUpperInvariant()]
    $pCount = $suffixCounts[$parCfg.PlusChar]
    if ($xCount -gt $pCount) { $majSuffix = $parCfg.XChar.ToUpperInvariant() }
    elseif ($pCount -gt $xCount) { $majSuffix = $parCfg.PlusChar }
    # If tie, leave majSuffix empty.
}

foreach ($row in $group) {
    try {
        $obsD = Get-ObservedCallDetailed -Row $row -Patterns $patternsForAssay -RegexCache $regexCache
        $expD = Get-ExpectedCallDetailed -Row $row -Rules $expectForAssay -RegexCache $regexCache

        $sid = (Get-RowField -Row $row -FieldName 'Sample ID')
        $assay = (Get-RowField -Row $row -FieldName 'Assay')

        $expTT = ''
        if (($sid + '').Trim()) { $expTT = Get-ExpectedTestTypeDerived -SampleId ($sid + '') }

        $expCall = ($expD.Call + '').Trim().ToUpperInvariant()
        $expSrc = 'RULE'
        if (-not $expCall) {

            $cc2 = Get-ControlCodeFromRow -Row $row -SampleTypeCodeTokenIndex $parCfg.SampleTypeCodeTokenIndex
            if ($cc2 -match '^\d+$') {
                $ccInt = -1
                try { $ccInt = [int]$cc2 } catch { $ccInt = -1 }

                if ($ccInt -eq 0) { $expCall = 'NEG'; $expSrc = 'CONTROL_CODE' }
                elseif ($ccInt -ge 1 -and $ccInt -le 5) { $expCall = 'POS'; $expSrc = 'CONTROL_CODE' }
            }

            if (-not $expCall) {
                $tt = (Get-RowField -Row $row -FieldName 'Test Type')
                $tt2 = ($tt + '')
                if ($tt2 -match '(?i)Negative\s+Control') { $expCall = 'NEG'; $expSrc = 'TESTTYPE' }
                elseif ($tt2 -match '(?i)Positive\s+Control') { $expCall = 'POS'; $expSrc = 'TESTTYPE' }
            }
        }

        $errInfo = Get-ErrorInfo -Row $row -ErrorLut $errLut -DelamPattern $delamPattern
        $dev = Classify-Deviation -Expected $expCall -Observed $obsD.Call

        $tokInfo = [pscustomobject]@{ SampleToken=''; BaseToken=''; ActualSuffix=''; SampleNum=''; SampleNumRaw='' }
        if (($sid + '').Trim()) {
            $tokInfo = Get-SampleTokenAndBase -SampleId ($sid + '') -TokenIndex $parCfg.TokenIndex -DelamPattern $delamPattern -ValidSuffixRegex $validSuffix -XChar $parCfg.XChar -PlusChar $parCfg.PlusChar

        $sidBasic = [pscustomobject]@{ Prefix=''; BagNo=''; SampleCode=''; RunToken=''; RunNoRaw=''; RunNo=''; RunSuffix=''; ReplacementLevel=0; DelamPresent=$false; DelamToken=''; DelamCodes=@() }
        if (($sid + '').Trim()) {
            $sidBasic = Parse-SampleIdBasic -SampleId ($sid + '') -DelamRegex $delamPattern -XChar $parCfg.XChar -PlusChar $parCfg.PlusChar
        }

        }

        # Reset per-row suffix variables
        $expectedSuffix = ''
        $suffixSource = ''
        $suffixCheck = ''

        # Always compute a numeric SN from the last character of the cartridge serial
        $snVal = (Get-RowField -Row $row -FieldName $parCfg.CartridgeField)
        $snNum = $null
        if (($snVal + '') -ne '') {
            $snStr = ($snVal + '').Trim()
            if ($snStr.Length -ge 1) {
                $lastChar = $snStr.Substring($snStr.Length - 1, 1)
                if ($lastChar -match '[0-9]') {
                    try { $snNum = [int]$lastChar } catch { $snNum = $null }
                }
            }
        }

        # Determine expected suffix using parity mapping or majority regardless of whether actual suffix exists.
        if ($useParity -and $snNum -ne $null -and $parityForX -ne $null) {
            # Use parity mapping: if remainder matches mapping, expect X; else expect +
            $expS = if (([int]($snNum % 2)) -eq $parityForX) { $parCfg.XChar.ToUpperInvariant() } else { $parCfg.PlusChar }
            $expectedSuffix = $expS
            $suffixSource = 'PARITY'
        } elseif ($majSuffix) {
            # Fall back to majority suffix when parity is not used or inconclusive
            $expectedSuffix = $majSuffix
            $suffixSource = 'MAJORITY'
        }

        # Compute suffix check status
        if ($expectedSuffix) {
            if ($tokInfo.ActualSuffix) {
                # If actual suffix exists, compare with expected
                $suffixCheck = if ($tokInfo.ActualSuffix -eq $expectedSuffix) { 'OK' } else { 'BAD' }
            } else {
                # If actual suffix missing but expected exists, mark as missing
                $suffixCheck = 'MISSING'
            }
        }

        $sampleNum = ''
        $sampleNumRaw = ''
        $sampleNumOk = ''
        $sampleNumWhy = ''

        $cc = Get-ControlCodeFromRow -Row $row -SampleTypeCodeTokenIndex $parCfg.SampleTypeCodeTokenIndex

        $rule = $null
        try { $rule = Get-SampleNumberRuleForRowCached -RuleBank $RuleBank -Assay $assay -ControlCode $cc -BagNo $sidBasic.BagNo } catch {}

        $snTokIndex = $parCfg.SampleNumberTokenIndex
        if ($rule) {
            $idxTxt = ((Get-RowField -Row $rule -FieldName 'SampleNumberTokenIndex') + '').Trim()
            if ($idxTxt) { try { $snTokIndex = [int]$idxTxt } catch {} }
        }

        $snInfo = [pscustomobject]@{ SampleToken=''; BaseToken=''; ActualSuffix=''; SampleNum=''; SampleNumRaw='' }
        if (($sid + '').Trim()) {
            $snInfo = Get-SampleTokenAndBase -SampleId ($sid + '') -TokenIndex $snTokIndex -DelamPattern $delamPattern -ValidSuffixRegex $validSuffix -XChar $parCfg.XChar -PlusChar $parCfg.PlusChar
        }
        $sampleNum = $snInfo.SampleNum
        $sampleNumRaw = $snInfo.SampleNumRaw

        if ($rule) {
            $rxTxt = ((Get-RowField -Row $rule -FieldName 'SampleNumberRegex') + '').Trim()
            $minTxt = ((Get-RowField -Row $rule -FieldName 'SampleNumberMin') + '').Trim()
            $maxTxt = ((Get-RowField -Row $rule -FieldName 'SampleNumberMax') + '').Trim()
            $padTxt = ((Get-RowField -Row $rule -FieldName 'SampleNumberPad') + '').Trim()

            $min = 0; $max = 0; $pad = 0
            try { $min = [int]$minTxt } catch {}
            try { $max = [int]$maxTxt } catch {}
            try { $pad = [int]$padTxt } catch {}

            if (-not $sampleNum) {
                $sampleNumOk = 'NO'
                $sampleNumWhy = 'No sample number'
            } else {
                $numInt = 0
                try { $numInt = [int]$sampleNum } catch { $numInt = 0 }

                $rxOk = $true
                if ($rxTxt) {
                    try { $rxOk = [regex]::IsMatch(($snInfo.BaseToken + ''), $rxTxt, [System.Text.RegularExpressions.RegexOptions]::IgnoreCase) } catch { $rxOk = $true }
                }

                $padOk = $true
                if ($pad -gt 0 -and ($sampleNumRaw + '').Length -ne $pad) { $padOk = $false }

                if ($rxOk -and $padOk -and $min -gt 0 -and $max -gt 0 -and $numInt -ge $min -and $numInt -le $max) {
                    $sampleNumOk = 'YES'
                } else {
                    $sampleNumOk = 'NO'
                    $sampleNumWhy = 'Out of range/regex/pad'
                }
            }
        }

        $results.Add([pscustomobject]@{
            SampleId         = $sid
            SamplePrefix     = $sidBasic.Prefix
            BagNo            = $sidBasic.BagNo
            SampleCode       = $sidBasic.SampleCode
            RunNo            = $sidBasic.RunNo
            RunNoRaw         = $sidBasic.RunNoRaw
            RunSuffix        = $sidBasic.RunSuffix
            ReplacementLevel = $sidBasic.ReplacementLevel
            DelamPresent     = $sidBasic.DelamPresent
            DelamCodes       = ($sidBasic.DelamCodes -join ',')
            CartridgeSN      = (Get-RowField -Row $row -FieldName $parCfg.CartridgeField)
            Assay            = $assay
            AssayVersion     = (Get-RowField -Row $row -FieldName 'Assay Version')
            ReagentLotId     = (Get-RowField -Row $row -FieldName 'Reagent Lot ID')
            TestType         = (Get-RowField -Row $row -FieldName 'Test Type')
            ExpectedTestType = $expTT
            ControlCode      = $cc
            SampleToken      = $tokInfo.SampleToken
            BaseToken        = $tokInfo.BaseToken
            ActualSuffix     = $tokInfo.ActualSuffix
            ExpectedSuffix   = $expectedSuffix
            SuffixCheck      = $suffixCheck
            SuffixSource     = $suffixSource
            SampleNum        = $sampleNum
            SampleNumOK      = $sampleNumOk
            SampleNumWhy     = $sampleNumWhy
            Status           = (Get-RowField -Row $row -FieldName 'Status')
            TestResult       = (Get-RowField -Row $row -FieldName 'Test Result')
            ErrorText        = (Get-RowField -Row $row -FieldName 'Error')
            MaxPressure      = $errInfo.MaxPressure
            PressureFlag     = $errInfo.PressureFlag
            ErrorCode        = $errInfo.ErrorCode
            ErrorName        = $errInfo.ErrorName
            IsKnownCode      = $errInfo.IsKnownCode
            GeneratesRetest  = $errInfo.GeneratesRetest
            ObservedCall     = $obsD.Call
            ObservedWhy      = $obsD.Reason
            ExpectedCall     = $expCall
            ExpectedWhy      = $expD.Reason
            ExpectedSource   = $expSrc
            Deviation        = $dev
            ModuleSN         = (Get-RowField -Row $row -FieldName 'Module S/N')
            StartTime        = (Get-RowField -Row $row -FieldName 'Start Time')
            RuleFlags        = ''
        })

    } catch {
        $sid2 = ''
        try { $sid2 = (Get-RowField -Row $row -FieldName 'Sample ID') } catch {}
        throw ("RuleEngine row-fel (Sample ID=" + $sid2 + "): " + $_.Exception.Message)
    }
}
```

}

function _Append-RuleFlag {
param([pscustomobject]$row, [string]$flag)
$f = (($row.RuleFlags + ‘’)).Trim()
if (-not $f) { $row.RuleFlags = $flag; return }
$parts = $f.Split(’|’)
if ($parts -contains $flag) { return }
$row.RuleFlags = ($f + ‘|’ + $flag)
}

$distinctAssays = @($results | ForEach-Object { ($*.Assay + ‘’).Trim() } | Where-Object { $* } | Sort-Object -Unique)
$distinctAssayVersions = @($results | ForEach-Object { ($*.AssayVersion + ‘’).Trim() } | Where-Object { $* } | Sort-Object -Unique)
$distinctReagentLots = @($results | ForEach-Object { ($*.ReagentLotId + ‘’).Trim() } | Where-Object { $* } | Sort-Object -Unique)

$majorAssay = ‘’
if ($distinctAssays.Count -gt 1) {
try { $majorAssay = ($results | Group-Object Assay | Sort-Object Count -Descending | Select-Object -First 1).Name } catch {}
foreach ($r in $results) {
$a = ((($r.Assay + ‘’)).Trim())
if ($majorAssay -and $a -and $a -ne $majorAssay) { _Append-RuleFlag -row $r -flag ‘DQ_ASSAY_OUTLIER’ }
}
}

$majorVer = ‘’
if ($distinctAssayVersions.Count -gt 1) {
try { $majorVer = ($results | Group-Object AssayVersion | Sort-Object Count -Descending | Select-Object -First 1).Name } catch {}
foreach ($r in $results) {
$v = ((($r.AssayVersion + ‘’)).Trim())
if ($majorVer -and $v -and $v -ne $majorVer) { _Append-RuleFlag -row $r -flag ‘DQ_ASSAYVER_OUTLIER’ }
}
}

$dupSample = @($results | Where-Object { (($*.SampleId + ‘’).Trim()) } | Group-Object SampleId | Where-Object { $*.Count -gt 1 })
if ($dupSample.Count -gt 0) {
$dupSet = @{}
foreach ($g in $dupSample) { $dupSet[$g.Name] = $true }
foreach ($r in $results) {
$sid = ((($r.SampleId + ‘’)).Trim())
if ($sid -and $dupSet.ContainsKey($sid)) { _Append-RuleFlag -row $r -flag ‘DQ_DUP_SAMPLEID’ }
}
}

$dupCart = @($results | Where-Object { (($*.CartridgeSN + ‘’).Trim()) } | Group-Object CartridgeSN | Where-Object { $*.Count -gt 1 })
if ($dupCart.Count -gt 0) {
$dupSet = @{}
foreach ($g in $dupCart) { $dupSet[$g.Name] = $true }
foreach ($r in $results) {
$csn = ((($r.CartridgeSN + ‘’)).Trim())
if ($csn -and $dupSet.ContainsKey($csn)) { _Append-RuleFlag -row $r -flag ‘DQ_DUP_CARTSN’ }
}
}

$useStrictTestType = $false
try {
$ttAll = @($results | Where-Object { (($*.ExpectedTestType + ‘’)).Trim() -and (($*.TestType + ‘’)).Trim() })
$ttCtl = @($ttAll | Where-Object { (($*.ExpectedTestType + ‘’)).Trim().ToUpperInvariant() -ne ‘SPECIMEN’ })
if ($ttCtl.Count -ge 5) {
$ttMatch = @($ttCtl | Where-Object { (($*.TestType + ‘’)).Trim().ToUpperInvariant() -eq (($_.ExpectedTestType + ‘’)).Trim().ToUpperInvariant() }).Count
$ttRate = 0.0
try { $ttRate = [double]$ttMatch / [double]$ttCtl.Count } catch { $ttRate = 0.0 }
if ($ttRate -ge 0.80) { $useStrictTestType = $true }
}
} catch { $useStrictTestType = $false }

if ($useStrictTestType) {
foreach ($r in $results) {
$rf = ((($r.RuleFlags + ‘’)).Trim())
if ($rf) {
$p = $rf.Split(’|’)
if ($p -contains ‘DQ_ASSAY_OUTLIER’ -or $p -contains ‘DQ_ASSAYVER_OUTLIER’) { continue }
}

```
    $act = ((($r.TestType + '')).Trim())
    $ass = ((($r.Assay + '')).Trim())

    $pol = $null
    try { $pol = Get-TestTypePolicyForAssayCached -RuleBank $RuleBank -Assay $ass } catch { $pol = $null }

    if ($pol) {
        $allowed = @()
        $raw = (($pol.AllowedTestTypes + '')).Trim()
        if ($raw) {

            if ($raw -like '*|*') { $allowed = @($raw.Split('|') | ForEach-Object { ($_ + '').Trim() } | Where-Object { $_ }) }
            else { $allowed = @($raw.Split(',') | ForEach-Object { ($_ + '').Trim() } | Where-Object { $_ }) }
        }

        if ($allowed -and ($allowed | Where-Object { $_ -match 'Control' }).Count -gt 0) {
            if (-not ($allowed | Where-Object { $_ -ieq 'Specimen' })) { $allowed += 'Specimen' }
        }

        if (-not $allowed -or -not $act) {
            _Append-RuleFlag -row $r -flag 'TESTTYPE_MISMATCH'
        } else {
            $ok = $false
            foreach ($t in $allowed) {
                if ($act.ToUpperInvariant() -eq ($t + '').Trim().ToUpperInvariant()) { $ok = $true; break }
            }
            if (-not $ok) { _Append-RuleFlag -row $r -flag 'TESTTYPE_MISMATCH' }
        }
    } else {

        $exp = ((($r.ExpectedTestType + '')).Trim())
        if ($act -and $exp -and ($act.ToUpperInvariant() -ne $exp.ToUpperInvariant())) {
            _Append-RuleFlag -row $r -flag 'TESTTYPE_MISMATCH'
        }
    }
}
```

}

foreach ($r in $results) {
$rf = ((($r.RuleFlags + ‘’)).Trim())
$isOutlier = $false
if ($rf) {
$p = $rf.Split(’|’)
if ($p -contains ‘DQ_ASSAY_OUTLIER’ -or $p -contains ‘DQ_ASSAYVER_OUTLIER’) { $isOutlier = $true }
}
if ($isOutlier) { continue }
$sc = ((($r.SuffixCheck + ‘’)).Trim().ToUpperInvariant())
if ($sc -and $sc -ne ‘OK’) { *Append-RuleFlag -row $r -flag (’SUFFIX*’ + $sc) }
$snok = ((($r.SampleNumOK + ‘’)).Trim().ToUpperInvariant())
if ($snok -eq ‘NO’) { _Append-RuleFlag -row $r -flag ‘SAMPLENUM_BAD’ }
}

$useStrictPrefix = $false
try {
$p0 = @($results | Where-Object { (($*.SampleCode + ‘’)).Trim() -eq ‘0’ -and (($*.SamplePrefix + ‘’)).Trim() })
$pP = @($results | Where-Object { (($*.SampleCode + ‘’)).Trim() -match ‘^[1-5]$’ -and (($*.SamplePrefix + ‘’)).Trim() })

```
$ok0 = @($p0 | Where-Object { (($_.SamplePrefix + '')).Trim().ToUpperInvariant() -eq 'NEG' }).Count
$okP = @($pP | Where-Object { (($_.SamplePrefix + '')).Trim().ToUpperInvariant() -eq 'POS' }).Count

$r0 = 0.0; $rP = 0.0
if ($p0.Count -gt 0) { try { $r0 = [double]$ok0 / [double]$p0.Count } catch { $r0 = 0.0 } }
if ($pP.Count -gt 0) { try { $rP = [double]$okP / [double]$pP.Count } catch { $rP = 0.0 } }

if ($p0.Count -ge 3 -and $pP.Count -ge 3 -and $r0 -ge 0.80 -and $rP -ge 0.80) {
    $useStrictPrefix = $true
} elseif ($p0.Count -ge 10 -and $pP.Count -eq 0 -and $r0 -ge 0.90) {
    $useStrictPrefix = $true
} elseif ($pP.Count -ge 10 -and $p0.Count -eq 0 -and $rP -ge 0.90) {
    $useStrictPrefix = $true
}
```

} catch { $useStrictPrefix = $false }

foreach ($r in $results) {

```
$rf = ((($r.RuleFlags + '')).Trim())
if ($rf) {
    $p = $rf.Split('|')
    if ($p -contains 'DQ_ASSAY_OUTLIER' -or $p -contains 'DQ_ASSAYVER_OUTLIER') { continue }
}

$sidp = ((($r.SamplePrefix + '')).Trim().ToUpperInvariant())
$scode = ((($r.SampleCode + '')).Trim())
if ($useStrictPrefix -and $scode -match '^\d+$') {
    $si = 0; try { $si = [int]$scode } catch { $si = 0 }
    if ($si -eq 0) {
        if ($sidp -and $sidp -ne 'NEG') { _Append-RuleFlag -row $r -flag 'PREFIX_BAD' }
    } elseif ($si -ge 1 -and $si -le 5) {
        if ($sidp -and $sidp -ne 'POS') { _Append-RuleFlag -row $r -flag 'PREFIX_BAD' }
    }
}

$rnRaw = ((($r.RunNoRaw + '')).Trim())
$rn = ((($r.RunNo + '')).Trim())
if ($rnRaw) {
    if ($rnRaw.Length -ne 2) { _Append-RuleFlag -row $r -flag 'RUNNO_BAD' }
    $ni = 0; try { $ni = [int]$rn } catch { $ni = 0 }
    if ($ni -lt 1 -or $ni -gt 20) { _Append-RuleFlag -row $r -flag 'RUNNO_BAD' }
}

$dl = $false
try { $dl = [bool]$r.DelamPresent } catch { $dl = $false }
if ($dl) { _Append-RuleFlag -row $r -flag 'DELAM_PRESENT' }

$rl = 0
try { $rl = [int]$r.ReplacementLevel } catch { $rl = 0 }
if ($rl -ge 1) { _Append-RuleFlag -row $r -flag ('REPL_A' + $rl) }
```

}

$hotModules = @{}
$byModErr = @($results | Where-Object { (($*.ModuleSN + ‘’).Trim()) -and (($*.ObservedCall + ‘’).Trim().ToUpperInvariant() -eq ‘ERROR’) } | Group-Object ModuleSN)
foreach ($g in $byModErr) {
if ($g.Count -ge 3) { $hotModules[$g.Name] = $g.Count }
}
if ($hotModules.Count -gt 0) {
foreach ($r in $results) {
$m = ((($r.ModuleSN + ‘’)).Trim())
if ($m -and $hotModules.ContainsKey($m)) { _Append-RuleFlag -row $r -flag ‘MODULE_ERR_HOTSPOT’ }
}
}

$qc = [pscustomobject]@{
DistinctAssays = $distinctAssays
DistinctAssayVersions = $distinctAssayVersions
DistinctReagentLots = $distinctReagentLots
DuplicateSampleIdCount = ($dupSample | ForEach-Object { $*.Name } | Select-Object -Unique).Count
DuplicateCartridgeSnCount = ($dupCart | ForEach-Object { $*.Name } | Select-Object -Unique).Count
HotModuleCount = $hotModules.Count
DelamCount = @($results | Where-Object { try { [bool]$*.DelamPresent } catch { $false } }).Count
ReplacementCount = @($results | Where-Object { try { [int]$*.ReplacementLevel -ge 1 } catch { $false } }).Count
BadPrefixCount = @($results | Where-Object { (($*.RuleFlags + ‘’) -split ‘|’) -contains ‘PREFIX_BAD’ }).Count
BadRunNoCount  = @($results | Where-Object { (($*.RuleFlags + ‘’) -split ‘|’) -contains ‘RUNNO_BAD’ }).Count
}

# —————————————————————

$summary = [pscustomobject]@{
Total = $results.Count
ObservedCounts = @{}
DeviationCounts = @{}
RetestYes = 0
MinorFunctionalError = 0
InstrumentError = 0
DelamCount = 0
ReplacementCount = 0
}

foreach ($r in $results) {
if (-not $summary.ObservedCounts.ContainsKey($r.ObservedCall)) { $summary.ObservedCounts[$r.ObservedCall] = 0 }
$summary.ObservedCounts[$r.ObservedCall]++

```
if (-not $summary.DeviationCounts.ContainsKey($r.Deviation)) { $summary.DeviationCounts[$r.Deviation] = 0 }
$summary.DeviationCounts[$r.Deviation]++

$rt = ($r.GeneratesRetest + '').Trim().ToUpperInvariant()
if ($rt -in @('YES','Y','TRUE','1')) { $summary.RetestYes++ }
# Extra QC counts used by summary:
$dl = $false
try { $dl = [bool]$r.DelamPresent } catch { $dl = $false }
if ($dl) { $summary.DelamCount++ }

$rl = 0
try { $rl = [int]$r.ReplacementLevel } catch { $rl = 0 }
if ($rl -ge 1) { $summary.ReplacementCount++ }

if ((($r.Deviation + '')).Trim().ToUpperInvariant() -eq 'ERROR') {
    # Använd IsKnownCode som redan beräknats i Get-ErrorInfo
    $isKnown = $false
    try { $isKnown = [bool]$r.IsKnownCode } catch { $isKnown = $false }
    $isMtbInd = $false
    try {
        if ((($r.Assay + '') -match '(?i)MTB') -and ((($r.TestResult + '') -match '(?i)INDETERMINATE'))) { $isMtbInd = $true }
    } catch {}
    if ($isMtbInd) {
        $summary.MinorFunctionalError++
    } elseif ($isKnown) {
        $summary.MinorFunctionalError++
    } else {
        $summary.InstrumentError++
    }
}
```

}

$assayList = @($results | ForEach-Object { ($*.Assay + ‘’).Trim() } | Where-Object { $* } | Select-Object -Unique)
$ttMatched = 0
$ttMissing = @()
$ttDetails = @()
foreach ($a in $assayList) {
$pol = $null
try { $pol = Get-TestTypePolicyForAssayCached -RuleBank $RuleBank -Assay $a } catch { $pol = $null }
if ($pol) {
$ttMatched++
$pat = ((Get-RowField -Row $pol -FieldName ‘AssayPattern’) + ‘’).Trim()
$allowed = ((Get-RowField -Row $pol -FieldName ‘AllowedTestTypes’) + ‘’).Trim()
if (-not $allowed) { $allowed = ((Get-RowField -Row $pol -FieldName ‘TestTypes’) + ‘’).Trim() }
$ttDetails += ($a + ’ => ’ + $allowed + ’ (pattern=’ + $pat + ‘)’)
} else {
$ttMissing += $a
}
}

$sidTotal = @($results | Where-Object { (($*.SampleId + ‘’)).Trim().Length -gt 0 }).Count
$sidOk = @($results | Where-Object {
(($*.SampleId + ‘’)).Trim().Length -gt 0 -and
(($*.BagNo + ‘’)).Trim().Length -gt 0 -and
(($*.SampleCode + ‘’)).Trim().Length -gt 0 -and
(($*.RunNo + ‘’)).Trim().Length -gt 0 -and
(($*.RunSuffix + ‘’)).Trim().Length -gt 0
}).Count

$snTotal = $sidTotal
$snCovered = @($results | Where-Object {
(($*.SampleId + ‘’)).Trim().Length -gt 0 -and
(($*.SampleNumOK + ‘’)).Trim().Length -gt 0
}).Count

$summary | Add-Member -NotePropertyName ‘TestTypePolicyAssaysTotal’ -NotePropertyValue $assayList.Count -Force
$summary | Add-Member -NotePropertyName ‘TestTypePolicyAssaysMatched’ -NotePropertyValue $ttMatched -Force
$summary | Add-Member -NotePropertyName ‘TestTypePolicyAssaysMissing’ -NotePropertyValue ($ttMissing -join ’, ’) -Force
$summary | Add-Member -NotePropertyName ‘TestTypePolicyDetails’ -NotePropertyValue $ttDetails -Force

$summary | Add-Member -NotePropertyName ‘SampleIdParseTotal’ -NotePropertyValue $sidTotal -Force
$summary | Add-Member -NotePropertyName ‘SampleIdParseOk’ -NotePropertyValue $sidOk -Force
$summary | Add-Member -NotePropertyName ‘SampleNumberRuleTotal’ -NotePropertyValue $snTotal -Force
$summary | Add-Member -NotePropertyName ‘SampleNumberRuleCovered’ -NotePropertyValue $snCovered -Force

# —————————————––

$top = @($results | Where-Object { $_.Deviation -in @(‘FP’,‘FN’,‘ERROR’,‘MISMATCH’) } | Select-Object -First 50)

return [pscustomobject]@{ Rows = $results.ToArray(); Summary = $summary; TopDeviations = $top; QC = $qc }

}

# ============================================================================

# Write-RuleEngineDebugSheet - FÖRBÄTTRAD VERSION

# ============================================================================

# Skapar “CSV Sammanfattning”-fliken med professionell layout

# 

# Förbättringar:

# 1. Strukturerad sammanfattnings-sektion med tydliga rubriker

# 2. Färgkodning av data-rader baserat på Deviation-typ

# 3. Villkorlig formatering för viktiga kolumner

# 4. Förbättrade svenska översättningar

# 5. Visuell separation mellan sektioner

# 

# Teknisk miljö: PowerShell 5.1, EPPlus 4.5.3.3

# ============================================================================

function Write-RuleEngineDebugSheet {
param(
[Parameter(Mandatory)][object]$Pkg,
[Parameter(Mandatory)][pscustomobject]$RuleEngineResult,
[Parameter(Mandatory=$false)][bool]$IncludeAllRows = $false
)

# ============================================================================

# FÄRGDEFINITIONER (EPPlus-kompatibla)

# ============================================================================

$Colors = @{
# Rubriker och sektioner
HeaderBg       = [System.Drawing.Color]::FromArgb(68, 84, 106)    # Mörkblå
HeaderFg       = [System.Drawing.Color]::White
SectionBg      = [System.Drawing.Color]::FromArgb(217, 225, 242)  # Ljusblå
SectionFg      = [System.Drawing.Color]::FromArgb(0, 32, 96)      # Mörkblå text

```
# Status-färger
OkBg           = [System.Drawing.Color]::FromArgb(198, 239, 206)  # Ljusgrön
OkFg           = [System.Drawing.Color]::FromArgb(0, 97, 0)       # Mörkgrön

# Major Functional (FP/FN) - MÖRKRÖD
MajorBg        = [System.Drawing.Color]::FromArgb(192, 0, 0)      # Mörkröd
MajorFg        = [System.Drawing.Color]::White                    # Vit text

# Minor Functional / Max Pressure ≥90 - LJUSRÖD
MinorBg        = [System.Drawing.Color]::FromArgb(255, 199, 206)  # Ljusröd
MinorFg        = [System.Drawing.Color]::FromArgb(156, 0, 6)      # Mörkröd text

# Varningar (Instrument Error, övriga) - GUL
WarningBg      = [System.Drawing.Color]::FromArgb(255, 235, 156)  # Ljusgul
WarningFg      = [System.Drawing.Color]::FromArgb(156, 101, 0)    # Mörkorange

# Tabell
# Sätt tabellhuvudets bakgrund till samma mörkblå färg som huvudrubriken för att harmonisera med flikens färg
TableHeaderBg  = [System.Drawing.Color]::FromArgb(68, 84, 106)   # Mörkblå (samma som HeaderBg)
TableHeaderFg  = [System.Drawing.Color]::White
TableAltRow    = [System.Drawing.Color]::FromArgb(242, 242, 242)  # Ljusgrå

# Summering
SummaryGoodBg  = [System.Drawing.Color]::FromArgb(198, 239, 206)  # Ljusgrön
```

}

# ============================================================================

# RADERA GAMMALT BLAD OCH SKAPA NYTT

# ============================================================================

try {
$old = $Pkg.Workbook.Worksheets[‘CSV Sammanfattning’]
if ($old) { $Pkg.Workbook.Worksheets.Delete($old) }
} catch {}

$ws = $Pkg.Workbook.Worksheets.Add(‘CSV Sammanfattning’)

# Sätt standardfont

try {
$ws.Cells.Style.Font.Name = ‘Calibri’
$ws.Cells.Style.Font.Size = 10
} catch {}

# ============================================================================

# KOLUMNDEFINITIONER (14 kolumner)

# ============================================================================

$headers = @(
‘Sample ID’,
‘Error Code’,
‘Avvikelse’,
‘Notering’,
‘Förväntat X/+’,
‘Cartridge S/N’,
‘Module S/N’,
‘Förväntad Test Type’,
‘Status’,
‘Error Type’,
‘Ersätts?’,
‘Max Pressure (PSI)’,
‘Test Result’,
‘Error’
)

# ============================================================================

# HJÄLPFUNKTIONER

# ============================================================================

# Svensk översättning av Deviation

function _SvDeviation([string]$d) {
$t = (($d + ‘’)).Trim().ToUpperInvariant()
switch ($t) {
‘OK’       { return ‘OK’ }
‘FP’       { return ‘Falskt positiv’ }
‘FN’       { return ‘Falskt negativ’ }
‘ERROR’    { return ‘Fel’ }
‘MISMATCH’ { return ‘Mismatch’ }
‘UNKNOWN’  { return ‘Okänt’ }
default    { return ($d + ‘’) }
}
}

# Svensk översättning av SuffixCheck

function _SvSuffixCheck([string]$s) {
$t = (($s + ‘’)).Trim().ToUpperInvariant()
switch ($t) {
‘OK’      { return ‘OK’ }
‘BAD’     { return ‘FEL’ }
‘MISSING’ { return ‘SAKNAS’ }
default   { return ($s + ‘’) }
}
}

# Svensk översättning av RuleFlags

function _SvRuleFlags([string]$s) {
$t = (($s + ‘’)).Trim()
if (-not $t) { return ‘’ }

```
$map = @{
    'TESTTYPE_MISMATCH'   = 'Fel Test Type'
    'SUFFIX_BAD'          = 'Fel suffix'
    'DQ_DUP_SAMPLEID'     = 'Dubblett Sample ID'
    'DQ_DUP_CARTSN'       = 'Dubblett Cart S/N'
    'DQ_ASSAYVER_OUTLIER' = 'Assay Version (outlier)'
    'DQ_ASSAY_OUTLIER'    = 'Assay (outlier)'
    'RUNNO_BAD'           = 'Fel Rep-Nr'
    'DELAM_PRESENT'       = 'Delam'
    'REPL_A1'             = 'Ers. A1'
    'REPL_A2'             = 'Ers. A2'
    'REPL_A3'             = 'Ers. A3'
}

$tokens = @($t -split '[|,;]+' | ForEach-Object { ($_.Trim()) } | Where-Object { $_ })
if (-not $tokens -or $tokens.Count -eq 0) { return $t }

$out = foreach ($tok in $tokens) {
    if ($map.ContainsKey($tok)) { $map[$tok] } else { $tok }
}
return ($out -join ', ')
```

}

# Skriv sektionsrubrik med styling

function Write-SectionHeader {
param([int]$Row, [string]$Text, [int]$ColSpan = 4)

```
$ws.Cells.Item($Row, 1).Value = $Text
$rng = $ws.Cells[$Row, 1, $Row, $ColSpan]
$rng.Merge = $true
$rng.Style.Font.Bold = $true
$rng.Style.Font.Size = 11
$rng.Style.Fill.PatternType = [OfficeOpenXml.Style.ExcelFillStyle]::Solid
$rng.Style.Fill.BackgroundColor.SetColor($Colors.SectionBg)
$rng.Style.Font.Color.SetColor($Colors.SectionFg)
$rng.Style.Border.Bottom.Style = [OfficeOpenXml.Style.ExcelBorderStyle]::Medium
$rng.Style.Border.Bottom.Color.SetColor($Colors.SectionFg)
```

}

# Skriv nyckel-värde par med valfri formatering

function Write-KV {
param(
[int]$Row,
[string]$Key,
$Value,
[int]$Col = 1,
[switch]$Highlight,
[switch]$Warning,
[switch]$Good,
[switch]$Major,
[switch]$Minor,
[switch]$Neutral
)

```
$ws.Cells.Item($Row, $Col).Value = $Key
$ws.Cells.Item($Row, $Col).Style.Font.Bold = $true
$ws.Cells.Item($Row, $Col + 1).Value = $Value

if ($Major) {
    # Major Functional (FP/FN) - Mörkröd bakgrund, vit text
    $ws.Cells.Item($Row, $Col + 1).Style.Fill.PatternType = [OfficeOpenXml.Style.ExcelFillStyle]::Solid
    $ws.Cells.Item($Row, $Col + 1).Style.Fill.BackgroundColor.SetColor($Colors.MajorBg)
    $ws.Cells.Item($Row, $Col + 1).Style.Font.Color.SetColor($Colors.MajorFg)
    $ws.Cells.Item($Row, $Col + 1).Style.Font.Bold = $true
}
elseif ($Minor) {
    # Minor Functional - Ljusröd bakgrund
    $ws.Cells.Item($Row, $Col + 1).Style.Fill.PatternType = [OfficeOpenXml.Style.ExcelFillStyle]::Solid
    $ws.Cells.Item($Row, $Col + 1).Style.Fill.BackgroundColor.SetColor($Colors.MinorBg)
    $ws.Cells.Item($Row, $Col + 1).Style.Font.Color.SetColor($Colors.MinorFg)
    $ws.Cells.Item($Row, $Col + 1).Style.Font.Bold = $true
}
elseif ($Neutral) {
    # Neutral grå bakgrund används t.ex. i Fel och varningar
    $ws.Cells.Item($Row, $Col + 1).Style.Fill.PatternType = [OfficeOpenXml.Style.ExcelFillStyle]::Solid
    $ws.Cells.Item($Row, $Col + 1).Style.Fill.BackgroundColor.SetColor($Colors.TableAltRow)
    $ws.Cells.Item($Row, $Col + 1).Style.Font.Bold = $true
}
elseif ($Highlight -or $Warning) {
    # Gul bakgrund (Instrument Error, övriga)
    $ws.Cells.Item($Row, $Col + 1).Style.Fill.PatternType = [OfficeOpenXml.Style.ExcelFillStyle]::Solid
    $ws.Cells.Item($Row, $Col + 1).Style.Fill.BackgroundColor.SetColor($Colors.WarningBg)
    $ws.Cells.Item($Row, $Col + 1).Style.Font.Bold = $true
}
elseif ($Good) {
    $ws.Cells.Item($Row, $Col + 1).Style.Fill.PatternType = [OfficeOpenXml.Style.ExcelFillStyle]::Solid
    $ws.Cells.Item($Row, $Col + 1).Style.Fill.BackgroundColor.SetColor($Colors.SummaryGoodBg)
}
```

}

# ============================================================================

# HUVUDRUBRIK

# ============================================================================

$row = 1
$ws.Cells.Item($row, 1).Value = ‘CSV Sammanfattning’
$titleRng = $ws.Cells[$row, 1, $row, 8]
$titleRng.Merge = $true
$titleRng.Style.Font.Bold = $true
$titleRng.Style.Font.Size = 14
$titleRng.Style.Fill.PatternType = [OfficeOpenXml.Style.ExcelFillStyle]::Solid
$titleRng.Style.Fill.BackgroundColor.SetColor($Colors.HeaderBg)
$titleRng.Style.Font.Color.SetColor($Colors.HeaderFg)
$ws.Row($row).Height = 25
$row += 2

# ============================================================================

# HÄMTA DATA

# ============================================================================

$sum = $RuleEngineResult.Summary
$qc  = $RuleEngineResult.QC
$allRows = @($RuleEngineResult.Rows)

# ============================================================================

# SEKTION 1: ÖVERGRIPANDE STATISTIK

# ============================================================================

Write-SectionHeader -Row $row -Text ‘📈 Övergripande statistik’ -ColSpan 8
$row++

# Rad 1: Totalt, Assay, Version, Lot

$assayTxt = ‘’
if ($qc -and $qc.DistinctAssays -and $qc.DistinctAssays.Count -eq 1) {
$assayTxt = $qc.DistinctAssays[0]
}
elseif ($qc -and $qc.DistinctAssays -and $qc.DistinctAssays.Count -gt 1) {
$assayTxt = “⚠ Flera ($($qc.DistinctAssays.Count))”
}

$verTxt = ‘’
if ($qc -and $qc.DistinctAssayVersions -and $qc.DistinctAssayVersions.Count -eq 1) {
$verTxt = $qc.DistinctAssayVersions[0]
}
elseif ($qc -and $qc.DistinctAssayVersions -and $qc.DistinctAssayVersions.Count -gt 1) {
$verTxt = “⚠ Flera ($($qc.DistinctAssayVersions.Count))”
}

$lotTxt = ‘’
if ($qc -and $qc.DistinctReagentLots -and $qc.DistinctReagentLots.Count -eq 1) {
$lotTxt = $qc.DistinctReagentLots[0]
}
elseif ($qc -and $qc.DistinctReagentLots -and $qc.DistinctReagentLots.Count -gt 1) {
$lotTxt = “⚠ Flera ($($qc.DistinctReagentLots.Count))”
}

Write-KV -Row $row -Key ‘Totalt tester’ -Value $sum.Total -Col 1
Write-KV -Row $row -Key ‘Assay’ -Value $assayTxt -Col 3 -Warning:($assayTxt -like ‘*Flera*’)
Write-KV -Row $row -Key ‘Assay Version’ -Value $verTxt -Col 5 -Warning:($verTxt -like ‘*Flera*’)
Write-KV -Row $row -Key ‘Reagent Lot’ -Value $lotTxt -Col 7 -Warning:($lotTxt -like ‘*Flera*’)
$row += 2

# ============================================================================

# SEKTION 2: AVVIKELSER (DEVIATION)

# ============================================================================

Write-SectionHeader -Row $row -Text ‘Avvikelseanalys’ -ColSpan 8
$row++

# Beräkna OK-antal (utan procent)

$okCount = 0
if ($sum.DeviationCounts.ContainsKey(‘OK’)) { $okCount = $sum.DeviationCounts[‘OK’] }

Write-KV -Row $row -Key ‘✓ Godkända (OK)’ -Value $okCount -Col 1 -Good
$row++

# FP - Falskt positiv (Major Functional) - Mörkröd

$fpCount = 0
if ($sum.DeviationCounts.ContainsKey(‘FP’)) { $fpCount = $sum.DeviationCounts[‘FP’] }
if ($fpCount -gt 0) {
Write-KV -Row $row -Key ‘❌ Falskt positiv (Major Functional)’ -Value $fpCount -Col 1 -Major
$row++
}

# FN - Falskt negativ (Major Functional) - Mörkröd

$fnCount = 0
if ($sum.DeviationCounts.ContainsKey(‘FN’)) { $fnCount = $sum.DeviationCounts[‘FN’] }
if ($fnCount -gt 0) {
Write-KV -Row $row -Key ‘❌ Falskt negativ (Major Functional)’ -Value $fnCount -Col 1 -Major
$row++
}

# Minor Functional - Ljusröd

if ($sum -and $sum.MinorFunctionalError -ne $null -and $sum.MinorFunctionalError -gt 0) {
Write-KV -Row $row -Key ‘⚠ Minor Functional’ -Value $sum.MinorFunctionalError -Col 1 -Minor
$row++
}

# Visa om inga avvikelser (om OK = Total och inga FP/FN/Minor)

if ($fpCount -eq 0 -and $fnCount -eq 0 -and ($sum.MinorFunctionalError -eq $null -or $sum.MinorFunctionalError -eq 0)) {
$ws.Cells.Item($row, 1).Value = ‘✓ Inga avvikelser hittades’
$ws.Cells.Item($row, 1).Style.Font.Italic = $true
$ws.Cells.Item($row, 1).Style.Font.Color.SetColor($Colors.OkFg)
$row++
}
$row++

# ============================================================================

# SEKTION 3: FEL OCH VARNINGAR

# ============================================================================

Write-SectionHeader -Row $row -Text ‘⚠ Fel och varningar’ -ColSpan 8
$row++

# Instrument Error - GUL bakgrund

if ($sum -and $sum.InstrumentError -ne $null -and $sum.InstrumentError -gt 0) {
Write-KV -Row $row -Key ‘Instrument Error’ -Value $sum.InstrumentError -Neutral
$row++
}

# Delam och ersättningar

if ($sum -and $sum.DelamCount -ne $null -and $sum.DelamCount -gt 0) {
Write-KV -Row $row -Key ‘Delamineringar’ -Value $sum.DelamCount -Neutral
$row++
}
if ($sum -and $sum.ReplacementCount -ne $null -and $sum.ReplacementCount -gt 0) {
Write-KV -Row $row -Key ‘Ersättningar (A/AA/AAA)’ -Value $sum.ReplacementCount -Neutral
$row++
}

# Omkörning

if ($sum.RetestYes -gt 0) {
Write-KV -Row $row -Key ‘Behöver omkörning (YES)’ -Value $sum.RetestYes -Neutral
$row++
}

# Dubbletter

if ($qc) {
if ($qc.DuplicateSampleIdCount -gt 0) {
Write-KV -Row $row -Key ‘Dubbletter Sample ID’ -Value $qc.DuplicateSampleIdCount -Neutral
$row++
}
if ($qc.DuplicateCartridgeSnCount -gt 0) {
Write-KV -Row $row -Key ‘Dubbletter Cartridge S/N’ -Value $qc.DuplicateCartridgeSnCount -Neutral
$row++
}
if ($qc.HotModuleCount -gt 0) {
Write-KV -Row $row -Key ‘Moduler med ≥3 fel’ -Value $qc.HotModuleCount -Neutral
$row++
}
}

# Max Pressure ≥ 90 PSI - Ljusröd bakgrund

$pressureGE90 = @($allRows | Where-Object {
$p = $null
try { $p = [double]$_.MaxPressure } catch { $p = $null }
return ($null -ne $p -and $p -ge 90)
}).Count
if ($pressureGE90 -gt 0) {
Write-KV -Row $row -Key ‘Max Pressure ≥ 90 PSI’ -Value $pressureGE90 -Neutral
$row++
}

# Max Pressure Failure utan Error Code - lägg till i Fel och varningar

$pressureFailNoError = @($allRows | Where-Object {
$p = $null
try { $p = [double]$*.MaxPressure } catch { $p = $null }
$hasError = ((($*.ErrorCode + ‘’)).Trim().Length -gt 0)
return ($null -ne $p -and $p -ge 90 -and -not $hasError)
}).Count
if ($pressureFailNoError -gt 0) {
Write-KV -Row $row -Key ‘Max Pressure Failure (utan Error Code)’ -Value $pressureFailNoError -Neutral
$row++
}

$row++

# ============================================================================

# SEKTION 4: OBSERVERADE RESULTAT

# ============================================================================

Write-SectionHeader -Row $row -Text ‘Observerade resultat’ -ColSpan 8
$row++

foreach ($k in @(‘POS’,‘NEG’,‘ERROR’,‘UNKNOWN’)) {
if ($sum.ObservedCounts.ContainsKey($k) -and $sum.ObservedCounts[$k] -gt 0) {
$icon = switch ($k) {
‘POS’     { ‘✓’ }
‘NEG’     { ‘✓’ }
‘ERROR’   { ‘❌’ }
‘UNKNOWN’ { ‘❓’ }
default   { ‘’ }
}
$good = ($k -in @(‘POS’,‘NEG’))
Write-KV -Row $row -Key “$icon Observerat $k” -Value $sum.ObservedCounts[$k] -Good:$good -Warning:(-not $good -and $sum.ObservedCounts[$k] -gt 0)
$row++
}
}

$row += 2

# ============================================================================

# DETALJTABELL

# ============================================================================

# Filtrera rader

$rowsToWrite = $allRows
if (-not $IncludeAllRows) {
$rowsToWrite = @($allRows | Where-Object {
$dev = (($_.Deviation + ‘’)).Trim()
$hasDeviation = ($dev.Length -gt 0 -and $dev -ne ‘OK’)

```
    $obs = (($_.ObservedCall + '')).Trim().ToUpperInvariant()
    $observedErr = ($obs -eq 'ERROR')

    $pressureFlag = $false
    try { $pressureFlag = [bool]$_.PressureFlag } catch { $pressureFlag = $false }

    $hasErrorCode = ((($_.ErrorCode + '')).Trim().Length -gt 0)

    $st = (($_.Status + '')).Trim()
    $statusNotDone = ($st.Length -gt 0 -and $st -ne 'Done')

    $retestTrue = $false
    $rt = (($_.GeneratesRetest + '')).Trim().ToUpperInvariant()
    if ($rt -in @('YES','Y','TRUE','1')) { $retestTrue = $true }

    $rf = (($_.RuleFlags + '')).Trim()
    $hasRuleFlags = ($rf.Length -gt 0)

    return ($hasDeviation -or $observedErr -or $pressureFlag -or $hasErrorCode -or $statusNotDone -or $retestTrue -or $hasRuleFlags)
})
```

}

# Tabell-rubrik

$tableInfoRow = $row
$deviationCount = $rowsToWrite.Count
$tableInfoText = if ($deviationCount -eq 0) {
“Detaljerad avvikelselista - Inga avvikelser att visa”
} else {
“Detaljerad avvikelselista - $deviationCount rader”
}

$ws.Cells.Item($row, 1).Value = $tableInfoText
$infoRng = $ws.Cells[$row, 1, $row, 6]
$infoRng.Merge = $true
$infoRng.Style.Font.Bold = $true
$infoRng.Style.Font.Size = 11
$row++

$tableHeaderRow = $row

# Skriv headers

for ($c = 1; $c -le $headers.Count; $c++) {
$ws.Cells.Item($tableHeaderRow, $c).Value = $headers[$c - 1]
}

# Styla header-rad

$headerRange = $ws.Cells[$tableHeaderRow, 1, $tableHeaderRow, $headers.Count]
$headerRange.Style.Font.Bold = $true
$headerRange.Style.Fill.PatternType = [OfficeOpenXml.Style.ExcelFillStyle]::Solid
$headerRange.Style.Fill.BackgroundColor.SetColor($Colors.TableHeaderBg)
$headerRange.Style.Font.Color.SetColor($Colors.TableHeaderFg)
$headerRange.Style.HorizontalAlignment = [OfficeOpenXml.Style.ExcelHorizontalAlignment]::Center

# AutoFilter och FreezePanes

try { $ws.Cells[$tableHeaderRow, 1, $tableHeaderRow, $headers.Count].AutoFilter = $true } catch {}
try { $ws.View.FreezePanes($tableHeaderRow + 1, 1) } catch {}

# Om inga rader

if (-not $rowsToWrite -or $rowsToWrite.Count -eq 0) {
$ws.Cells.Item($tableHeaderRow + 1, 1).Value = ‘✓ Inga avvikelser hittades - alla tester OK!’
$ws.Cells.Item($tableHeaderRow + 1, 1).Style.Font.Italic = $true
$ws.Cells.Item($tableHeaderRow + 1, 1).Style.Font.Color.SetColor($Colors.OkFg)
$noDevRng = $ws.Cells[$tableHeaderRow + 1, 1, $tableHeaderRow + 1, 6]
$noDevRng.Merge = $true

```
try {
    $rAll = $ws.Cells[1, 1, ($tableHeaderRow + 1), $headers.Count]
    if (Get-Command Safe-AutoFitColumns -ErrorAction SilentlyContinue) {
        Safe-AutoFitColumns -Ws $ws -Range $rAll -Context 'CSV Sammanfattning'
    } else {
        $rAll.AutoFitColumns() | Out-Null
    }
} catch {}

$ws.TabColor = [System.Drawing.Color]::Green
return $ws
```

}

# ============================================================================

# SKRIV DATA MED BULK-OPERATION

# ============================================================================

$rowCount = $rowsToWrite.Count
$colCount = $headers.Count
$data = New-Object ‘object[,]’ $rowCount, $colCount

for ($i = 0; $i -lt $rowCount; $i++) {
$r = $rowsToWrite[$i]

```
# Kolumn 1: Sample ID
$data[$i, 0] = ($r.SampleId + '')

# Kolumn 2: Error Code och Kolumn 3: Avvikelse
$rawDev = (($r.Deviation + '')).Trim().ToUpperInvariant()
$errCode = (($r.ErrorCode + '')).Trim()
$isKnown = $false
$isMtbInd = $false
# Kontrollera om felkoden finns i regelbanken (numerisk kod)
if ($rawDev -eq 'ERROR' -and $errCode -match '^\d{4,5}$') {
    try { $isKnown = $errLut.Codes.ContainsKey($errCode) } catch { $isKnown = $false }
}
# MTB Indeterminate: behandlas som minor functional
try {
    if ((($r.Assay + '') -match '(?i)MTB') -and ((($r.TestResult + '') -match '(?i)INDETERMINATE'))) { $isMtbInd = $true }
} catch {}
switch ($rawDev) {
        'ERROR' {
    # Använd IsKnownCode för att avgöra om felet är känt (Minor) eller okänt (Instrument)
    # IsKnownCode sätts i Get-ErrorInfo - true om felkoden finns i RuleBank
    $isKnown = $false
    try { $isKnown = [bool]$r.IsKnownCode } catch { $isKnown = $false }
    if ($isMtbInd) {
        # MTB indeterminate behandlas som Minor Functional oavsett felkod
        $data[$i, 2] = 'Minor Functional'
    } elseif ($isKnown) {
        # Felkoden finns i RuleBank = Minor Functional
        $data[$i, 2] = 'Minor Functional'
    } else {
        # Felkoden finns INTE i RuleBank = Instrument Error
        $data[$i, 2] = 'Instrument Error'
    }
    # Visa faktisk felkod i Error Code-kolumnen
    $data[$i, 1] = ($r.ErrorCode + '')
}
    'FP' {
        $data[$i, 2] = 'Major Functional'
        $data[$i, 1] = 'Falskt positiv'
    }
    'FN' {
        $data[$i, 2] = 'Major Functional'
        $data[$i, 1] = 'Falskt negativ'
    }
    'MISMATCH' {
        $data[$i, 2] = 'Instrument Error'
        $data[$i, 1] = 'Mismatch'
    }
    'UNKNOWN' {
        $data[$i, 2] = 'Instrument Error'
        $data[$i, 1] = 'Okänt'
    }
    'OK' {
        $data[$i, 2] = 'OK'
        $data[$i, 1] = ($r.ErrorCode + '')
    }
    default {
        $data[$i, 2] = 'OK'
        $data[$i, 1] = ($r.ErrorCode + '')
    }
}

# Kolumn 4: Notering (RuleFlags)
$data[$i, 3] = (_SvRuleFlags ($r.RuleFlags + ''))

# Kolumn 5: Förväntat X/+
$sc = (($r.SuffixCheck + '')).Trim().ToUpperInvariant()
if ($sc -and $sc -ne 'OK') {
    $expectedSuffix = ($r.ExpectedSuffix + '')
    if (-not $expectedSuffix) { $expectedSuffix = '' }
    $data[$i, 4] = $expectedSuffix
} else {
    $data[$i, 4] = 'OK'
}

# Kolumn 6: Cartridge S/N
$data[$i, 5] = ($r.CartridgeSN + '')

# Kolumn 7: Module S/N
$data[$i, 6] = ($r.ModuleSN + '')

# Kolumn 8: Förväntad Test Type
$expTestType = ($r.ExpectedTestType + '')
$obsTestType = (($r.TestType + '')).Trim()
if ($expTestType -and $obsTestType -and ($expTestType -ne $obsTestType)) {
    $data[$i, 7] = $expTestType
} else {
    $data[$i, 7] = 'OK'
}

# Kolumn 9: Status
$data[$i, 8] = ($r.Status + '')

# Kolumn 10: Error Type
$data[$i, 9] = ($r.ErrorName + '')

# Kolumn 11: Ersätts?
$rt = (($r.GeneratesRetest + '')).Trim().ToUpperInvariant()
if ($rt -in @('YES','Y','TRUE','1')) {
    $data[$i, 10] = 'Ja'
} elseif ($rt) {
    $data[$i, 10] = 'Nej'
} else {
    $data[$i, 10] = ''
}

# Kolumn 12: Max Pressure (PSI)
if ($null -ne $r.MaxPressure) {
    $data[$i, 11] = $r.MaxPressure
} else {
    $data[$i, 11] = ''
}

# Kolumn 13: Test Result
$data[$i, 12] = ($r.TestResult + '')

# Kolumn 14: Error (feltext)
$data[$i, 13] = ($r.ErrorText + '')
```

}

$startRow = $tableHeaderRow + 1
$endRow = $startRow + $rowCount - 1
$rng = $ws.Cells[$startRow, 1, $endRow, $colCount]
$rng.Value = $data

# ============================================================================

# FÄRGKODNING AV DATA-RADER (baserat på Deviation)

# ============================================================================

for ($i = 0; $i -lt $rowCount; $i++) {
$dataRow = $startRow + $i
$r = $rowsToWrite[$i]
$dev = (($r.Deviation + ‘’)).Trim().ToUpperInvariant()

```
$rowRange = $ws.Cells[$dataRow, 1, $dataRow, $colCount]

# Avvikelse-kolumnen (kolumn C, index 3)
$devCell = $ws.Cells.Item($dataRow, 3)

# Error Code-kolumnen (kolumn B, index 2) - för Major Functional
$errorCodeCell = $ws.Cells.Item($dataRow, 2)

switch ($dev) {
    'FP' {
        # Major Functional - Mörkröd bakgrund, vit text
        $devCell.Style.Fill.PatternType = [OfficeOpenXml.Style.ExcelFillStyle]::Solid
        $devCell.Style.Fill.BackgroundColor.SetColor($Colors.MajorBg)
        $devCell.Style.Font.Color.SetColor($Colors.MajorFg)
        $devCell.Style.Font.Bold = $true
        # Färgmarkera även Error Code-kolumnen
        $errorCodeCell.Style.Fill.PatternType = [OfficeOpenXml.Style.ExcelFillStyle]::Solid
        $errorCodeCell.Style.Fill.BackgroundColor.SetColor($Colors.MajorBg)
        $errorCodeCell.Style.Font.Color.SetColor($Colors.MajorFg)
        $errorCodeCell.Style.Font.Bold = $true
    }
    'FN' {
        # Major Functional - Mörkröd bakgrund, vit text
        $devCell.Style.Fill.PatternType = [OfficeOpenXml.Style.ExcelFillStyle]::Solid
        $devCell.Style.Fill.BackgroundColor.SetColor($Colors.MajorBg)
        $devCell.Style.Font.Color.SetColor($Colors.MajorFg)
        $devCell.Style.Font.Bold = $true
        # Färgmarkera även Error Code-kolumnen
        $errorCodeCell.Style.Fill.PatternType = [OfficeOpenXml.Style.ExcelFillStyle]::Solid
        $errorCodeCell.Style.Fill.BackgroundColor.SetColor($Colors.MajorBg)
        $errorCodeCell.Style.Font.Color.SetColor($Colors.MajorFg)
        $errorCodeCell.Style.Font.Bold = $true
    }
    'ERROR' {
        # Differentiera mellan Minor Functional och Instrument Error
        $disp = ($data[$i, 2] + '')
        if ($disp -eq 'Instrument Error') {
            # Instrument Error - använd varningsfärg (gul)
            $devCell.Style.Fill.PatternType = [OfficeOpenXml.Style.ExcelFillStyle]::Solid
            $devCell.Style.Fill.BackgroundColor.SetColor($Colors.WarningBg)
            $devCell.Style.Font.Color.SetColor($Colors.WarningFg)
        } else {
            # Minor Functional - ljusröd bakgrund
            $devCell.Style.Fill.PatternType = [OfficeOpenXml.Style.ExcelFillStyle]::Solid
            $devCell.Style.Fill.BackgroundColor.SetColor($Colors.MinorBg)
            $devCell.Style.Font.Color.SetColor($Colors.MinorFg)
        }
        $devCell.Style.Font.Bold = $true
    }
    'MISMATCH' {
        # Varning - Gul bakgrund
        $devCell.Style.Fill.PatternType = [OfficeOpenXml.Style.ExcelFillStyle]::Solid
        $devCell.Style.Fill.BackgroundColor.SetColor($Colors.WarningBg)
        $devCell.Style.Font.Color.SetColor($Colors.WarningFg)
        $devCell.Style.Font.Bold = $true
    }
    'UNKNOWN' {
        # Okänt/Instrument Error - Gul bakgrund
        $devCell.Style.Fill.PatternType = [OfficeOpenXml.Style.ExcelFillStyle]::Solid
        $devCell.Style.Fill.BackgroundColor.SetColor($Colors.WarningBg)
        $devCell.Style.Font.Color.SetColor($Colors.WarningFg)
        $devCell.Style.Font.Bold = $true
    }
}

# Markera "Ersätts? = Ja" med gul bakgrund (kolumn 11)
$ersattsVal = ($data[$i, 10] + '').Trim()
if ($ersattsVal -eq 'Ja') {
    $ersattsCell = $ws.Cells.Item($dataRow, 11)
    $ersattsCell.Style.Fill.PatternType = [OfficeOpenXml.Style.ExcelFillStyle]::Solid
    $ersattsCell.Style.Fill.BackgroundColor.SetColor($Colors.WarningBg)
    $ersattsCell.Style.Font.Bold = $true
}

# Markera högt tryck (≥90 PSI) med ljusröd (Minor-nivå) på kolumn 12
try {
    $pressure = $null
    if ($r.MaxPressure -ne $null) { $pressure = [double]$r.MaxPressure }
    if ($pressure -ne $null -and $pressure -ge 90) {
        $pressCell = $ws.Cells.Item($dataRow, 12)
        $pressCell.Style.Fill.PatternType = [OfficeOpenXml.Style.ExcelFillStyle]::Solid
        $pressCell.Style.Fill.BackgroundColor.SetColor($Colors.MinorBg)
        $pressCell.Style.Font.Color.SetColor($Colors.MinorFg)
        $pressCell.Style.Font.Bold = $true
    }
} catch {}

# Varannan rad med ljusgrå bakgrund (zebra-ränder) för rader utan avvikelse-färg
if ($i % 2 -eq 1 -and $dev -eq 'OK') {
    $rowRange.Style.Fill.PatternType = [OfficeOpenXml.Style.ExcelFillStyle]::Solid
    $rowRange.Style.Fill.BackgroundColor.SetColor($Colors.TableAltRow)
}
```

}

# ============================================================================

# RAMAR OCH FORMATERING

# ============================================================================

# Lägg till tunna ramar runt alla dataceller

try {
$tableRange = $ws.Cells[$tableHeaderRow, 1, $endRow, $colCount]
$tableRange.Style.Border.Top.Style    = [OfficeOpenXml.Style.ExcelBorderStyle]::Thin
$tableRange.Style.Border.Bottom.Style = [OfficeOpenXml.Style.ExcelBorderStyle]::Thin
$tableRange.Style.Border.Left.Style   = [OfficeOpenXml.Style.ExcelBorderStyle]::Thin
$tableRange.Style.Border.Right.Style  = [OfficeOpenXml.Style.ExcelBorderStyle]::Thin
$tableRange.Style.Border.Top.Color.SetColor([System.Drawing.Color]::LightGray)
$tableRange.Style.Border.Bottom.Color.SetColor([System.Drawing.Color]::LightGray)
$tableRange.Style.Border.Left.Color.SetColor([System.Drawing.Color]::LightGray)
$tableRange.Style.Border.Right.Color.SetColor([System.Drawing.Color]::LightGray)
} catch {}

# Tjockare ram runt header

try {
$headerRange.Style.Border.Bottom.Style = [OfficeOpenXml.Style.ExcelBorderStyle]::Medium
$headerRange.Style.Border.Bottom.Color.SetColor([System.Drawing.Color]::FromArgb(68, 84, 106))
} catch {}

# ============================================================================

# KOLUMNBREDDER

# ============================================================================

try {
$rAll = $ws.Cells[1, 1, $endRow, $colCount]
if (Get-Command Safe-AutoFitColumns -ErrorAction SilentlyContinue) {
Safe-AutoFitColumns -Ws $ws -Range $rAll -Context ‘CSV Sammanfattning’
} else {
$rAll.AutoFitColumns() | Out-Null
}
} catch {}

# Sätt minbredd för vissa kolumner

try {
$ws.Column(1).Width = [Math]::Max($ws.Column(1).Width, 15)   # Sample ID
$ws.Column(3).Width = [Math]::Max($ws.Column(3).Width, 14)   # Avvikelse
$ws.Column(7).Width = [Math]::Max($ws.Column(7).Width, 14)   # Cartridge S/N
# Justera bredder efter att kolumner tagits bort: feltextkolumnen finns nu på kolumn 14
$ws.Column(14).Width = [Math]::Max($ws.Column(14).Width, 30) # Error (kan vara lång)
} catch {}

# ============================================================================

# FLIK-FÄRG

# ============================================================================

# Grön om inga avvikelser, annars orange

if ($deviationCount -eq 0) {
$ws.TabColor = [System.Drawing.Color]::Green
} elseif ($deviationCount -le 5) {
$ws.TabColor = [System.Drawing.Color]::Orange
} else {
$ws.TabColor = [System.Drawing.Color]::Red
}

return $ws
}