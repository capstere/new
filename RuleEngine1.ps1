function Import-RuleCsv {
    param([Parameter(Mandatory)][string]$Path)
    if (-not (Test-Path -LiteralPath $Path)) { return @() }

    $delim = ','
    try { $delim = Get-CsvDelimiter -Path $Path } catch {}

    try {
        $lines = Get-Content -LiteralPath $Path -ErrorAction Stop
        if (-not $lines -or $lines.Count -lt 1) { return @() }
        return @(ConvertFrom-Csv -InputObject ($lines -join "`n") -Delimiter $delim)
    } catch {
        try { return @(Import-Csv -LiteralPath $Path -Delimiter $delim) } catch { return @() }
    }
}

function _RuleEngine_Log {
    param(
        [Parameter(Mandatory)][string]$Text,
        [ValidateSet('Info','Warn','Error')][string]$Severity = 'Info'
    )
    try {
        $cmd = Get-Command -Name Gui-Log -ErrorAction SilentlyContinue
        if ($cmd) { Gui-Log -Text $Text -Severity $Severity -Category 'RuleEngine' }
    } catch {}
}

function Test-RuleBankIntegrity {
    param(
        [Parameter(Mandatory)][pscustomobject]$RuleBank,
        [Parameter(Mandatory=$false)][string]$Source = ''
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
                if (-not $src) { $src = 'RuleBank' }
                throw ("RuleBank (Load-RuleBank): Tabell '" + $tableName + "' saknar kolumn: " + $c + " (" + $src + ")")
            }
        }
    }

    foreach ($t in @('ResultCallPatterns','SampleExpectationRules','ErrorCodes','MissingSamplesConfig','SampleIdMarkers','ParityCheckConfig','SampleNumberRules','TestTypePolicy')) {
        $null = _EnsureArray $t
    }

    _RequireColumns 'ResultCallPatterns'     @('Assay','Call','MatchType','Pattern','Enabled','Priority')
    _RequireColumns 'SampleExpectationRules' @('Assay','SampleIdMatchType','SampleIdPattern','Expected','Enabled','Priority')
    _RequireColumns 'ErrorCodes'             @('ErrorCode','Name','GeneratesRetest')
    _RequireColumns 'SampleIdMarkers'        @('AssayPattern','MarkerType','Marker','Enabled')
    _RequireColumns 'ParityCheckConfig'      @('AssayPattern','Enabled','CartridgeField','SampleTokenIndex','SuffixX','SuffixPlus','MinValidCartridgeSNPercent','Priority')
    _RequireColumns 'SampleNumberRules'      @('AssayPattern','SampleTypeCode','BagNoPattern','SampleNumberTokenIndex','SampleNumberRegex','SampleNumberMin','SampleNumberMax','SampleNumberPad','Enabled','Priority')
    _RequireColumns 'TestTypePolicy'         @('AssayPattern','AllowedTestTypes','Enabled','Priority')

    return $true
}

function Load-RuleBank {
    param([Parameter(Mandatory)][string]$RuleBankDir)
    $requireCompiled = $false
    try {
        if (Get-Command Get-ConfigValue -ErrorAction SilentlyContinue) {
            $requireCompiled = [bool](Get-ConfigValue -Name 'RuleBankRequireCompiled' -Default $false)
        } else {
            $cfg = $null
            if ($global:Config) { $cfg = $global:Config }
            elseif (Get-Variable -Name Config -Scope Script -ErrorAction SilentlyContinue) { $cfg = (Get-Variable -Name Config -Scope Script -ValueOnly -ErrorAction SilentlyContinue) }

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
    } catch { $requireCompiled = $false }

    if (-not (Test-Path -LiteralPath $RuleBankDir)) {
        throw ("RuleBank (Load-RuleBank): Directory not found: " + $RuleBankDir)
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

    $expectedTables = @('ResultCallPatterns','SampleExpectationRules','ErrorCodes','MissingSamplesConfig','SampleIdMarkers','ParityCheckConfig','SampleNumberRules','TestTypePolicy')

    function _HasKey([object]$dict, [string]$key) {
        try {
            if ($dict -is [hashtable]) { return $dict.ContainsKey($key) }
            if ($dict -is [System.Collections.IDictionary]) { return $dict.Contains($key) }
        } catch {}
        return $false
    }

    $compiledCandidates = @(
        (Join-Path $RuleBankDir 'RuleBank.compiled.ps1'),
        (Join-Path $RuleBankDir 'build\RuleBank.compiled.ps1'),
        (Join-Path $RuleBankDir 'RuleBank.compiled.psd1'),
        (Join-Path $RuleBankDir 'build\RuleBank.compiled.psd1')
    )

    foreach ($cp in $compiledCandidates) {
        if (-not (Test-Path -LiteralPath $cp)) { continue }

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
    }

    if ($requireCompiled) {
        throw ("RuleBank (Load-RuleBank): Compiled artifact missing. Expected RuleBank.compiled.ps1 in: {0}" -f $RuleBankDir)
    }

    # ---- CSV fallback ----
    $map = @(
        @{ Key='ResultCallPatterns';      File='01_ResultCallPatterns.csv' },
        @{ Key='SampleExpectationRules';  File='02_SampleExpectationRules.csv' },
        @{ Key='ErrorCodes';              File='03_ErrorCodes.csv' },
        @{ Key='MissingSamplesConfig';    File='04_MissingSamplesConfig.csv' },
        @{ Key='SampleIdMarkers';         File='05_SampleIdMarkers.csv' },
        @{ Key='ParityCheckConfig';       File='06_ParityCheckConfig.csv' },
        @{ Key='SampleNumberRules';       File='07_SampleNumberRules.csv' },
        @{ Key='TestTypePolicy';          File='08_TestTypePolicy.csv' }
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
    Test-RuleBankIntegrity -RuleBank $rbObj2 -Source 'csv'
    try {
        $cnt = @()
        foreach ($t in $expectedTables) {
            $cnt += ("{0}={1}" -f $t, (@($rbObj2.$t).Count))
        }
        _RuleEngine_Log -Text ("🧠 RuleBank laddad från CSV. " + ($cnt -join ', ')) -Severity 'Info'
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
            $mt = ((Get-RowField -Row $r -FieldName 'MatchType') + '').Trim().ToUpperInvariant()
            if ($mt -ne 'REGEX') { continue }
            $pat = ((Get-RowField -Row $r -FieldName 'Pattern') + '')
            if (-not ($pat.Trim())) { continue }
            if (-not $compiled.RegexCache.ContainsKey($pat)) {
                try {
                    $compiled.RegexCache[$pat] = New-Object System.Text.RegularExpressions.Regex($pat, [System.Text.RegularExpressions.RegexOptions]::IgnoreCase)
                } catch {
                    # Invalid regex: cache a regex that never matches to preserve "false" outcome deterministically
                    $compiled.RegexCache[$pat] = New-Object System.Text.RegularExpressions.Regex('a\A', [System.Text.RegularExpressions.RegexOptions]::IgnoreCase)
                }
            }
        }
    } catch {}

    try { $RuleBank | Add-Member -NotePropertyName 'Compiled' -NotePropertyValue ([pscustomobject]$compiled) -Force } catch { $RuleBank.Compiled = [pscustomobject]$compiled }
    return $RuleBank
}

function Get-ResultCallPatternsForAssay {
    param(
        [Parameter(Mandatory)][pscustomobject]$RuleBank,
        [Parameter(Mandatory)][string]$Assay
    )
    $aKey = (($Assay + '').Trim())
    if (-not $aKey) { $aKey = '(blank)' }

    $c = $RuleBank.Compiled
    if (-not $c) { return @($RuleBank.ResultCallPatterns) }

    if ($c.PatternsByAssay.ContainsKey($aKey)) { return @($c.PatternsByAssay[$aKey]) }

    $list = New-Object System.Collections.Generic.List[object]
    foreach ($r in @($RuleBank.ResultCallPatterns)) {
        if (-not $r) { continue }
        if (-not (Test-RuleEnabled $r)) { continue }
        $ruleAssay = (Get-RowField -Row $r -FieldName 'Assay')
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
    $aKey = (($Assay + '').Trim())
    if (-not $aKey) { $aKey = '(blank)' }

    $c = $RuleBank.Compiled
    if (-not $c) { return @($RuleBank.SampleExpectationRules) }

    if ($c.ExpectRulesByAssay.ContainsKey($aKey)) { return @($c.ExpectRulesByAssay[$aKey]) }

    $list = New-Object System.Collections.Generic.List[object]
    foreach ($r in @($RuleBank.SampleExpectationRules)) {
        if (-not $r) { continue }
        if (-not (Test-RuleEnabled $r)) { continue }
        $ruleAssay = (Get-RowField -Row $r -FieldName 'Assay')
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

    $t = ($Text + '')
    $p = ($Pattern + '')
    $m = ($MatchType + '').Trim().ToUpperInvariant()
    if (-not $m) { $m = 'CONTAINS' }

    try {
        switch ($m) {
            'REGEX'  {
                if (($RegexCache -is [hashtable]) -and $RegexCache.ContainsKey($p)) {
                    return $RegexCache[$p].IsMatch($t)
                }
                return [regex]::IsMatch($t, $p, [System.Text.RegularExpressions.RegexOptions]::IgnoreCase)
            }
            'EQUALS' {
                return (($t.Trim()).ToUpperInvariant() -eq ($p.Trim()).ToUpperInvariant())
            }
            'PREFIX' {
                $tt = ($t.Trim()).ToUpperInvariant()
                $pp = ($p.Trim()).ToUpperInvariant()
                if (-not $pp) { return $true }
                return ($tt.Length -ge $pp.Length -and $tt.Substring(0, $pp.Length) -eq $pp)
            }
            'SUFFIX' {
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
    $aKey = (($Assay + '').Trim())
    if (-not $aKey) { $aKey = '(blank)' }

    $c = $RuleBank.Compiled
    if (-not $c) { return (Get-TestTypePolicyForAssay -Assay $Assay -Policies $RuleBank.TestTypePolicy) }

    if ($c.PolicyByAssay.ContainsKey($aKey)) { return $c.PolicyByAssay[$aKey] }

    $pol = $null
    foreach ($p in @($RuleBank.TestTypePolicy)) {
        try {
            if (((Get-RowField -Row $p -FieldName 'Enabled') + '').Trim().Length -gt 0 -and ((Get-RowField -Row $p -FieldName 'Enabled') + '').Trim().ToUpperInvariant() -in @('FALSE','0','NO','N')) { continue }
            $pat = ((Get-RowField -Row $p -FieldName 'AssayPattern') + '')
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
        [Parameter(Mandatory=$false)][string]$BagNo = ''
    )

    $aKey = (($Assay + '').Trim()); if (-not $aKey) { $aKey = '(blank)' }
    $ccKey = (($ControlCode + '').Trim())
    $bnKey = (($BagNo + '').Trim())
    $key = $aKey + '|' + $ccKey + '|' + $bnKey

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

    if ($null -eq $Row) { return '' }
    try {
        if ($Row -is [hashtable]) {
            if ($Row.ContainsKey($FieldName) -and $null -ne $Row[$FieldName]) { return $Row[$FieldName] }
            return ''
        }
        if ($Row -is [System.Collections.IDictionary]) {
            if ($Row.Contains($FieldName) -and $null -ne $Row[$FieldName]) { return $Row[$FieldName] }
            return ''
        }
    } catch {}

    try {
        $p = $Row.PSObject.Properties[$FieldName]
        if ($p -and $null -ne $p.Value) { return $p.Value }
    } catch {}

    return ''
}

function Test-RuleEnabled {
    param([Parameter(Mandatory)][object]$Rule)
    $en = (Get-RowField -Row $Rule -FieldName 'Enabled')
    if ($en -eq $null) { return $true }
    $s = ($en + '').Trim().ToUpperInvariant()
    if (-not $s) { return $true }
    return ($s -in @('TRUE','1','YES','Y'))
}

function Test-AssayMatch {
    param(
        [Parameter(Mandatory)][string]$RuleAssay,
        [Parameter(Mandatory)][string]$RowAssay
    )
    $ra = ($RuleAssay + '').Trim()
    if (-not $ra -or $ra -eq '*') { return $true }

    $row = ($RowAssay + '').Trim()
    if ($ra -like '*[*?]*') {
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
            if ((($p.Enabled + '')).Trim().Length -gt 0 -and ($p.Enabled + '').Trim().ToUpperInvariant() -in @('FALSE','0','NO','N')) { continue }
            $pat = ($p.AssayPattern + '')
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

    $t = ($Text + '')
    $p = ($Pattern + '')
    $m = ($MatchType + '').Trim().ToUpperInvariant()
    if (-not $m) { $m = 'CONTAINS' }

    try {
        switch ($m) {
            'REGEX'  { return [regex]::IsMatch($t, $p, [System.Text.RegularExpressions.RegexOptions]::IgnoreCase) }
            'EQUALS' { return (($t.Trim()).ToUpperInvariant() -eq ($p.Trim()).ToUpperInvariant()) }
            'PREFIX' {
                $tt = ($t.Trim()).ToUpperInvariant()
                $pp = ($p.Trim()).ToUpperInvariant()
                if (-not $pp) { return $true }
                return ($tt.Length -ge $pp.Length -and $tt.Substring(0, $pp.Length) -eq $pp)
            }
            'SUFFIX' {
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
$status = (Get-RowField -Row $Row -FieldName 'Status')
    $errTxt = (Get-RowField -Row $Row -FieldName 'Error')
    $testResult = (Get-RowField -Row $Row -FieldName 'Test Result')
    $assay = (Get-RowField -Row $Row -FieldName 'Assay')

    if (($errTxt + '').Trim()) {
        return [pscustomobject]@{ Call='ERROR'; Reason='Error column populated' }
    }
    $st = ($status + '').Trim()
    if ($st -and ($st -ine 'Done')) {
        return [pscustomobject]@{ Call='ERROR'; Reason=("Status=" + $st) }
    }

    $tr = ($testResult + '').Trim()
    if (-not $tr) { return [pscustomobject]@{ Call='UNKNOWN'; Reason='Blank Test Result' } }

    $ass = ($assay + '')
    if ($ass -match 'MTB') {
        if ($tr -match '(?i)\bMTB\s+TRACE\s+DETECTED\b') {
            return [pscustomobject]@{ Call='POS'; Reason='MTB Trace detected (override)' }
        }
        if ($tr -match '(?i)\bMTB\s+DETECTED\b') {
            return [pscustomobject]@{ Call='POS'; Reason='MTB detected (override)' }
        }
        if ($tr -match '(?i)\bMTB\s+NOT\s+DETECTED\b') {
            return [pscustomobject]@{ Call='NEG'; Reason='MTB not detected (override)' }
        }
    }

    $hasErr = $false
    $hasNeg = $false
    $hasPos = $false
    $isMixed = $false
    if (-not ($ass -match '(?i)MTB')) {
        $u = ($tr + '').ToUpperInvariant()
        $u = [regex]::Replace($u, '\s+', ' ').Trim()

        $hasErr = ($u -match '\bINVALID\b') -or ($u -match 'NO\s+RESULT') -or ($u -match '\bERROR\b')
        $hasNeg = ($u -match 'NOT\s+DETECTED') -or ($u -match '\bNEGATIVE\b')

        $uNoNotDetected = ($u -replace 'NOT\s+DETECTED', '')
        $hasPos = ($uNoNotDetected -match '\bDETECTED\b') -or ($uNoNotDetected -match '\bPOSITIVE\b')

        $isMixed = ($hasPos -and $hasNeg)
    }

    foreach ($r in $Patterns) {
        if (-not (Test-RuleEnabled $r)) { continue }
        $ruleAssay = (Get-RowField -Row $r -FieldName 'Assay')
        if (-not (Test-AssayMatch -RuleAssay $ruleAssay -RowAssay $assay)) { continue }

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
    }

    if (-not ($ass -match '(?i)MTB')) {
        $u = ($tr + '').ToUpperInvariant()
        $u = [regex]::Replace($u, '\s+', ' ').Trim()

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
    }

    return [pscustomobject]@{ Call='UNKNOWN'; Reason='No pattern matched' }
}
function Get-ExpectedCallDetailed {
    param(
        [Parameter(Mandatory)][object]$Row,
        [Parameter(Mandatory=$false)][object[]]$Rules = @(),
        [Parameter(Mandatory=$false)][object]$RegexCache = $null
    )

    if (-not $Rules) { $Rules = @() }
$sampleId = (Get-RowField -Row $Row -FieldName 'Sample ID')
    $assay    = (Get-RowField -Row $Row -FieldName 'Assay')
    $sid = ($sampleId + '').Trim()
    if (-not $sid) { return [pscustomobject]@{ Call=''; Reason='Blank Sample ID' } }

    foreach ($r in $Rules) {
        if (-not (Test-RuleEnabled $r)) { continue }
        $ruleAssay = (Get-RowField -Row $r -FieldName 'Assay')
        if (-not (Test-AssayMatch -RuleAssay $ruleAssay -RowAssay $assay)) { continue }

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
    }

    return [pscustomobject]@{ Call=''; Reason='No expectation rule matched' }
}

function Get-ExpectedTestTypeDerived {
    param([Parameter(Mandatory)][string]$SampleId)

    $parts = $SampleId.Split('_')
    if ($parts.Count -ge 3) {
        $tc = $parts[2]
        switch -Regex ($tc) {
            '^0$' { return 'Negative Control 1' }
            '^1$' { return 'Positive Control 1' }
            '^2$' { return 'Positive Control 2' }
            '^3$' { return 'Positive Control 3' }
            '^4$' { return 'Positive Control 4' }
            '^5$' { return 'Positive Control 5' }
            default { }
        }
    }
    return 'Specimen'
}

function Build-ErrorCodeLookup {
    param([Parameter(Mandatory)][object[]]$ErrorCodes)
    $lut = @{
        Codes = @{}
        Defaults = New-Object System.Collections.Generic.List[object]
        NamedBlanks = New-Object System.Collections.Generic.List[object]
    }

    foreach ($r in $ErrorCodes) {
        $code = ((Get-RowField -Row $r -FieldName 'ErrorCode') + '').Trim()
        $name = (Get-RowField -Row $r -FieldName 'Name')
        $ret  = (Get-RowField -Row $r -FieldName 'GeneratesRetest')

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
    }

    return $lut
}

function Get-ErrorInfo {
    param(
        [Parameter(Mandatory)][object]$Row,
        [Parameter(Mandatory)][hashtable]$ErrorLut,
        [Parameter(Mandatory)][string]$DelamPattern
    )

    $errTxt = (Get-RowField -Row $Row -FieldName 'Error')
    $mpTxt  = (Get-RowField -Row $Row -FieldName 'Max Pressure (PSI)')

    $code = ''
    $hasErr = (($errTxt + '').Trim().Length -gt 0)

    if ($hasErr) {
        if (($errTxt + '') -match '(\d{4,5})') { $code = $Matches[1] }
    }

    $name = ''
    $retest = ''

    if ($hasErr) {
        if ($code -and $ErrorLut.Codes.ContainsKey($code)) {
            $name   = $ErrorLut.Codes[$code].Name
            $retest = $ErrorLut.Codes[$code].GeneratesRetest
        } else {
            $picked = $null
            try {
                foreach ($d in $ErrorLut.Defaults) {
                    if (($d.Name + '') -match '(?i)Delamination' -and ($errTxt + '') -match $DelamPattern) {
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
        if (($mpTxt + '').Trim()) { $pressure = [double]($mpTxt + '') }
    } catch {}

    $pressureFlag = $false
    if ($pressure -ne $null -and $pressure -ge 90) { $pressureFlag = $true }

    # If pressure flag and there is a named blank describing it, surface the name
    if ($pressureFlag -and -not $name) {
        try {
            foreach ($b in $ErrorLut.NamedBlanks) {
                if (($b.Name + '') -match '(?i)Max\s+Pressure') {
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
    }
}

function Classify-Deviation {
    param(
        [AllowEmptyString()][string]$Expected,
        [AllowEmptyString()][string]$Observed
    )
    $e = ($Expected + '').Trim().ToUpperInvariant()
    $o = ($Observed + '').Trim().ToUpperInvariant()

    if (-not $e) { return 'UNKNOWN' }
    if ($o -eq 'ERROR') { return 'ERROR' }
    if ($o -eq 'UNKNOWN' -or -not $o) { return 'UNKNOWN' }
    if ($e -eq $o) { return 'OK' }
    if ($o -eq 'MIXED') {
    if ($e -eq 'POS') { return 'FN' }
    if ($e -eq 'NEG') { return 'FP' }
    return 'MISMATCH'
    }
    if ($e -eq 'POS' -and $o -eq 'NEG') { return 'FN' }
    if ($e -eq 'NEG' -and $o -eq 'POS') { return 'FP' }
    return 'MISMATCH'
}

function Split-CsvLineQuoted {
    param(
        [Parameter(Mandatory)][string]$Line,
        [Parameter(Mandatory)][string]$Delimiter
    )
    $d = [regex]::Escape($Delimiter)
    $rx = $d + '(?=(?:(?:[^"]*"){2})*[^"]*$)'
    return [regex]::Split($Line, $rx)
}

function Get-HeaderFromTestSummaryFile {
    param([Parameter(Mandatory)][string]$CsvPath)

    if (-not (Test-Path -LiteralPath $CsvPath)) { return @() }

    $delim = ','
    try { $delim = Get-CsvDelimiter -Path $CsvPath } catch {}

    $lines = @()
    try { $lines = Get-Content -LiteralPath $CsvPath -ErrorAction Stop } catch { return @() }

    # Test Summary: header is line 8 (index 7)
    if (-not $lines -or $lines.Count -lt 8) { return @() }
    $hdrLine = $lines[7]
    if (-not $hdrLine) { return @() }

    $headers = Split-CsvLineQuoted -Line $hdrLine -Delimiter $delim
    $headers = @($headers | ForEach-Object { (($_ + '') -replace '^"|"$','').Trim() })
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

        $o = [ordered]@{}
        $max = [Math]::Min($Headers.Count, $arr.Count)
        for ($i=0; $i -lt $max; $i++) {
            $h = $Headers[$i]
            if (-not $h) { continue }
            $o[$h] = $arr[$i]
        }
        $out.Add([pscustomobject]$o)
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
    $aKey = (($Assay + '').Trim())
    if (-not $aKey) { $aKey = '(blank)' }
    $tKey = (($MarkerType + '').Trim().ToUpperInvariant())
    $mKey = ($aKey + '|' + $tKey)

    try {
        if ($RuleBank.Compiled -and $RuleBank.Compiled.MarkerByAssayType) {
            $mc = $RuleBank.Compiled.MarkerByAssayType
            if ($mc.ContainsKey($mKey)) {
                $v = $mc[$mKey]
                if ($v -eq '__MISS__') { return '' }
                return ($v + '')
            }
        }
    } catch {}

    foreach ($r in $RuleBank.SampleIdMarkers) {
        if (-not (Test-RuleEnabled $r)) { continue }

        $ap = ((Get-RowField -Row $r -FieldName 'AssayPattern') + '').Trim()
        if (-not (Test-AssayMatch -RuleAssay $ap -RowAssay $Assay)) { continue }

        $mt = ((Get-RowField -Row $r -FieldName 'MarkerType') + '').Trim()
        if (-not $mt) { continue }
        if ($mt -ine $MarkerType) { continue }

        $m = ((Get-RowField -Row $r -FieldName 'Marker') + '').Trim()
        try { if ($RuleBank.Compiled -and $RuleBank.Compiled.MarkerByAssayType) { $RuleBank.Compiled.MarkerByAssayType[$mKey] = $m } } catch {}
        return $m
    }
    try { if ($RuleBank.Compiled -and $RuleBank.Compiled.MarkerByAssayType) { $RuleBank.Compiled.MarkerByAssayType[$mKey] = '__MISS__' } } catch {}

    return ''
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
        CartridgeField = 'Cartridge S/N'
        TokenIndex = 3
        XChar = 'X'
        PlusChar = '+'
        NumericRatioThreshold = 0.60
        DelaminationMarkerType = 'DelaminationCodeRegex'
        DelamRegex = 'D\d{1,2}[A-Z]?'
        ValidSuffixRegex = 'X|\+'
        SampleTypeCodeTokenIndex = 2
        SampleNumberTokenIndex = 3
    }

    $delam = Get-MarkerValue -RuleBank $RuleBank -Assay $Assay -MarkerType 'DelaminationCodeRegex'
    if ($delam) { $cfg.DelamRegex = $delam }

    $suffix = Get-MarkerValue -RuleBank $RuleBank -Assay $Assay -MarkerType 'SuffixChars'
    if ($suffix) {
        while ($suffix -like '*\\*') { $suffix = $suffix.Replace('\\','\') }
        $cfg.ValidSuffixRegex = $suffix
    }

    $stIdx = Get-IntMarkerValue -RuleBank $RuleBank -Assay $Assay -MarkerType 'SampleTypeCodeTokenIndex' -Default 2
    $snIdx = Get-IntMarkerValue -RuleBank $RuleBank -Assay $Assay -MarkerType 'SampleNumberTokenIndex' -Default 3
    $cfg.SampleTypeCodeTokenIndex = $stIdx
    $cfg.SampleNumberTokenIndex = $snIdx

    foreach ($r in $RuleBank.ParityCheckConfig) {
        if (-not (Test-RuleEnabled $r)) { continue }

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

    $sid = (Get-RowField -Row $Row -FieldName 'Sample ID')
    if (($sid + '').Trim()) {
        $parts = ($sid + '').Split('_')
        if ($parts.Count -gt $SampleTypeCodeTokenIndex) {
            $cc = ($parts[$SampleTypeCodeTokenIndex] + '').Trim()
            if ($cc -match '^\d+$') { return $cc }
        }
        if ($parts.Count -ge 3) {
            $cc2 = ($parts[2] + '').Trim()
            if ($cc2 -match '^\d+$') { return $cc2 }
        }
    }

    $tt = (Get-RowField -Row $Row -FieldName 'Test Type')
    if (($tt + '') -match '(?i)Negative\s+Control') { return '0' }
    if (($tt + '') -match '(?i)Positive\s+Control\s+(\d+)') { return $Matches[1] }

    return ''
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

    $tok = ''
    $base = ''

    $parts = $SampleId.Split('_')
    if ($parts.Count -gt $TokenIndex) {
        $tok = ($parts[$TokenIndex] + '').Trim()
    }

    if (-not $tok) { return [pscustomobject]@{ SampleToken=''; BaseToken=''; ActualSuffix=''; SampleNum=''; SampleNumRaw=''; } }

    # strip trailing delamination code if present INSIDE token
    $rx = "([_-]?(?:" + $DelamPattern + "))$"
    try {
        $base = [regex]::Replace($tok, $rx, '', [System.Text.RegularExpressions.RegexOptions]::IgnoreCase)
    } catch {
        $base = $tok
    }

    $base = ($base + '').Trim()

    $act = ''
    if ($base.Length -ge 1) {
        $last = $base.Substring($base.Length - 1, 1)
        if ($last -match ('^(?:' + $ValidSuffixRegex + ')$')) {
            $u = $last.ToUpperInvariant()
            if ($u -eq $XChar.ToUpperInvariant()) { $act = $XChar.ToUpperInvariant() }
            elseif ($last -eq $PlusChar) { $act = $PlusChar }
            else { $act = $u }
        }
    }

    $numRaw = ''
    $num = ''
    if ($base -match '^(\d{1,4})') {
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
        Prefix = ''
        BagNo = ''
        SampleCode = ''
        RunToken = ''
        RunNoRaw = ''
        RunNo = ''
        RunSuffix = ''
        ReplacementLevel = 0
        DelamPresent = $false
        DelamToken = ''
        DelamCodes = @()
    }

    $sid = ($SampleId + '').Trim()
    if (-not $sid) { return [pscustomobject]$out }

    $parts = $sid.Split('_')
    if ($parts.Count -ge 1) { $out.Prefix = (($parts[0] + '').Trim()).ToUpperInvariant() }
    if ($parts.Count -ge 2) { $out.BagNo = (($parts[1] + '').Trim()) }
    if ($parts.Count -ge 3) { $out.SampleCode = (($parts[2] + '').Trim()) }
    if ($parts.Count -ge 4) { $out.RunToken = (($parts[3] + '').Trim()) }

    if ($parts.Count -ge 5) {
        $dt = (($parts[4] + '').Trim())
        if ($dt) {
            $out.DelamToken = $dt
            if ($dt -match '^(?i)D') { $out.DelamPresent = $true }
        }
    }

    if (-not $out.DelamPresent -and $DelamRegex) {
        try {
            $rx = '(?i)(?:^|[_-])(' + $DelamRegex + ')(?:$|[,;_ -])'
            if ([regex]::IsMatch($sid, $rx)) { $out.DelamPresent = $true }
        } catch {}
    }

    $rt = ($out.RunToken + '').Trim()
    if ($rt.Length -ge 1) {
        $last = $rt.Substring($rt.Length - 1, 1)
        if ($last -eq $PlusChar -or $last.ToUpperInvariant() -eq $XChar.ToUpperInvariant()) {
            $out.RunSuffix = $last.ToUpperInvariant()
            $core = $rt.Substring(0, $rt.Length - 1)

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
    }

    # Parse delamination codes list (if present)
    if ($out.DelamToken) {
        $codes = @()
        foreach ($c in ($out.DelamToken -split ',')) {
            $t = ($c + '').Trim()
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
        [Parameter(Mandatory=$false)][string]$BagNo = '',
        [Parameter(Mandatory)][object[]]$Rules
    )

    foreach ($r in $Rules) {
        if (-not (Test-RuleEnabled $r)) { continue }

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

if (-not $RuleBank) { throw 'RuleEngine: RuleBank is null.' }

Test-RuleBankIntegrity -RuleBank $RuleBank -Source 'runtime'

try { $RuleBank = Compile-RuleBank -RuleBank $RuleBank } catch {}

    $needsConvert = $false
    try {
        if ($CsvObjects[0] -is [object[]]) { $needsConvert = $true }
        else {
            $p1 = $CsvObjects[0].PSObject.Properties.Match('Sample ID')
            if ($p1.Count -eq 0) { $needsConvert = $true }
        }
    } catch { $needsConvert = $true }

    if ($needsConvert) {
        if (-not $CsvPath) { throw 'RuleEngine: CsvPath is required to convert field-array rows to objects.' }
        $hdr = Get-HeaderFromTestSummaryFile -CsvPath $CsvPath
        if (-not $hdr -or $hdr.Count -lt 5) { throw 'RuleEngine: Could not read CSV header (line 8).' }
        $CsvObjects = Convert-FieldRowsToObjects -FieldRows $CsvObjects -Headers $hdr
        if (-not $CsvObjects -or $CsvObjects.Count -eq 0) {
            return [pscustomobject]@{ Rows=@(); Summary=[pscustomobject]@{ Total=0; ObservedCounts=@{}; DeviationCounts=@{}; RetestYes=0 }; TopDeviations=@() }
        }
    }

    $byAssay = @{}
    foreach ($row in $CsvObjects) {
        $a = (Get-RowField -Row $row -FieldName 'Assay')
        $key = (($a + '').Trim())
        if (-not $key) { $key = '(blank)' }
        if (-not $byAssay.ContainsKey($key)) { $byAssay[$key] = New-Object System.Collections.Generic.List[object] }
        $byAssay[$key].Add($row)
    }

    $results = New-Object System.Collections.Generic.List[object]
    $errLut = Build-ErrorCodeLookup -ErrorCodes $RuleBank.ErrorCodes

    foreach ($assayKey in $byAssay.Keys) {
        $group = $byAssay[$assayKey]
        if (-not $group -or $group.Count -eq 0) { continue }

        $parCfg = Get-ParityConfigForAssay -RuleBank $RuleBank -Assay $assayKey

        $patternsForAssay = @(Get-ResultCallPatternsForAssay -RuleBank $RuleBank -Assay $assayKey)
        $expectForAssay   = @(Get-ExpectationRulesForAssay -RuleBank $RuleBank -Assay $assayKey)
        if (-not $patternsForAssay) { $patternsForAssay = @() }
        if (-not $expectForAssay) { $expectForAssay = @() }

        $regexCache = $null
        try { if ($RuleBank.Compiled -and $RuleBank.Compiled.RegexCache) { $regexCache = $RuleBank.Compiled.RegexCache } } catch { $regexCache = $null }
        $delamPattern = $parCfg.DelamRegex
        $validSuffix = $parCfg.ValidSuffixRegex

        $numeric = New-Object System.Collections.Generic.List[long]
        $parityPairs = New-Object System.Collections.Generic.List[object]
        $suffixCounts = @{}
        $suffixCounts[$parCfg.XChar.ToUpperInvariant()] = 0
        $suffixCounts[$parCfg.PlusChar] = 0

        foreach ($row in $group) {
            $sn = (Get-RowField -Row $row -FieldName $parCfg.CartridgeField)
            $snN = $null
            if (($sn + '').Trim() -match '^\d+$') {
                try { $snN = [long]($sn + '') } catch {}
                try { $numeric.Add([long]($sn + '')) } catch {}
            }

            $sid = (Get-RowField -Row $row -FieldName 'Sample ID')
            if (($sid + '').Trim()) {
                $t = Get-SampleTokenAndBase -SampleId ($sid + '') -TokenIndex $parCfg.TokenIndex -DelamPattern $delamPattern -ValidSuffixRegex $validSuffix -XChar $parCfg.XChar -PlusChar $parCfg.PlusChar
                if ($t.ActualSuffix -and $suffixCounts.ContainsKey($t.ActualSuffix)) {
                    $suffixCounts[$t.ActualSuffix]++
                    if ($snN -ne $null) {
                        try { $parityPairs.Add([pscustomobject]@{ SN=$snN; Sfx=$t.ActualSuffix }) } catch {}
                    }
                }
            }
        }

        $numRatio = 0.0
        try { $numRatio = [double]$numeric.Count / [double]$group.Count } catch {}

        $useParity = ($parCfg.UseParity -and $numeric.Count -gt 0 -and $numRatio -ge $parCfg.NumericRatioThreshold)

        $minSn = $null
        $parityForX = $null
        if ($useParity) {
            try { $minSn = ($numeric | Measure-Object -Minimum).Minimum } catch { $minSn = $null }
            if ($parityPairs) {
                $m0 = 0; $m1 = 0; $tot = 0
                foreach ($pp in $parityPairs) {
                    $sn = $null
                    try { $sn = [long]$pp.SN } catch { $sn = $null }
                    if ($sn -eq $null) { continue }
                    $sfx = (($pp.Sfx + '')).Trim()
                    if (-not $sfx) { continue }
                    $par = [int]($sn % 2)

                    $exp0 = if ($par -eq 0) { $parCfg.XChar.ToUpperInvariant() } else { $parCfg.PlusChar }
                    $exp1 = if ($par -eq 1) { $parCfg.XChar.ToUpperInvariant() } else { $parCfg.PlusChar }

                    $tot++
                    if ($sfx -eq $exp0) { $m0++ }
                    if ($sfx -eq $exp1) { $m1++ }
                }

                if ($tot -gt 0) {
                    if ($m0 -eq $tot) { $parityForX = 0 }
                    elseif ($m1 -eq $tot) { $parityForX = 1 }
                }
            }
if ($parityForX -eq $null) { $useParity = $false }
        }

        $majSuffix = ''
        if (-not $useParity) {
            $xCount = $suffixCounts[$parCfg.XChar.ToUpperInvariant()]
            $pCount = $suffixCounts[$parCfg.PlusChar]
            $totS = 0
            try { $totS = [int]$xCount + [int]$pCount } catch { $totS = 0 }
            if ($totS -gt 0) {
                if ($xCount -eq $totS) { $majSuffix = $parCfg.XChar.ToUpperInvariant() }
                elseif ($pCount -eq $totS) { $majSuffix = $parCfg.PlusChar }
            }
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

                $expectedSuffix = ''
                $suffixSource = ''
                $suffixCheck = ''

                $snVal = (Get-RowField -Row $row -FieldName $parCfg.CartridgeField)
                $snNum = $null
                if (($snVal + '').Trim() -match '^\d+$') { try { $snNum = [long]($snVal + '') } catch {} }

                if ($tokInfo.ActualSuffix) {
                    if ($useParity -and $snNum -ne $null -and $parityForX -ne $null) {
                        $expS = if (([int]($snNum % 2)) -eq $parityForX) { $parCfg.XChar.ToUpperInvariant() } else { $parCfg.PlusChar }
                        $expectedSuffix = $expS
                        $suffixSource = 'PARITY'
                    } elseif ($majSuffix) {
                        $expectedSuffix = $majSuffix
                        $suffixSource = 'MAJORITY'
                    }

                    if ($expectedSuffix) {
                        $suffixCheck = if ($tokInfo.ActualSuffix -eq $expectedSuffix) { 'OK' } else { 'BAD' }
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
    }

    function _Append-RuleFlag {
        param([pscustomobject]$row, [string]$flag)
        $f = (($row.RuleFlags + '')).Trim()
        if (-not $f) { $row.RuleFlags = $flag; return }
        $parts = $f.Split('|')
        if ($parts -contains $flag) { return }
        $row.RuleFlags = ($f + '|' + $flag)
    }

    $distinctAssays = @($results | ForEach-Object { ($_.Assay + '').Trim() } | Where-Object { $_ } | Sort-Object -Unique)
    $distinctAssayVersions = @($results | ForEach-Object { ($_.AssayVersion + '').Trim() } | Where-Object { $_ } | Sort-Object -Unique)
    $distinctReagentLots = @($results | ForEach-Object { ($_.ReagentLotId + '').Trim() } | Where-Object { $_ } | Sort-Object -Unique)

    $majorAssay = ''
    if ($distinctAssays.Count -gt 1) {
        try { $majorAssay = ($results | Group-Object Assay | Sort-Object Count -Descending | Select-Object -First 1).Name } catch {}
        foreach ($r in $results) {
            $a = ((($r.Assay + '')).Trim())
            if ($majorAssay -and $a -and $a -ne $majorAssay) { _Append-RuleFlag -row $r -flag 'DQ_ASSAY_OUTLIER' }
        }
    }

    $majorVer = ''
    if ($distinctAssayVersions.Count -gt 1) {
        try { $majorVer = ($results | Group-Object AssayVersion | Sort-Object Count -Descending | Select-Object -First 1).Name } catch {}
        foreach ($r in $results) {
            $v = ((($r.AssayVersion + '')).Trim())
            if ($majorVer -and $v -and $v -ne $majorVer) { _Append-RuleFlag -row $r -flag 'DQ_ASSAYVER_OUTLIER' }
        }
    }

    $dupSample = @($results | Where-Object { (($_.SampleId + '').Trim()) } | Group-Object SampleId | Where-Object { $_.Count -gt 1 })
    if ($dupSample.Count -gt 0) {
        $dupSet = @{}
        foreach ($g in $dupSample) { $dupSet[$g.Name] = $true }
        foreach ($r in $results) {
            $sid = ((($r.SampleId + '')).Trim())
            if ($sid -and $dupSet.ContainsKey($sid)) { _Append-RuleFlag -row $r -flag 'DQ_DUP_SAMPLEID' }
        }
    }

    $dupCart = @($results | Where-Object { (($_.CartridgeSN + '').Trim()) } | Group-Object CartridgeSN | Where-Object { $_.Count -gt 1 })
    if ($dupCart.Count -gt 0) {
        $dupSet = @{}
        foreach ($g in $dupCart) { $dupSet[$g.Name] = $true }
        foreach ($r in $results) {
            $csn = ((($r.CartridgeSN + '')).Trim())
            if ($csn -and $dupSet.ContainsKey($csn)) { _Append-RuleFlag -row $r -flag 'DQ_DUP_CARTSN' }
        }
    }

    $useStrictTestType = $false
    try {
        $ttAll = @($results | Where-Object { (($_.ExpectedTestType + '')).Trim() -and (($_.TestType + '')).Trim() })
        $ttCtl = @($ttAll | Where-Object { (($_.ExpectedTestType + '')).Trim().ToUpperInvariant() -ne 'SPECIMEN' })
        if ($ttCtl.Count -ge 5) {
            $ttMatch = @($ttCtl | Where-Object { (($_.TestType + '')).Trim().ToUpperInvariant() -eq (($_.ExpectedTestType + '')).Trim().ToUpperInvariant() }).Count
            $ttRate = 0.0
            try { $ttRate = [double]$ttMatch / [double]$ttCtl.Count } catch { $ttRate = 0.0 }
            if ($ttRate -ge 0.80) { $useStrictTestType = $true }
        }
    } catch { $useStrictTestType = $false }

    if ($useStrictTestType) {
        foreach ($r in $results) {
            $rf = ((($r.RuleFlags + '')).Trim())
            if ($rf) {
                $p = $rf.Split('|')
                if ($p -contains 'DQ_ASSAY_OUTLIER' -or $p -contains 'DQ_ASSAYVER_OUTLIER') { continue }
            }

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
   }

    foreach ($r in $results) {
        $rf = ((($r.RuleFlags + '')).Trim())
        $isOutlier = $false
        if ($rf) {
            $p = $rf.Split('|')
            if ($p -contains 'DQ_ASSAY_OUTLIER' -or $p -contains 'DQ_ASSAYVER_OUTLIER') { $isOutlier = $true }
        }
        if ($isOutlier) { continue }
        $sc = ((($r.SuffixCheck + '')).Trim().ToUpperInvariant())
        if ($sc -and $sc -ne 'OK') { _Append-RuleFlag -row $r -flag ('SUFFIX_' + $sc) }
        $snok = ((($r.SampleNumOK + '')).Trim().ToUpperInvariant())
        if ($snok -eq 'NO') { _Append-RuleFlag -row $r -flag 'SAMPLENUM_BAD' }
    }

    $useStrictPrefix = $false
    try {
        $p0 = @($results | Where-Object { (($_.SampleCode + '')).Trim() -eq '0' -and (($_.SamplePrefix + '')).Trim() })
        $pP = @($results | Where-Object { (($_.SampleCode + '')).Trim() -match '^[1-5]$' -and (($_.SamplePrefix + '')).Trim() })

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
    } catch { $useStrictPrefix = $false }

    foreach ($r in $results) {

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
    }

    $hotModules = @{}
    $byModErr = @($results | Where-Object { (($_.ModuleSN + '').Trim()) -and (($_.ObservedCall + '').Trim().ToUpperInvariant() -eq 'ERROR') } | Group-Object ModuleSN)
    foreach ($g in $byModErr) {
        if ($g.Count -ge 3) { $hotModules[$g.Name] = $g.Count }
    }
    if ($hotModules.Count -gt 0) {
        foreach ($r in $results) {
            $m = ((($r.ModuleSN + '')).Trim())
            if ($m -and $hotModules.ContainsKey($m)) { _Append-RuleFlag -row $r -flag 'MODULE_ERR_HOTSPOT' }
        }
    }

    $qc = [pscustomobject]@{
        DistinctAssays = $distinctAssays
        DistinctAssayVersions = $distinctAssayVersions
        DistinctReagentLots = $distinctReagentLots
        DuplicateSampleIdCount = ($dupSample | ForEach-Object { $_.Name } | Select-Object -Unique).Count
        DuplicateCartridgeSnCount = ($dupCart | ForEach-Object { $_.Name } | Select-Object -Unique).Count
        HotModuleCount = $hotModules.Count
        DelamCount = @($results | Where-Object { try { [bool]$_.DelamPresent } catch { $false } }).Count
        ReplacementCount = @($results | Where-Object { try { [int]$_.ReplacementLevel -ge 1 } catch { $false } }).Count
        BadPrefixCount = @($results | Where-Object { (($_.RuleFlags + '') -split '\|') -contains 'PREFIX_BAD' }).Count
        BadRunNoCount  = @($results | Where-Object { (($_.RuleFlags + '') -split '\|') -contains 'RUNNO_BAD' }).Count
    }
    # ---------------------------------------------------------------
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
            $code = ((($r.ErrorCode + '')).Trim())
            $known = $false
            if ($code -match '^\d{4,5}$') { try { $known = $errLut.Codes.ContainsKey($code) } catch { $known = $false } }

            $isMtbInd = $false
            try {
                if ((($r.Assay + '') -match '(?i)MTB') -and ((($r.TestResult + '') -match '(?i)INDETERMINATE'))) { $isMtbInd = $true }
            } catch {}

            if ($known -or $isMtbInd) { $summary.MinorFunctionalError++ }
            else { $summary.InstrumentError++ }
        }

    }

    $assayList = @($results | ForEach-Object { ($_.Assay + '').Trim() } | Where-Object { $_ } | Select-Object -Unique)
    $ttMatched = 0
    $ttMissing = @()
    $ttDetails = @()
    foreach ($a in $assayList) {
        $pol = $null
        try { $pol = Get-TestTypePolicyForAssayCached -RuleBank $RuleBank -Assay $a } catch { $pol = $null }
        if ($pol) {
            $ttMatched++
            $pat = ((Get-RowField -Row $pol -FieldName 'AssayPattern') + '').Trim()
            $allowed = ((Get-RowField -Row $pol -FieldName 'AllowedTestTypes') + '').Trim()
            if (-not $allowed) { $allowed = ((Get-RowField -Row $pol -FieldName 'TestTypes') + '').Trim() }
            $ttDetails += ($a + ' => ' + $allowed + ' (pattern=' + $pat + ')')
        } else {
            $ttMissing += $a
        }
    }

    $sidTotal = @($results | Where-Object { (($_.SampleId + '')).Trim().Length -gt 0 }).Count
    $sidOk = @($results | Where-Object {
        (($_.SampleId + '')).Trim().Length -gt 0 -and
        (($_.BagNo + '')).Trim().Length -gt 0 -and
        (($_.SampleCode + '')).Trim().Length -gt 0 -and
        (($_.RunNo + '')).Trim().Length -gt 0 -and
        (($_.RunSuffix + '')).Trim().Length -gt 0
    }).Count

    $snTotal = $sidTotal
    $snCovered = @($results | Where-Object {
        (($_.SampleId + '')).Trim().Length -gt 0 -and
        (($_.SampleNumOK + '')).Trim().Length -gt 0
    }).Count

    $summary | Add-Member -NotePropertyName 'TestTypePolicyAssaysTotal' -NotePropertyValue $assayList.Count -Force
    $summary | Add-Member -NotePropertyName 'TestTypePolicyAssaysMatched' -NotePropertyValue $ttMatched -Force
    $summary | Add-Member -NotePropertyName 'TestTypePolicyAssaysMissing' -NotePropertyValue ($ttMissing -join ', ') -Force
    $summary | Add-Member -NotePropertyName 'TestTypePolicyDetails' -NotePropertyValue $ttDetails -Force

    $summary | Add-Member -NotePropertyName 'SampleIdParseTotal' -NotePropertyValue $sidTotal -Force
    $summary | Add-Member -NotePropertyName 'SampleIdParseOk' -NotePropertyValue $sidOk -Force
    $summary | Add-Member -NotePropertyName 'SampleNumberRuleTotal' -NotePropertyValue $snTotal -Force
    $summary | Add-Member -NotePropertyName 'SampleNumberRuleCovered' -NotePropertyValue $snCovered -Force
    # -------------------------------------------
$top = @($results | Where-Object { $_.Deviation -in @('FP','FN','ERROR','MISMATCH') } | Select-Object -First 50)

    return [pscustomobject]@{ Rows = $results.ToArray(); Summary = $summary; TopDeviations = $top; QC = $qc }
}

function Write-RuleEngineDebugSheet {
    param(
        [Parameter(Mandatory)][object]$Pkg,
        [Parameter(Mandatory)][pscustomobject]$RuleEngineResult,
        [Parameter(Mandatory=$false)][bool]$IncludeAllRows = $false
    )

    try {
        $old = $Pkg.Workbook.Worksheets['CSV Sammanfattning']
        if ($old) { $Pkg.Workbook.Worksheets.Delete($old) }
    } catch {}

    $ws = $Pkg.Workbook.Worksheets.Add('CSV Sammanfattning')

    $headers = @(
        'Sample ID',
        'Error Code',
        'Avvikelse',
        'Flagga',
        'Förväntat X/+',
        'Suffix-kontroll',
        'Cartridge S/N',
        'Module S/N',
        'Start Time',
        'Test Type',
        'Förväntad Test Type',
        'Regler/Övrigt',
        'Status',
        'Error Type',
        'Ersätts?',
        'Max Pressure (PSI)',
        'Test Result',
        'Error'
    )

    $row = 1
    $ws.Cells.Item($row,1).Value = 'Sammanfattning felsökning'
    $ws.Cells.Item($row,1).Style.Font.Bold = $true
    $row++

    $sum = $RuleEngineResult.Summary
    $qc  = $RuleEngineResult.QC
    $allRows = @($RuleEngineResult.Rows)

     function _KV {
         param([int]$r, [string]$k, $v, [int]$c = 1)
     $ws.Cells.Item($r,$c).Value = $k
     $ws.Cells.Item($r,$c+1).Value = $v
     $ws.Cells.Item($r,$c).Style.Font.Bold = $true
     }

    $assayTxt = ''
    if ($qc -and $qc.DistinctAssays -and $qc.DistinctAssays.Count -eq 1) { $assayTxt = $qc.DistinctAssays[0] }
    elseif ($qc -and $qc.DistinctAssays -and $qc.DistinctAssays.Count -gt 1) { $assayTxt = 'Flera (' + ($qc.DistinctAssays.Count) + ')' }

    $verTxt = ''
    if ($qc -and $qc.DistinctAssayVersions -and $qc.DistinctAssayVersions.Count -eq 1) { $verTxt = $qc.DistinctAssayVersions[0] }
    elseif ($qc -and $qc.DistinctAssayVersions -and $qc.DistinctAssayVersions.Count -gt 1) { $verTxt = 'Flera (' + ($qc.DistinctAssayVersions.Count) + ')' }

    $lotTxt = ''
    if ($qc -and $qc.DistinctReagentLots -and $qc.DistinctReagentLots.Count -eq 1) { $lotTxt = $qc.DistinctReagentLots[0] }
    elseif ($qc -and $qc.DistinctReagentLots -and $qc.DistinctReagentLots.Count -gt 1) { $lotTxt = 'Flera (' + ($qc.DistinctReagentLots.Count) + ')' }

    _KV -r $row -k 'Totalt tester' -v $sum.Total -c 1
    _KV -r $row -k 'Assay' -v $assayTxt -c 3
    _KV -r $row -k 'Assay Version' -v $verTxt -c 5
    _KV -r $row -k 'Reagent Lot' -v $lotTxt -c 7
    $row++

    $ok = 0; if ($sum.DeviationCounts.ContainsKey('OK')) { $ok = $sum.DeviationCounts['OK'] }
    _KV -r $row -k 'Tester GK' -v $ok; $row++

    foreach ($k in @('FP','FN','ERROR','MISMATCH','UNKNOWN')) {
        if ($sum.DeviationCounts.ContainsKey($k)) {
            $label = switch ($k) {
                'FP' { 'FP' }
                'FN' { 'FN' }
                'ERROR' { 'Deviation ERROR (totalt)' }
                'MISMATCH' { 'Mismatch' }
                'UNKNOWN' { 'Okänt (UNKNOWN)' }
                default { $k }
            }
            _KV -r $row -k $label -v $sum.DeviationCounts[$k]; $row++
        }
    }

    if ($sum -and ($sum.MinorFunctionalError -ne $null -or $sum.InstrumentError -ne $null)) {
        _KV -r $row -k 'Minor Functional Error' -v $sum.MinorFunctionalError; $row++
        _KV -r $row -k 'Instrument Error' -v $sum.InstrumentError; $row++
    }

    if ($sum -and $sum.DelamCount -ne $null) { _KV -r $row -k 'Delamineringar (Sample ID)' -v $sum.DelamCount; $row++ }
    if ($sum -and $sum.ReplacementCount -ne $null) { _KV -r $row -k 'Ersättningar (A/AA/AAA)' -v $sum.ReplacementCount; $row++ }

    foreach ($k in @('POS','NEG','ERROR','UNKNOWN')) {
        if ($sum.ObservedCounts.ContainsKey($k)) {
            _KV -r $row -k ('Observerat ' + $k) -v $sum.ObservedCounts[$k]; $row++
        }
    }

    _KV -r $row -k 'Omkörning (YES)' -v $sum.RetestYes; $row++

    if ($qc) {
        _KV -r $row -k 'Dubbletter av Sample ID' -v $qc.DuplicateSampleIdCount; $row++
        _KV -r $row -k 'Dubbletter av Cartridge S/N' -v $qc.DuplicateCartridgeSnCount; $row++
        _KV -r $row -k 'Moduler med error (≥3 fel)' -v $qc.HotModuleCount; $row++

        if ($qc.DistinctAssays.Count -gt 1) { _KV -r $row -k 'Varning: flera assay' -v ($qc.DistinctAssays -join ', '); $row++ }
        if ($qc.DistinctAssayVersions.Count -gt 1) { _KV -r $row -k 'Varning: flera versioner' -v ($qc.DistinctAssayVersions -join ', '); $row++ }
        if ($qc.DistinctReagentLots.Count -gt 1) { _KV -r $row -k 'Varning: flera reagent lots' -v ($qc.DistinctReagentLots -join ', '); $row++ }
    }

    $pressureGE90 = @($allRows | Where-Object {
        $p = $null
        try { $p = [double]$_.MaxPressure } catch { $p = $null }
        return ($null -ne $p -and $p -ge 90)
    }).Count
    _KV -r $row -k 'Max Pressure ≥ 90 PSI' -v $pressureGE90; $row++

    $row++
    $tableHeaderRow = $row

    $rowsToWrite = $allRows
    if (-not $IncludeAllRows) {
        $rowsToWrite = @($allRows | Where-Object {
            $dev = (($_.Deviation + '')).Trim()
            $hasDeviation = ($dev.Length -gt 0 -and $dev -ne 'OK')

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
    }

    for ($c=1; $c -le $headers.Count; $c++) {
         $ws.Cells.Item($tableHeaderRow,$c).Value = $headers[$c-1]
         $ws.Cells.Item($tableHeaderRow,$c).Style.Font.Bold = $true
    }

    try { $ws.Cells[$tableHeaderRow,1,$tableHeaderRow,$headers.Count].AutoFilter = $true } catch {}
    try { $ws.View.FreezePanes($tableHeaderRow+1, 1) } catch {}

    if (-not $rowsToWrite -or $rowsToWrite.Count -eq 0) {
        $ws.Cells.Item($tableHeaderRow+1,1).Value = 'No deviations found'
        $ws.Cells.Item($tableHeaderRow+1,1).Style.Font.Italic = $true
         try {
             $r0 = $ws.Cells[1,1,($tableHeaderRow+1),$headers.Count]
             if (Get-Command Safe-AutoFitColumns -ErrorAction SilentlyContinue) {
                 Safe-AutoFitColumns -Ws $ws -Range $r0 -Context 'CSV Sammanfattning'
             } else {
                 $r0.AutoFitColumns() | Out-Null
             }
         } catch {}
         return $ws
     }

    $rOut = $tableHeaderRow + 1

    function _SvDeviation([string]$d) {
        $t = (($d + '')).Trim().ToUpperInvariant()
        switch ($t) {
            'OK' { return 'OK' }
            'FP' { return 'Falskt positiv' }
            'FN' { return 'Falskt negativ' }
            'ERROR' { return 'Fel' }
            'MISMATCH' { return 'Mismatch' }
            'UNKNOWN' { return 'Okänt' }
            default { return ($d + '') }
        }
    }

    function _SvSuffixCheck([string]$s) {
        $t = (($s + '')).Trim().ToUpperInvariant()
        switch ($t) {
            'OK' { return 'OK' }
            'BAD' { return 'FEL' }
            'MISSING' { return 'SAKNAS' }
            default { return ($s + '') }
        }
    }

$rowCount = $rowsToWrite.Count
$colCount = $headers.Count
$data = New-Object 'object[,]' $rowCount, $colCount

for ($i = 0; $i -lt $rowCount; $i++) {
    $r = $rowsToWrite[$i]

    $data[$i,0]  = ($r.SampleId + '')
    $data[$i,1]  = ($r.ErrorCode + '')
    $data[$i,2]  = (_SvDeviation ($r.Deviation + ''))
    $data[$i,3]  = ($r.RuleFlags + '')
    $data[$i,4]  = ($r.ExpectedSuffix + '')
    $data[$i,5]  = (_SvSuffixCheck ($r.SuffixCheck + ''))
    $data[$i,6]  = ($r.CartridgeSN + '')
    $data[$i,7]  = ($r.ModuleSN + '')
    $data[$i,8]  = ($r.StartTime + '')
    $data[$i,9]  = ($r.TestType + '')
    $data[$i,10] = ($r.ExpectedTestType + '')
    $data[$i,11] = ($r.ObservedWhy + '')
    $data[$i,12] = ($r.Status + '')
    $data[$i,13] = ($r.ErrorName + '')

    $rt = (($r.GeneratesRetest + '')).Trim().ToUpperInvariant()
    if ($rt -in @('YES','Y','TRUE','1')) { $data[$i,14] = 'Ja' }
    elseif ($rt) { $data[$i,14] = 'Nej' }
    else { $data[$i,14] = '' }

	    $data[$i,15] = $(if ($null -ne $r.MaxPressure) { $r.MaxPressure } else { '' })
    $data[$i,16] = ($r.TestResult + '')
    $data[$i,17] = ($r.ErrorText + '')
}

$startRow = $tableHeaderRow + 1
$endRow = $startRow + $rowCount - 1
$rng = $ws.Cells[$startRow, 1, $endRow, $colCount]
$rng.Value = $data


    try {
        $rAll = $ws.Cells[1,1,$endRow,$colCount]
        if (Get-Command Safe-AutoFitColumns -ErrorAction SilentlyContinue) {
            Safe-AutoFitColumns -Ws $ws -Range $rAll -Context 'CSV Sammanfattning'
        } else {
            $rAll.AutoFitColumns() | Out-Null
        }
    } catch {}
    return $ws
}