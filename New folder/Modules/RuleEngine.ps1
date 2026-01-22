# RuleEngine.ps1 (Windows PowerShell 5.1)
# Shadow-mode RuleEngine + RuleBank (CSV-driven)
# - Does NOT change legacy report logic unless explicitly wired
# - Produces per-row: ObservedCall, ExpectedCall, Deviation, ErrorCode/RetestPolicy
# - Writes RuleEngine_Debug worksheet (shadow)

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

function Load-RuleBank {
    param([Parameter(Mandatory)][string]$RuleBankDir)

    $rb = [ordered]@{
        Dir = $RuleBankDir
        ResultCallPatterns = @()
        SampleExpectationRules = @()
        ErrorCodes = @()
        MissingSamplesConfig = @()
        SampleIdMarkers = @()
        ParityCheckConfig = @()
        SampleNumberRules = @()
    }

    if (-not (Test-Path -LiteralPath $RuleBankDir)) { return [pscustomobject]$rb }

    $map = @{
        ResultCallPatterns      = '01_ResultCallPatterns.csv'
        SampleExpectationRules  = '02_SampleExpectationRules.csv'
        ErrorCodes              = '03_ErrorCodes.csv'
        MissingSamplesConfig    = '04_MissingSamplesConfig.csv'
        SampleIdMarkers         = '05_SampleIdMarkers.csv'
        ParityCheckConfig       = '06_ParityCheckConfig.csv'
        SampleNumberRules       = '07_SampleNumberRules.csv'
    }

    foreach ($k in $map.Keys) {
        $p = Join-Path $RuleBankDir $map[$k]
        $rb[$k] = Import-RuleCsv -Path $p
    }

    # Priority sort where applicable
    try { $rb.ResultCallPatterns = @($rb.ResultCallPatterns | Sort-Object { [int]($_.Priority) } -Descending) } catch {}
    try { $rb.SampleExpectationRules = @($rb.SampleExpectationRules | Sort-Object { [int]($_.Priority) } -Descending) } catch {}
    try { $rb.ParityCheckConfig = @($rb.ParityCheckConfig | Sort-Object { [int]($_.Priority) } -Descending) } catch {}
    try { $rb.SampleIdMarkers = @($rb.SampleIdMarkers | Sort-Object { [int]($_.Priority) } -Descending) } catch {}
    try { $rb.SampleNumberRules = @($rb.SampleNumberRules | Sort-Object { [int]($_.Priority) } -Descending) } catch {}

    return [pscustomobject]$rb
}

function Get-RowField {
    param(
        [Parameter(Mandatory)][object]$Row,
        [Parameter(Mandatory)][string]$FieldName
    )
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
    <#
      Supports:
      - empty / '*' => match all
      - wildcard patterns (contains '*' or '?') via -like
      - otherwise case-insensitive equals
    #>
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
        [Parameter(Mandatory)][object[]]$Patterns
    )

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

    # MTB-family override: treat MTB DETECTED/TRACE DETECTED as POS even if RIF contains NOT DETECTED/INDETERMINATE.
    # Prevent classic false INDET/NEG due to substring collisions.
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

    foreach ($r in $Patterns) {
        if (-not (Test-RuleEnabled $r)) { continue }
        $ruleAssay = (Get-RowField -Row $r -FieldName 'Assay')
        if (-not (Test-AssayMatch -RuleAssay $ruleAssay -RowAssay $assay)) { continue }

        $pat  = (Get-RowField -Row $r -FieldName 'Pattern')
        if (-not ($pat + '').Trim()) { continue }
        $mt   = (Get-RowField -Row $r -FieldName 'MatchType')

        if (Match-Text -Text $tr -Pattern $pat -MatchType $mt) {
            $call = ((Get-RowField -Row $r -FieldName 'Call') + '').Trim().ToUpperInvariant()
            if ($call) {
                $note = (Get-RowField -Row $r -FieldName 'Note')
                $why = if (($note + '').Trim()) { $note } else { ("Matched " + $mt + ": " + $pat) }
                return [pscustomobject]@{ Call=$call; Reason=$why }
            }
        }
    }

    return [pscustomobject]@{ Call='UNKNOWN'; Reason='No pattern matched' }
}

function Get-ExpectedCallDetailed {
    param(
        [Parameter(Mandatory)][object]$Row,
        [Parameter(Mandatory)][object[]]$Rules
    )

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

        if (Match-Text -Text $sid -Pattern $pat -MatchType $mtype) {
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

    # Supports:
    # - numeric codes (4-5 digits)
    # - multiple defaults via '####' (we choose based on content when possible)
    # - "blank code" rows (used for derived flags like pressure) exposed as NamedBlanks
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
            # Choose default:
            # 1) if error text indicates delamination and a "Delamination" default exists -> use that
            # 2) else use the last default (typically "Other error codes")
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
    if ($pressure -ne $null -and $pressure -gt 90) { $pressureFlag = $true }

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
    <#
      05_SampleIdMarkers.csv schema:
        AssayPattern,MarkerType,Marker,SampleTokenIndex,Enabled,Note
      Returns the first enabled marker matching assay pattern + marker type (by RuleBank ordering / priority).
    #>
    param(
        [Parameter(Mandatory)][pscustomobject]$RuleBank,
        [Parameter(Mandatory)][string]$Assay,
        [Parameter(Mandatory)][string]$MarkerType
    )

    foreach ($r in $RuleBank.SampleIdMarkers) {
        if (-not (Test-RuleEnabled $r)) { continue }

        $ap = ((Get-RowField -Row $r -FieldName 'AssayPattern') + '').Trim()
        if (-not (Test-AssayMatch -RuleAssay $ap -RowAssay $Assay)) { continue }

        $mt = ((Get-RowField -Row $r -FieldName 'MarkerType') + '').Trim()
        if (-not $mt) { continue }
        if ($mt -ine $MarkerType) { continue }

        $m = ((Get-RowField -Row $r -FieldName 'Marker') + '').Trim()
        return $m
    }

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
    <#
      06_ParityCheckConfig.csv schema:
        AssayPattern,Enabled,CartridgeField,SampleTokenIndex,SuffixX,SuffixPlus,DelaminationMarkerType,MinValidCartridgeSNPercent,Note,Priority
    #>
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

    # Pull defaults from markers (assay-aware)
    $delam = Get-MarkerValue -RuleBank $RuleBank -Assay $Assay -MarkerType 'DelaminationCodeRegex'
    if ($delam) { $cfg.DelamRegex = $delam }

    $suffix = Get-MarkerValue -RuleBank $RuleBank -Assay $Assay -MarkerType 'SuffixChars'
    if ($suffix) {
        # Some CSV exports double-escape backslashes (e.g. X|\+). Normalize "\" -> "\" for regex use.
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

    # Refresh delam regex by configured marker type
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
        # Legacy: token2 is commonly sample type code
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

function Get-SampleNumberRuleForRow {
    <#
      07_SampleNumberRules.csv schema:
        AssayPattern,SampleTypeCode,SampleNumberTokenIndex,SampleNumberRegex,SampleNumberMin,SampleNumberMax,SampleNumberPad,Enabled,Note,Priority
    #>
    param(
        [Parameter(Mandatory)][string]$Assay,
        [Parameter(Mandatory)][string]$ControlCode,
        [Parameter(Mandatory)][object[]]$Rules
    )

    foreach ($r in $Rules) {
        if (-not (Test-RuleEnabled $r)) { continue }

        $ap = ((Get-RowField -Row $r -FieldName 'AssayPattern') + '').Trim()
        if (-not (Test-AssayMatch -RuleAssay $ap -RowAssay $Assay)) { continue }

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

    # Convert field-arrays (from Import-CsvRows) to objects using header from file
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

    # Group by assay: parity + suffix expectations are assay-configurable.
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
        $delamPattern = $parCfg.DelamRegex
        $validSuffix = $parCfg.ValidSuffixRegex

        # Pre-scan for parity + majority suffix
        $numeric = New-Object System.Collections.Generic.List[long]
        $suffixCounts = @{}
        $suffixCounts[$parCfg.XChar.ToUpperInvariant()] = 0
        $suffixCounts[$parCfg.PlusChar] = 0

        foreach ($row in $group) {
            $sn = (Get-RowField -Row $row -FieldName $parCfg.CartridgeField)
            if (($sn + '').Trim() -match '^\d+$') {
                try { $numeric.Add([long]($sn + '')) } catch {}
            }

            $sid = (Get-RowField -Row $row -FieldName 'Sample ID')
            if (($sid + '').Trim()) {
                $t = Get-SampleTokenAndBase -SampleId ($sid + '') -TokenIndex $parCfg.TokenIndex -DelamPattern $delamPattern -ValidSuffixRegex $validSuffix -XChar $parCfg.XChar -PlusChar $parCfg.PlusChar
                if ($t.ActualSuffix -and $suffixCounts.ContainsKey($t.ActualSuffix)) { $suffixCounts[$t.ActualSuffix]++ }
            }
        }

        $numRatio = 0.0
        try { $numRatio = [double]$numeric.Count / [double]$group.Count } catch {}

        $useParity = ($parCfg.UseParity -and $numeric.Count -gt 0 -and $numRatio -ge $parCfg.NumericRatioThreshold)

        $minSn = $null
        $parityForX = $null
        if ($useParity) {
            try {
                $minSn = ($numeric | Measure-Object -Minimum).Minimum
                $parityForX = [int]($minSn % 2)
            } catch {
                $useParity = $false
            }
        }

        $majSuffix = ''
        if (-not $useParity) {
            $xCount = $suffixCounts[$parCfg.XChar.ToUpperInvariant()]
            $pCount = $suffixCounts[$parCfg.PlusChar]
            if ($xCount -gt $pCount) { $majSuffix = $parCfg.XChar.ToUpperInvariant() }
            elseif ($pCount -gt $xCount) { $majSuffix = $parCfg.PlusChar }
        }

        foreach ($row in $group) {
            try {
                $obsD = Get-ObservedCallDetailed -Row $row -Patterns $RuleBank.ResultCallPatterns
                $expD = Get-ExpectedCallDetailed -Row $row -Rules $RuleBank.SampleExpectationRules

                $sid = (Get-RowField -Row $row -FieldName 'Sample ID')
                $assay = (Get-RowField -Row $row -FieldName 'Assay')

                $expTT = ''
                if (($sid + '').Trim()) { $expTT = Get-ExpectedTestTypeDerived -SampleId ($sid + '') }

                # ExpectedCall fallback: derive from Test Type / control code when no explicit rule matched
                $expCall = ($expD.Call + '').Trim().ToUpperInvariant()
                $expSrc = 'RULE'
                if (-not $expCall) {
                    $tt = (Get-RowField -Row $row -FieldName 'Test Type')
                    $tt2 = ($tt + '')
                    if ($tt2 -match '(?i)Negative\s+Control') { $expCall = 'NEG'; $expSrc = 'TESTTYPE' }
                    elseif ($tt2 -match '(?i)Positive\s+Control') { $expCall = 'POS'; $expSrc = 'TESTTYPE' }
                    else {
                        $cc2 = Get-ControlCodeFromRow -Row $row -SampleTypeCodeTokenIndex $parCfg.SampleTypeCodeTokenIndex
                        if ($cc2 -match '^\d+$') {
                            if ([int]$cc2 -eq 0) { $expCall = 'NEG'; $expSrc = 'CONTROL_CODE' }
                            elseif ([int]$cc2 -ge 1) { $expCall = 'POS'; $expSrc = 'CONTROL_CODE' }
                        }
                    }
                }

                $errInfo = Get-ErrorInfo -Row $row -ErrorLut $errLut -DelamPattern $delamPattern
                $dev = Classify-Deviation -Expected $expCall -Observed $obsD.Call

                # Suffix / parity validation (naming robustness)
                $tokInfo = [pscustomobject]@{ SampleToken=''; BaseToken=''; ActualSuffix=''; SampleNum=''; SampleNumRaw='' }
                if (($sid + '').Trim()) {
                    $tokInfo = Get-SampleTokenAndBase -SampleId ($sid + '') -TokenIndex $parCfg.TokenIndex -DelamPattern $delamPattern -ValidSuffixRegex $validSuffix -XChar $parCfg.XChar -PlusChar $parCfg.PlusChar
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

                # Sample number validation (config-driven)
                $sampleNum = ''
                $sampleNumRaw = ''
                $sampleNumOk = ''
                $sampleNumWhy = ''

                $cc = Get-ControlCodeFromRow -Row $row -SampleTypeCodeTokenIndex $parCfg.SampleTypeCodeTokenIndex

                $rule = $null
                try { $rule = Get-SampleNumberRuleForRow -Assay $assay -ControlCode $cc -Rules $RuleBank.SampleNumberRules } catch {}

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



    # --- QC post-processing (file-level validations + per-row flags) ---
    function _Append-RuleFlag {
        param([pscustomobject]$row, [string]$flag)
        $f = (($row.RuleFlags + '')).Trim()
        if (-not $f) { $row.RuleFlags = $flag; return }
        $parts = $f.Split('|')
        if ($parts -contains $flag) { return }
        $row.RuleFlags = ($f + '|' + $flag)
    }

    # Distinct counts (file-level)
    $distinctAssays = @($results | ForEach-Object { ($_.Assay + '').Trim() } | Where-Object { $_ } | Sort-Object -Unique)
    $distinctAssayVersions = @($results | ForEach-Object { ($_.AssayVersion + '').Trim() } | Where-Object { $_ } | Sort-Object -Unique)
    $distinctReagentLots = @($results | ForEach-Object { ($_.ReagentLotId + '').Trim() } | Where-Object { $_ } | Sort-Object -Unique)

    # Duplicate Sample IDs
    $dupSample = @($results | Where-Object { (($_.SampleId + '').Trim()) } | Group-Object SampleId | Where-Object { $_.Count -gt 1 })
    if ($dupSample.Count -gt 0) {
        $dupSet = @{}
        foreach ($g in $dupSample) { $dupSet[$g.Name] = $true }
        foreach ($r in $results) {
            $sid = ((($r.SampleId + '')).Trim())
            if ($sid -and $dupSet.ContainsKey($sid)) { _Append-RuleFlag -row $r -flag 'DQ_DUP_SAMPLEID' }
        }
    }

    # Duplicate Cartridge S/N
    $dupCart = @($results | Where-Object { (($_.CartridgeSN + '').Trim()) } | Group-Object CartridgeSN | Where-Object { $_.Count -gt 1 })
    if ($dupCart.Count -gt 0) {
        $dupSet = @{}
        foreach ($g in $dupCart) { $dupSet[$g.Name] = $true }
        foreach ($r in $results) {
            $csn = ((($r.CartridgeSN + '')).Trim())
            if ($csn -and $dupSet.ContainsKey($csn)) { _Append-RuleFlag -row $r -flag 'DQ_DUP_CARTSN' }
        }
    }

    # Test Type mismatch (derived expected vs actual)
    foreach ($r in $results) {
        $act = ((($r.TestType + '')).Trim())
        $exp = ((($r.ExpectedTestType + '')).Trim())
        if ($act -and $exp -and ($act.ToUpperInvariant() -ne $exp.ToUpperInvariant())) {
            _Append-RuleFlag -row $r -flag 'TESTTYPE_MISMATCH'
        }
    }

    # Suffix/sample number quick flags (so exception-only debug includes these)
    foreach ($r in $results) {
        $sc = ((($r.SuffixCheck + '')).Trim().ToUpperInvariant())
        if ($sc -and $sc -ne 'OK') { _Append-RuleFlag -row $r -flag ('SUFFIX_' + $sc) }
        $snok = ((($r.SampleNumOK + '')).Trim().ToUpperInvariant())
        if ($snok -eq 'NO') { _Append-RuleFlag -row $r -flag 'SAMPLENUM_BAD' }
    }

    # Module hotspot: >=3 ERROR rows on same Module S/N
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

    # Attach QC stats to result (used by Debug sheet summary)
    $qc = [pscustomobject]@{
        DistinctAssays = $distinctAssays
        DistinctAssayVersions = $distinctAssayVersions
        DistinctReagentLots = $distinctReagentLots
        DuplicateSampleIdCount = ($dupSample | ForEach-Object { $_.Name } | Select-Object -Unique).Count
        DuplicateCartridgeSnCount = ($dupCart | ForEach-Object { $_.Name } | Select-Object -Unique).Count
        HotModuleCount = $hotModules.Count
    }
    # ---------------------------------------------------------------
    $summary = [pscustomobject]@{
        Total = $results.Count
        ObservedCounts = @{}
        DeviationCounts = @{}
        RetestYes = 0
    }

    foreach ($r in $results) {
        if (-not $summary.ObservedCounts.ContainsKey($r.ObservedCall)) { $summary.ObservedCounts[$r.ObservedCall] = 0 }
        $summary.ObservedCounts[$r.ObservedCall]++

        if (-not $summary.DeviationCounts.ContainsKey($r.Deviation)) { $summary.DeviationCounts[$r.Deviation] = 0 }
        $summary.DeviationCounts[$r.Deviation]++

        $rt = ($r.GeneratesRetest + '').Trim().ToUpperInvariant()
        if ($rt -in @('YES','Y','TRUE','1')) { $summary.RetestYes++ }
    }

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
        $old = $Pkg.Workbook.Worksheets['RuleEngine_Debug']
        if ($old) { $Pkg.Workbook.Worksheets.Delete($old) }
    } catch {}

    $ws = $Pkg.Workbook.Worksheets.Add('RuleEngine_Debug')

    # Column layout inspired by RuleEngine_Debug_FUTURE.xlsx (Swedish, fewer columns, more QC-friendly)
    $headers = @(
        'Prov-ID',
        'Felkod',
        'Avvikelse',
        'Flaggor',
        'Förväntat suffix',
        'Suffixkontroll',
        'Kassett S/N',
        'Modul S/N',
        'Starttid',
        'Testtyp',
        'Förväntad testtyp',
        'Observerat (varför)',
        'Status',
        'Felnamn',
        'Omkörning',
        'Maxtryck (PSI)',
        'Testresultat',
        'Feltext'
    )

    # -------------------- Summary (always) --------------------
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

    # Row 2: multiple key/value pairs on one line
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
                'ERROR' { 'Instrumentfel (ERROR)' }
                'MISMATCH' { 'Mismatch' }
                'UNKNOWN' { 'Okänt (UNKNOWN)' }
                default { $k }
            }
            _KV -r $row -k $label -v $sum.DeviationCounts[$k]; $row++
        }
    }

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

    # Pressure >= 90 summary (derived from rows to stay source-of-truth even if pressure rule changes)
    $pressureGE90 = @($allRows | Where-Object {
        $p = $null
        try { $p = [double]$_.MaxPressure } catch { $p = $null }
        return ($null -ne $p -and $p -ge 90)
    }).Count
    _KV -r $row -k 'Tryck ≥ 90 PSI' -v $pressureGE90; $row++

    # Leave one blank row before table
    $row++
    $tableHeaderRow = $row

    # -------------------- Exception-only filter --------------------
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

    # -------------------- Table header --------------------
    for ($c=1; $c -le $headers.Count; $c++) {
         $ws.Cells.Item($tableHeaderRow,$c).Value = $headers[$c-1]
         $ws.Cells.Item($tableHeaderRow,$c).Style.Font.Bold = $true
    }

    try { $ws.Cells[$tableHeaderRow,1,$tableHeaderRow,$headers.Count].AutoFilter = $true } catch {}
    try { $ws.View.FreezePanes($tableHeaderRow+1, 1) } catch {}

    # If filtering yields nothing: keep header + deterministic message row
    if (-not $rowsToWrite -or $rowsToWrite.Count -eq 0) {
        $ws.Cells.Item($tableHeaderRow+1,1).Value = 'No deviations found'
        $ws.Cells.Item($tableHeaderRow+1,1).Style.Font.Italic = $true
         try { if ($ws.Dimension) { $ws.Cells[$ws.Dimension.Address].AutoFitColumns() } } catch {}
         return $ws
     }

    # -------------------- Write rows --------------------
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

    foreach ($r in $rowsToWrite) {
        $ws.Cells.Item($rOut,1).Value  = ($r.SampleId + '')
        $ws.Cells.Item($rOut,2).Value  = ($r.ErrorCode + '')
        $ws.Cells.Item($rOut,3).Value  = (_SvDeviation ($r.Deviation + ''))
        $ws.Cells.Item($rOut,4).Value  = ($r.RuleFlags + '')
        $ws.Cells.Item($rOut,5).Value  = ($r.ExpectedSuffix + '')
        $ws.Cells.Item($rOut,6).Value  = (_SvSuffixCheck ($r.SuffixCheck + ''))
        $ws.Cells.Item($rOut,7).Value  = ($r.CartridgeSN + '')
        $ws.Cells.Item($rOut,8).Value  = ($r.ModuleSN + '')
        $ws.Cells.Item($rOut,9).Value  = ($r.StartTime + '')
        $ws.Cells.Item($rOut,10).Value = ($r.TestType + '')
        $ws.Cells.Item($rOut,11).Value = ($r.ExpectedTestType + '')
        $ws.Cells.Item($rOut,12).Value = ($r.ObservedWhy + '')
        $ws.Cells.Item($rOut,13).Value = ($r.Status + '')
        $ws.Cells.Item($rOut,14).Value = ($r.ErrorName + '')

        $rt = (($r.GeneratesRetest + '')).Trim().ToUpperInvariant()
        if ($rt -in @('YES','Y','TRUE','1')) { $ws.Cells.Item($rOut,15).Value = 'Ja' }
        elseif ($rt) { $ws.Cells.Item($rOut,15).Value = 'Nej' }
        else { $ws.Cells.Item($rOut,15).Value = '' }

        $ws.Cells.Item($rOut,16).Value = $(if ($null -ne $r.MaxPressure) { $r.MaxPressure } else { '' })
        $ws.Cells.Item($rOut,17).Value = ($r.TestResult + '')
        $ws.Cells.Item($rOut,18).Value = ($r.ErrorText + '')

        $rOut++
    }

    try { if ($ws.Dimension) { $ws.Cells[$ws.Dimension.Address].AutoFitColumns() } } catch {}
    return $ws
}