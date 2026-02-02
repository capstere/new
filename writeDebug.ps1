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
    $ws.View.ShowGridLines = $false

    # Column layout inspired by RuleEngine_Debug_FUTURE.xlsx (Swedish, fewer columns, more QC-friendly)
    $headers = @(
        'Sampe ID',
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

    # -------------------- Summary (always) --------------------
    $row = 1
    $ws.Cells.Item($row,1).Value = 'Sammanfattning felsökning'
    $ws.Cells.Item($row,1).Style.Font.Bold = $true
    $ws.Cells.Item($row,1).Style.Font.Size = 14
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

    # Row 2: Key metrics
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
                'FP' { 'Falskt Positiv (FP)' }
                'FN' { 'Falskt Negativ (FN)' }
                'ERROR' { 'Deviation ERROR (totalt)' }
                'MISMATCH' { 'Mismatch' }
                'UNKNOWN' { 'Okänt (UNKNOWN)' }
                default { $k }
            }
            $countValue = $sum.DeviationCounts[$k]
            _KV -r $row -k $label -v $countValue
            
            # Färgmarkera höga tal
            if ($countValue -gt 0 -and $k -in @('FP','FN','ERROR')) {
                $ws.Cells.Item($row,2).Style.Fill.PatternType = 'Solid'
                $ws.Cells.Item($row,2).Style.Fill.BackgroundColor.SetColor([System.Drawing.Color]::FromArgb(255,199,206))
                $ws.Cells.Item($row,2).Style.Font.Bold = $true
            }
            $row++
        }
    }

    # Split ERROR into categories
    if ($sum -and ($sum.MinorFunctionalError -ne $null -or $sum.InstrumentError -ne $null)) {
        _KV -r $row -k 'Minor Functional Error' -v $sum.MinorFunctionalError
        if ($sum.MinorFunctionalError -gt 0) {
            $ws.Cells.Item($row,2).Style.Fill.PatternType = 'Solid'
            $ws.Cells.Item($row,2).Style.Fill.BackgroundColor.SetColor([System.Drawing.Color]::FromArgb(255,235,156))
            $ws.Cells.Item($row,2).Style.Font.Bold = $true
        }
        $row++
        
        _KV -r $row -k 'Instrument Error' -v $sum.InstrumentError
        if ($sum.InstrumentError -gt 0) {
            $ws.Cells.Item($row,2).Style.Fill.PatternType = 'Solid'
            $ws.Cells.Item($row,2).Style.Fill.BackgroundColor.SetColor([System.Drawing.Color]::FromArgb(255,97,89))
            $ws.Cells.Item($row,2).Style.Font.Color.SetColor([System.Drawing.Color]::White)
            $ws.Cells.Item($row,2).Style.Font.Bold = $true
        }
        $row++
    }

    # Extra summary signals
    if ($sum -and $sum.DelamCount -ne $null -and $sum.DelamCount -gt 0) { 
        _KV -r $row -k 'Delamineringar (Sample ID)' -v $sum.DelamCount
        $ws.Cells.Item($row,2).Style.Fill.PatternType = 'Solid'
        $ws.Cells.Item($row,2).Style.Fill.BackgroundColor.SetColor([System.Drawing.Color]::FromArgb(255,235,156))
        $row++ 
    }
    
    if ($sum -and $sum.ReplacementCount -ne $null -and $sum.ReplacementCount -gt 0) { 
        _KV -r $row -k 'Ersättningar (A/AA/AAA)' -v $sum.ReplacementCount
        $ws.Cells.Item($row,2).Style.Fill.PatternType = 'Solid'
        $ws.Cells.Item($row,2).Style.Fill.BackgroundColor.SetColor([System.Drawing.Color]::FromArgb(217,210,233))
        $row++ 
    }

    foreach ($k in @('POS','NEG','ERROR','UNKNOWN')) {
        if ($sum.ObservedCounts.ContainsKey($k)) {
            _KV -r $row -k ('Observerat ' + $k) -v $sum.ObservedCounts[$k]; $row++
        }
    }

    _KV -r $row -k 'Omkörning (YES)' -v $sum.RetestYes
    if ($sum.RetestYes -gt 0) {
        $ws.Cells.Item($row,2).Style.Fill.PatternType = 'Solid'
        $ws.Cells.Item($row,2).Style.Fill.BackgroundColor.SetColor([System.Drawing.Color]::FromArgb(217,210,233))
        $ws.Cells.Item($row,2).Style.Font.Bold = $true
    }
    $row++

    if ($qc) {
        _KV -r $row -k 'Dubbletter av Sample ID' -v $qc.DuplicateSampleIdCount; $row++
        _KV -r $row -k 'Dubbletter av Cartridge S/N' -v $qc.DuplicateCartridgeSnCount; $row++
        _KV -r $row -k 'Moduler med error (≥3 fel)' -v $qc.HotModuleCount; $row++

        if ($qc.DistinctAssays.Count -gt 1) { _KV -r $row -k 'Varning: flera assay' -v ($qc.DistinctAssays -join ', '); $row++ }
        if ($qc.DistinctAssayVersions.Count -gt 1) { _KV -r $row -k 'Varning: flera versioner' -v ($qc.DistinctAssayVersions -join ', '); $row++ }
        if ($qc.DistinctReagentLots.Count -gt 1) { _KV -r $row -k 'Varning: flera reagent lots' -v ($qc.DistinctReagentLots -join ', '); $row++ }
    }

    # Pressure >= 90 summary
    $pressureGE90 = @($allRows | Where-Object {
        $p = $null
        try { $p = [double]$_.MaxPressure } catch { $p = $null }
        return ($null -ne $p -and $p -ge 90)
    }).Count
    _KV -r $row -k 'Max Pressure ≥ 90 PSI' -v $pressureGE90
    if ($pressureGE90 -gt 0) {
        $ws.Cells.Item($row,2).Style.Fill.PatternType = 'Solid'
        $ws.Cells.Item($row,2).Style.Fill.BackgroundColor.SetColor([System.Drawing.Color]::FromArgb(255,235,156))
        $ws.Cells.Item($row,2).Style.Font.Bold = $true
    }
    $row++

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
        $ws.Cells.Item($tableHeaderRow,$c).Style.WrapText = $true
        $ws.Cells.Item($tableHeaderRow,$c).Style.HorizontalAlignment = 'Center'
        $ws.Cells.Item($tableHeaderRow,$c).Style.VerticalAlignment = 'Center'
        $ws.Cells.Item($tableHeaderRow,$c).Style.Fill.PatternType = 'Solid'
        $ws.Cells.Item($tableHeaderRow,$c).Style.Fill.BackgroundColor.SetColor([System.Drawing.Color]::FromArgb(31,78,121))
        $ws.Cells.Item($tableHeaderRow,$c).Style.Font.Color.SetColor([System.Drawing.Color]::White)
    }

    try { $ws.Cells[$tableHeaderRow,1,$tableHeaderRow,$headers.Count].AutoFilter = $true } catch {}
    try { $ws.View.FreezePanes($tableHeaderRow+1, 1) } catch {}

    # If filtering yields nothing: keep header + deterministic message row
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

    # -------------------- Write rows (bulk data) --------------------
    $rOut = $tableHeaderRow + 1
    $startRow = $rOut
    $endRow = $startRow + $rowsToWrite.Count - 1

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

    function _SvErrorName([string]$e) {
        # Translate common error names for better color mapping
        $t = (($e + '')).Trim().ToLowerInvariant()
        switch -Regex ($t) {
            'delaminat' { return 'Delamination' }
            'pressure' { return 'Pressure' }
            'temperature' { return 'Temperature' }
            'cartridge' { return 'Cartridge' }
            'module' { return 'Module' }
            'reagent' { return 'Reagent' }
            default { return ($e + '') }
        }
    }

    # Bulk write for performance (EPPlus: assign object[,] to a range)
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

    $rng = $ws.Cells[$startRow, 1, $endRow, $colCount]
    $rng.Value = $data

    # -------------------- Styling + Conditional Formatting --------------------
    try {
        # Create unique table name to avoid collisions
        $tblName = 'CsvSummaryTable_' + ([Guid]::NewGuid().ToString('N').Substring(0,6))
        $tbl = $ws.Tables.Add($rng, $tblName)
        $tbl.ShowHeader = $false
        $tbl.ShowFilter = $false
        $tbl.TableStyle = [OfficeOpenXml.Table.TableStyles]::Medium6
    } catch { }

    # Dynamically calculate last column letter (robust for future column additions)
    $lastColLetter = [OfficeOpenXml.ExcelCellAddress]::GetColumnLetter($colCount)
    $rowRangeAddr = ("A{0}:{1}{2}" -f $startRow, $lastColLetter, $endRow)

    # Column ranges for targeted conditional formatting
    $deviationColAddr = ("C{0}:C{1}" -f $startRow, $endRow)
    $errCodeColAddr   = ("B{0}:B{1}" -f $startRow, $endRow)
    $sfxColAddr       = ("F{0}:F{1}" -f $startRow, $endRow)
    $errTypeColAddr   = ("N{0}:N{1}" -f $startRow, $endRow)
    $retestAddr       = ("O{0}:O{1}" -f $startRow, $endRow)
    $pressAddr        = ("P{0}:P{1}" -f $startRow, $endRow)

    # Helper function for conditional formatting (robust EPPlus usage)
    function _AddCfExpr {
        param(
            [string]$addr,
            [string]$formula,
            [System.Drawing.Color]$bg,
            [System.Drawing.Color]$fg = $null,
            [bool]$bold = $false
        )
        try {
            $cf = $ws.ConditionalFormatting.AddExpression($ws.Cells[$addr])
            $cf.Formula = $formula
            $cf.Style.Fill.PatternType = [OfficeOpenXml.Style.ExcelFillStyle]::Solid
            $cf.Style.Fill.BackgroundColor.SetColor($bg)
            if ($fg) { $cf.Style.Font.Color.SetColor($fg) }
            if ($bold) { $cf.Style.Font.Bold = $true }
        } catch { }
    }

    # ===== CONDITIONAL FORMATTING RULES (Prioritized) =====

    # 1. Whole-row styling by Avvikelse (Column C) — Highest priority
    _AddCfExpr -addr $rowRangeAddr -formula ('$C{0}="Fel"' -f $startRow) `
        -bg ([System.Drawing.Color]::FromArgb(255,97,89)) `
        -fg ([System.Drawing.Color]::White) -bold $true

    _AddCfExpr -addr $rowRangeAddr -formula ('$C{0}="Falskt positiv"' -f $startRow) `
        -bg ([System.Drawing.Color]::FromArgb(255,199,206)) `
        -fg ([System.Drawing.Color]::FromArgb(156,0,6))

    _AddCfExpr -addr $rowRangeAddr -formula ('$C{0}="Falskt negativ"' -f $startRow) `
        -bg ([System.Drawing.Color]::FromArgb(255,199,206)) `
        -fg ([System.Drawing.Color]::FromArgb(156,0,6))

    _AddCfExpr -addr $rowRangeAddr -formula ('$C{0}="Mismatch"' -f $startRow) `
        -bg ([System.Drawing.Color]::FromArgb(255,242,204))

    _AddCfExpr -addr $rowRangeAddr -formula ('$C{0}="Okänt"' -f $startRow) `
        -bg ([System.Drawing.Color]::FromArgb(242,242,242))

    # 2. Error Code column (B) — Non-blank = visible error
    _AddCfExpr -addr $errCodeColAddr -formula ('LEN(TRIM($B{0}))>0' -f $startRow) `
        -bg ([System.Drawing.Color]::FromArgb(255,199,206)) `
        -fg ([System.Drawing.Color]::FromArgb(156,0,6)) -bold $true

    # 3. Suffix-kontroll (F)
    _AddCfExpr -addr $sfxColAddr -formula ('$F{0}="FEL"' -f $startRow) `
        -bg ([System.Drawing.Color]::FromArgb(255,97,89)) `
        -fg ([System.Drawing.Color]::White) -bold $true

    _AddCfExpr -addr $sfxColAddr -formula ('$F{0}="SAKNAS"' -f $startRow) `
        -bg ([System.Drawing.Color]::FromArgb(255,242,204)) `
        -fg ([System.Drawing.Color]::FromArgb(156,101,0)) -bold $true

    # 4. Error Type (N) — categorized by name
    _AddCfExpr -addr $errTypeColAddr -formula ('SEARCH("Delaminat",$N{0})' -f $startRow) `
        -bg ([System.Drawing.Color]::FromArgb(217,210,233)) `
        -fg ([System.Drawing.Color]::FromArgb(62,37,92)) -bold $true

    _AddCfExpr -addr $errTypeColAddr -formula ('SEARCH("Pressure",$N{0})' -f $startRow) `
        -bg ([System.Drawing.Color]::FromArgb(255,235,156)) `
        -fg ([System.Drawing.Color]::FromArgb(156,101,0)) -bold $true

    _AddCfExpr -addr $errTypeColAddr -formula ('SEARCH("Temperature",$N{0})' -f $startRow) `
        -bg ([System.Drawing.Color]::FromArgb(255,235,156)) `
        -fg ([System.Drawing.Color]::FromArgb(156,101,0)) -bold $true

    # 5. Retest (O) — "Ja" = critical action needed
    _AddCfExpr -addr $retestAddr -formula ('$O{0}="Ja"' -f $startRow) `
        -bg ([System.Drawing.Color]::FromArgb(217,210,233)) `
        -fg ([System.Drawing.Color]::FromArgb(62,37,92)) -bold $true

    # 6. Max Pressure >= 90 (P) — Instrument stress signal
    _AddCfExpr -addr $pressAddr -formula ('AND(ISNUMBER($P{0}),$P{0}>=90)' -f $startRow) `
        -bg ([System.Drawing.Color]::FromArgb(255,235,156)) `
        -fg ([System.Drawing.Color]::FromArgb(156,101,0)) -bold $true

    # ===== BORDER STYLING =====
    try {
        $rng.Style.Border.Top.Style    = [OfficeOpenXml.Style.ExcelBorderStyle]::Thin
        $rng.Style.Border.Bottom.Style = [OfficeOpenXml.Style.ExcelBorderStyle]::Thin
        $rng.Style.Border.Left.Style   = [OfficeOpenXml.Style.ExcelBorderStyle]::Thin
        $rng.Style.Border.Right.Style  = [OfficeOpenXml.Style.ExcelBorderStyle]::Thin
    } catch { }

    # ===== AUTO-FIT COLUMNS =====
    try {
        $rAll = $ws.Cells[1,1,$endRow,$colCount]
        if (Get-Command Safe-AutoFitColumns -ErrorAction SilentlyContinue) {
            Safe-AutoFitColumns -Ws $ws -Range $rAll -Context 'CSV Sammanfattning'
        } else {
            $rAll.AutoFitColumns() | Out-Null
        }
    } catch { }

    return $ws
}