#requires -Version 5.1
<#
  Information2Writer.ps1  (EPPlus writer only)
  - Writes ONE sheet: "Information2"
  - Focus: human-friendly summary first, then compact tables.
  - Compatible with EPPlus 4.5.3.3.
#>

Set-StrictMode -Off

function _I2_Trunc {
    param([string]$Text, [int]$Max = 140)
    if ([string]::IsNullOrEmpty($Text)) { return '' }
    $t = ($Text + '').Replace("`r"," ").Replace("`n"," ")
    $t = ($t -replace '\s+',' ').Trim()
    if ($t.Length -le $Max) { return $t }
    return ($t.Substring(0,$Max-1) + '…')
}

function _I2_DisplayColumnName {
    param([string]$Key)
    $k = ($Key + '').Trim()
    switch ($k) {
        'Assay'       { return 'Assay' }
        'AssayVer'    { return 'Assay Version' }
        'ReagentLot'  { return 'Reagent Lot ID (LSP)' }
        'SampleId'    { return 'Sample ID' }
        'CartridgeSn' { return 'Cartridge S/N' }
        'ModuleSn'    { return 'Module S/N' }
        'TestType'    { return 'Test Type' }
        'Status'      { return 'Status' }
        'TestResult'  { return 'Test Result' }
        'Error'       { return 'Error' }
        'ErrorCode'   { return 'Error Code' }
        'MaxPressure' { return 'Max Pressure (PSI)' }
        'WorkCenter'  { return 'Work Center' }
        default       { return $k }
    }
}

function _I2_SetHeader {
    param(
        [object]$Ws,
        [string]$Range,
        [string]$Text,
        [int]$FontSize = 12
    )
    $Ws.Cells[$Range].Value = $Text
    $Ws.Cells[$Range].Style.Font.Bold = $true
    $Ws.Cells[$Range].Style.Font.Size = $FontSize
}

function _I2_StyleKey {
    param([object]$Cell)
    $Cell.Style.Font.Bold = $true
    $Cell.Style.HorizontalAlignment = 2 # Center
}

function _I2_ApplyStatusStyle {
    param([object]$Cell, [string]$Status)
    try {
        $Cell.Style.Font.Bold = $true
        $Cell.Style.Font.Size = 14
        $Cell.Style.Fill.PatternType = 'Solid'
        switch (($Status + '').ToUpperInvariant()) {
            'PASS' { $Cell.Style.Fill.BackgroundColor.Color = [System.Drawing.ColorTranslator]::FromHtml('#c6efce') }
            'WARN' { $Cell.Style.Fill.BackgroundColor.Color = [System.Drawing.ColorTranslator]::FromHtml('#ffe699') }
            default { $Cell.Style.Fill.BackgroundColor.Color = [System.Drawing.ColorTranslator]::FromHtml('#ffc7ce') }
        }
    } catch {}
}

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
        if (-not $Evaluation) {
            $Evaluation = [pscustomobject]@{
                OverallStatus          = 'FAIL'
                OverallSeverity        = 'Error'
                Findings               = @()
                AffectedTests          = @()
                AffectedTestsTruncated = 0
                ErrorSummary           = @()
                PressureStats          = [pscustomobject]@{ Max=$null; OverWarn=0; OverFail=0; WarnThreshold=$null; FailThreshold=$null }
                Debug                  = @()
                UniqueErrorCodes       = 0
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
                UniqueSampleIds      = 0
                UniqueCartridgeSN    = 0
                DuplicateCounts      = [ordered]@{ SampleId=0; CartridgeSN=0 }
                AssayCanonicalSource = ''
                ParseErrors          = @()
                ColumnMissing        = @()
            }
        }

        # Clean sheet (EPPlus 4.5.3.3 safe)
        try { $Worksheet.Cells.Clear() } catch {}
        try { $Worksheet.ConditionalFormatting.Clear() } catch {}
        try { $Worksheet.DataValidations.Clear() } catch {}
        try { $Worksheet.Tables.Clear() } catch {}

        $Worksheet.Cells.Style.Font.Name = 'Arial'
        $Worksheet.Cells.Style.Font.Size = 10

        $ps = if ($Evaluation.PressureStats) { $Evaluation.PressureStats } else { [pscustomobject]@{ Max=$null; OverWarn=0; OverFail=0; WarnThreshold=$null; FailThreshold=$null } }
        $sevCounts = if ($Evaluation.SeverityCounts) { $Evaluation.SeverityCounts } else { @{ Error=0; Warn=0; Info=0 } }


        # Performance toggle (Config.ps1)
        $enableCF = $true
        try {
            if ($Config -and $Config.Performance -and $Config.Performance.EnableConditionalFormatting -ne $null) {
                $enableCF = [bool]$Config.Performance.EnableConditionalFormatting
            }
        } catch { $enableCF = $true }

        $missingCols = @()
        try { if ($Context.ColumnMissing) { $missingCols = @($Context.ColumnMissing | Sort-Object) } } catch {}

        $csvLeaf = if ($CsvPath) { Split-Path $CsvPath -Leaf } else { '—' }
        $lots = ''
        try { $lots = ($Context.ReagentLotIds -join ', ') } catch { $lots = '' }
        $wcs = ''
        try { $wcs = ($Context.WorkCenters -join ', ') } catch { $wcs = '' }

        $dupTxt = ''
        try {
            $d1 = if ($Context.DuplicateCounts -and $Context.DuplicateCounts.SampleId -ne $null) { [int]$Context.DuplicateCounts.SampleId } else { 0 }
            $d2 = if ($Context.DuplicateCounts -and $Context.DuplicateCounts.CartridgeSN -ne $null) { [int]$Context.DuplicateCounts.CartridgeSN } else { 0 }
            $dupTxt = ("Sample ID={0}, Cartridge S/N={1}" -f $d1, $d2)
        } catch { $dupTxt = '' }

        # ---------------------------------------------------------
        # Header / at-a-glance block (compact, non-repeating)
        # ---------------------------------------------------------
        $r = 1
        $Worksheet.Cells["A$r"].Value = "Information2 - QC Summary"
        $Worksheet.Cells["A$r:H$r"].Merge = $true
        $Worksheet.Cells["A$r"].Style.Font.Bold = $true
        $Worksheet.Cells["A$r"].Style.Font.Size = 16
        $r++

        # Row 2: Overall status + key counters
        $Worksheet.Cells["A$r"].Value = "Overall"
        _I2_StyleKey -Cell $Worksheet.Cells["A$r"]

        $Worksheet.Cells["B$r"].Value = ($Evaluation.OverallStatus + '')
        _I2_ApplyStatusStyle -Cell $Worksheet.Cells["B$r"] -Status $Evaluation.OverallStatus

        $Worksheet.Cells["C$r"].Value = "Worst severity"
        _I2_StyleKey -Cell $Worksheet.Cells["C$r"]
        $Worksheet.Cells["D$r"].Value = ($Evaluation.OverallSeverity + '')

        $Worksheet.Cells["E$r"].Value = "Affected rows"
        _I2_StyleKey -Cell $Worksheet.Cells["E$r"]
        $Worksheet.Cells["F$r"].Value = [int]$Evaluation.ErrorRowCount

        $Worksheet.Cells["G$r"].Value = "Error codes"
        _I2_StyleKey -Cell $Worksheet.Cells["G$r"]
        $Worksheet.Cells["H$r"].Value = [int]$Evaluation.UniqueErrorCodes
        $r++

        # Row 3: Identity
        $Worksheet.Cells["A$r"].Value = "Assay"
        _I2_StyleKey -Cell $Worksheet.Cells["A$r"]
        $Worksheet.Cells["B$r"].Value = (if ($Context.PSObject.Properties.Match('AssayDisplayName').Count -gt 0 -and $Context.AssayDisplayName) { $Context.AssayDisplayName } else { $Context.AssayCanonical })

        $Worksheet.Cells["C$r"].Value = "Assay version"
        _I2_StyleKey -Cell $Worksheet.Cells["C$r"]
        $Worksheet.Cells["D$r"].Value = ($Context.AssayVersion + '')

        $Worksheet.Cells["E$r"].Value = "Lot(s)"
        _I2_StyleKey -Cell $Worksheet.Cells["E$r"]
        $Worksheet.Cells["F$r"].Value = (if ($Context.PSObject.Properties.Match('ReagentLotDisplay').Count -gt 0 -and $Context.ReagentLotDisplay) { $Context.ReagentLotDisplay } else { $lots })

        $Worksheet.Cells["G$r"].Value = "WorkCenter(s)"
        _I2_StyleKey -Cell $Worksheet.Cells["G$r"]
        $Worksheet.Cells["H$r"].Value = (if ($Context.PSObject.Properties.Match('WorkCentersDisplay').Count -gt 0 -and $Context.WorkCentersDisplay) { ($Context.WorkCentersDisplay -join ', ') } else { $wcs })
        $r++

        # Row 4: Run details + high-level signals
        $Worksheet.Cells["A$r"].Value = "CSV"
        _I2_StyleKey -Cell $Worksheet.Cells["A$r"]
        $Worksheet.Cells["B$r"].Value = $csvLeaf

        $Worksheet.Cells["C$r"].Value = "Generated"
        _I2_StyleKey -Cell $Worksheet.Cells["C$r"]
        $Worksheet.Cells["D$r"].Value = (Get-Date).ToString('yyyy-MM-dd HH:mm')

        $Worksheet.Cells["E$r"].Value = "Version"
        _I2_StyleKey -Cell $Worksheet.Cells["E$r"]
        $Worksheet.Cells["F$r"].Value = ($ScriptVersion + '')

        $Worksheet.Cells["G$r"].Value = "Match source"
        _I2_StyleKey -Cell $Worksheet.Cells["G$r"]
        $Worksheet.Cells["H$r"].Value = ($Context.AssayCanonicalSource + '')
        $r++

        # Row 5: Counts
        $Worksheet.Cells["A$r"].Value = "Row count"
        _I2_StyleKey -Cell $Worksheet.Cells["A$r"]
        $Worksheet.Cells["B$r"].Value = [int]$Context.TotalTests

        $Worksheet.Cells["C$r"].Value = "Unique SampleID"
        _I2_StyleKey -Cell $Worksheet.Cells["C$r"]
        $Worksheet.Cells["D$r"].Value = [int]$Context.UniqueSampleIds

        $Worksheet.Cells["E$r"].Value = "Unique CartridgeSN"
        _I2_StyleKey -Cell $Worksheet.Cells["E$r"]
        $Worksheet.Cells["F$r"].Value = [int]$Context.UniqueCartridgeSN

        $Worksheet.Cells["G$r"].Value = "Duplicates"
        _I2_StyleKey -Cell $Worksheet.Cells["G$r"]
        $Worksheet.Cells["H$r"].Value = $dupTxt
        $r++

        # Row 6: Pressure + baseline
        $Worksheet.Cells["A$r"].Value = "Max pressure (max)"
        _I2_StyleKey -Cell $Worksheet.Cells["A$r"]
        $Worksheet.Cells["B$r"].Value = if ($ps.Max -ne $null) { [Math]::Round([double]$ps.Max,2) } else { '—' }

        $Worksheet.Cells["C$r"].Value = ">= Warn / Fail"
        _I2_StyleKey -Cell $Worksheet.Cells["C$r"]
        $Worksheet.Cells["D$r"].Value = ("{0} / {1}" -f [int]$ps.OverWarn, [int]$ps.OverFail)

        $Worksheet.Cells["E$r"].Value = "Baseline (exp / delta)"
        _I2_StyleKey -Cell $Worksheet.Cells["E$r"]
        $Worksheet.Cells["F$r"].Value = ("{0} / {1}" -f [int]$Evaluation.BaselineExpected, [int]$Evaluation.BaselineDelta)

        $Worksheet.Cells["G$r"].Value = "Findings (E/W/I)"
        _I2_StyleKey -Cell $Worksheet.Cells["G$r"]
        $Worksheet.Cells["H$r"].Value = ("{0}/{1}/{2}" -f [int]$sevCounts.Error, [int]$sevCounts.Warn, [int]$sevCounts.Info)
        $r++

        # Row 7: Missing columns (only if any)
        $Worksheet.Cells["A$r"].Value = "Missing columns"
        _I2_StyleKey -Cell $Worksheet.Cells["A$r"]
        if ($missingCols.Count -gt 0) {
            $Worksheet.Cells["B$r:H$r"].Merge = $true
            $Worksheet.Cells["B$r"].Value = (($missingCols | ForEach-Object { _I2_DisplayColumnName $_ }) -join ', ')
        } else {
            $Worksheet.Cells["B$r"].Value = "—"
        }
        $r += 2

        # ---------------------------------------------------------
        # Findings (compact table)
        # ---------------------------------------------------------
        _I2_SetHeader -Ws $Worksheet -Range ("A{0}" -f $r) -Text "Findings (sorted by severity)" -FontSize 12
        $Worksheet.Cells["A$r:H$r"].Merge = $true
        $r++

        $Worksheet.Cells["A$r"].Value = "Severity"
        $Worksheet.Cells["B$r"].Value = "Title"
        $Worksheet.Cells["C$r"].Value = "Count"
        $Worksheet.Cells["D$r"].Value = "RuleId"
        $Worksheet.Cells["E$r"].Value = "Example"
        $Worksheet.Cells["F$r"].Value = "Evidence (short)"
        $Worksheet.Cells["A$r:F$r"].Style.Font.Bold = $true
        $Worksheet.Cells["A$r:F$r"].Style.WrapText = $true
        $r++

        $findStart = $r
        $allFind = @()
        try { $allFind = @($Evaluation.Findings) } catch { $allFind = @() }

        # Keep table readable: cap to 50 findings, still enough.
        $maxFind = 50
        $findShown = 0
        foreach ($f in $allFind) {
            if (-not $f) { continue }
            if ($findShown -ge $maxFind) { break }
            $Worksheet.Cells["A$r"].Value = ($f.Severity + '')
            $Worksheet.Cells["B$r"].Value = ($f.Title + '')
            $Worksheet.Cells["C$r"].Value = [int]$f.Count
            $Worksheet.Cells["D$r"].Value = ($f.RuleId + '')
            $Worksheet.Cells["E$r"].Value = ($f.Example + '')
            $Worksheet.Cells["F$r"].Value = _I2_Trunc ($f.Evidence + '') 160
            $r++
            $findShown++
        }
        $findEnd = $r - 1
        if ($findEnd -ge $findStart) {
            $Worksheet.Cells["A$findStart:F$findEnd"].AutoFilter = $true
            $Worksheet.Cells["B$findStart:F$findEnd"].Style.WrapText = $true

            # Conditional formatting: severity color in column A
            if ($enableCF) {
                try {
                $addr = "A$findStart:A$findEnd"

                $cfErr = $Worksheet.ConditionalFormatting.AddExpression($addr)
                $cfErr.Formula = 'ISNUMBER(SEARCH("Error",A1))'
                $cfErr.Style.Fill.PatternType = 'Solid'
                $cfErr.Style.Fill.BackgroundColor.Color = [System.Drawing.ColorTranslator]::FromHtml('#ffc7ce')

                $cfWarn = $Worksheet.ConditionalFormatting.AddExpression($addr)
                $cfWarn.Formula = 'ISNUMBER(SEARCH("Warn",A1))'
                $cfWarn.Style.Fill.PatternType = 'Solid'
                $cfWarn.Style.Fill.BackgroundColor.Color = [System.Drawing.ColorTranslator]::FromHtml('#ffe699')
                } catch {}
            }
        } else {
            $Worksheet.Cells["A$r"].Value = "—"
            $r++
        }
        $r += 2

        # ---------------------------------------------------------
        # Error codes (if present)
        # ---------------------------------------------------------
        _I2_SetHeader -Ws $Worksheet -Range ("A{0}" -f $r) -Text "Error codes (summary)" -FontSize 12
        $Worksheet.Cells["A$r:H$r"].Merge = $true
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
        $errShown = 0
        $maxErr = 60
        foreach ($err in @($Evaluation.ErrorSummary)) {
            if (-not $err) { continue }
            if ($errShown -ge $maxErr) { break }
            $Worksheet.Cells["A$r"].Value = ($err.ErrorCode + '')
            $Worksheet.Cells["B$r"].Value = ($err.Name + '')
            $Worksheet.Cells["C$r"].Value = ($err.Classification + '')
            $Worksheet.Cells["D$r"].Value = [int]$err.Count
            $Worksheet.Cells["E$r"].Value = ($err.ExampleSampleID + '')
            $r++
            $errShown++
        }
        $errEnd = $r - 1
        if ($errEnd -ge $errStart) {
            $Worksheet.Cells["A$errStart:E$errEnd"].AutoFilter = $true
            $Worksheet.Cells["A$errStart:E$errEnd"].Style.WrapText = $true
        } else {
            $Worksheet.Cells["A$r"].Value = "—"
            $Worksheet.Cells["A$r:E$r"].Merge = $true
            $r++
        }
        $r += 2

        # ---------------------------------------------------------
        # Affected tests (Warn/Error) - keep but cap for usability
        # ---------------------------------------------------------
        _I2_SetHeader -Ws $Worksheet -Range ("A{0}" -f $r) -Text "Affected tests (Warn / Error)" -FontSize 12
        $Worksheet.Cells["A$r:H$r"].Merge = $true
        $r++

        $Worksheet.Cells["A$r"].Value = "Severity"
        $Worksheet.Cells["B$r"].Value = "RuleId"
        $Worksheet.Cells["C$r"].Value = "Sample ID"
        $Worksheet.Cells["D$r"].Value = "Cartridge S/N"
        $Worksheet.Cells["E$r"].Value = "Module S/N"
        $Worksheet.Cells["F$r"].Value = "Status"
        $Worksheet.Cells["G$r"].Value = "MaxP (PSI)"
        $Worksheet.Cells["H$r"].Value = "ErrorCode"
        $Worksheet.Cells["A$r:H$r"].Style.Font.Bold = $true
        $Worksheet.Cells["A$r:H$r"].Style.WrapText = $true

        $affStart = ++$r
        $maxAff = 250
        $affShown = 0
        foreach ($row in @($Evaluation.AffectedTests)) {
            if (-not $row) { continue }
            if ($affShown -ge $maxAff) { break }
            $Worksheet.Cells["A$r"].Value = ($row.Severity + '')
            $Worksheet.Cells["B$r"].Value = ($row.PrimaryRule + '')
            $Worksheet.Cells["C$r"].Value = ($row.SampleID + '')
            $Worksheet.Cells["D$r"].Value = ($row.CartridgeSN + '')
            $Worksheet.Cells["E$r"].Value = ($row.ModuleSN + '')
            $Worksheet.Cells["F$r"].Value = ($row.Status + '')
            $Worksheet.Cells["G$r"].Value = if ($row.MaxPressure -ne $null) { [Math]::Round([double]$row.MaxPressure,2) } else { $null }
            $Worksheet.Cells["H$r"].Value = ($row.ErrorCode + '')
            $r++
            $affShown++
        }
        $affEnd = $r - 1

        if ($affShown -lt @($Evaluation.AffectedTests).Count) {
            $Evaluation.AffectedTestsTruncated = @($Evaluation.AffectedTests).Count - $affShown
        }

        if ($affEnd -ge $affStart) {
            $Worksheet.Cells["A$affStart:H$affEnd"].AutoFilter = $true
            $Worksheet.View.FreezePanes($affStart, 1)

            # CF: Severity + Status + Pressure thresholds via expressions (robust)
            if ($enableCF) {
                try {
                    $addrA = "A$affStart:A$affEnd"
                $addrF = "F$affStart:F$affEnd"
                $addrG = "G$affStart:G$affEnd"

                $cfAffErr = $Worksheet.ConditionalFormatting.AddExpression($addrA)
                $cfAffErr.Formula = 'ISNUMBER(SEARCH("Error",A1))'
                $cfAffErr.Style.Fill.PatternType = 'Solid'
                $cfAffErr.Style.Fill.BackgroundColor.Color = [System.Drawing.ColorTranslator]::FromHtml('#ffc7ce')

                $cfAffWarn = $Worksheet.ConditionalFormatting.AddExpression($addrA)
                $cfAffWarn.Formula = 'ISNUMBER(SEARCH("Warn",A1))'
                $cfAffWarn.Style.Fill.PatternType = 'Solid'
                $cfAffWarn.Style.Fill.BackgroundColor.Color = [System.Drawing.ColorTranslator]::FromHtml('#ffe699')

                $cfDone = $Worksheet.ConditionalFormatting.AddExpression($addrF)
                $cfDone.Formula = 'ISNUMBER(SEARCH("Done",F1))'
                $cfDone.Style.Fill.PatternType = 'Solid'
                $cfDone.Style.Fill.BackgroundColor.Color = [System.Drawing.ColorTranslator]::FromHtml('#c6efce')

                # Pressure thresholds
                if ($ps.FailThreshold -ne $null) {
                    $cfMpFail = $Worksheet.ConditionalFormatting.AddExpression($addrG)
                    $cfMpFail.Formula = ("AND(ISNUMBER(G1),G1>={0})" -f ([double]$ps.FailThreshold).ToString([System.Globalization.CultureInfo]::InvariantCulture))
                    $cfMpFail.Style.Fill.PatternType = 'Solid'
                    $cfMpFail.Style.Fill.BackgroundColor.Color = [System.Drawing.ColorTranslator]::FromHtml('#ffc7ce')
                }
                if ($ps.WarnThreshold -ne $null) {
                    $cfMpWarn = $Worksheet.ConditionalFormatting.AddExpression($addrG)
                    $cfMpWarn.Formula = ("AND(ISNUMBER(G1),G1>={0})" -f ([double]$ps.WarnThreshold).ToString([System.Globalization.CultureInfo]::InvariantCulture))
                    $cfMpWarn.Style.Fill.PatternType = 'Solid'
                    $cfMpWarn.Style.Fill.BackgroundColor.Color = [System.Drawing.ColorTranslator]::FromHtml('#ffe699')
                }
                } catch {}
            }
        } else {
            $Worksheet.Cells["A$r"].Value = "—"
            $Worksheet.Cells["A$r:H$r"].Merge = $true
            $r++
        }

        # Note if truncated
        if ($Evaluation.AffectedTestsTruncated -and [int]$Evaluation.AffectedTestsTruncated -gt 0) {
            $r += 1
            $Worksheet.Cells["A$r"].Value = ("NOTE: Affected tests truncated. {0} rows not shown." -f [int]$Evaluation.AffectedTestsTruncated)
            $Worksheet.Cells["A$r:H$r"].Merge = $true
            $Worksheet.Cells["A$r"].Style.Font.Italic = $true
            $r++
        }

        # Column widths (lightweight + stable)
        try {
            $Worksheet.Column(1).Width = 14
            $Worksheet.Column(2).Width = 40
            $Worksheet.Column(3).Width = 18
            $Worksheet.Column(4).Width = 18
            $Worksheet.Column(5).Width = 18
            $Worksheet.Column(6).Width = 36
            $Worksheet.Column(7).Width = 16
            $Worksheet.Column(8).Width = 18
        } catch {}

        # Border for the top block (A2:H7)
        try {
            $Worksheet.Cells["A2:H7"].Style.Border.Top.Style = 1
            $Worksheet.Cells["A2:H7"].Style.Border.Bottom.Style = 1
            $Worksheet.Cells["A2:H7"].Style.Border.Left.Style = 1
            $Worksheet.Cells["A2:H7"].Style.Border.Right.Style = 1
        } catch {}

    } catch {
        if (Get-Command Log-Exception -ErrorAction SilentlyContinue) {
            Log-Exception -Message "Information2 failed to build" -ErrorRecord $_ -Severity 'Warn'
        } elseif (Get-Command Gui-Log -ErrorAction SilentlyContinue) {
            Gui-Log ("Information2 failed: {0}" -f $_.Exception.Message) 'Warn'
        }

        try {
            $Worksheet.Cells.Clear()
            $Worksheet.Cells["A1"].Value = "Information2 failed"
            $Worksheet.Cells["A1"].Style.Font.Bold = $true
            $Worksheet.Cells["A2"].Value = $_.Exception.Message
            $Worksheet.Cells["A2"].Style.WrapText = $true
        } catch {}
    }
}
