<#
    Module: DataHelpers.ps1
    Purpose: Low-level helpers for loading and parsing CSV, Worksheet, Seal Test, and Excel templates.
    Platform: PowerShell 5.1 + EPPlus 4.5.3.3 (.NET 3.5)
    Notes:
      - No business rule changes; this module reorganises existing logic for clarity.
      - Heavy operations are explicitly marked to aid future optimisation.
#>

<#
    .SYNOPSIS
        Locates or downloads EPPlus.dll compatible with .NET 3.5.

    .DESCRIPTION
        Searches common module folders and the script directory for EPPlus. If not found,
        attempts a NuGet download (heavy network call) and copies the DLL locally.

    .PARAMETER Version
        Version string expected (default 4.5.3.3).
    .PARAMETER SourceDllPath
        Optional explicit DLL to prefer (site-specific UNC path).
    .PARAMETER LocalFolder
        Temporary folder used when downloading from NuGet.

    .OUTPUTS
        String path to EPPlus.dll, or $null if unavailable.

    .NOTES
        Heavy operation when NuGet download is required; logged for traceability.
#>
function Ensure-EPPlus {
    param(
        [string] $Version = "4.5.3.3",
        [string] $SourceDllPath = "N:\QC\QC-1\IPT\Skiftspecifika dokument\PQC analyst\JESPER\Scripts\Modules\EPPlus\EPPlus.4.5.3.3\.5.3.3\lib\net35\EPPlus.dll",
        [string] $LocalFolder = "$env:TEMP\EPPlus"
    )
    try {
        $candidatePaths = @()
        if ($SourceDllPath) { $candidatePaths += $SourceDllPath }
        $localScriptDll = Join-Path $PSScriptRoot 'EPPlus.dll'
        $candidatePaths += $localScriptDll
        $userModRoot = Join-Path ([Environment]::GetFolderPath('MyDocuments')) 'WindowsPowerShell\Modules'
        if (Test-Path $userModRoot) {
            Get-ChildItem -Path (Join-Path $userModRoot 'EPPlus') -Directory -ErrorAction SilentlyContinue | ForEach-Object {
                $candidatePaths += Join-Path $_.FullName 'lib\net45\EPPlus.dll'
                $candidatePaths += Join-Path $_.FullName 'lib\net40\EPPlus.dll'
            }
        }

        $progFiles = $env:ProgramFiles
        $systemModRoot = Join-Path $progFiles 'WindowsPowerShell\Modules'
        if (Test-Path $systemModRoot) {
            Get-ChildItem -Path (Join-Path $systemModRoot 'EPPlus') -Directory -ErrorAction SilentlyContinue | ForEach-Object {
                $candidatePaths += Join-Path $_.FullName 'lib\net45\EPPlus.dll'
                $candidatePaths += Join-Path $_.FullName 'lib\net40\EPPlus.dll'
            }
        }
        foreach ($cand in $candidatePaths) {
            if (-not [string]::IsNullOrWhiteSpace($cand) -and (Test-Path -LiteralPath $cand)) { return $cand }
        }

        $nugetUrl = "https://www.nuget.org/api/v2/package/EPPlus/$Version"
        try {
            $guid = [Guid]::NewGuid().ToString()
            $tempDir = Join-Path $env:TEMP "EPPlus_$guid"
            New-Item -ItemType Directory -Path $tempDir -Force | Out-Null
            $zipPath  = Join-Path $tempDir 'EPPlus.zip'
            $reqParams = @{ Uri = $nugetUrl; OutFile = $zipPath; UseBasicParsing = $true; Headers = @{ 'User-Agent' = 'DocMerge/1.0' } }
            Invoke-WebRequest @reqParams -ErrorAction Stop | Out-Null 
            if (-not ([System.AppDomain]::CurrentDomain.GetAssemblies() | Where-Object { $_.GetName().Name -eq 'System.IO.Compression.FileSystem' })) {
                Add-Type -AssemblyName 'System.IO.Compression.FileSystem' -ErrorAction SilentlyContinue
            }
            [System.IO.Compression.ZipFile]::ExtractToDirectory($zipPath, $tempDir)
            $extractedRoot = Join-Path $tempDir 'lib'
            if (Test-Path $extractedRoot) {
                $dllCandidates = Get-ChildItem -Path (Join-Path $extractedRoot 'net45'), (Join-Path $extractedRoot 'net40') -Filter 'EPPlus.dll' -Recurse -ErrorAction SilentlyContinue | Select-Object -First 1
                if ($dllCandidates) {
                    try {
                        if (-not (Test-Path -LiteralPath $localScriptDll)) {
                            Copy-Item -Path $dllCandidates.FullName -Destination $localScriptDll -Force -ErrorAction SilentlyContinue
                        }
                    } catch { Gui-Log "⚠️ Kunde inte kopiera EPPlus.dll: $($_.Exception.Message)" 'Warn' }
                    return $localScriptDll
                }
            }
        } catch {
            Gui-Log "❌ EPPlus: Kunde inte hämta EPPlus ($Version): $($_.Exception.Message)" 'Error'
        }
    } catch {
        Gui-Log "❌ Ensure-EPPlus fel: $($_.Exception.Message)" 'Error'
    }
    Gui-Log "❌ EPPlus.dll hittades inte. Installera EPPlus $Version manuellt." 'Error'
    return $null
}

<#
    .SYNOPSIS
        Loads EPPlus into the current AppDomain if not already present.

    .OUTPUTS
        Boolean success indicator.

    .NOTES
        Relies on Ensure-EPPlus to resolve the DLL path.
#>
function Load-EPPlus {
    if ([System.AppDomain]::CurrentDomain.GetAssemblies() | Where-Object { $_.GetName().Name -eq 'EPPlus' }) { return $true }
    $dllPath = Ensure-EPPlus -Version '4.5.3.3'
    if (-not $dllPath) { return $false }
    try {
        $bytes = [System.IO.File]::ReadAllBytes($dllPath)
        [System.Reflection.Assembly]::Load($bytes) | Out-Null
        return $true
    } catch {
        Gui-Log "❌ EPPlus-fel: $($_.Exception.Message)" 'Error'
        return $false
    }
}

<#
    .SYNOPSIS
        Returns the data worksheets from a Seal Test package (skips instructions tab).

    .PARAMETER Pkg
        EPPlus ExcelPackage.

    .OUTPUTS
        Array of worksheets excluding the first "Worksheet Instructions" tab.
#>
function Get-DataSheets { param([OfficeOpenXml.ExcelPackage]$Pkg)
    $all = @($Pkg.Workbook.Worksheets | Where-Object { $_.Name -ne "Worksheet Instructions" })
    if ($all.Count -gt 1) { return $all | Select-Object -Skip 1 } else { return @() }
}

<#
    .SYNOPSIS
        Safe URL encoder helper for SharePoint links.

    .PARAMETER s
        Raw string to encode.
    .OUTPUTS
        Encoded string or original on failure.
#>
function UrlEncode([string]$s){ try { [System.Uri]::EscapeDataString($s) } catch { $s } }

<#
    .SYNOPSIS
        Reads Batch Number from a Seal Test workbook (cell D2 on data sheets).

    .PARAMETER Path
        Seal Test NEG/POS file path.

    .OUTPUTS
        Batch number string or $null if not found.

    .NOTES
        Relies on Load-EPPlus; lightweight scan over data sheets only.
#>
function Get-BatchNumberFromSealFile([string]$Path){
    if (-not (Test-Path -LiteralPath $Path)) { return $null }
    if (-not (Load-EPPlus)) { return $null }
    $pkg = $null
    try {
        $pkg = New-Object OfficeOpenXml.ExcelPackage (New-Object IO.FileInfo($Path))
        foreach ($ws in (Get-DataSheets $pkg)) {
            $txt = ($ws.Cells['D2'].Text + '').Trim()   # "Batch Number"
            if ($txt) { return $txt }
        }
    } catch {
        Gui-Log "⚠️ Get-BatchNumberFromSealFile: $($_.Exception.Message)" 'Warn'
    } finally { if ($pkg) { try { $pkg.Dispose() } catch {} } }
    return $null
}

<#
    .SYNOPSIS
        Creates SharePoint batch link metadata based on Seal Test files and LSP.

    .PARAMETER SealPosPath
        Path to Seal Test POS file.
    .PARAMETER SealNegPath
        Path to Seal Test NEG file.
    .PARAMETER Lsp
        LSP identifier (used for placeholder replacement in template).

    .OUTPUTS
        PSCustomObject with Batch, Url, LinkText.
#>
function Get-BatchLinkInfo {
    param(
        [string]$SealPosPath,
        [string]$SealNegPath,
        [string]$Lsp
    )

    $batch = $null
    try { if ($SealPosPath) { $batch = Get-BatchNumberFromSealFile $SealPosPath } } catch {}
    if (-not $batch) {
        try { if ($SealNegPath) { $batch = Get-BatchNumberFromSealFile $SealNegPath } } catch {}
    }

    $batchEsc = if ($batch) { [uri]::EscapeDataString($batch) } else { '' }
    $lspEsc   = if ($Lsp)   { [uri]::EscapeDataString($Lsp) }   else { '' }

    $url = if ($SharePointBatchLinkTemplate) {
        ($SharePointBatchLinkTemplate -replace '\{BatchNumber\}', $batchEsc) -replace '\{LSP\}', $lspEsc
    } else {
        "https://danaher.sharepoint.com/sites/CEP-Sweden-Production-Management/Lists/Cepheid%20%20Production%20orders/AllItems.aspx?view=7&q=$batchEsc"
    }
    $linkText = if ($batch) { "Öppna $batch" } else { 'Ingen batch funnen' }

    return [pscustomobject]@{
        Batch    = $batch
        Url      = $url
        LinkText = $linkText
    }
}

<#
    .SYNOPSIS
        Creates/overwrites the "SharePoint Info" worksheet with provided data.

    .PARAMETER Pkg
        Destination ExcelPackage.
    .PARAMETER Rows
        Collection of PSObjects to render; supports key/value layout (Rubrik/Värde) or table layout.
    .PARAMETER DesiredOrder
        Optional column order for table rendering.
    .PARAMETER Batch
        Batch identifier for status text.

    .OUTPUTS
        Boolean indicating success.

    .NOTES
        Deletes any existing "SharePoint Info" sheet before writing.
#>
function Write-SPSheet-Safe {
    param(
        [OfficeOpenXml.ExcelPackage]$Pkg,
        [object]$Rows,                    
        [string[]]$DesiredOrder,            
        [string]$Batch
    )
    if (-not $Pkg) { return $false }
    $Rows = @($Rows)
    $name = "SharePoint Info"
    $wsOld = $Pkg.Workbook.Worksheets[$name]
    if ($wsOld) { $Pkg.Workbook.Worksheets.Delete($wsOld) }
    $ws = $Pkg.Workbook.Worksheets.Add($name)
    if ($Rows.Count -eq 0 -or $Rows[0] -eq $null) {
        $ws.Cells[1,1].Value = "No rows found (Batch=$Batch)"
        return $true
    }
    $isKV = ($Rows[0].psobject.Properties.Name -contains 'Rubrik') -and `
            ($Rows[0].psobject.Properties.Name -contains 'Värde')
    if ($isKV) {
        $ws.Cells[1,1].Value = "SharePoint Information"
        $ws.Cells[1,2].Value = ""
        $ws.Cells["A1:B1"].Merge = $true
        $ws.Cells["A1"].Style.Font.Bold = $true
        $ws.Cells["A1"].Style.Font.Size = 12
        $ws.Cells["A1"].Style.Font.Color.SetColor([System.Drawing.Color]::White)
        $ws.Cells["A1"].Style.Fill.PatternType = "Solid"
        $ws.Cells["A1"].Style.Fill.BackgroundColor.SetColor([System.Drawing.Color]::DarkBlue)
        $ws.Cells["A1"].Style.HorizontalAlignment = "Center"
        $ws.Cells["A1"].Style.VerticalAlignment   = "Center"

        $r = 2
        foreach ($row in $Rows) {
            $ws.Cells[$r,1].Value = $row.Rubrik
            $ws.Cells[$r,2].Value = $row.'Värde'
            $r++
        }
        $lastRow = $r-1
        $ws.Cells["A2:A$lastRow"].Style.Font.Bold = $true
        $ws.Cells["A2:A$lastRow"].Style.Fill.PatternType = "Solid"
        $ws.Cells["A2:A$lastRow"].Style.Fill.BackgroundColor.SetColor([System.Drawing.Color]::Gainsboro)
        $ws.Cells["B2:B$lastRow"].Style.Fill.PatternType = "Solid"
        $ws.Cells["B2:B$lastRow"].Style.Fill.BackgroundColor.SetColor([System.Drawing.Color]::WhiteSmoke)

        $rng = $ws.Cells["A1:B$lastRow"]
        $rng.Style.Font.Name = "Arial"
        $rng.Style.Font.Size = 10
        $rng.Style.HorizontalAlignment = "Left"
        $rng.Style.VerticalAlignment   = "Center"
        $rng.Style.Border.Top.Style    = "Thin"
        $rng.Style.Border.Bottom.Style = "Thin"
        $rng.Style.Border.Left.Style   = "Thin"
        $rng.Style.Border.Right.Style  = "Thin"
        $rng.Style.Border.BorderAround("Medium")
        try { $rng.AutoFitColumns() | Out-Null } catch {}
    }
    else {
        $cols = @()
        if ($DesiredOrder) { $cols += $DesiredOrder }
        foreach ($k in $Rows[0].psobject.Properties.Name) {
            if ($cols -notcontains $k) { $cols += $k }
        }

        for ($c=0; $c -lt $cols.Count; $c++) {
            $ws.Cells[1,$c+1].Value = $cols[$c]
            $ws.Cells[1,$c+1].Style.Font.Bold = $true
        }
        $r = 2
        foreach ($row in $Rows) {
            for ($c=0; $c -lt $cols.Count; $c++) {
                $ws.Cells[$r,$c+1].Value = $row.$($cols[$c])
            }
            $r++
        }
        try {
            if ($ws.Dimension) {
                $maxR = [Math]::Min($ws.Dimension.End.Row, 2000)
                $ws.Cells[$ws.Dimension.Start.Row,$ws.Dimension.Start.Column,$maxR,$ws.Dimension.End.Column].AutoFitColumns() | Out-Null
            }
        } catch {}
    }
    return $true
}

<#
    .SYNOPSIS
        Applies a framed border around a row range for STF summary styling.

    .DESCRIPTION
        Resets border styles for B..H cells on the specified row, then draws medium
        borders at the table edges and thin internal borders.

    .PARAMETER ws
        Target Excel worksheet.
    .PARAMETER row
        Row number to style.
    .PARAMETER firstRow
        First data row (used to set top border thickness).
    .PARAMETER lastRow
        Last data row (used to set bottom border thickness).
#>
function Set-RowBorder {
    param ($ws, [int] $row, [int] $firstRow, [int] $lastRow)
    foreach ($col in 'B','C','D','E','F','G','H') {
        $ws.Cells["$col$row"].Style.Border.Left.Style   = "None"
        $ws.Cells["$col$row"].Style.Border.Right.Style  = "None"
        $ws.Cells["$col$row"].Style.Border.Top.Style    = "None"
        $ws.Cells["$col$row"].Style.Border.Bottom.Style = "None"
    }
    $ws.Cells["B$row"].Style.Border.Left.Style  = "Medium"
    $ws.Cells["H$row"].Style.Border.Right.Style = "Medium"
    foreach ($col in 'B','C','D','E','F','G') { $ws.Cells["$col$row"].Style.Border.Right.Style = "Thin" }
    $topStyle = if ($row -eq $firstRow) { "Medium" } else { "Thin" }
    $bottomStyle = if ($row -eq $lastRow)  { "Medium" } else { "Thin" }
    foreach ($col in 'B','C','D','E','F','G','H') {
        $ws.Cells["$col$row"].Style.Border.Top.Style    = $topStyle
        $ws.Cells["$col$row"].Style.Border.Bottom.Style = $bottomStyle
    }
}

<#
    .SYNOPSIS
        Convenience wrapper to set common style attributes on a cell.

    .PARAMETER cell
        EPPlus ExcelRange to style.
    .PARAMETER bold
        Switch to enable bold font.
    .PARAMETER bg
        Optional hex background color (without '#').
    .PARAMETER border
        Optional border style string (e.g., 'Medium', 'Thin').
    .PARAMETER fontColor
        Optional hex font color (without '#').
#>
function Style-Cell { param($cell,$bold,$bg,$border,$fontColor)
    if ($bold) { $cell.Style.Font.Bold = $true }
    if ($bg)   { $cell.Style.Fill.PatternType = "Solid"; $cell.Style.Fill.BackgroundColor.SetColor([System.Drawing.ColorTranslator]::FromHtml("#$bg")) }
    if ($fontColor) { $cell.Style.Font.Color.SetColor([System.Drawing.ColorTranslator]::FromHtml("#$fontColor")) }
    if ($border) { $cell.Style.Border.Top.Style=$border; $cell.Style.Border.Bottom.Style=$border; $cell.Style.Border.Left.Style=$border; $cell.Style.Border.Right.Style=$border }
}

<#
    .SYNOPSIS
        Checks whether a file is locked for read/write.

    .PARAMETER Path
        Target file path.

    .OUTPUTS
        Boolean: $true if locked/unwritable, $false if free.
#>
function Test-FileLocked { param([Parameter(Mandatory=$true)][string]$Path)
    try { $fs = [IO.File]::Open($Path,'Open','ReadWrite','None'); $fs.Close(); return $false } catch { return $true }
}

<#
    .SYNOPSIS
        Detects whether a CSV uses semicolon or comma as delimiter.

    .PARAMETER Path
        CSV file path.

    .OUTPUTS
        Delimiter character as string.
#>
function Get-CsvDelimiter { param([string]$Path)
    try {
        $first = Get-Content -LiteralPath $Path -Encoding Default -TotalCount 30 | Where-Object { $_ -and $_.Trim() } | Select-Object -First 1
        if (-not $first) { return ';' }
        $sc = ($first -split ';').Count; $cc = ($first -split ',').Count
        if ($cc -gt $sc -and $cc -ge 2) { return ',' } else { return ';' }
    } catch {
        Gui-Log "⚠️ Get-CsvDelimiter misslyckades: $($_.Exception.Message)" 'Warn'
        return ';'
    }
}

<#
    .SYNOPSIS
        Creates a VB TextFieldParser configured for delimited CSVs.

    .PARAMETER Path
        CSV file path.
    .PARAMETER Delimiter
        Delimiter character detected via Get-CsvDelimiter.

    .OUTPUTS
        Microsoft.VisualBasic.FileIO.TextFieldParser or $null on failure.
#>
function New-TextFieldParser { param([string]$Path,[string]$Delimiter)
    try {
        $tp = New-Object Microsoft.VisualBasic.FileIO.TextFieldParser($Path, [System.Text.Encoding]::Default)
        $tp.TextFieldType = [Microsoft.VisualBasic.FileIO.FieldType]::Delimited
        $tp.SetDelimiters($Delimiter)
        $tp.HasFieldsEnclosedInQuotes = $true
        $tp.TrimWhiteSpace = $true
        return $tp
    } catch {
        Gui-Log "⚠️ New-TextFieldParser: $($_.Exception.Message)" 'Warn'
        return $null
    }
}

<#
    .SYNOPSIS
        Extracts the assay name from a Tests Summary CSV.

    .DESCRIPTION
        Scans rows starting at StartRow and returns the first non-empty value
        in column 1 that is not the literal header "assay".

    .PARAMETER Path
        CSV path.
    .PARAMETER StartRow
        Row index (1-based) to begin scanning, defaults to 10.

    .OUTPUTS
        String assay name or $null.
#>
function Get-AssayFromCsv { param([string]$Path,[int]$StartRow=10)
    if (-not (Test-Path -LiteralPath $Path)) { return $null }
    $tp = $null; $delim=Get-CsvDelimiter $Path; $row=0
    try {
        $tp = New-TextFieldParser -Path $Path -Delimiter $delim
        if (-not $tp) { return $null }
        while (-not $tp.EndOfData) {
            $row++; $f = $tp.ReadFields()
            if ($row -lt $StartRow) { continue }
            if (-not $f -or $f.Length -lt 1) { continue }
            $a=([string]$f[0]).Trim()
            if ($a -and $a -notmatch '^(?i)\s*assay\s*$') { return $a }
        }
    } catch {
        Gui-Log "⚠️ Get-AssayFromCsv: $($_.Exception.Message)" 'Warn'
    } finally { if ($tp){$tp.Close()} }
    return $null
}

<#
    .SYNOPSIS
        Imports data rows from a Tests Summary CSV starting at a given row.

    .PARAMETER Path
        CSV path.
    .PARAMETER StartRow
        Row index to begin reading (default 10).

    .OUTPUTS
        Array of string arrays (fields).

    .NOTES
        Heavy operation on large CSVs; uses TextFieldParser to preserve quoting.
#>
function Import-CsvRows { param([string]$Path,[int]$StartRow=10)
    if (-not (Test-Path -LiteralPath $Path)) { return @() }
    $delim=Get-CsvDelimiter $Path; $tp=$null; $rows=@()
    try {
        $tp = New-TextFieldParser -Path $Path -Delimiter $delim
        if (-not $tp) { return @() }
        $r=0
        while (-not $tp.EndOfData) {
            $r++; $f=$tp.ReadFields()
            if ($r -lt $StartRow) { continue }
            if (-not $f -or ($f -join '').Trim().Length -eq 0) { continue }
            $rows += ,$f
        }
    } catch {
        Gui-Log "⚠️ Import-CsvRows: $($_.Exception.Message)" 'Warn'
    } finally { if ($tp){$tp.Close()} }
    return ,@($rows)
}

<#
    .SYNOPSIS
        Splits a CSV line while respecting quoted fields.

    .PARAMETER Line
        Raw CSV line.

    .OUTPUTS
        Array of fields.
#>
function ConvertTo-CsvFields {
    param([string]$Line)
    return [regex]::Split($Line, ',(?=(?:[^"]*"[^"]*")*[^"]*$)')
}

<#
    .SYNOPSIS
        Calculates CSV statistics (counts, duplicates, instrument summary).

    .DESCRIPTION
        Reads a Tests Summary CSV (or supplied lines) and returns counts of tests,
        duplicates, detected LSPs, and instrument usage grouped by type.

    .PARAMETER Path
        CSV file path.
    .PARAMETER Lines
        Optional pre-loaded lines to avoid re-reading from disk.

    .OUTPUTS
        PSCustomObject with TestCount, DupCount, Duplicates, LspValues, LspOK, InstrumentByType.

    .NOTES
        Heavy operation for large CSVs; reused across Information sheet and logging.
#>
function Get-CsvStats {
    param(
        [string]$Path,
        [string[]]$Lines
    )
    $out = [ordered]@{
        TestCount       = 0
        DupCount        = 0
        Duplicates      = @()
        LspValues       = @()
        LspOK           = $null
        InstrumentByType= [ordered]@{}
    }
    $lines = $Lines
    if (-not $lines) {
        if (-not (Test-Path -LiteralPath $Path)) { return [pscustomobject]$out }
        try { $lines = Get-Content -LiteralPath $Path } catch { Gui-Log "⚠️ Get-CsvStats: $($_.Exception.Message)" 'Warn'; return [pscustomobject]$out }
    }
    if (-not $lines -or $lines.Count -lt 8) { return [pscustomobject]$out }
    $header = ConvertTo-CsvFields $lines[7]
    $dataLines = @()
    if ($lines.Count -gt 9) { $dataLines = $lines[9..($lines.Count-1)] }
    $csv = $null
    try {
        if ($dataLines.Count -gt 0) {
            $csv = ConvertFrom-Csv -InputObject ($dataLines -join "`n") -Header $header
        }
    } catch { $csv = $null }

    $cartSnList = New-Object System.Collections.Generic.List[string]
    $lspList    = New-Object System.Collections.Generic.List[string]
    $instrList  = New-Object System.Collections.Generic.List[string] 
    if ($csv) {
        foreach ($row in $csv) {
            $cart = $row.'Cartridge S/N'
            $lsp  = $row.'Reagent Lot ID'
            $ins  = $row.'Instrument S/N'
            if (-not $cart) { try { $cart = $row.Item(3) } catch {} }
            if (-not $ins)  { try { $ins  = $row.Item(6) } catch {} }
            if ($cart) { $cartSnList.Add( ($cart + '').Trim() ) }
            if ($lsp)  { $lspList.Add(  ($lsp  + '').Trim() ) }
            if ($ins)  { $instrList.Add(($ins  + '').Trim() ) }
        }
    } else {
        foreach ($ln in $dataLines) {
            if (-not $ln -or -not $ln.Trim()) { continue }
            $f = ConvertTo-CsvFields $ln
            if ($f.Count -lt 7) { continue }
            $cartSnList.Add( ($f[3] + '').Trim() )
            $lspList.Add(    ($f[4] + '').Trim() )
            $instrList.Add(  ($f[6] + '').Trim() )
        }
    }
    $cartSn = $cartSnList | Where-Object { $_ -and $_ -ne '' }
    $out.TestCount = @($cartSn).Count
    $dups = $cartSn | Group-Object | Where-Object { $_.Count -gt 1 }
    if ($dups) {
        $out.DupCount   = $dups.Count
        $out.Duplicates = $dups | ForEach-Object { "$($_.Name) x$($_.Count)" }
    }
    $lspClean = $lspList | ForEach-Object {
        $m = [regex]::Match($_, '(?<!\d)(\d{5})(?!\d)')
        if ($m.Success) { $m.Groups[1].Value } else { $null }
    } | Where-Object { $_ } | Select-Object -Unique
    $out.LspValues = $lspClean
    if ($lspClean.Count -gt 0) {
        $out.LspOK = ($lspClean.Count -eq 1)
    }
    $lut = @{}
    foreach ($k in $script:GXINF_Map.Keys) {
        $codes = $script:GXINF_Map[$k].Split(',') | ForEach-Object { $_.Trim() } | Where-Object { $_ }
        foreach ($code in $codes) { $lut[$code] = $k }
    }
    foreach ($ins in ($instrList | Where-Object { $_ })) {
        $t = $null
        if ($lut.ContainsKey($ins)) { $t = $lut[$ins] } else { $t = 'Unknown' }
        if (-not $out.InstrumentByType.Contains($t)) { $out.InstrumentByType[$t] = 0 }
        $out.InstrumentByType[$t]++
    }
    return [pscustomobject]$out
}

<#
    .SYNOPSIS
        Formats Infinity SP presence counts into a compressed string.

    .PARAMETER Counts
        Hashtable where key=SP number, value=count.

    .OUTPUTS
        Human readable summary such as "#01-03 x5".
#>
function Format-SpPresenceGrandTotalStrict {
    param([hashtable]$Counts) 
    if (-not $Counts -or $Counts.Count -eq 0) { return '—' }
    $present = @(
        $Counts.GetEnumerator() |
        Where-Object { $_.Value -gt 0 } |
        ForEach-Object { [int]$_.Key } |
        Sort-Object
    )
    if ($present.Count -eq 0) { return '—' }
    $parts = New-Object System.Collections.Generic.List[string]
    $i = 0
    while ($i -lt $present.Count) {
        $start = $present[$i]; $end = $start
        $j = $i + 1
        while ($j -lt $present.Count -and $present[$j] -eq $end + 1) {
            $end = $present[$j]; $j++
        }
        if ($start -eq $end) { $parts.Add( ('#{0:00}' -f $start) ) }
        else { $parts.Add( ('#{0:00}-{1:00}' -f $start, $end) ) }
        $i = $j
    }
    $total = (
        $Counts.GetEnumerator() |
        Where-Object { $_.Value -gt 0 } |
        Measure-Object -Property Value -Sum
    ).Sum
    return ( ($parts -join '+') + (' x{0}' -f $total) )
}

if (Get-Command Get-InfinitySpFromCsvStrict -ErrorAction SilentlyContinue) {
    Remove-Item Function:\Get-InfinitySpFromCsvStrict -ErrorAction SilentlyContinue
}

<#
    .SYNOPSIS
        Counts Infinity sample positions from a CSV for a given instrument set.

    .DESCRIPTION
        Parses the Tests Summary CSV (or provided lines) to detect sample IDs of
        Infinity instruments, summarising present SP numbers.

    .PARAMETER Path
        CSV path.
    .PARAMETER InstrumentSerials
        Array of instrument serial numbers to match.
    .PARAMETER Lines
        Optional pre-loaded lines to avoid re-reading.

    .OUTPUTS
        String summary or '—' if nothing found.

    .NOTES
        Heavy loop over CSV rows; avoid calling repeatedly.
#>
function Get-InfinitySpFromCsvStrict {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$Path,
        [Alias('InfinitySerials')][string[]]$InstrumentSerials,
        [string[]]$Lines
    )
    if (-not $Lines -and -not (Test-Path -LiteralPath $Path)) { return '—' }
    $useConvertFn = $false
    try { $useConvertFn = [bool](Get-Command ConvertTo-CsvFields -ErrorAction Stop) } catch {}
    function Split-CsvSmart([string]$ln) {
        if ($ln -like '*;*' -and $ln -notlike '*,*') {
            return [regex]::Split($ln, ';(?=(?:[^"]*"[^"]*")*[^"]*$)')
        } else {
            return [regex]::Split($ln, ',(?=(?:[^"]*"[^"]*")*[^"]*$)')
        }
    }

    $serSet = New-Object System.Collections.Generic.HashSet[string]([StringComparer]::OrdinalIgnoreCase)
    foreach ($s in ($InstrumentSerials | Where-Object { $_ })) { $null = $serSet.Add( ($s + '').Trim().Trim('"') ) }
    if ($serSet.Count -eq 0) { return '—' } 

    $lines = $Lines
    if (-not $lines) {
        try { $lines = Get-Content -LiteralPath $Path } catch { Gui-Log "⚠️ Get-InfinitySpFromCsvStrict: $($_.Exception.Message)" 'Warn'; return '—' }
    }
    if (-not $lines -or $lines.Count -lt 10) { return '—' }

    $headerIndex = 7
    $dataStart   = 9

    $hdr = if ($useConvertFn) { ConvertTo-CsvFields $lines[$headerIndex] } else { Split-CsvSmart $lines[$headerIndex] }
    $colCount = $hdr.Count

    $idxInstr = -1
    for ($i=0; $i -lt $colCount; $i++) {
        $h = (($hdr[$i] + '') -replace '[\uFEFF\u200B]','').Trim().ToLowerInvariant()
        if ($h -match 'instrument' -and ($h -match 's/?n' -or $h -match 'serial')) { $idxInstr = $i; break }
    }
    if ($idxInstr -lt 0) { $idxInstr = 6 }  # G
    $idxSample = 2
    $counts = @{}

    for ($r = $dataStart; $r -lt $lines.Count; $r++) {
        $ln = $lines[$r]; if (-not $ln -or -not $ln.Trim()) { continue }
        $f = if ($useConvertFn) { ConvertTo-CsvFields $ln } else { Split-CsvSmart $ln }
        if ($f.Count -le [Math]::Max($idxInstr,$idxSample)) { continue }

        $instr  = ($f[$idxInstr] + '').Trim().Trim('"')
        if (-not $serSet.Contains($instr)) { continue }  # inte Infinity → skippa

        $sample = ($f[$idxSample] + '').Trim().Trim('"')
        if (-not $sample) { continue }

        $m = [regex]::Match($sample, '_(\d{2})_')
         if ($m.Success) {
             $nRaw = 0
             if ([int]::TryParse($m.Groups[1].Value, [ref]$nRaw)) {

                $sp = $nRaw 
                if ($sp -ge 0 -and $sp -le 10) {
                     if (-not $counts.ContainsKey($sp)) { $counts[$sp] = 0 }
                     $counts[$sp]++
                 }
             }
         }
    }

    if ($counts.Count -eq 0) { return '—' }
    return (Format-SpPresenceGrandTotalStrict -Counts $counts)
 }

<#
    .SYNOPSIS
        Normalizes assay names for lookup.

    .PARAMETER s
        Raw assay text.

    .OUTPUTS
        Lower-cased alphanumeric string with single spaces.
#>
function Normalize-Assay { param([string]$s)
    if ([string]::IsNullOrWhiteSpace($s)) { return $null }
    $x=$s.ToLowerInvariant(); $x=[regex]::Replace($x,'[^a-z0-9]+',' '); $x=$x.Trim() -replace '\s+',' '; return $x
}
$AssayMap = @(

    @{ Tab='MTB ULTRA';            Aliases=@('Xpert MTB-RIF Ultra') }
    @{ Tab='MTB RIF';              Aliases=@('Xpert MTB-RIF Assay G4') }
    @{ Tab='MTB JP';               Aliases=@('Xpert MTB-RIF JP IVD') }
    @{ Tab='MTB XDR';              Aliases=@('Xpert MTB-XDR') }
    @{ Tab='FLUVID | FLUVID+';     Aliases=@('Xpress SARS-CoV-2_Flu_RSV plus','Xpert Xpress_SARS-CoV-2_Flu_RSV') }
    @{ Tab='SARS-COV-2 Plus';      Aliases=@('Xpert Xpress CoV-2 plus') }
    @{ Tab='CTNG | CTNG JP';       Aliases=@('Xpert CT_NG','Xpert CT_CE') }
    @{ Tab='C.DIFF | C.DIFF JP';   Aliases=@('Xpert C.difficile G3','Xpert C.difficile BT') }
    @{ Tab='HPV';                  Aliases=@('Xpert HPV HR','Xpert HPV v2 HR') }
    @{ Tab='HBV VL';               Aliases=@('Xpert HBV Viral Load') }
    @{ Tab='HCV VL';               Aliases=@('Xpert HCV Viral Load','Xpert_HCV Viral Load') }
    @{ Tab='HCV VL FS';            Aliases=@('Xpert HCV VL Fingerstick') }
    @{ Tab='HIV VL';               Aliases=@('Xpert HIV-1 Viral Load','Xpert_HIV-1 Viral Load') }
    @{ Tab='HIV VL XC';            Aliases=@('Xpert HIV-1 Viral Load XC') }
    @{ Tab='HIV QA';               Aliases=@('Xpert HIV-1 Qual','Xpert_HIV-1 Qual') }
    @{ Tab='HIV QA XC';            Aliases=@('Xpert HIV-1 Qual XC PQC','Xpert HIV-1 Qual XC') }
    @{ Tab='SARS-COV-2';           Aliases=@('Xpert Xpress SARS-CoV-2 CE-IVD','Xpert Xpress SARS-CoV-2') }
    @{ Tab='FLU RSV';              Aliases=@('Xpert Xpress Flu-RSV','Xpress Flu IPT_EAT off') }
    @{ Tab='MRSA SA';              Aliases=@('Xpert SA Nasal Complete G3','Xpert MRSA-SA SSTI G3') }
    @{ Tab='MRSA NxG';             Aliases=@('Xpert MRSA NxG') }
    @{ Tab='NORO';                 Aliases=@('Xpert Norovirus') }
    @{ Tab='VAN AB';               Aliases=@('Xpert vanA vanB') }
    @{ Tab='GBS';                  Aliases=@('Xpert GBS LB XC','Xpert Xpress GBS','Xpert Xpress GBS US-IVD') }
    @{ Tab='STREP A';              Aliases=@('Xpert Xpress Strep A') }
    @{ Tab='CARBA R';              Aliases=@('Xpert Carba-R','Xpert_Carba-R') }
    @{ Tab='EBOLA';                Aliases=@('Xpert Ebola EUA','Xpert Ebola CE-IVD') }
    @{ Tab='Respiratory Panel';    Aliases=@('Respiratory Panel IUO') }

)

$AssayIndex = @{}
foreach($row in $AssayMap){ foreach($a in $row.Aliases){ $k=Normalize-Assay $a; if($k -and -not $AssayIndex.ContainsKey($k)){ $AssayIndex[$k]=$row.Tab } } }

<#
    .SYNOPSIS
        Maps an assay name to a Control Material worksheet tab.

    .PARAMETER AssayName
        Assay name from CSV or Seal Test.

    .OUTPUTS
        Tab name string or $null if no mapping found.

    .NOTES
        Falls back to slangassay.xlsx if provided; heavy operation when reading that file.
#>
function Get-ControlTabName {
    param([string]$AssayName)
    $k = Normalize-Assay $AssayName
    if ($k -and $AssayIndex.ContainsKey($k)) { return $AssayIndex[$k] }

    if (Test-Path $SlangAssayPath) {
        try {
            # NOTE: Heavy operation - opens slangassay.xlsx to resolve aliases.
            $mapPkg = New-Object OfficeOpenXml.ExcelPackage (New-Object IO.FileInfo($SlangAssayPath))
            $ws = $mapPkg.Workbook.Worksheets['Slang till Assay']; if (-not $ws) { $ws = $mapPkg.Workbook.Worksheets[1] }
            if ($ws -and $ws.Dimension) {
                for ($r=2; $r -le $ws.Dimension.End.Row; $r++){
                    $sheet=$ws.Cells[$r,1].Text.Trim()
                    $aliases=@($ws.Cells[$r,2].Text,$ws.Cells[$r,3].Text,$ws.Cells[$r,4].Text) | Where-Object { $_ -and $_.Trim() }
                    foreach($al in $aliases){ if (Normalize-Assay $AssayName -eq (Normalize-Assay $al)) { $mapPkg.Dispose(); return $sheet } }
                }
            }
            $mapPkg.Dispose()
        } catch {
            Gui-Log "⚠️ Get-ControlTabName: $($_.Exception.Message)" 'Warn'
        }
    }

    return $null

}

$MinitabMap = @(

    @{ Aliases=@('Xpert MTB-RIF Ultra');                           Macro='%D12547-MTBU-SWE' }
    @{ Aliases=@('Xpert MTB-RIF Assay G4');                        Macro='%D12547-MTB-SWE' }
    @{ Aliases=@('Xpress SARS-CoV-2_Flu_RSV plus','Xpert Xpress_SARS-CoV-2_Flu_RSV'); Macro='%D12547-XP3COV2FLURSV-SWE' }
    @{ Aliases=@('Xpert Xpress CoV-2 plus');                        Macro='%D12547-XP3SARSCOV2-SWE' }
    @{ Aliases=@('CT_NG','Xpert CT_CE');                            Macro='%D12547-CTNG-SWE' }
    @{ Aliases=@('Xpert C.difficile G3','Xpert C.difficile BT');    Macro='%D12547-CDCE-SWE' }
    @{ Aliases=@('Xpert HPV HR','Xpert HPV v2 HR');                 Macro='%D12547-HPV-SWE' }
    @{ Aliases=@('Xpert HBV Viral Load');                           Macro='%D12547-HBVVL-SWE' }
    @{ Aliases=@('Xpert HCV Viral Load','Xpert_HCV Viral Load');    Macro='%D12547-HCVVL-SWE' }
    @{ Aliases=@('Xpert HCV VL Fingerstick');                       Macro='%D12547-FSHCV-SWE' }
    @{ Aliases=@('Xpert HIV-1 Viral Load','Xpert_HIV-1 Viral Load'); Macro='%D12547-HIVVL-SWE' }
    @{ Aliases=@('Xpert HIV-1 Qual','Xpert_HIV-1 Qual');            Macro='%D12547-HIVQA-SWE' }
    @{ Aliases=@('Xpert Xpress SARS-CoV-2 CE-IVD','Xpert Xpress SARS-CoV-2'); Macro='%D12547-XPRSARSCOV2-SWE' }
    @{ Aliases=@('Xpert Xpress Flu-RSV');                           Macro='%D12547-XPFLURSV-SWE' }
    @{ Aliases=@('Xpress Flu IPT_EAT off');                         Macro='%D12547-FLUNG-SWE' } 
    @{ Aliases=@('Xpert Norovirus');                                Macro='%D12547-NORO-SWE' }
    @{ Aliases=@('Xpert vanA vanB');                                Macro='%D12547-VAB-SWE' }
    @{ Aliases=@('Xpert Xpress Strep A');                           Macro='%D12547-STREPA-SWE' }
    @{ Aliases=@('Xpert Carba-R','Xpert_Carba-R');                  Macro='%D12547-CARBAR-SWE' }
    @{ Aliases=@('Xpert Ebola EUA','Xpert Ebola CE-IVD');           Macro='%D12547-EBOLA-SWE' }
    @{ Aliases=@('Xpert SA Nasal Complete G3','Xpert MRSA-SA SSTI G3'); Macro='%D12547-SACOMP-SWE' }

    # N/A-gruppen:
    @{ Aliases=@('Xpert GBS LB XC','Xpert Xpress GBS','Xpert Xpress GBS US-IVD'); Macro=$null }
    @{ Aliases=@('Xpert HIV-1 Qual XC PQC','Xpert HIV-1 Qual XC');  Macro=$null }
    @{ Aliases=@('Xpert HIV-1 Viral Load XC');                      Macro=$null }
    @{ Aliases=@('Xpert MTB-RIF JP IVD');                           Macro=$null }
    @{ Aliases=@('Xpert MTB-XDR');                                  Macro=$null }
    @{ Aliases=@('Xpert MRSA NxG');                                 Macro=$null }
)

$MinitabIndex = @{}
foreach ($row in $MinitabMap) { foreach ($a in $row.Aliases) { $k = Normalize-Assay $a; if ($k -and -not $MinitabIndex.ContainsKey($k)) { $MinitabIndex[$k] = $row.Macro } } }

<#
    .SYNOPSIS
        Resolves the Minitab macro name for a given assay.

    .PARAMETER AssayName
        Assay name from CSV/Seal Test.

    .OUTPUTS
        Macro string or $null.
#>
function Get-MinitabMacro { param([string]$AssayName)
    if ([string]::IsNullOrWhiteSpace($AssayName)) { return $null }
    $k = Normalize-Assay $AssayName
    if ($k -and $MinitabIndex.ContainsKey($k)) { return $MinitabIndex[$k] }
    return $null
}

<#
    .SYNOPSIS
        Finds the column index for the observation column in Seal Test sheets.

    .PARAMETER ws
        Worksheet to inspect.

    .OUTPUTS
        Column index (int), defaults to 13 if not detected.
#>
function Find-ObservationCol { param($ws)

    $default = 13
    if (-not $ws -or -not $ws.Dimension) { return $default }
    $maxR = [Math]::Min(5, $ws.Dimension.End.Row)
    $maxC = $ws.Dimension.End.Column

    for ($r=1; $r -le $maxR; $r++) {
        for ($c=1; $c -le $maxC; $c++) {
            $t = ($ws.Cells[$r,$c].Text + '').Trim()
            if ($t -match '^(?i)\s*(obs|observation)') { return $c }
        }
    }
    return $default
}

if (-not (Get-Command Extract-WorksheetHeader -ErrorAction SilentlyContinue)) {
    <#
        .SYNOPSIS
            Parses header/footer metadata from a Worksheet.xlsx file.

        .PARAMETER Pkg
            ExcelPackage for the worksheet file.

        .OUTPUTS
            PSCustomObject with PartNo, BatchNo, CartridgeNo, DocumentNumber, Attachment, Rev, Effective.

        .NOTES
            Heavy: iterates all worksheets to combine header/footer data.
    #>
    function Extract-WorksheetHeader {
        param([OfficeOpenXml.ExcelPackage]$Pkg)
        $result = [pscustomobject]@{
            PartNo         = $null
            BatchNo        = $null
            CartridgeNo    = $null
            DocumentNumber = $null
            Attachment     = $null
            Rev            = $null
            Effective      = $null
        }
        if (-not $Pkg) { return $result }
        try {
            foreach ($ws in $Pkg.Workbook.Worksheets) {
                if (-not $ws.Dimension) { continue }
                if ($ws.Name -match '(?i)for reference only') { continue }
                $left  = (($ws.HeaderFooter.OddHeader.LeftAlignedText  + '') -replace '\r?\n',' ')
                if (-not $left)  { $left  = (($ws.HeaderFooter.EvenHeader.LeftAlignedText  + '') -replace '\r?\n',' ') }
                $right = (($ws.HeaderFooter.OddHeader.RightAlignedText + '') -replace '\r?\n',' ')
                if (-not $right) { $right = (($ws.HeaderFooter.EvenHeader.RightAlignedText + '') -replace '\r?\n',' ') }
                $raw = (($left + ' ' + $right) + '').Trim()
                if (-not $raw) { continue }
                $parsed = Parse-WorksheetHeaderRaw -Raw $raw
                if (-not $result.PartNo         -and $parsed.PartNo)         { $result.PartNo         = $parsed.PartNo }
                if (-not $result.BatchNo        -and $parsed.BatchNo)        { $result.BatchNo        = $parsed.BatchNo }
                if (-not $result.CartridgeNo    -and $parsed.CartridgeNo)    { $result.CartridgeNo    = $parsed.CartridgeNo }
                if (-not $result.DocumentNumber -and $parsed.DocumentNumber) { $result.DocumentNumber = $parsed.DocumentNumber }
                if (-not $result.Attachment     -and $parsed.Attachment)     { $result.Attachment     = $parsed.Attachment }
                if (-not $result.Rev            -and $parsed.Rev)            { $result.Rev            = $parsed.Rev }
                if (-not $result.Effective      -and $parsed.Effective)      { $result.Effective      = $parsed.Effective }
            }
        } catch { Gui-Log "⚠️ Extract-WorksheetHeader: $($_.Exception.Message)" 'Warn' }
        if (-not $result.CartridgeNo) {
            foreach ($ws in $Pkg.Workbook.Worksheets) {
                if (-not $ws.Dimension) { continue }
                if ($ws.Name -match '(?i)for reference only') { continue }
                $left  = (($ws.HeaderFooter.OddHeader.LeftAlignedText  + '') -replace '\r?\n',' ').Trim()
                $right = (($ws.HeaderFooter.OddHeader.RightAlignedText + '') -replace '\r?\n',' ').Trim()
                $m  = [regex]::Match($left,  '(?<!\d)(\d{5})(?!\d)')
                $m2 = if (-not $m.Success) { [regex]::Match($right,'(?<!\d)(\d{5})(?!\d)') } else { $null }
                if ($m.Success) { $result.CartridgeNo = $m.Groups[1].Value; break }
                if ($m2 -and $m2.Success) { $result.CartridgeNo = $m2.Groups[1].Value; break }
            }
        }
        return $result
    }
}

if (-not (Get-Command Get-WorksheetHeaderPerSheet -ErrorAction SilentlyContinue)) {
    <#
        .SYNOPSIS
            Builds a per-sheet header summary used for cross-tab validation.

        .PARAMETER Pkg
            ExcelPackage containing Worksheet tabs.

        .OUTPUTS
            List of PSCustomObjects with parsed header values and label presence flags.
    #>
    function Get-WorksheetHeaderPerSheet {
        param([OfficeOpenXml.ExcelPackage]$Pkg)

        $rxPart      = '(?i)^\s*Part\s*(?:No|Number)\.?\s*(?:\(s\))?\s*$'
        $rxBatch     = '(?i)^\s*Batch\s*(?:No|Number)(?:\s*\(s\))?\.?\s*$'
        $rxCartridge = '(?i)^\s*Cartridge\s*(?:No|Number)?\s*(?:\(?LSP\)?)?\.?\s*$'
        $rxDoc       = '(?i)^\s*Document\s*(?:No|Number|#)\s*$'
        $rxRev       = '(?i)^\s*Rev(?:ision)?\.?\s*$'
        $rxEff       = '(?i)^\s*Effective(?:\s*Date)?\s*$'
        function Get-HeaderFooterText([Object]$ws) {
            $left  = Normalize-HeaderText ((($ws.HeaderFooter.OddHeader.LeftAlignedText  + '') -replace '\r?\n',' '))
            if (-not $left)  { $left  = Normalize-HeaderText ((($ws.HeaderFooter.EvenHeader.LeftAlignedText  + '') -replace '\r?\n',' ')) }
            $right = Normalize-HeaderText ((($ws.HeaderFooter.OddHeader.RightAlignedText + '') -replace '\r?\n',' '))
            if (-not $right) { $right = Normalize-HeaderText ((($ws.HeaderFooter.EvenHeader.RightAlignedText + '') -replace '\r?\n',' ')) }
            return @($left,$right)
        }

        function Try-FromHeaderFooter([string[]]$lr, [string]$kind) {
            $left,$right = $lr
            $raw = (($left + ' ' + $right) + '').Trim()
            if (-not $raw) {
                return @{ Has = $false; Val = $null }
            }
            $parsed = Parse-WorksheetHeaderRaw -Raw $raw
            $has = $false
            $val = $null

            switch ($kind) {
                'Part' {
                    $has = ($raw -match '(?i)\bPart\b')
                    $val = $parsed.PartNo
                }
                'Batch' {
                    $has = ($raw -match '(?i)\bBatch\b')
                    $val = $parsed.BatchNo
                }
                'Cartridge' {
                    $has = ($raw -match '(?i)\bCartridge\b')
                    $val = $parsed.CartridgeNo
                }
                'Doc' {
                    $has = ($raw -match '(?i)Document\s*Number')
                    $val = $parsed.DocumentNumber
                }
                'REV' {
                    $has = ($raw -match '(?i)\bRev(?:ision)?\b')
                    $val = $parsed.Rev
                }
                'EFF' {
                    $has = ($raw -match '(?i)\bEffective(\s*Date)?\b')
                    $val = $parsed.Effective
                }
            }
            return @{ Has = [bool]$has; Val = $val }
        }

        function Try-FromCells([Object]$ws, [string]$rxLabel, [int]$maxR=30, [int]$maxC=20, [string]$canon='') {
            $has=$false; $val=$null
            for($r=1;$r -le $maxR;$r++){
                for($c=1;$c -le $maxC;$c++){
                    $t     = Normalize-HeaderText (""+$ws.Cells[$r,$c].Text)
                    if([string]::IsNullOrWhiteSpace($t)){ continue }
                    if($t -match $rxLabel){
                        $has=$true
$right = Normalize-HeaderText (""+$ws.Cells[$r,[Math]::Min($c+1,$maxC)].Text)
$down  = Normalize-HeaderText (""+$ws.Cells[[Math]::Min($r+1,$maxR),$c].Text)
                        $val = if($right){$right}elseif($down){$down}else{''}
                        if($canon -eq 'Cartridge' -and $val){
                        elseif ($canon -eq 'EFF' -and $val) {
    $dt = $null
    if (Try-Parse-HeaderDate $val ([ref]$dt)) { $val = $dt.ToString('yyyy-MM-dd') }
}
                            $m=[regex]::Match($val,'(?<!\d)(\d{5})(?!\d)'); if($m.Success){$val=$m.Groups[1].Value}
                        }
                        return @{Has=$has; Val=$val}
                    }
                }
            }
            return @{Has=$has; Val=$val}
        }
        $rows=@()
        if(-not $Pkg){ return $rows }

foreach ($ws in $Pkg.Workbook.Worksheets) {
    if ($ws.Name -eq 'Worksheet Instructions' -or
        $ws.Name -match '(?i)for reference only') {
        continue
    }

            $partLabel=$false; $batchLabel=$false; $cartLabel=$false; $docLabel=$false; $revLabel=$false; $effLabel=$false
            $part=$null; $batch=$null; $cart=$null; $doc=$null; $rev=$null; $eff=$null
            $lr = Get-HeaderFooterText $ws
            $r1 = Try-FromHeaderFooter $lr 'Part';      $partLabel=$partLabel -or $r1.Has; if($r1.Val){$part=$r1.Val}
            $r2 = Try-FromHeaderFooter $lr 'Batch';     $batchLabel=$batchLabel -or $r2.Has; if($r2.Val){$batch=$r2.Val}
            $r3 = Try-FromHeaderFooter $lr 'Cartridge'; $cartLabel=$cartLabel -or $r3.Has; if($r3.Val){$cart=$r3.Val}
            $r4 = Try-FromHeaderFooter $lr 'Doc';       $docLabel=$docLabel -or $r4.Has; if($r4.Val){$doc=$r4.Val}
            $r5 = Try-FromHeaderFooter $lr 'REV';       $revLabel=$revLabel -or $r5.Has; if($r5.Val){$rev=$r5.Val}
            $r6 = Try-FromHeaderFooter $lr 'EFF';       $effLabel=$effLabel -or $r6.Has; if($r6.Val){$eff=$r6.Val}
            if(-not $part){  $c=Try-FromCells $ws $rxPart      30 20 'Part';      $partLabel=$partLabel -or $c.Has; if($c.Val){$part=$c.Val} }
            if(-not $batch){ $c=Try-FromCells $ws $rxBatch     30 20 'Batch';     $batchLabel=$batchLabel -or $c.Has; if($c.Val){$batch=$c.Val} }
            if(-not $cart){  $c=Try-FromCells $ws $rxCartridge 30 20 'Cartridge'; $cartLabel=$cartLabel -or $c.Has; if($c.Val){$cart=$c.Val} }
            if(-not $doc){   $c=Try-FromCells $ws $rxDoc       30 20;             $docLabel=$docLabel -or $c.Has; if($c.Val){$doc=$c.Val} }
            if(-not $rev){   $c=Try-FromCells $ws $rxRev       30 20;             $revLabel=$revLabel -or $c.Has; if($c.Val){$rev=$c.Val} }
            if(-not $eff){   $c=Try-FromCells $ws $rxEff       30 20 'EFF';       $effLabel=$effLabel -or $c.Has; if($c.Val){$eff=$c.Val} }
            $rows += [pscustomobject]@{
                Sheet              = $ws.Name
                PartNo             = $part
                BatchNo            = $batch
                CartridgeNo        = $cart
                DocumentNumber     = $doc
                Rev                = $rev
                Effective          = $eff
                HasPartNoLabel     = [bool]$partLabel
                HasBatchNoLabel    = [bool]$batchLabel
                HasCartridgeLabel  = [bool]$cartLabel
                HasDocumentLabel   = [bool]$docLabel
                HasRevLabel        = [bool]$revLabel
                HasEffectiveLabel  = [bool]$effLabel
            }
        }
        return $rows
    }
}

<#
    .SYNOPSIS
        Extracts pipette and instrument calibration data from a Test Summary sheet.

    .DESCRIPTION
        Scans for pipette and instrument tables (Table 2/3) and returns identifiers
        with calibration due dates. Heavy EPPlus traversal across the worksheet.

    .PARAMETER Worksheet
        EPPlus worksheet representing a Test Summary tab.

    .OUTPUTS
        PSCustomObject with WorksheetName, Pipettes[], Instruments[].
#>
function Get-TestSummaryEquipmentFromWorksheet {
    param(
        [OfficeOpenXml.ExcelWorksheet]$Worksheet
    )

    $result = [pscustomobject]@{
        WorksheetName = $null
        Pipettes      = @()
        Instruments   = @()
    }

    if (-not $Worksheet -or -not $Worksheet.Dimension) {
        return $result
    }

    $result.WorksheetName = $Worksheet.Name

    $maxR = $Worksheet.Dimension.End.Row
    $maxC = $Worksheet.Dimension.End.Column

    # Hjälp: hämta trim:ad text
    $getCellText = {
        param(
            $Ws,
            [int]$Row,
            [int]$Col
        )
        if ($Row -le 0 -or $Col -le 0) { return '' }
        if (-not $Ws.Dimension) { return '' }
        $t = $Ws.Cells[$Row,$Col].Text
        if ($t) { return $t.Trim() }
        return ''
    }

    $pipHeaderRows = @()
    for ($r = 1; $r -le $maxR; $r++) {
        $txt = & $getCellText $Worksheet $r 1
        if (-not $txt) { continue }

        # Ex: "Table 2; Pipette Information"
        if ($txt -match '(?i)table\s*2\b.*pipett(e)?\b') {
            $pipHeaderRows += $r
        }
    }

    $pipPairs = @()  # @{ IdRow = <row>; CalRow = <row> }

    foreach ($hdrRow in $pipHeaderRows) {
        $searchMax = [Math]::Min($maxR, $hdrRow + 12)

        for ($r = $hdrRow + 1; $r -le $searchMax; $r++) {
            $label = & $getCellText $Worksheet $r 1
            if ($label -match '(?i)cepheid\s*id') {

                # Hitta "Calibration Due Date" strax under
                $calRow = $null
                for ($cRow = $r + 1; $cRow -le [Math]::Min($maxR, $r + 6); $cRow++) {
                    $lab2 = & $getCellText $Worksheet $cRow 1
                    if ($lab2 -match '(?i)calibration\s+due\s+date') {
                        $calRow = $cRow
                        break
                    }
                }

                if ($calRow) {
                    $pipPairs += [pscustomobject]@{
                        IdRow  = $r
                        CalRow = $calRow
                    }
                }
            }
        }
    }

    $pipettes = @()

    # Regex för pipett-ID: "Nr. 1", "nr 7", "no 26", "nr.3" etc.
    $pipetteIdRegex = '(?i)\b(?:nr|no)\.?\s*\d+\b'

    foreach ($pair in $pipPairs) {
        for ($c = 2; $c -le $maxC; $c += 2) {
            $idTxt = & $getCellText $Worksheet $pair.IdRow $c
            if (-not $idTxt) { continue }

            if ($idTxt -match '^(?i)N/?A$') { continue }
            if ($idTxt -notmatch $pipetteIdRegex) { continue }

            $dueCell = $Worksheet.Cells[$pair.CalRow,$c]
            $dueVal  = $dueCell.Value
            $dueOut  = $null

            if ($dueVal -is [datetime]) {
                $dueOut = $dueVal
            }
            elseif ($dueVal -is [double] -or $dueVal -is [int]) {
                try {
                    $base   = Get-Date '1899-12-30'
                    $dueOut = $base.AddDays([double]$dueVal)
                } catch {
                    $dueOut = ($dueCell.Text + '').Trim()
                }
            }
            else {
                $txt = ($dueCell.Text + '').Trim()
                if ($txt) { $dueOut = $txt }
            }

            $pipettes += [pscustomobject]@{
                Id                 = $idTxt
                CalibrationDueDate = $dueOut
            }
        }
    }

    if ($pipettes -and $pipettes.Count -gt 0) {
        $uniq = @()
        $seen = @{}
        foreach ($p in $pipettes) {
            if (-not $p.Id) { continue }
            if ($seen.ContainsKey($p.Id)) { continue }
            $seen[$p.Id] = $true
            $uniq += $p
        }
        $pipettes = $uniq
    }

    $result.Pipettes = $pipettes

    $instHeaderRows = @()
    for ($r = 1; $r -le $maxR; $r++) {
        $txt = & $getCellText $Worksheet $r 1
        if (-not $txt) { continue }

        # Bred match: "Table 3; GeneXperts/Infinitys Used", "Table 3; GeneXperts Used:" etc.
        if ($txt -match '(?i)table\s*3\b.*(gene\s*experts?|genexperts?|genex?perts?|infinitys?)') {
            $instHeaderRows += $r
        }
    }

    $instPairs = @()

    foreach ($hdrRow in $instHeaderRows) {
        $searchMax = [Math]::Min($maxR, $hdrRow + 20)

        for ($r = $hdrRow + 1; $r -le $searchMax; $r++) {
            $label = & $getCellText $Worksheet $r 1
            if ($label -match '(?i)cepheid\s*id') {

                $calRow = $null
                for ($cRow = $r + 1; $cRow -le [Math]::Min($maxR, $r + 6); $cRow++) {
                    $lab2 = & $getCellText $Worksheet $cRow 1
                    if ($lab2 -match '(?i)calibration\s+due\s+date') {
                        $calRow = $cRow
                        break
                    }
                }

                if ($calRow) {
                    $instPairs += [pscustomobject]@{
                        IdRow  = $r
                        CalRow = $calRow
                    }
                }
            }
        }
    }

    $insts = @()

    foreach ($pair in $instPairs) {
        for ($c = 2; $c -le $maxC; $c += 2) {
            $idTxt = & $getCellText $Worksheet $pair.IdRow $c
            if (-not $idTxt) { continue }
            if ($idTxt -match '^(?i)N/?A$') { continue }

            # Bara GX/Infinity
            if ($idTxt -notmatch '(?i)^(infinity|gx\s*\d+|gx\d+)') {
                continue
            }

            $dueCell = $Worksheet.Cells[$pair.CalRow,$c]
            $dueVal  = $dueCell.Value
            $dueOut  = $null

            if ($dueVal -is [datetime]) {
                $dueOut = $dueVal
            }
            elseif ($dueVal -is [double] -or $dueVal -is [int]) {
                try {
                    $base   = Get-Date '1899-12-30'
                    $dueOut = $base.AddDays([double]$dueVal)
                } catch {
                    $dueOut = ($dueCell.Text + '').Trim()
                }
            }
            else {
                $txt = ($dueCell.Text + '').Trim()
                if ($txt) { $dueOut = $txt }
            }

            $insts += [pscustomobject]@{
                Id                 = $idTxt
                CalibrationDueDate = $dueOut
            }
        }
    }

    if ($insts -and $insts.Count -gt 0) {
        $uniq = @()
        $seen = @{}
        foreach ($i in $insts) {
            if (-not $i.Id) { continue }
            if ($seen.ContainsKey($i.Id)) { continue }
            $seen[$i.Id] = $true
            $uniq += $i
        }
        $insts = $uniq
    }

    $result.Instruments = $insts

    return $result
}

<#
    .SYNOPSIS
        Finds the best Test Summary sheet containing equipment data.

    .PARAMETER Pkg
        ExcelPackage containing potential Test Summary sheets.

    .OUTPUTS
        PSCustomObject with WorksheetName, Pipettes[], Instruments[] (or empty defaults).
#>
function Get-TestSummaryEquipment {
    param(
        [OfficeOpenXml.ExcelPackage]$Pkg
    )

    $empty = [pscustomobject]@{
        WorksheetName = $null
        Pipettes      = @()
        Instruments   = @()
    }

    if (-not $Pkg) { return $empty }

    $best = $null

    try {
        foreach ($ws in $Pkg.Workbook.Worksheets) {
            if (-not $ws.Dimension) { continue }
            if ($ws.Name -match '(?i)worksheet\s*instructions') { continue }

            $info = Get-TestSummaryEquipmentFromWorksheet -Worksheet $ws
            if (-not $info) { continue }

            $pipCount  = if ($info.Pipettes)   { $info.Pipettes.Count }   else { 0 }
            $instCount = if ($info.Instruments){ $info.Instruments.Count } else { 0 }

            if (-not $best) { $best = $info }

            if ($pipCount -gt 0 -or $instCount -gt 0) {
                if ($ws.Name -match '(?i)test\s*summary') {
                    return $info
                }
                $best = $info
            }
        }
    } catch {
        return $empty
    }

    if ($best) { return $best }
    return $empty
}

if (-not (Get-Command Compare-WorksheetHeaderSet -ErrorAction SilentlyContinue)) {
    <#
        .SYNOPSIS
            Compares header fields across Worksheet tabs to find deviations.

        .PARAMETER Rows
            Collection of header info objects (Sheet, PartNo, BatchNo etc.).

        .OUTPUTS
            PSCustomObject with Issues, Summary, Details string.
    #>
    function Compare-WorksheetHeaderSet {
        param([Parameter(Mandatory)][object[]]$Rows)

        $devIssues  = 0
        $reqIssues  = 0
        $detailsLst = New-Object System.Collections.Generic.List[string]

function _canon([string]$raw, [string]$type) {
    if ([string]::IsNullOrWhiteSpace($raw)) { return $null }
    $txt = Normalize-HeaderText $raw
            switch ($type) {
                'Part'      { $m=[regex]::Match($txt,'(?i)\b(\d{3}-\d{4})\b'); return $(if($m.Success){$m.Groups[1].Value}else{$txt.ToUpper()}) }
                'Batch'     { $m=[regex]::Match($txt,'(?<!\d)(\d{10})(?!\d)');  return $(if($m.Success){$m.Groups[1].Value}else{$txt.ToUpper()}) }
                'Cartridge' { $m=[regex]::Match($txt,'(?<!\d)(\d{5})(?!\d)');   return $(if($m.Success){$m.Groups[1].Value}else{$txt.ToUpper()}) }
                'Doc'       { $m=[regex]::Match($txt,'(?i)\b(D\d+)\b');         return $(if($m.Success){$m.Groups[1].Value.ToUpper()}else{$txt.ToUpper()}) }
                'REV'       { return $txt.ToUpper() }
                'EFF' {
    $dt = $null
    if (Try-Parse-HeaderDate $txt ([ref]$dt)) { return $dt.ToString('yyyy-MM-dd') }
    return $txt
}
                default     { return $txt }
            }
        }
        $keys = @(

            @{K='PartNo';         T='Part';      Required=$true;  Label='HasPartNoLabel'    },
            @{K='BatchNo';        T='Batch';     Required=$true;  Label='HasBatchNoLabel'   },
            @{K='CartridgeNo';    T='Cartridge'; Required=$true;  Label='HasCartridgeLabel' },
            @{K='DocumentNumber'; T='Doc';       Required=$false; Label='HasDocumentLabel'  },
            @{K='Rev';            T='REV';       Required=$false; Label='HasRevLabel'       },
            @{K='Effective';      T='EFF';       Required=$false; Label='HasEffectiveLabel' }
        )

 

        foreach ($entry in $keys) {
            $k=$entry.K; $t=$entry.T; $req=$entry.Required; $labelProp=$entry.Label
            $vals = foreach($r in $Rows){
                [pscustomobject]@{
                    Sheet  = $r.Sheet
                    Canon  = _canon ($r.$k) $t
                    Raw    = if([string]::IsNullOrWhiteSpace($r.$k)){ $null } else { $r.$k.Trim() }
                    HasLbl = [bool]($r.$labelProp)
                }
            }
            $nonEmpty = $vals | Where-Object { $_.Canon }
            $useSet   = if($nonEmpty.Count -gt 0){$nonEmpty}else{$vals}
            $byVal    = $useSet | Group-Object Canon | Sort-Object Count -Descending
            $major    = if($byVal.Count -gt 0){ $byVal[0].Name } else { $null }
            $deviants = @()
            if($major){
                $deviants = $vals | Where-Object { $_.Canon -and ($_.Canon -ne $major) } | Select-Object -ExpandProperty Sheet -Unique
            }

            if($deviants.Count -gt 0){

                $devIssues++

                $majShow = if($major){"'$major'"}else{'(empty)'}

                $line = "- $k majoritet=$majShow | avvikande flikar: " + ($deviants -join ', ')

                [void]$detailsLst.Add($line)

            }

 

            if($req){

                $missingSheets = $vals | Where-Object { $_.HasLbl -and -not $_.Canon } | Select-Object -ExpandProperty Sheet -Unique

                if($missingSheets.Count -gt 0){

                    $reqIssues++

                    $line = "*Saknas* - $k " + ($missingSheets -join ', ')

                    [void]$detailsLst.Add($line)

                }

            }

        }

 

        $issuesTotal = $devIssues + $reqIssues

        $summary = if($issuesTotal -eq 0){ 'OK' } else { "Avvikelser: $issuesTotal (avvikare=$devIssues, saknas=$reqIssues)" }

 

        return [pscustomobject]@{

            Issues  = $issuesTotal

            Summary = $summary

            Details = ($detailsLst -join "`r`n")

        }

    }

}

if (-not (Get-Command Extract-SealTestHeader -ErrorAction SilentlyContinue)) {

    function Extract-SealTestHeader {

        param([OfficeOpenXml.ExcelPackage]$Pkg)

 

        $result = [pscustomobject]@{ DocumentNumber=$null; Rev=$null; Effective=$null }

        if (-not $Pkg) { return $result }

 

        foreach ($ws in $Pkg.Workbook.Worksheets) {

            if ($ws.Name -eq 'Worksheet Instructions') { continue }

            $right = (($ws.HeaderFooter.OddHeader.RightAlignedText + '') -replace '\r?\n',' ').Trim()

            if (-not $right) { $right = (($ws.HeaderFooter.EvenHeader.RightAlignedText + '') -replace '\r?\n',' ').Trim() }

 

            if (-not $result.DocumentNumber -and $right -match '(?i)\bDocument\s*(?:No|Number|#)\s*[:#]?\s*(D\d+)\b') { $result.DocumentNumber = $matches[1] }

            if (-not $result.Rev            -and $right -match '(?i)\bRev(?:ision)?\.?\s*[:#]?\s*([A-Z]{1,3}(?:\.\d+)?)\b') { $result.Rev = $matches[1] }

            if (-not $result.Effective      -and $right -match '(?i)\bEffective\s*[:#]?\s*([0-9]{1,2}[\/\-][0-9]{1,2}[\/\-][0-9]{4}|[0-9]{4}[\/\-][0-9]{2}[\/\-][0-9]{2})') { $result.Effective = $matches[1] }

 

            if ($result.DocumentNumber -and $result.Rev -and $result.Effective) { break }

        }

        return $result

    }

}

<#
    .SYNOPSIS
        Normalizes identifiers (Batch/Part/Cartridge) to canonical formats.

    .PARAMETER Value
        Raw identifier string.
    .PARAMETER Type
        One of Batch, Part, Cartridge controlling normalization rules.

    .OUTPUTS
        Normalized string or $null.
#>
function Normalize-Id {

    param([string]$Value, [ValidateSet('Batch','Part','Cartridge')] [string]$Type)

    if ([string]::IsNullOrWhiteSpace($Value)) { return $null }

    $v = $Value.Trim()

    switch ($Type) {

        'Batch'     { return ($v -replace '[^\d]', '') }                      # 10 siffror

        'Part'      { return (($v -replace '[^0-9A-Za-z\-]', '')).ToUpper() } # t.ex. 700-5702

        'Cartridge' { return ($v -replace '[^\d]', '') }                      # 5 siffror

        default     { return $v }

    }

}

<#
    .SYNOPSIS
        Cleans header text by removing BOM/line-break artifacts and normalizing spacing.

    .PARAMETER s
        Raw text from Excel header/footer or cells.

    .OUTPUTS
        Trimmed, normalized string.
#>
function Normalize-HeaderText {

    param([string]$s)

    if ([string]::IsNullOrEmpty($s)) { return '' }

    # Ta bort BOM/zero-width

    $s = $s -replace "[\uFEFF\u200B]", ""

    # Ta bort inbäddade radbrytningsmarkörer från Excel-export

    # t.ex. 700-4370_x000D_ Batch No(s).: 1001502746_x000D_ ...

    $s = $s -replace "_x000D_", " "

    # NBSP/figur/narrow NBSP → vanligt mellanslag

    $s = $s -replace "[\u00A0\u2007\u202F]", " "

 

    # Normalisera olika streck till '-'

    $s = $s -replace "[\u2013\u2014\u2212]", "-"

 

    # Kollapsa whitespace och trimma

    $s = ($s -replace "\s+", " ").Trim()

    return $s

}

<#
    .SYNOPSIS
        Parses Seal Test/Worksheet header text for document metadata.

    .PARAMETER Raw
        Concatenated header/footer text.

    .OUTPUTS
        PSCustomObject with PartNo, BatchNo, CartridgeNo, DocumentNumber, Attachment, Rev, Effective.
#>
function Parse-WorksheetHeaderRaw {

    param(

        [string]$Raw

    )

 

    $obj = [pscustomobject]@{

        PartNo         = $null

        BatchNo        = $null

        CartridgeNo    = $null

        DocumentNumber = $null

        Attachment     = $null

        Rev            = $null

        Effective      = $null

    }

 

    if (-not $Raw) { return $obj }

 

    # Grundstädning

    $txt = $Raw + ''

 

    # Ta bort header-koder (sidnr m.m.)

    $txt = $txt -replace '&P', ''

    $txt = $txt -replace '&N', ''

    $txt = $txt -replace '&L', ' '

 

    # Normalisera (tar också bort _x000D_)

    $txt = Normalize-HeaderText $txt

 

    # Document Number + ev. Attachment

    $m = [regex]::Match(

        $txt,

        '(?i)\bDocument\s*Number\b\s*[:#]?\s*(D\d{5})(?:\s+Attachment[:\s]+([0-9A-Za-z\.]+))?'

    )

    if ($m.Success) {

        $obj.DocumentNumber = $m.Groups[1].Value.Trim()

        if ($m.Groups.Count -ge 3 -and $m.Groups[2].Value) {

            $obj.Attachment = $m.Groups[2].Value.Trim()

        }

    }

 

    # Fristående "Attachment" om det inte fångades ovan

    if (-not $obj.Attachment) {

        $m = [regex]::Match($txt, '(?i)\bAttachment\b\s*:?\s*([0-9A-Za-z\.]+)')

        if ($m.Success) {

            $obj.Attachment = $m.Groups[1].Value.Trim()

        }

    }

 

    # Rev (t.ex. F, AW, E.1)

    $m = [regex]::Match($txt, '(?i)\bRev(?:ision)?\b\s*:?\s*([A-Z0-9]+(?:\.[0-9]+)?)')

    if ($m.Success) {

        $obj.Rev = $m.Groups[1].Value.Trim()

    }

 

    # Effective (t.ex. 4/14/2025 eller 09/20/2023)

    $m = [regex]::Match($txt, '(\d{1,2}/\d{1,2}/\d{4})')

    if ($m.Success) {

        $obj.Effective = $m.Groups[1].Value.Trim()

    }

 

    # Part No (t.ex. 700-4370)

    $m = [regex]::Match($txt, '(?i)Part\s*No\.\s*:\s*([0-9A-Z\-]+)')

    if ($m.Success) {

        $obj.PartNo = $m.Groups[1].Value.Trim()

    }

 

    # Batch No (t.ex. 1001502746)

    $m = [regex]::Match($txt, '(?i)Batch\s*No\(s\)\.\s*:\s*(\d{6,12})')

    if ($m.Success) {

        $obj.BatchNo = $m.Groups[1].Value.Trim()

    }

 

    # Cartridge (LSP, t.ex. 78704)

    $m = [regex]::Match($txt, '(?i)Cartridge\s*No\.\s*\(LSP\)\s*:\s*(\d{5,7})')

    if ($m.Success) {

        $obj.CartridgeNo = $m.Groups[1].Value.Trim()

    }

 

    return $obj

}

<#
    .SYNOPSIS
        Attempts to parse dates from mixed Excel header values.

    .PARAMETER cellValue
        Incoming value (numeric OADate, DateTime, or string).
    .PARAMETER out
        [ref] output DateTime when parse succeeds.

    .OUTPUTS
        Boolean indicating parse success.
#>
function Try-Parse-HeaderDate {

    param([object]$cellValue, [ref]$out)

    $out.Value = $null

    if ($null -eq $cellValue) { return $false }

 

    if ($cellValue -is [double] -or $cellValue -is [int]) {

        try { $out.Value = [DateTime]::FromOADate([double]$cellValue); return $true } catch {}

    }

    if ($cellValue -is [DateTime]) { $out.Value = [DateTime]$cellValue; return $true }

 

    $s = Normalize-HeaderText ([string]$cellValue)

    if ([string]::IsNullOrWhiteSpace($s)) { return $false }

 

    $cultures = @('sv-SE','en-US','en-GB')

    foreach ($c in $cultures) {

        try {

            $dt = [DateTime]::Parse($s, [Globalization.CultureInfo]::GetCultureInfo($c), [Globalization.DateTimeStyles]::AssumeLocal)

            $out.Value = $dt; return $true

        } catch {}

    }

    $formats = @(

        'yyyy-MM-dd','yyyy/MM/dd','dd/MM/yyyy','d/M/yyyy','MM/dd/yyyy','M/d/yyyy',

        'dd-MM-yyyy','d-M-yyyy','dd.M.yyyy','d.M.yyyy','dd-MMM-yyyy','MMM-yy'

    )

    foreach ($fmt in $formats) {

        try {

            $dt = [DateTime]::ParseExact($s, $fmt, [Globalization.CultureInfo]::InvariantCulture, [Globalization.DateTimeStyles]::AssumeLocal)

            $out.Value = $dt; return $true

        } catch {}

    }

    return $false

}

<#
    .SYNOPSIS
        Derives a consensus identifier across Worksheet/NEG/POS sources.

    .PARAMETER Ws
        Value from Worksheet.
    .PARAMETER Pos
        Value from Seal Test POS.
    .PARAMETER Neg
        Value from Seal Test NEG.
    .PARAMETER Type
        Identifier type (Batch/Part/Cartridge) controlling normalization.

    .OUTPUTS
        Hashtable with Value, Source (WS/POS/NEG), and Note.
#>
function Get-ConsensusValue {

    param(

        [string]$Ws,  [string]$Pos, [string]$Neg,

        [ValidateSet('Batch','Part','Cartridge')] [string]$Type

    )

 

    $nWs  = Normalize-Id $Ws  $Type

    $nPos = Normalize-Id $Pos $Type

    $nNeg = Normalize-Id $Neg $Type

 

    $present = @{}

    if ($nWs)  { $present['WS']  = $nWs  }

    if ($nPos) { $present['POS'] = $nPos }

    if ($nNeg) { $present['NEG'] = $nNeg }

 

    $chosenNorm = $null

    $sources    = @()

 

    $counts = @{}

    foreach ($k in $present.Keys) {

        $val = $present[$k]

        if (-not $counts.ContainsKey($val)) { $counts[$val] = 0 }

        $counts[$val]++

    }

    $maxCount = 0; $maxVal = $null

    foreach ($kv in $counts.GetEnumerator()) {

        if ($kv.Value -gt $maxCount) { $maxCount = $kv.Value; $maxVal = $kv.Key }

    }

    if ($maxCount -ge 2) {

        $chosenNorm = $maxVal

        foreach ($k in $present.Keys) { if ($present[$k] -eq $chosenNorm) { $sources += $k } }

    }

 

    if (-not $chosenNorm -and $nWs) {

        $posMatch = ($nPos -and $nPos -eq $nWs)

        $negMatch = ($nNeg -and $nNeg -eq $nWs)

        if (($posMatch -and -not $negMatch) -or ($negMatch -and -not $posMatch)) {

            $chosenNorm = $nWs

            $sources = @('WS')

            if ($posMatch) { $sources += 'POS' }

            if ($negMatch) { $sources += 'NEG' }

        }

    }

 

    if (-not $chosenNorm -and $nPos -and $nNeg -and $nPos -eq $nNeg) {

        $chosenNorm = $nPos

        $sources = @('POS','NEG')

    }

 

    if (-not $chosenNorm) {

        if ($nWs)      { $chosenNorm = $nWs;  $sources = @('WS')  }

        elseif ($nPos) { $chosenNorm = $nPos; $sources = @('POS') }

        elseif ($nNeg) { $chosenNorm = $nNeg; $sources = @('NEG') }

    }

 

    $orig = @{ WS=$Ws; POS=$Pos; NEG=$Neg }

    $pretty = $null

    foreach ($pref in @('WS','POS','NEG')) {

        if ($present.ContainsKey($pref) -and $present[$pref] -eq $chosenNorm) { $pretty = $orig[$pref]; break }

    }

 

    $note = $null

    if ($sources -and $sources.Count -gt 0) {

        $note = "Consensus: " + ($sources -join '+')

        $others = @()

        foreach ($k in @('WS','POS','NEG')) {

            if ($present.ContainsKey($k) -and ($sources -notcontains $k)) {

                $val = $orig[$k]

                if (-not [string]::IsNullOrEmpty($val)) { $others += ($k + '=' + $val) }

            }

        }

        if ($others.Count -gt 0) { $note += " | Others: " + ($others -join ', ') }

    }

 

    return @{ Value=$pretty; Source=($sources -join '+'); Note=$note }

}

<#
    .SYNOPSIS
        Adds a styled hyperlink to an Excel cell.

    .PARAMETER Cell
        Target ExcelRange cell.
    .PARAMETER Text
        Display text.
    .PARAMETER Url
        Destination URL.
#>
function Add-Hyperlink {
    param([OfficeOpenXml.ExcelRange]$Cell,[string]$Text,[string]$Url)
    try {
        $Cell.Value = $Text
        $Cell.Hyperlink = [Uri]$Url
        $Cell.Style.Font.UnderLine = $true
        $Cell.Style.Font.Color.SetColor([System.Drawing.Color]::FromArgb(0,102,204))
    } catch {}
}

<#
    .SYNOPSIS
        Finds the first cell matching a regex within a bounded range.

    .PARAMETER Ws
        Worksheet to scan.
    .PARAMETER Rx
        Regex to match against normalized cell text.
    .PARAMETER MaxRows
        Maximum rows to scan (default 200).
    .PARAMETER MaxCols
        Maximum columns to scan (default 40).

    .OUTPUTS
        Hashtable with Row/Col/Text or $null if not found.
#>
function Find-RegexCell {
    param([OfficeOpenXml.ExcelWorksheet]$Ws,[regex]$Rx,[int]$MaxRows=200,[int]$MaxCols=40)
    if (-not $Ws -or -not $Ws.Dimension) { return $null }
    $rMax = [Math]::Min($Ws.Dimension.End.Row, $MaxRows)
    $cMax = [Math]::Min($Ws.Dimension.End.Column, $MaxCols)
    for ($r=1; $r -le $rMax; $r++) {
        for ($c=1; $c -le $cMax; $c++) {
            $t = Normalize-HeaderText ($Ws.Cells[$r,$c].Text + '')
            if ($t -and $Rx.IsMatch($t)) { return @{Row=$r;Col=$c;Text=$t} }
        }
    }
    return $null
}

<#
    .SYNOPSIS
        Extracts document number and revision from Seal Test header/footer.

    .PARAMETER Pkg
        ExcelPackage for Seal Test workbook.

    .OUTPUTS
        PSCustomObject with Raw, DocNo, Rev.
#>
function Get-SealHeaderDocInfo {
    param([OfficeOpenXml.ExcelPackage]$Pkg)
    $result = [pscustomobject]@{ Raw=''; DocNo=''; Rev='' }
    if (-not $Pkg) { return $result }
    $ws = $Pkg.Workbook.Worksheets | Where-Object { $_.Name -ne 'Worksheet Instructions' } | Select-Object -First 1
    if (-not $ws) { return $result }
    try {
        $lt = ($ws.HeaderFooter.OddHeader.LeftAlignedText + '').Trim()
        if (-not $lt) { $lt = ($ws.HeaderFooter.EvenHeader.LeftAlignedText + '').Trim() }
        $result.Raw = $lt
        $rx = [regex]'(?i)(?:document\s*(?:no|nr|#|number)\s*[:#]?\s*([A-Z0-9\-_\.\/]+))?.*?(?:rev(?:ision)?\.?\s*[:#]?\s*([A-Z0-9\-_\.]+))?'
        $m = $rx.Match($lt)
        if ($m.Success) {
            if ($m.Groups[1].Value) { $result.DocNo = $m.Groups[1].Value.Trim() }
            if ($m.Groups[2].Value) { $result.Rev   = $m.Groups[2].Value.Trim() }
        }
    } catch {}
    return $result
}

<#
    .SYNOPSIS
        Locates a label in column A and returns its row index.

    .PARAMETER Ws
        Worksheet to inspect.
    .PARAMETER Label
        Label text to search for (case-insensitive).

    .OUTPUTS
        Row index or $null if not found.
#>
function Find-InfoRow {
    param([OfficeOpenXml.ExcelWorksheet]$Ws, [string]$Label)
    if (-not $Ws -or -not $Ws.Dimension) { return $null }
    $maxRow = [Math]::Min($Ws.Dimension.End.Row, 300)
    for ($ri=1; $ri -le $maxRow; $ri++) {
        $txt = (($Ws.Cells[$ri,1].Text) + '').Trim()
        if (-not $txt) { continue }
        if ($txt.ToLowerInvariant() -eq $Label.ToLowerInvariant()) { return $ri }
    }
    return $null
}

<#
    .SYNOPSIS
        Finds the value to the right of a label cell within a bounded range.

    .PARAMETER Ws
        Worksheet to scan.
    .PARAMETER Label
        Label text to normalize and search for.
    .PARAMETER MaxRows
        Max rows to scan (default 200).
    .PARAMETER MaxCols
        Max columns to scan (default 120).

    .OUTPUTS
        String cell text to the right of the matched label, or $null.
#>
function Find-LabelValueRightward {
    param(
        [OfficeOpenXml.ExcelWorksheet]$Ws,
        [string]$Label,
        [int]$MaxRows = 200,
        [int]$MaxCols = 120
    )
    $normLbl = Normalize-HeaderText $Label
    $pat = '^(?i)\s*' + [regex]::Escape($normLbl).Replace('\ ', '\s*') + '\s*[:\.]*\s*$'
    $rx  = [regex]::new($pat, [Text.RegularExpressions.RegexOptions]::IgnoreCase)
    $hit = Find-RegexCell -Ws $Ws -Rx $rx -MaxRows $MaxRows -MaxCols $MaxCols
    if (-not $hit) { return $null }
    $cMax = [Math]::Min($Ws.Dimension.End.Column, $MaxCols)
    for ($c = $hit.Col + 1; $c -le $cMax; $c++) {
        $t = Normalize-HeaderText ($Ws.Cells[$hit.Row,$c].Text + '')
        if ($t) { return $t }
    }
    return $null
}

<#
    .SYNOPSIS
        Converts mixed Excel date representations into a DateTime where possible.

    .PARAMETER Value
        Incoming value (DateTime/double/string).

    .OUTPUTS
        DateTime or original value if conversion fails.
#>
function Convert-ToEqDate {
    param(
        [Parameter(Mandatory = $false)]
        $Value
    )

    if (-not $Value -or $Value -eq 'N/A') {
        return $null
    }

    if ($Value -is [datetime]) {
        return $Value
    }

    if ($Value -is [double] -or $Value -is [int]) {
        try {
            $base = Get-Date '1899-12-30'
            return $base.AddDays([double]$Value)
        } catch {
            return $Value
        }
    }

    try {
        return (Get-Date -Date $Value -ErrorAction Stop)
    } catch {
        return $Value
    }
}


<#
    .SYNOPSIS
        Builds the full assay report workbook based on selected source files.

    .DESCRIPTION
        Orchestrates Seal Test parsing, optional signature writing, CSV/Worksheet
        ingestion, Control Material copy, SharePoint enrichment, and saving the
        final Excel report. Behaviour mirrors the legacy button click logic.

    .PARAMETER CsvPath
        Optional Tests Summary CSV path.
    .PARAMETER SealNegPath
        Seal Test NEG workbook path.
    .PARAMETER SealPosPath
        Seal Test POS workbook path.
    .PARAMETER WorksheetPath
        Optional LSP Worksheet path.
    .PARAMETER Lsp
        LSP identifier.
    .PARAMETER SignerText
        Signature text to write (if SignSealTest is true).
    .PARAMETER SignSealTest
        Toggle to write signature into Seal Test sheets.
    .PARAMETER OverwriteSignature
        Toggle to overwrite existing signature cells.
    .PARAMETER IncludeSharePoint
        Toggle to attempt SharePoint enrichment.
    .PARAMETER SaveInLsp
        Toggle to save report in LSP folder (defaults to TEMP for regulated flow).
    .PARAMETER SharePointLinkLabel
        Optional status label to update with SharePoint link metadata.

    .OUTPUTS
        String path to the generated report (or $null on failure).

    .NOTES
        Heavy operation – runs the entire pipeline previously embedded in Main.ps1.
#>
function Invoke-AssayReportBuild {
    param(
        [string]$CsvPath,
        [Parameter(Mandatory)][string]$SealNegPath,
        [Parameter(Mandatory)][string]$SealPosPath,
        [string]$WorksheetPath,
        [string]$Lsp,
        [string]$SignerText,
        [bool]$SignSealTest,
        [bool]$OverwriteSignature,
        [bool]$IncludeSharePoint = $true,
        [bool]$SaveInLsp = $false,
        [System.Windows.Forms.ToolStripStatusLabel]$SharePointLinkLabel
    )
    if (-not (Assert-StartupReady)) { return }
    Gui-Log 'Skapar rapport…' -Immediate
    try {
    if (-not (Load-EPPlus)) { Gui-Log "❌ EPPlus kunde inte laddas – avbryter." 'Error'; return }
     
    $selCsv = $CsvPath
    $selNeg = $SealNegPath
    $selPos = $SealPosPath
    
    if (-not $selNeg -or -not $selPos) { Gui-Log "❌ Du måste välja en Seal NEG och en Seal POS." 'Error'; return }
    $lsp = ($Lsp + '').Trim()
    if (-not $lsp) { Gui-Log "⚠️ Ange ett LSP-nummer." 'Warn'; return }
    
    Gui-Log "📄 Neg-fil: $(Split-Path $selNeg -Leaf)" 'Info'
    Gui-Log "📄 Pos-fil: $(Split-Path $selPos -Leaf)" 'Info'
    if ($selCsv) { Gui-Log "📄 CSV: $(Split-Path $selCsv -Leaf)" 'Info' } else { Gui-Log "ℹ️ Ingen CSV vald." 'Info' }
    
    $negWritable = $true; $posWritable = $true
    if ($SignSealTest) {
        $negWritable = -not (Test-FileLocked $selNeg); if (-not $negWritable) { Gui-Log "🔒 NEG är låst (öppen i Excel?)." 'Warn' }
        $posWritable = -not (Test-FileLocked $selPos); if (-not $posWritable) { Gui-Log "🔒 POS är låst (öppen i Excel?)." 'Warn' }
    }
    $pkgNeg = $null; $pkgPos = $null; $pkgOut = $null
    try {
        try {
            $pkgNeg = New-Object OfficeOpenXml.ExcelPackage (New-Object IO.FileInfo($selNeg))
            $pkgPos = New-Object OfficeOpenXml.ExcelPackage (New-Object IO.FileInfo($selPos))
        } catch {
            Gui-Log "❌ Kunde inte öppna NEG/POS: $($_.Exception.Message)" 'Error'
            return
        }
     
        $templatePath = $Config.TemplatePath
        if (-not (Test-Path -LiteralPath $templatePath)) { Gui-Log "❌ Mallfilen 'output_template-v4.xlsx' saknas!" 'Error'; return }
        try {
            $pkgOut = New-Object OfficeOpenXml.ExcelPackage (New-Object IO.FileInfo($templatePath))
        } catch {
            Gui-Log "❌ Kunde inte läsa mall: $($_.Exception.Message)" 'Error'
            return
        }
    
        # ============================
        # === SIGNATUR I NEG/POS  ====
        # ============================
    
        $signToWrite = ($SignerText + '').Trim()
        if ($SignSealTest) {
            if (-not $signToWrite) { Gui-Log "❌ Ingen signatur angiven (B47). Avbryter."; return }
            if (-not (Confirm-SignatureInput -Text $signToWrite)) { Gui-Log "🛑 Signatur ej bekräftad. Avbryter."; return }
    
            $negWritten = 0; $posWritten = 0; $negSkipped = 0; $posSkipped = 0
            foreach ($ws in $pkgNeg.Workbook.Worksheets) {
                if ($ws.Name -eq 'Worksheet Instructions') { continue }
                $h3 = ($ws.Cells['H3'].Text + '').Trim()
                if ($h3 -match '^[0-9]') {
                    $existing = ($ws.Cells['B47'].Text + '').Trim()
                    if ($existing -and -not $OverwriteSignature) { $negSkipped++; continue }
                    $ws.Cells['B47'].Style.Numberformat.Format = '@'
                    $ws.Cells['B47'].Value = $signToWrite
                    $negWritten++
    
                } elseif ([string]::IsNullOrWhiteSpace($h3) -or $h3 -match '^(?i)(N\/\?A|NA|Tomt( innehåll)?)$') {
                    break
                }
            }
            foreach ($ws in $pkgPos.Workbook.Worksheets) {
                if ($ws.Name -eq 'Worksheet Instructions') { continue }
                $h3 = ($ws.Cells['H3'].Text + '').Trim()
                if ($h3 -match '^[0-9]') {
                    $existing = ($ws.Cells['B47'].Text + '').Trim()
                    if ($existing -and -not $OverwriteSignature) { $posSkipped++; continue }
                    $ws.Cells['B47'].Style.Numberformat.Format = '@'
                    $ws.Cells['B47'].Value = $signToWrite
                    $posWritten++
                } elseif ([string]::IsNullOrWhiteSpace($h3) -or $h3 -match '^(?i)(N\/\?A|NA|Tomt( innehåll)?)$') {
                    break
                }
            }
            try {
                if ($negWritten -eq 0 -and $negSkipped -eq 0 -and $posWritten -eq 0 -and $posSkipped -eq 0) {
                    Gui-Log "ℹ️ Inga databladsflikar efter flik 1 att sätta signatur i (ingen åtgärd)."
                } else {
                    if ($negWritten -gt 0 -and $negWritable) { $pkgNeg.Save() } elseif ($negWritten -gt 0) { Gui-Log "🔒 Kunde inte spara NEG (låst)." 'Warn' }
                    if ($posWritten -gt 0 -and $posWritable) { $pkgPos.Save() } elseif ($posWritten -gt 0) { Gui-Log "🔒 Kunde inte spara POS (låst)." 'Warn' }
                    Gui-Log "🖊️ Signatur satt: NEG $negWritten blad (överhoppade $negSkipped), POS $posWritten blad (överhoppade $posSkipped)."
                }
            } catch {
                Gui-Log "⚠️ Kunde inte spara signatur i NEG/POS: $($_.Exception.Message)" 'Warn'
            }
        }
    
        # ============================
        # === CSV (Info/Control)  ====
        # ============================
    
        $csvRows = @(); $runAssay = $null
        if ($selCsv) {
            try { $csvRows = Import-CsvRows -Path $selCsv -StartRow 10 } catch {}
            try { $runAssay = Get-AssayFromCsv -Path $selCsv -StartRow 10 } catch {}
            if ($runAssay) { Gui-Log "🔎 Assay från CSV: $runAssay" }
        }
        $controlTab = $null
        if ($runAssay) { $controlTab = Get-ControlTabName -AssayName $runAssay }
        if ($controlTab) { Gui-Log "🧪 Control Material-flik: $controlTab" } else { Gui-Log "ℹ️ Ingen control-mappning (fortsätter utan)." }
    
        # ============================
        # === Läs avvikelser       ===
        # ============================
    
        $violationsNeg = @(); $violationsPos = @(); $failNegCount = 0; $failPosCount = 0
         foreach ($ws in $pkgNeg.Workbook.Worksheets) {
            if ($ws.Name -eq "Worksheet Instructions") { continue }
            if (-not $ws.Dimension) { continue }
            $obsC = Find-ObservationCol $ws
            for ($r = 3; $r -le 45; $r++) {
                $valK = $ws.Cells["K$r"].Value; $textL = $ws.Cells["L$r"].Text
                if ($valK -ne $null -and $valK -is [double]) {
                    if ($textL -eq "FAIL" -or $valK -le -3.0) {
                        $obsTxt = $ws.Cells[$r, $obsC].Text
                        $violationsNeg += [PSCustomObject]@{
                            Sheet      = $ws.Name
                            Cartridge  = $ws.Cells["H$r"].Text
                            InitialW   = $ws.Cells["I$r"].Value
                            FinalW     = $ws.Cells["J$r"].Value
                            WeightLoss = $valK
                            Status     = if ($textL -eq "FAIL") { "FAIL" } else { "Minusvärde" }
                            Obs        = $obsTxt
                        }
                        if ($textL -eq "FAIL") { $failNegCount++ }
                    }
                }
            }
        }
         foreach ($ws in $pkgPos.Workbook.Worksheets) {
            if ($ws.Name -eq "Worksheet Instructions") { continue }
            if (-not $ws.Dimension) { continue }
            $obsC = Find-ObservationCol $ws
            for ($r = 3; $r -le 45; $r++) {
                $valK = $ws.Cells["K$r"].Value; $textL = $ws.Cells["L$r"].Text
                if ($valK -ne $null -and $valK -is [double]) {
                    if ($textL -eq "FAIL" -or $valK -le -3.0) {
                        $obsTxt = $ws.Cells[$r, $obsC].Text
                        $violationsPos += [PSCustomObject]@{
                            Sheet      = $ws.Name
                            Cartridge  = $ws.Cells["H$r"].Text
                            InitialW   = $ws.Cells["I$r"].Value
                            FinalW     = $ws.Cells["J$r"].Value
                            WeightLoss = $valK
                            Status     = if ($textL -eq "FAIL") { "FAIL" } else { "Minusvärde" }
                            Obs        = $obsTxt
                        }
                        if ($textL -eq "FAIL") { $failPosCount++ }
                    }
                }
            }
        }
     
        # ============================
        # === Seal Test Info (blad) ==
        # ============================
    
        $wsOut1 = $pkgOut.Workbook.Worksheets["Seal Test Info"]
        if (-not $wsOut1) { Gui-Log "❌ Fliken 'Seal Test Info' saknas i mallen"; return }
    
        for ($row = 3; $row -le 15; $row++) {
            $wsOut1.Cells["D$row"].Value = $null
            try { $wsOut1.Cells["D$row"].Style.Fill.PatternType = [OfficeOpenXml.Style.ExcelFillStyle]::None } catch {}
        }
    
        $fields = @(
            @{ Label = "ROBAL";                         Cell = "F2"  }
            @{ Label = "Part Number";                   Cell = "B2"  }
            @{ Label = "Batch Number";                  Cell = "D2"  }
            @{ Label = "Cartridge Number (LSP)";        Cell = "B6"  }
            @{ Label = "PO Number";                     Cell = "B10" }
            @{ Label = "Assay Family";                  Cell = "D10" }
            @{ Label = "Weight Loss Spec";              Cell = "F10" }
            @{ Label = "Balance ID Number";             Cell = "B14" }
            @{ Label = "Balance Cal Due Date";          Cell = "D14" }
            @{ Label = "Vacuum Oven ID Number";         Cell = "B20" }
            @{ Label = "Vacuum Oven Cal Due Date";      Cell = "D20" }
            @{ Label = "Timer ID Number";               Cell = "B25" }
            @{ Label = "Timer Cal Due Date";            Cell = "D25" }
        )
    
        $forceText = @("ROBAL","Part Number","Batch Number","Cartridge Number (LSP)","PO Number","Assay Family","Balance ID Number","Vacuum Oven ID Number","Timer ID Number")
        $mismatchFields = $fields[0..6] | ForEach-Object { $_.Label }
    
        $row = 3
        foreach ($f in $fields) {
            $valNeg=''; $valPos=''
             foreach ($wsN in $pkgNeg.Workbook.Worksheets) {
                if ($wsN.Name -eq "Worksheet Instructions") { continue }
                $cell = $wsN.Cells[$f.Cell]
                if ($cell.Value -ne $null) { if ($cell.Value -is [datetime]) { $valNeg = $cell.Value.ToString('MMM-yy') } else { $valNeg = $cell.Text }; break }
            }
             foreach ($wsP in $pkgPos.Workbook.Worksheets) {
                if ($wsP.Name -eq "Worksheet Instructions") { continue }
                $cell = $wsP.Cells[$f.Cell]
                if ($cell.Value -ne $null) { if ($cell.Value -is [datetime]) { $valPos = $cell.Value.ToString('MMM-yy') } else { $valPos = $cell.Text }; break }
            }
    
            if ($forceText -contains $f.Label) {
                $wsOut1.Cells["B$row"].Style.Numberformat.Format = '@'
                $wsOut1.Cells["C$row"].Style.Numberformat.Format = '@'
            }
    
            $wsOut1.Cells["B$row"].Value = $valNeg
            $wsOut1.Cells["C$row"].Value = $valPos
            $wsOut1.Cells["B$row"].Style.Border.Right.Style = "Medium"
            $wsOut1.Cells["C$row"].Style.Border.Left.Style  = "Medium"
    
            if ($mismatchFields -contains $f.Label -and $valNeg -ne $valPos) {
                $wsOut1.Cells["D$row"].Value = "Mismatch"
    
                Style-Cell $wsOut1.Cells["D$row"] $true "FF0000" "Medium" "FFFFFF"
                Gui-Log "⚠️ Mismatch: $($f.Label) ($valNeg vs $valPos)"
            }
            $row++
        }
    
        # ============================
        # === Testare (B43)        ===
        # ============================
    
        $testersNeg = @(); $testersPos = @()
        foreach ($s in $pkgNeg.Workbook.Worksheets | Where-Object { $_.Name -ne "Worksheet Instructions" }) { $t=$s.Cells["B43"].Text; if ($t) { $testersNeg += ($t -split ",") } }
        foreach ($s in $pkgPos.Workbook.Worksheets | Where-Object { $_.Name -ne "Worksheet Instructions" }) { $t=$s.Cells["B43"].Text; if ($t) { $testersPos += ($t -split ",") } }
        $testersNeg = $testersNeg | ForEach-Object { $_.Trim() } | Where-Object { $_ } | Sort-Object -Unique
        $testersPos = $testersPos | ForEach-Object { $_.Trim() } | Where-Object { $_ } | Sort-Object -Unique
    
        $wsOut1.Cells["B16"].Value = "Name of Tester"
        $wsOut1.Cells["B16:C16"].Merge = $true
        $wsOut1.Cells["B16"].Style.HorizontalAlignment = "Center"
    
        $maxTesters = [Math]::Max($testersNeg.Count, $testersPos.Count)
        $initialRows = 11
        if ($maxTesters -lt $initialRows) { $wsOut1.DeleteRow(17 + $maxTesters, $initialRows - $maxTesters) }
        if ($maxTesters -gt $initialRows) {
            $rowsToAdd = $maxTesters - $initialRows
            $lastRow = 16 + $initialRows
            for ($i = 1; $i -le $rowsToAdd; $i++) { $wsOut1.InsertRow($lastRow + 1, 1, $lastRow) }
        }
        for ($i = 0; $i -lt $maxTesters; $i++) {
            $rowIndex = 17 + $i
            $wsOut1.Cells["A$rowIndex"].Value = $null
            $wsOut1.Cells["B$rowIndex"].Value = if ($i -lt $testersNeg.Count) { $testersNeg[$i] } else { "N/A" }
            $wsOut1.Cells["C$rowIndex"].Value = if ($i -lt $testersPos.Count) { $testersPos[$i] } else { "N/A" }
    
            $topStyle    = if ($i -eq 0) { "Medium" } else { "Thin" }
            $bottomStyle = if ($i -eq $maxTesters - 1) { "Medium" } else { "Thin" }
            foreach ($col in @("B","C")) {
                $cell = $wsOut1.Cells["$col$rowIndex"]
                $cell.Style.Border.Top.Style    = $topStyle
                $cell.Style.Border.Bottom.Style = $bottomStyle
                $cell.Style.Border.Left.Style   = "Medium"
                $cell.Style.Border.Right.Style  = "Medium"
                $cell.Style.Fill.PatternType = "Solid"
                $cell.Style.Fill.BackgroundColor.SetColor([System.Drawing.ColorTranslator]::FromHtml("#CCFFFF"))
            }
        }
    
        # ============================
        # === Signatur-jämförelse  ===
        # ============================
    
        $negSigSet = Get-SignatureSetForDataSheets -Pkg $pkgNeg
        $posSigSet = Get-SignatureSetForDataSheets -Pkg $pkgPos
        $negSet = New-Object 'System.Collections.Generic.HashSet[string]'
        $posSet = New-Object 'System.Collections.Generic.HashSet[string]'
        foreach ($n in $negSigSet.NormSet) { [void]$negSet.Add($n) }
        foreach ($p in $posSigSet.NormSet) { [void]$posSet.Add($p) }
        $hasNeg = ($negSet.Count -gt 0)
        $hasPos = ($posSet.Count -gt 0)
        $onlyNeg = @(); $onlyPos = @(); $sigMismatch = $false
        if ($hasNeg -and $hasPos) {
            foreach ($n in $negSet) { if (-not $posSet.Contains($n)) { $onlyNeg += $n } }
            foreach ($p in $posSet) { if (-not $negSet.Contains($p)) { $onlyPos += $p } }
            $sigMismatch = ($onlyNeg.Count -gt 0 -or $onlyPos.Count -gt 0)
        } else {
            $sigMismatch = $false
        }
    
        $mismatchSheets = @()
        if ($sigMismatch) {
            foreach ($k in $onlyNeg) {
                $raw = if ($negSigSet.RawByNorm.ContainsKey($k)) { $negSigSet.RawByNorm[$k] } else { $k }
                $where = if ($negSigSet.Occ.ContainsKey($k)) { ($negSigSet.Occ[$k] -join ', ') } else { '—' }
                $mismatchSheets += ("NEG: " + $raw + "  [Blad: " + $where + "]")
            }
    
            foreach ($k in $onlyPos) {
                $raw = if ($posSigSet.RawByNorm.ContainsKey($k)) { $posSigSet.RawByNorm[$k] } else { $k }
                $where = if ($posSigSet.Occ.ContainsKey($k)) { ($posSigSet.Occ[$k] -join ', ') } else { '—' }
                $mismatchSheets += ("POS: " + $raw + "  [Blad: " + $where + "]")
            }
            Gui-Log "⚠️ Mismatch: Print Full Name, Sign, and Date (NEG vs POS)"
        }
    
        function Set-MergedWrapAutoHeight {
            param([OfficeOpenXml.ExcelWorksheet]$Sheet,[int]$RowIndex,[int]$ColStart=2,[int]$ColEnd=3,[string]$Text)
            $rng = $Sheet.Cells[$RowIndex, $ColStart, $RowIndex, $ColEnd]
            $rng.Style.WrapText = $true
            $rng.Style.Fill.PatternType = [OfficeOpenXml.Style.ExcelFillStyle]::None
            $Sheet.Row($RowIndex).CustomHeight = $false
    
            try {
                $wChars = [Math]::Floor(($Sheet.Column($ColStart).Width + $Sheet.Column($ColEnd).Width) - 2); if ($wChars -lt 1) { $wChars = 1 }
                $segments = $Text -split "(\r\n|\n|\r)"; $lineCount = 0
                foreach ($seg in $segments) { if (-not $seg) { $lineCount++ } else { $lineCount += [Math]::Ceiling($seg.Length / $wChars) } }
                if ($lineCount -lt 1) { $lineCount = 1 }
                $targetHeight = [Math]::Max(15, [Math]::Ceiling(15 * $lineCount * 2.15))
                if ($Sheet.Row($RowIndex).Height -lt $targetHeight) {
                    $Sheet.Row($RowIndex).Height = $targetHeight
                    $Sheet.Row($RowIndex).CustomHeight = $true
                }
            } catch { $Sheet.Row($RowIndex).CustomHeight = $false }
        }
    
        $signRow = 17 + $maxTesters + 3
        $displaySignNeg = $null; $displaySignPos = $null
        if ($signToWrite) { $displaySignNeg = $signToWrite; $displaySignPos = $signToWrite }
        else {
            $displaySignNeg = if ($negSigSet.RawFirst) { $negSigSet.RawFirst } else { '—' }
            $displaySignPos = if ($posSigSet.RawFirst) { $posSigSet.RawFirst } else { '—' }
        }
    
        $wsOut1.Cells["B$signRow"].Style.Numberformat.Format = '@'
        $wsOut1.Cells["C$signRow"].Style.Numberformat.Format = '@'
        $wsOut1.Cells["B$signRow"].Value = $displaySignNeg
        $wsOut1.Cells["C$signRow"].Value = $displaySignPos
        foreach ($col in @('B','C')) {
    
            $cell = $wsOut1.Cells["${col}$signRow"]
            Style-Cell $cell $false 'CCFFFF' 'Medium' $null
            $cell.Style.HorizontalAlignment = 'Center'
        }
    
        try { $wsOut1.Column(2).Width = 40; $wsOut1.Column(3).Width = 40 } catch {}
        try { $wsOut1.Column(4).Width = 10 } catch {}
    
        if ($sigMismatch) {
            $mismatchCell = $wsOut1.Cells["D$signRow"]
            $mismatchCell.Value = 'Mismatch'
            Style-Cell $mismatchCell $true 'FF0000' 'Medium' 'FFFFFF'
    
            if ($mismatchSheets.Count -gt 0) {
                for ($j = 0; $j -lt $mismatchSheets.Count; $j++) {
                    $rowIdx = $signRow + 1 + $j
                    try { $wsOut1.Cells["B$rowIdx:C$rowIdx"].Merge = $true } catch {}
                    $text = $mismatchSheets[$j]
                    $wsOut1.Cells["B$rowIdx"].Value = $text
                    foreach ($mc in $wsOut1.Cells["B$rowIdx:C$rowIdx"]) { Style-Cell $mc $false 'CCFFFF' 'Medium' $null }
                    $wsOut1.Cells["B$rowIdx:C$rowIdx"].Style.HorizontalAlignment = 'Center'
    
                    if ($text -like 'NEG:*' -or $text -like 'POS:*') {
                        Set-MergedWrapAutoHeight -Sheet $wsOut1 -RowIndex $rowIdx -ColStart 2 -ColEnd 3 -Text $text
                    }
                }
            }
        }
     
        # ============================
        # === STF Sum              ===
        # ============================
    
        $wsOut2 = $pkgOut.Workbook.Worksheets["STF Sum"]
        if (-not $wsOut2) { Gui-Log "❌ Fliken 'STF Sum' saknas i mallen!"; return }
        $totalRows = $violationsNeg.Count + $violationsPos.Count
        $currentRow = 2
    
        if ($totalRows -eq 0) {
            Gui-Log "✅ Seal Test hittades"
            $wsOut2.Cells["B1:H1"].Value = $null
            $wsOut2.Cells["A1"].Value = "Inga STF hittades!"
            Style-Cell $wsOut2.Cells["A1"] $true "D9EAD3" "Medium" "006100"
            $wsOut2.Cells["A1"].Style.HorizontalAlignment = "Left"
            if ($wsOut2.Dimension -and $wsOut2.Dimension.End.Row -gt 1) { $wsOut2.DeleteRow(2, $wsOut2.Dimension.End.Row - 1) }
    
        } else {
            Gui-Log "❗ $failNegCount avvikelser i NEG, $failPosCount i POS"
            $oldDataRows = 0
            if ($wsOut2.Dimension) { $oldDataRows = $wsOut2.Dimension.End.Row - 1; if ($oldDataRows -lt 0) { $oldDataRows = 0 } }
            if ($totalRows -lt $oldDataRows) { $wsOut2.DeleteRow(2 + $totalRows, $oldDataRows - $totalRows) }
            elseif ($totalRows -gt $oldDataRows) { $wsOut2.InsertRow(2 + $oldDataRows, $totalRows - $oldDataRows, 1 + $oldDataRows) }
    
            $currentRow = 2
            foreach ($v in $violationsNeg) {
    
                $wsOut2.Cells["A$currentRow"].Value = "NEG"
                $wsOut2.Cells["B$currentRow"].Value = $v.Sheet
                $wsOut2.Cells["C$currentRow"].Value = $v.Cartridge
                $wsOut2.Cells["D$currentRow"].Value = $v.InitialW
                $wsOut2.Cells["E$currentRow"].Value = $v.FinalW
                $wsOut2.Cells["F$currentRow"].Value = [Math]::Round($v.WeightLoss, 1)
                $wsOut2.Cells["G$currentRow"].Value = $v.Status
                $wsOut2.Cells["H$currentRow"].Value = if ([string]::IsNullOrWhiteSpace($v.Obs)) { 'NA' } else { $v.Obs }
                Style-Cell $wsOut2.Cells["A$currentRow"] $true "B5E6A2" "Medium" $null
                $wsOut2.Cells["C$currentRow:E$currentRow"].Style.Fill.PatternType = "Solid"
                $wsOut2.Cells["C$currentRow:E$currentRow"].Style.Fill.BackgroundColor.SetColor([System.Drawing.ColorTranslator]::FromHtml("#CCFFFF"))
                $wsOut2.Cells["F$currentRow:G$currentRow"].Style.Fill.PatternType = "Solid"
                $wsOut2.Cells["F$currentRow:G$currentRow"].Style.Fill.BackgroundColor.SetColor([System.Drawing.ColorTranslator]::FromHtml("#FFFF99"))
                $wsOut2.Cells["H$currentRow"].Style.Fill.PatternType = "Solid"
                $wsOut2.Cells["H$currentRow"].Style.Fill.BackgroundColor.SetColor([System.Drawing.ColorTranslator]::FromHtml("#D9D9D9"))
     
                if ($v.Status -in @("FAIL","Minusvärde")) {
                    $wsOut2.Cells["F$currentRow"].Style.Font.Bold = $true
                    $wsOut2.Cells["F$currentRow"].Style.Font.Color.SetColor([System.Drawing.Color]::Red)
                    $wsOut2.Cells["G$currentRow"].Style.Font.Bold = $true
                    $wsOut2.Cells["G$currentRow"].Style.Font.Color.SetColor([System.Drawing.Color]::Red)
                }
                Set-RowBorder -ws $wsOut2 -row $currentRow -firstRow 2 -lastRow ($totalRows + 1)
                $currentRow++
            }
            foreach ($v in $violationsPos) {
                $wsOut2.Cells["A$currentRow"].Value = "POS"
                $wsOut2.Cells["B$currentRow"].Value = $v.Sheet
                $wsOut2.Cells["C$currentRow"].Value = $v.Cartridge
                $wsOut2.Cells["D$currentRow"].Value = $v.InitialW
                $wsOut2.Cells["E$currentRow"].Value = $v.FinalW
                $wsOut2.Cells["F$currentRow"].Value = [Math]::Round($v.WeightLoss, 1)
                $wsOut2.Cells["G$currentRow"].Value = $v.Status
                $wsOut2.Cells["H$currentRow"].Value = if ($v.Obs) { $v.Obs } else { 'NA' }
     
                Style-Cell $wsOut2.Cells["A$currentRow"] $true "FFB3B3" "Medium" $null
                $wsOut2.Cells["C$currentRow:E$currentRow"].Style.Fill.PatternType = "Solid"
                $wsOut2.Cells["C$currentRow:E$currentRow"].Style.Fill.BackgroundColor.SetColor([System.Drawing.ColorTranslator]::FromHtml("#CCFFFF"))
                $wsOut2.Cells["F$currentRow:G$currentRow"].Style.Fill.PatternType = "Solid"
                $wsOut2.Cells["F$currentRow:G$currentRow"].Style.Fill.BackgroundColor.SetColor([System.Drawing.ColorTranslator]::FromHtml("#FFFF99"))
                $wsOut2.Cells["H$currentRow"].Style.Fill.PatternType = "Solid"
                $wsOut2.Cells["H$currentRow"].Style.Fill.BackgroundColor.SetColor([System.Drawing.ColorTranslator]::FromHtml("#D9D9D9"))
    
                if ($v.Status -in @("FAIL","Minusvärde")) {
                    $wsOut2.Cells["F$currentRow"].Style.Font.Bold = $true
                    $wsOut2.Cells["F$currentRow"].Style.Font.Color.SetColor([System.Drawing.Color]::Red)
                    $wsOut2.Cells["G$currentRow"].Style.Font.Bold = $true
                    $wsOut2.Cells["G$currentRow"].Style.Font.Color.SetColor([System.Drawing.Color]::Red)
                }
                Set-RowBorder -ws $wsOut2 -row $currentRow -firstRow 2 -lastRow ($totalRows + 1)
                $currentRow++
            }
    
            $wsOut2.Cells.Style.WrapText = $false
            $wsOut2.Cells["A1"].Style.HorizontalAlignment = "Left"
            try { $wsOut2.Cells[2,6,([Math]::Max($currentRow-1,2)),6].Style.Numberformat.Format = '0.0' } catch {}
            if ($wsOut2.Dimension) { $wsOut2.Cells[$wsOut2.Dimension.Address].AutoFitColumns() }
        }
    
    # ============================
    # === Information-blad     ===
    # ============================
    
    try {
    if (-not (Get-Command Add-Hyperlink -ErrorAction SilentlyContinue)) {
        function Add-Hyperlink {
            param([OfficeOpenXml.ExcelRange]$Cell,[string]$Text,[string]$Url)
            try {
                $Cell.Value = $Text
                $Cell.Hyperlink = [Uri]$Url
                $Cell.Style.Font.UnderLine = $true
                $Cell.Style.Font.Color.SetColor([System.Drawing.Color]::FromArgb(0,102,204))
            } catch {}
        }
    }
    if (-not (Get-Command Find-RegexCell -ErrorAction SilentlyContinue)) {
        function Find-RegexCell {
            param([OfficeOpenXml.ExcelWorksheet]$Ws,[regex]$Rx,[int]$MaxRows=200,[int]$MaxCols=40)
            if (-not $Ws -or -not $Ws.Dimension) { return $null }
            $rMax = [Math]::Min($Ws.Dimension.End.Row, $MaxRows)
            $cMax = [Math]::Min($Ws.Dimension.End.Column, $MaxCols)
            for ($r=1; $r -le $rMax; $r++) {
                for ($c=1; $c -le $cMax; $c++) {
            $t = Normalize-HeaderText ($Ws.Cells[$r,$c].Text + '')
            if ($t -and $Rx.IsMatch($t)) { return @{Row=$r;Col=$c;Text=$t} }
                }
            }
            return $null
        }
    }
     
    if (-not (Get-Command Get-SealHeaderDocInfo -ErrorAction SilentlyContinue)) {
        function Get-SealHeaderDocInfo {
            param([OfficeOpenXml.ExcelPackage]$Pkg)
            $result = [pscustomobject]@{ Raw=''; DocNo=''; Rev='' }
            if (-not $Pkg) { return $result }
            $ws = $Pkg.Workbook.Worksheets | Where-Object { $_.Name -ne 'Worksheet Instructions' } | Select-Object -First 1
            if (-not $ws) { return $result }
            try {
                $lt = ($ws.HeaderFooter.OddHeader.LeftAlignedText + '').Trim()
                if (-not $lt) { $lt = ($ws.HeaderFooter.EvenHeader.LeftAlignedText + '').Trim() }
                $result.Raw = $lt
                $rx = [regex]'(?i)(?:document\s*(?:no|nr|#|number)\s*[:#]?\s*([A-Z0-9\-_\.\/]+))?.*?(?:rev(?:ision)?\.?\s*[:#]?\s*([A-Z0-9\-_\.]+))?'
                $m = $rx.Match($lt)
                if ($m.Success) {
                    if ($m.Groups[1].Value) { $result.DocNo = $m.Groups[1].Value.Trim() }
                    if ($m.Groups[2].Value) { $result.Rev   = $m.Groups[2].Value.Trim() }
                }
            } catch {}
            return $result
        }
    }
    
    $wsInfo = $pkgOut.Workbook.Worksheets['Information']
    if (-not $wsInfo) {
        $wsInfo = $pkgOut.Workbook.Worksheets.Add('Information')
    }
    try { $wsInfo.Cells.Style.Font.Name='Arial'; $wsInfo.Cells.Style.Font.Size=11 } catch {}
    try {
        $csvLines = $null
        $csvStats = $null
        if ($selCsv -and (Test-Path -LiteralPath $selCsv)) {
            try { $csvLines = Get-Content -LiteralPath $selCsv } catch { Gui-Log ("⚠️ Kunde inte läsa CSV: " + $_.Exception.Message) 'Warn' }
            try { $csvStats = Get-CsvStats -Path $selCsv -Lines $csvLines } catch { Gui-Log ("⚠️ Get-CsvStats: " + $_.Exception.Message) 'Warn' }
        }
        if (-not $csvStats) {
            $csvStats = [pscustomobject]@{
                TestCount    = 0
                DupCount     = 0
                Duplicates   = @()
                LspValues    = @()
                LspOK        = $null
                InstrumentByType = [ordered]@{}
            }
        }
    
        $infSN = @()
        if ($script:GXINF_Map) {
            foreach ($k in $script:GXINF_Map.Keys) {
                if ($k -like 'Infinity-*') {
                    $infSN += ($script:GXINF_Map[$k].Split(',') | ForEach-Object { ($_ + '').Trim() } | Where-Object { $_ })
                }
            }
        }
    
        $infSN = $infSN | Select-Object -Unique
        $infSummary = '—'
    
        try {
            if ($selCsv -and (Test-Path -LiteralPath $selCsv) -and $infSN.Count -gt 0) {
                $infSummary = Get-InfinitySpFromCsvStrict -Path $selCsv -InfinitySerials $infSN -Lines $csvLines
            }
        } catch {
            Gui-Log ("Infinity SP fel: " + $_.Exception.Message) 'Warn'
        }
    
        $dupSampleCount = 0
        $dupSampleList  = @()
        if ($csvLines -and $csvLines.Count -gt 8) {
            try {
                $headerFields = ConvertTo-CsvFields $csvLines[7]
                $sampleIdx = -1
                for ($i=0; $i -lt $headerFields.Count; $i++) {
                    $hf = ($headerFields[$i] + '').Trim().ToLower()
                    if ($hf -match 'sample') { $sampleIdx = $i; break }
                }
                if ($sampleIdx -ge 0) {
                    $samples = @()
                    for ($r=9; $r -lt $csvLines.Count; $r++) {
                        $line = $csvLines[$r]
                        if (-not $line -or -not $line.Trim()) { continue }
                        $fields = ConvertTo-CsvFields $line
                        if ($fields.Count -gt $sampleIdx) {
                            $val = ($fields[$sampleIdx] + '').Trim()
                            if ($val) { $samples += $val }
                       }
                    }
    
                    if ($samples.Count -gt 0) {
                        $counts = @{}
                        foreach ($s in $samples) { if (-not $counts.ContainsKey($s)) { $counts[$s] = 0 }; $counts[$s]++ }
                        $dupList = @()
                        foreach ($entry in $counts.GetEnumerator()) {
                            if ($entry.Value -gt 1) {
                                $dupList += ("$($entry.Key) x$($entry.Value)")
                            }
                        }
                        $dupSampleCount = $dupList.Count
                        $dupSampleList  = $dupList
                    }
                }
            } catch {
                Gui-Log ("⚠️ Fel vid analys av Sample ID: " + $_.Exception.Message) 'Warn'
            }
        }
        $dupSampleText = if ($dupSampleCount -gt 0) {
            $show = ($dupSampleList | Select-Object -First 8) -join ', '
            "$dupSampleCount ($show)"
        } else { 'N/A' }
        $dupCartText = if ($csvStats.DupCount -gt 0) {
            $show = ($csvStats.Duplicates | Select-Object -First 8) -join ', '
            "$($csvStats.DupCount) ($show)"
        } else { 'N/A' }
        $lspSummary = ''
        try {
            if ($csvLines -and $csvLines.Count -gt 8) {
                $counts = @{}
                for ($rr = 9; $rr -lt $csvLines.Count; $rr++) {
                    $ln = $csvLines[$rr]
                    if (-not $ln -or -not $ln.Trim()) { continue }
                    $fs = ConvertTo-CsvFields $ln
                    if ($fs.Count -gt 4) {
                        $raw = ($fs[4] + '').Trim()
                        if ($raw) {
                            $mLsp = [regex]::Match($raw,'(\\d{5})')
                            $code = if ($mLsp.Success) { $mLsp.Groups[1].Value } else { $raw }
                            if (-not $counts.ContainsKey($code)) { $counts[$code] = 0 }
                            $counts[$code]++
                        }
                    }
                }
    
                if ($counts.Count -gt 0) {
                    $sorted = $counts.GetEnumerator() | Sort-Object Key
                    $lspSummaryParts = @()
                    foreach ($kvp in $sorted) {
                        $part = if ($kvp.Value -gt 1) { "$($kvp.Key) x$($kvp.Value)" } else { $kvp.Key }
                        $lspSummaryParts += $part
                    }
                    $total = $sorted.Count
                    if ($total -eq 1) {
                        $lspSummary = $sorted[0].Key
                    }
                    else {
                        $lspSummary = "$total (" + ($lspSummaryParts -join ', ') + ")"
                    }
                }
            }
        } catch {
            Gui-Log ("⚠️ Fel vid extraktion av LSP från CSV: " + $_.Exception.Message) 'Warn'
            $lspSummary = ''
        }
    
        $instText = if ($csvStats.InstrumentByType.Keys.Count -gt 0) {
            ($csvStats.InstrumentByType.GetEnumerator() | ForEach-Object { "$($_.Key)" } | Sort-Object) -join '; '
        } else { '' }
    
        function Find-InfoRow {
            param([OfficeOpenXml.ExcelWorksheet]$Ws, [string]$Label)
            if (-not $Ws -or -not $Ws.Dimension) { return $null }
            $maxRow = [Math]::Min($Ws.Dimension.End.Row, 300)
            for ($ri=1; $ri -le $maxRow; $ri++) {
                $txt = (($Ws.Cells[$ri,1].Text) + '').Trim()
                if (-not $txt) { continue }
                if ($txt.ToLowerInvariant() -eq $Label.ToLowerInvariant()) { return $ri }
            }
            return $null
        }
    
        $isNewLayout = $true
    
        try {
            $tmpRow = Find-InfoRow -Ws $wsInfo -Label 'CSV-Info'
            if ($tmpRow) { $isNewLayout = $true }
    
        } catch {}
    
        $rowCsvFile    = Find-InfoRow -Ws $wsInfo -Label 'CSV'
        $rowLsp        = Find-InfoRow -Ws $wsInfo -Label 'LSP'
        $rowAntal      = Find-InfoRow -Ws $wsInfo -Label 'Antal tester'
        $rowDupSample  = Find-InfoRow -Ws $wsInfo -Label 'Dubblett Sample ID'
        if (-not $rowDupSample) { $rowDupSample = Find-InfoRow -Ws $wsInfo -Label 'Dublett Sample ID' }
        $rowDupCart    = Find-InfoRow -Ws $wsInfo -Label 'Dubblett Cartridge S/N'
        if (-not $rowDupCart) { $rowDupCart = Find-InfoRow -Ws $wsInfo -Label 'Dublett Cartridge S/N' }
        $rowInst       = Find-InfoRow -Ws $wsInfo -Label 'Använda INF/GX'
        $rowBag = Find-InfoRow -Ws $wsInfo -Label 'Bag Numbers Tested Using Infinity'
    if (-not $rowBag) { $rowBag = Find-InfoRow -Ws $wsInfo -Label 'Bag Numbers Tested Using Infinity:' }
    if (-not $rowBag) { $rowBag = 14 } 
    
    $wsInfo.Cells["B$rowBag"].Style.Numberformat.Format = '@'
    $wsInfo.Cells["B$rowBag"].Value = $infSummary
    
        if ($isNewLayout) {
            $rowLsp = Find-InfoRow -Ws $wsInfo -Label 'LSP'
    
            if (-not $rowCsvFile)   { $rowCsvFile   = 8 }
            if (-not $rowLsp)       { $rowLsp       = 9 }
            if (-not $rowAntal)     { $rowAntal     = 10 }
            if (-not $rowDupSample) { $rowDupSample = 11 }
            if (-not $rowDupCart)   { $rowDupCart   = 12 }
            if (-not $rowInst)      { $rowInst      = 13 }
    
        }
    
        if ($selCsv) {
            $wsInfo.Cells["B$rowCsvFile"].Style.Numberformat.Format = '@'
            $wsInfo.Cells["B$rowCsvFile"].Value = (Split-Path $selCsv -Leaf)
        } else {
            $wsInfo.Cells["B$rowCsvFile"].Value = ''
        }
    
        if ($lspSummary -and $lspSummary -ne '') {
            $wsInfo.Cells["B$rowLsp"].Style.Numberformat.Format = '@'
            $wsInfo.Cells["B$rowLsp"].Value = $lspSummary
        } else {
            $wsInfo.Cells["B$rowLsp"].Style.Numberformat.Format = '@'
            $wsInfo.Cells["B$rowLsp"].Value = $lsp
        }
    
        $wsInfo.Cells["B$rowAntal"].Value = $csvStats.TestCount
        $wsInfo.Cells["B$rowAntal"].Style.Numberformat.Format = '@'
        $wsInfo.Cells["B$rowAntal"].Value = "$($csvStats.TestCount)"
    
        if ($rowDupSample) {
            $wsInfo.Cells["B$rowDupSample"].Value = $dupSampleText
    
        }
        if ($rowDupCart) {
            $wsInfo.Cells["B$rowDupCart"].Value = $dupCartText
        }
    
        $wsInfo.Cells["B$rowInst"].Value = $instText
    } catch {
        Gui-Log ("⚠️ CSV data-fel: " + $_.Exception.Message) 'Warn'
    }
    
    $assayForMacro = ''
    if ($runAssay) {
        $assayForMacro = $runAssay
    } elseif ($wsOut1) {
        $assayForMacro = ($wsOut1.Cells['D10'].Text + '').Trim()
    }
    
    $miniVal = ''
    if (Get-Command Get-MinitabMacro -ErrorAction SilentlyContinue) {
        $miniVal = Get-MinitabMacro -AssayName $assayForMacro
    }
    if (-not $miniVal) { $miniVal = 'N/A' }
    
    $hdNeg = $null; $hdPos = $null
    try { $hdNeg = Get-SealHeaderDocInfo -Pkg $pkgNeg } catch {}
    try { $hdPos = Get-SealHeaderDocInfo -Pkg $pkgPos } catch {}
    if (-not $hdNeg) { $hdNeg = [pscustomobject]@{Raw='';DocNo='';Rev=''} }
    if (-not $hdPos) { $hdPos = [pscustomobject]@{Raw='';DocNo='';Rev=''} }
    
    $wsInfo.Cells['B2'].Value = $ScriptVersion
    $wsInfo.Cells['B3'].Value = $env:USERNAME
    $wsInfo.Cells['B4'].Value = (Get-Date).ToString('yyyy-MM-dd HH:mm')
    $wsInfo.Cells['B5'].Value = if ($miniVal) { $miniVal } else { 'N/A' }
    $selLsp = $WorksheetPath
    $batchInfo = Get-BatchLinkInfo -SealPosPath $selPos -SealNegPath $selNeg -Lsp $lsp
    $batch = $batchInfo.Batch
    $wsInfo.Cells['A34'].Value = 'SharePoint Batch'
    $wsInfo.Cells['A34'].Style.Font.Bold = $true
    Add-Hyperlink -Cell $wsInfo.Cells['B34'] -Text $batchInfo.LinkText -Url $batchInfo.Url
    $linkMap = [ordered]@{
    
        'IPT App'      = 'https://apps.powerapps.com/play/e/default-771c9c47-7f24-44dc-958e-34f8713a8394/a/fd340dbd-bbbf-470b-b043-d2af4cb62c83'
        'MES Login'    = 'http://mes.cepheid.pri/camstarportal/?domain=CEPHEID.COM'
        'CSV Uploader' = 'http://auw2wgxtpap01.cepaws.com/Welcome.aspx'
        'BMRAM'        = 'https://cepheid62468.coolbluecloud.com/'
        'Agile'        = 'https://agileprod.cepheid.com/Agile/default/login-cms.jsp'
    }
    
    $rowLink = 35
    foreach ($key in $linkMap.Keys) {
        $wsInfo.Cells["A$rowLink"].Value = $key
        # Förkorta texten som visas i cellen till "LÄNK" enligt mallens stil
        Add-Hyperlink -Cell $wsInfo.Cells["B$rowLink"] -Text 'LÄNK' -Url $linkMap[$key]
        $rowLink++
    }
    
    # ----------------------------------------------------------------
    # WS (LSP Worksheet): hitta fil och skriv in i Information-bladet
    # ----------------------------------------------------------------
    try {
        if (-not $selLsp) {
            $probeDir = $null
            if ($selPos) { $probeDir = Split-Path -Parent $selPos }
            if (-not $probeDir -and $selNeg) { $probeDir = Split-Path -Parent $selNeg }
            if ($probeDir -and (Test-Path -LiteralPath $probeDir)) {
                $cand = Get-ChildItem -LiteralPath $probeDir -File -ErrorAction SilentlyContinue |
                        Where-Object {
                            ($_.Name -match '(?i)worksheet') -and ($_.Name -match [regex]::Escape($lsp)) -and ($_.Extension -match '^\.(xlsx|xlsm|xls)$')
                        } |
                        Sort-Object LastWriteTime -Descending | Select-Object -First 1
                if ($cand) {
                    $selLsp = $cand.FullName
                }
            }
        }
    
        function Find-LabelValueRightward {
        $normLbl = Normalize-HeaderText $Label
        $pat = '^(?i)\s*' + [regex]::Escape($normLbl).Replace('\ ', '\s*') + '\s*[:\.]*\s*$'
        $rx  = [regex]::new($pat, [Text.RegularExpressions.RegexOptions]::IgnoreCase)
    $hit = Find-RegexCell -Ws $Ws -Rx $rx -MaxRows $MaxRows -MaxCols $MaxCols
    if (-not $hit) { return $null }
    $cMax = [Math]::Min($Ws.Dimension.End.Column, $MaxCols)
    for ($c = $hit.Col + 1; $c -le $cMax; $c++) {
        $t = Normalize-HeaderText ($Ws.Cells[$hit.Row,$c].Text + '')
        if ($t) { return $t }
    }
    return $null
    }
    
        if ($selLsp -and (Test-Path -LiteralPath $selLsp)) {
            Gui-Log ("🔎 WS hittad: " + (Split-Path $selLsp -Leaf)) 'Info'
        } else {
            Gui-Log "ℹ️ Ingen WS-fil vald/hittad (LSP Worksheet). Hoppar över WS-extraktion." 'Info'
        }
    } catch {
        Gui-Log ("⚠️ WS-block fel: " + $_.Exception.Message) 'Warn'
    }
        try {
            $headerWs  = $null
            $headerNeg = $null
            $headerPos = $null
                 if ($selLsp -and (Test-Path -LiteralPath $selLsp)) {
            try {
                    $tmpPkg = New-Object OfficeOpenXml.ExcelPackage (New-Object IO.FileInfo($selLsp))                  
    $eqInfo = $null
    try {
    $eqInfo = Get-TestSummaryEquipment -Pkg $tmpPkg
    if ($eqInfo) {
        Gui-Log ("ℹ️ Utrustning hittad i WS '{0}': Pipetter={1}, Instrument={2}" -f $eqInfo.WorksheetName, ($eqInfo.Pipettes.Count), ($eqInfo.Instruments.Count)) 'Info'
    } else {
        Gui-Log "ℹ️ Utrustning gav tomt resultat." 'Info'
    }
    } catch {
    Gui-Log ("⚠️ Kunde inte extrahera utrustning från Test Summary: " + $_.Exception.Message) 'Warn'
    }
    
            $headerWs = Extract-WorksheetHeader -Pkg $tmpPkg
            $wsHeaderRows  = Get-WorksheetHeaderPerSheet -Pkg $tmpPkg
            $wsHeaderCheck = Compare-WorksheetHeaderSet   -Rows $wsHeaderRows
            try {
                if ($wsHeaderCheck.Issues -gt 0 -and $wsHeaderCheck.Summary) {
                    Gui-Log ("Worksheet header-avvikelser: {0} – se Information!" -f $wsHeaderCheck.Summary) 'Warn'
                } else {
                    Gui-Log "✅ Worksheet header korrekt" 'Info'
                }
            } catch {}
                $tmpPkg.Dispose()
                } catch {}
            }
            try { $headerNeg = Extract-SealTestHeader -Pkg $pkgNeg } catch {}
            try { $headerPos = Extract-SealTestHeader -Pkg $pkgPos } catch {}
            try {
                if ($selLsp -and (Test-Path -LiteralPath $selLsp)) {
                    $tmpPkg2 = New-Object OfficeOpenXml.ExcelPackage (New-Object IO.FileInfo($selLsp))
                    $wsLsp   = $tmpPkg2.Workbook.Worksheets | Where-Object { $_.Name -ne 'Worksheet Instructions' } | Select-Object -First 1
                    if ($wsLsp) {
                        if (-not $headerWs -or -not $headerWs.PartNo) {
                            $val = $null
                            $labels = @(
                                'Part No.', 'Part No.:', 'Part No', 'Part Number', 'Part Number:', 'Part Number.', 'Part Number.:'
                            )
                            foreach ($lbl in $labels) {
                                $val = Find-LabelValueRightward -Ws $wsLsp -Label $lbl
                                if ($val) { break }
                            }
                            if ($val) { $headerWs.PartNo = $val }
                        }
                        if (-not $headerWs -or -not $headerWs.BatchNo) {
                            $val = $null
                            $labels = @(
                                'Batch No(s)', 'Batch No(s).', 'Batch No(s):', 'Batch No(s).:',
                                'Batch No', 'Batch No.', 'Batch No:', 'Batch No.:' ,
                                'Batch Number', 'Batch Number.', 'Batch Number:', 'Batch Number.:'
                            )
                            foreach ($lbl in $labels) {
                                $val = Find-LabelValueRightward -Ws $wsLsp -Label $lbl
                                if ($val) { break }
                            }
                            if ($val) { $headerWs.BatchNo = $val }
                        }
                        if (-not $headerWs -or -not $headerWs.CartridgeNo -or $headerWs.CartridgeNo -eq '.') {
                            $val = $null
                            $labels = @(
                                'Cartridge No. (LSP)', 'Cartridge No. (LSP):', 'Cartridge No. (LSP) :',
                                'Cartridge No (LSP)', 'Cartridge No (LSP):', 'Cartridge No (LSP) :',
                                'Cartridge Number (LSP)', 'Cartridge Number (LSP):', 'Cartridge Number (LSP) :',
                                'Cartridge No.', 'Cartridge No.:', 'Cartridge No. :', 'Cartridge No :',
                                'Cartridge Number', 'Cartridge Number:', 'Cartridge Number :',
                                'Cartridge No', 'Cartridge No:', 'Cartridge No :'
                            )
                            foreach ($lbl in $labels) {
                                $val = Find-LabelValueRightward -Ws $wsLsp -Label $lbl
                                if ($val) { break }
                            }
                            if (-not $val) {
                                $rxCart = [regex]::new('(?i)Cartridge.*\(LSP\)')
                                $maxCols = [Math]::Min($wsLsp.Dimension.End.Column, 100)
                                $hitCart = Find-RegexCell -Ws $wsLsp -Rx $rxCart -MaxRows 200 -MaxCols $maxCols
                                if ($hitCart) {
                                    for ($c = $hitCart.Col + 1; $c -le $wsLsp.Dimension.End.Column; $c++) {
                                        $cellVal = ($wsLsp.Cells[$hitCart.Row, $c].Text + '').Trim()
                                        if ($cellVal) { $val = $cellVal; break }
                                    }
                                }
                            }
                            if ($val) { $headerWs.CartridgeNo = $val }
                        }
                        if (-not $headerWs -or -not $headerWs.Effective) {
                            $val = Find-LabelValueRightward -Ws $wsLsp -Label 'Effective'
                            if (-not $val) { $val = Find-LabelValueRightward -Ws $wsLsp -Label 'Effective Date' }
                            if ($val) { $headerWs.Effective = $val }
                        }
            }
            try {
    
                if ($selLsp -and (-not $headerWs -or -not $headerWs.CartridgeNo -or $headerWs.CartridgeNo -eq '.' -or $headerWs.CartridgeNo -eq '')) {
                    $fn = Split-Path $selLsp -Leaf
                    $m = [regex]::Matches($fn, '(?<!\d)(\d{5,7})(?!\d)')
                    if ($m.Count -gt 0) {
                        $headerWs.CartridgeNo = $m[0].Groups[1].Value
                    }
                }
            } catch {}
                    $tmpPkg2.Dispose()
                }
            } catch {}
            try {
                if ($pkgPos -and -not $headerPos.Effective) {
                    $wsPos = $pkgPos.Workbook.Worksheets | Where-Object { $_.Name -ne 'Worksheet Instructions' } | Select-Object -First 1
                    if ($wsPos) {
                        $val = Find-LabelValueRightward -Ws $wsPos -Label 'Effective'
                        if (-not $val) { $val = Find-LabelValueRightward -Ws $wsPos -Label 'Effective Date' }
                        if ($val) { $headerPos.Effective = $val }
                    }
                }
            } catch {}
            try {
                if ($pkgNeg -and -not $headerNeg.Effective) {
                    $wsNeg = $pkgNeg.Workbook.Worksheets | Where-Object { $_.Name -ne 'Worksheet Instructions' } | Select-Object -First 1
                    if ($wsNeg) {
                        $val = Find-LabelValueRightward -Ws $wsNeg -Label 'Effective'
                        if (-not $val) { $val = Find-LabelValueRightward -Ws $wsNeg -Label 'Effective Date' }
                        if ($val) { $headerNeg.Effective = $val }
                    }
                }
            } catch {}
            $wsBatch   = if ($headerWs -and $headerWs.BatchNo) { $headerWs.BatchNo } else { $null }
            $sealBatch = $batch
            if (-not $sealBatch) {
                try { if ($selPos) { $sealBatch = Get-BatchNumberFromSealFile $selPos } } catch {}
                if (-not $sealBatch) { try { if ($selNeg) { $sealBatch = Get-BatchNumberFromSealFile $selNeg } } catch {} }
            }
            $batchMatchFlag = $null
            if ($wsBatch -and $sealBatch) { $batchMatchFlag = ($wsBatch -eq $sealBatch) }
            $sealConsistentFlag = $null
            if ($headerNeg -and $headerPos) {
                if ($headerNeg.DocumentNumber -and $headerPos.DocumentNumber -and $headerNeg.Rev -and $headerPos.Rev -and $headerNeg.Effective -and $headerPos.Effective) {
                    $sealConsistentFlag = (($headerNeg.DocumentNumber -eq $headerPos.DocumentNumber) -and ($headerNeg.Rev -eq $headerPos.Rev) -and ($headerNeg.Effective -eq $headerPos.Effective))
                }
            }
            $noteStr = ''
            if ($headerNeg -and $headerNeg.DocumentNumber -and $headerNeg.DocumentNumber -ne 'D10552') { $noteStr += ("NEG DocNo (" + $headerNeg.DocumentNumber + ") != D10552; ") }
            if ($headerPos -and $headerPos.DocumentNumber -and $headerPos.DocumentNumber -ne 'D10552') { $noteStr += ("POS DocNo (" + $headerPos.DocumentNumber + ") != D10552; ") }
            $rowWsFile = Find-InfoRow -Ws $wsInfo -Label 'Worksheet'
            if (-not $rowWsFile) { $rowWsFile = 17 }
            $rowPart  = $rowWsFile + 1
            $rowBatch = $rowWsFile + 2
            $rowCart  = $rowWsFile + 3
            $rowDoc   = $rowWsFile + 4
            $rowRev   = $rowWsFile + 5
            $rowEff   = $rowWsFile + 6
            $rowPosFile = Find-InfoRow -Ws $wsInfo -Label 'Seal Test POS'
            if (-not $rowPosFile) {
                if ($rowWsFile) { $rowPosFile = $rowWsFile + 7 } else { $rowPosFile = 24 }
            }
            $rowPosDoc = $rowPosFile + 1
            $rowPosRev = $rowPosFile + 2
            $rowPosEff = $rowPosFile + 3
            $rowNegFile = Find-InfoRow -Ws $wsInfo -Label 'Seal Test NEG'
            if (-not $rowNegFile) {
                $rowNegFile = $rowPosFile + 4
            }
            $rowNegDoc = $rowNegFile + 1
            $rowNegRev = $rowNegFile + 2
            $rowNegEff = $rowNegFile + 3
            if ($selLsp) {
                $wsInfo.Cells["B$rowWsFile"].Style.Numberformat.Format = '@'
                $wsInfo.Cells["B$rowWsFile"].Value = (Split-Path $selLsp -Leaf)
            } else {
                $wsInfo.Cells["B$rowWsFile"].Value = ''
            }
    
            $consPart  = Get-ConsensusValue -Type 'Part'      -Ws $headerWs.PartNo      -Pos $headerPos.PartNumber   -Neg $headerNeg.PartNumber
            $consBatch = Get-ConsensusValue -Type 'Batch'     -Ws $headerWs.BatchNo     -Pos $headerPos.BatchNumber  -Neg $headerNeg.BatchNumber
            $consCart  = Get-ConsensusValue -Type 'Cartridge' -Ws $headerWs.CartridgeNo -Pos $headerPos.CartridgeNo  -Neg $headerNeg.CartridgeNo
     
            if (-not $consCart.Value -and $selLsp) {
                $fnCart = Split-Path $selLsp -Leaf
                $mCart  = [regex]::Match($fnCart,'(?<!\d)(\d{5,7})(?!\d)')
                if ($mCart.Success) {
                    $consCart = @{
                        Value  = $mCart.Groups[1].Value
                        Source = 'FILENAME'
                        Note   = 'Filename fallback'
                    }
                }
            }
    
            if ($consPart.Value)  { $wsInfo.Cells["B$rowPart"].Value = $consPart.Value }  else { $wsInfo.Cells["B$rowPart"].Value = '' }
            if ($consBatch.Value) { $wsInfo.Cells["B$rowBatch"].Value = $consBatch.Value } else { $wsInfo.Cells["B$rowBatch"].Value = '' }
            if ($consCart.Value)  { $wsInfo.Cells["B$rowCart"].Value = $consCart.Value }  else { $wsInfo.Cells["B$rowCart"].Value = '' }
    
            $batchMismatch = $false
            try {
                if ($headerNeg -and $headerPos -and $headerNeg.BatchNumber -and $headerPos.BatchNumber) {
                    $normNeg = Normalize-Id -Value $headerNeg.BatchNumber -Type 'Batch'
                    $normPos = Normalize-Id -Value $headerPos.BatchNumber -Type 'Batch'
                    if ($normNeg -and $normPos -and $normNeg -ne $normPos) { $batchMismatch = $true }
                }
            } catch {}
     
            if ($batchMismatch) {
                try { if ($consPart.Note)  { [void]$wsInfo.Cells["B$rowPart"].AddComment($consPart.Note,  'DocMerge') } } catch {}
                try { if ($consBatch.Note) { [void]$wsInfo.Cells["B$rowBatch"].AddComment($consBatch.Note, 'DocMerge') } } catch {}
                try { if ($consCart.Note)  { [void]$wsInfo.Cells["B$rowCart"].AddComment($consCart.Note,  'DocMerge') } } catch {}
            }
    
            try {
                if ($wsHeaderCheck -and $wsHeaderCheck.Details) {
                    $linesDev = ($wsHeaderCheck.Details -split "`r?`n")
                    $devPart  = $null
                    $devBatch = $null
                    $devCart  = $null
                    foreach ($ln in $linesDev) {
                        if ($ln -match '^-\s*PartNo[^:]*:\s*(.+)$') {
                            $devPart = $matches[1].Trim()
                        } elseif ($ln -match '^-\s*BatchNo[^:]*:\s*(.+)$') {
                            $devBatch = $matches[1].Trim()
                        } elseif ($ln -match '^-\s*CartridgeNo[^:]*:\s*(.+)$') {
                            $devCart = $matches[1].Trim()
                        }
                    }
                    if ($devPart) {
                        $wsInfo.Cells["C$rowPart"].Style.Numberformat.Format = '@'
                        $wsInfo.Cells["C$rowPart"].Value = 'Avvikande flik: ' + $devPart
                    }
                    if ($devBatch) {
                        $wsInfo.Cells["C$rowBatch"].Style.Numberformat.Format = '@'
                        $wsInfo.Cells["C$rowBatch"].Value = 'Avvikande flik: ' + $devBatch
                    }
                    if ($devCart) {
                        $wsInfo.Cells["C$rowCart"].Style.Numberformat.Format = '@'
                        $wsInfo.Cells["C$rowCart"].Value = 'Avvikande flik: ' + $devCart
                    }
                }
            } catch {}
    
            if ($headerWs) {
                $doc = $headerWs.DocumentNumber
                if ($doc) {
                    $doc = ($doc -replace '(?i)\s+(?:Rev(?:ision)?|Effective|p\.)\b.*$', '').Trim()
                }
                if ($headerWs.Attachment -and ($doc -notmatch '(?i)\bAttachment\s+\w+\b')) {
                    $doc = "$doc Attachment $($headerWs.Attachment)"
                }
                $wsInfo.Cells["B$rowDoc"].Value = $doc
                $wsInfo.Cells["B$rowRev"].Value = $headerWs.Rev
                $wsInfo.Cells["B$rowEff"].Value = $headerWs.Effective
            } else {
                $wsInfo.Cells["B$rowDoc"].Value = ''
                $wsInfo.Cells["B$rowRev"].Value = ''
                $wsInfo.Cells["B$rowEff"].Value = ''
            }
     
            if ($selPos) {
                $wsInfo.Cells["B$rowPosFile"].Style.Numberformat.Format = '@'
                $wsInfo.Cells["B$rowPosFile"].Value = (Split-Path $selPos -Leaf)
            } else {
                $wsInfo.Cells["B$rowPosFile"].Value = ''
            }
    
            if ($headerPos) {
                $docPos = $headerPos.DocumentNumber
                if ($docPos) { $docPos = ($docPos -replace '(?i)\s+(?:Rev(?:ision)?|Effective|p\.)\b.*$','').Trim() }
                $wsInfo.Cells["B$rowPosDoc"].Value = $docPos
                $wsInfo.Cells["B$rowPosRev"].Value = $headerPos.Rev
                $wsInfo.Cells["B$rowPosEff"].Value = $headerPos.Effective
            } else {
                $wsInfo.Cells["B$rowPosDoc"].Value = ''
                $wsInfo.Cells["B$rowPosRev"].Value = ''
                $wsInfo.Cells["B$rowPosEff"].Value = ''
            }
            if ($selNeg) {
                $wsInfo.Cells["B$rowNegFile"].Style.Numberformat.Format = '@'
                $wsInfo.Cells["B$rowNegFile"].Value = (Split-Path $selNeg -Leaf)
            } else {
                $wsInfo.Cells["B$rowNegFile"].Value = ''
            }
            # Seal Test NEG metadata
            if ($headerNeg) {
                # NEG: ta bort ev. "Rev/Effective" som följt med
                $docNeg = $headerNeg.DocumentNumber
                if ($docNeg) { $docNeg = ($docNeg -replace '(?i)\s+(?:Rev(?:ision)?|Effective|p\.)\b.*$','').Trim() }
                $wsInfo.Cells["B$rowNegDoc"].Value = $docNeg
                $wsInfo.Cells["B$rowNegRev"].Value = $headerNeg.Rev
                $wsInfo.Cells["B$rowNegEff"].Value = $headerNeg.Effective
            } else {
                $wsInfo.Cells["B$rowNegDoc"].Value = ''
                $wsInfo.Cells["B$rowNegRev"].Value = ''
                $wsInfo.Cells["B$rowNegEff"].Value = ''
            }
            # Töm eventuella överflödiga rader nedanför tabellen – ej nödvändig då layout definierad i mall
        } catch {
            Gui-Log ("⚠️ Header summary fel: " + $_.Exception.Message) 'Warn'
        }
    } catch {
    Gui-Log "⚠️ Information-blad fel: $($_.Exception.Message)" 'Warn'
    }
    
        # ============================
        # === Equipment-blad       ===
        # ============================
        try {
            if (Test-Path -LiteralPath $UtrustningListPath) {
                $srcPkg = New-Object OfficeOpenXml.ExcelPackage (New-Object IO.FileInfo($UtrustningListPath))
                try {
                    $srcWs = $srcPkg.Workbook.Worksheets['Sheet1']
                    if (-not $srcWs) {
                        $srcWs = $srcPkg.Workbook.Worksheets[1]
                    }
    
                    if ($srcWs) {
                        $wsEq = $pkgOut.Workbook.Worksheets['Infinity/GX']
                        if ($wsEq) {
                            $pkgOut.Workbook.Worksheets.Delete($wsEq)
                        }
                        $wsEq = $pkgOut.Workbook.Worksheets.Add('Infinity/GX', $srcWs)
    
                        if ($wsEq.Dimension) {
                            foreach ($cell in $wsEq.Cells[$wsEq.Dimension.Address]) {
                                if ($cell.Formula -or $cell.FormulaR1C1) {
                                    $val = $cell.Value
                                    $cell.Formula     = $null
                                    $cell.FormulaR1C1 = $null
                                    $cell.Value       = $val
                                }
                            }
                            $colCount = $srcWs.Dimension.End.Column
                            for ($c = 1; $c -le $colCount; $c++) {
                                try {
                                    $wsEq.Column($c).Width = $srcWs.Column($c).Width
                                } catch {
                                }
                            }
                        }
    
                        if ($eqInfo) {
                            $wsName = $null
                            if ($selLsp) {
                                $wsName = Split-Path $selLsp -Leaf
                            } elseif ($eqInfo.PSObject.Properties['WorksheetName'] -and $eqInfo.WorksheetName) {
                                $wsName = $eqInfo.WorksheetName
                            } elseif ($headerWs -and $headerWs.PSObject.Properties['WorksheetName'] -and $headerWs.WorksheetName) {
                                $wsName = $headerWs.WorksheetName
                            } else {
                                $wsName = 'Test Summary'
                            }
    
                            $cellHeaderPip = $wsEq.Cells['A14']
                            $cellHeaderPip.Value = "PIPETTER hämtade från $wsName"
                            $cellHeaderPip.Style.Font.Bold = $true
                            $cellHeaderPip.Style.HorizontalAlignment = [OfficeOpenXml.Style.ExcelHorizontalAlignment]::Center
                            $cellHeaderPip.Style.VerticalAlignment   = [OfficeOpenXml.Style.ExcelVerticalAlignment]::Center
    
                            $cellHeaderInst = $wsEq.Cells['A18']
                            $cellHeaderInst.Value = "INSTRUMENT hämtade från $wsName"
                            $cellHeaderInst.Style.Font.Bold = $true
                            $cellHeaderInst.Style.HorizontalAlignment = [OfficeOpenXml.Style.ExcelHorizontalAlignment]::Center
                            $cellHeaderInst.Style.VerticalAlignment   = [OfficeOpenXml.Style.ExcelVerticalAlignment]::Center
    
                            function Convert-ToEqDate {
                                param(
                                    [Parameter(Mandatory = $false)]
                                    $Value
                                )
    
                                if (-not $Value -or $Value -eq 'N/A') {
                                    return $null
                                }
    
                                if ($Value -is [datetime]) {
                                    return $Value
                                }
    
                                if ($Value -is [double] -or $Value -is [int]) {
                                    try {
                                        $base = Get-Date '1899-12-30'
                                        return $base.AddDays([double]$Value)
                                    } catch {
                                        return $Value
                                    }
                                }
    
                                try {
                                    return (Get-Date -Date $Value -ErrorAction Stop)
                                } catch {
                                    return $Value
                                }
                            }
    
                            $pipetteIdCells  = @('B15','D15','F15','H15','J15','L15')
                            $pipetteDueCells = @('B16','D16','F16','H16','J16','L16')
    
                            $pipettes = @()
                            if ($eqInfo.PSObject.Properties['Pipettes'] -and $eqInfo.Pipettes) {
                                $pipettes = @($eqInfo.Pipettes)
                            }
    
                            for ($i = 0; $i -lt $pipetteIdCells.Count; $i++) {
                                $cellId  = $wsEq.Cells[$pipetteIdCells[$i]]
                                $cellDue = $wsEq.Cells[$pipetteDueCells[$i]]
    
                                if ($i -lt $pipettes.Count) {
                                    $p = $pipettes[$i]
    
                                    $id  = $null
                                    $due = $null
    
                                    if ($p -is [string]) {
                                        $id = $p
                                    } else {
                                        $idCandidates = @()
                                        foreach ($propName in 'Id','CepheidId','Name','PipetteId') {
                                            if ($p.PSObject.Properties[$propName]) {
                                                $idCandidates += $p.$propName
                                            }
                                        }
                                        $id = $idCandidates | Where-Object { $_ } | Select-Object -First 1
    
                                        $dueCandidates = @()
                                        foreach ($propName in 'CalibrationDueDate','DueDate','CalDue') {
                                            if ($p.PSObject.Properties[$propName]) {
                                                $dueCandidates += $p.$propName
                                            }
                                        }
                                        $due = $dueCandidates | Where-Object { $_ } | Select-Object -First 1
                                    }
    
                                    if ([string]::IsNullOrWhiteSpace($id) -or $id -eq 'N/A') {
                                        $cellId.Value  = 'N/A'
                                        $cellDue.Value = 'N/A'
                                    } else {
                                        $cellId.Value = $id
    
                                        $dt = Convert-ToEqDate -Value $due
                                        if ($dt -is [datetime]) {
                                            $cellDue.Value = $dt
                                            $cellDue.Style.Numberformat.Format = 'mmm-yy'
                                        } elseif ($dt) {
                                            $cellDue.Value = $dt
                                        } else {
                                            $cellDue.Value = 'N/A'
                                        }
                                    }
                                } else {
                                    $cellId.Value  = 'N/A'
                                    $cellDue.Value = 'N/A'
                                }
    
                                foreach ($c in @($cellId,$cellDue)) {
                                    $c.Style.HorizontalAlignment = [OfficeOpenXml.Style.ExcelHorizontalAlignment]::Center
                                    $c.Style.VerticalAlignment   = [OfficeOpenXml.Style.ExcelVerticalAlignment]::Center
                                }
                            }
    
                            $instIdCells  = @(
                                'B19','D19','F19','H19','J19','L19',
                                'B21','D21','F21','H21','J21','L21'
                            )
                            $instDueCells = @(
                                'B20','D20','F20','H20','J20','L20',
                                'B22','D22','F22','H22','J22','L22'
                            )
    
                            $instruments = @()
                            if ($eqInfo.PSObject.Properties['Instruments'] -and $eqInfo.Instruments) {
                                $instruments = @($eqInfo.Instruments)
                            }
    
                            for ($i = 0; $i -lt $instIdCells.Count; $i++) {
                                $cellId  = $wsEq.Cells[$instIdCells[$i]]
                                $cellDue = $wsEq.Cells[$instDueCells[$i]]
    
                                if ($i -lt $instruments.Count) {
                                    $inst = $instruments[$i]
    
                                    $id  = $null
                                    $due = $null
    
                                    if ($inst -is [string]) {
                                        $id = $inst
                                    } else {
                                        $idCandidates = @()
                                        foreach ($propName in 'Id','CepheidId','Name','InstrumentId') {
                                            if ($inst.PSObject.Properties[$propName]) {
                                                $idCandidates += $inst.$propName
                                            }
                                        }
                                        $id = $idCandidates | Where-Object { $_ } | Select-Object -First 1
    
                                        $dueCandidates = @()
                                        foreach ($propName in 'CalibrationDueDate','DueDate','CalDue') {
                                            if ($inst.PSObject.Properties[$propName]) {
                                                $dueCandidates += $inst.$propName
                                            }
                                        }
                                        $due = $dueCandidates | Where-Object { $_ } | Select-Object -First 1
                                    }
    
                                    if ([string]::IsNullOrWhiteSpace($id) -or $id -eq 'N/A') {
                                        $cellId.Value  = 'N/A'
                                        $cellDue.Value = 'N/A'
                                    } else {
                                        $cellId.Value = $id
    
                                        $dt = Convert-ToEqDate -Value $due
                                        if ($dt -is [datetime]) {
                                            $cellDue.Value = $dt
                                            $cellDue.Style.Numberformat.Format = 'mmm-yy'
                                        } elseif ($dt) {
                                            $cellDue.Value = $dt
                                        } else {
                                            $cellDue.Value = 'N/A'
                                        }
                                    }
                                } else {
                                    $cellId.Value  = 'N/A'
                                    $cellDue.Value = 'N/A'
                                }
    
                                foreach ($c in @($cellId,$cellDue)) {
                                    $c.Style.HorizontalAlignment = [OfficeOpenXml.Style.ExcelHorizontalAlignment]::Center
                                    $c.Style.VerticalAlignment   = [OfficeOpenXml.Style.ExcelVerticalAlignment]::Center
                                }
                            }
    
                            if ($pipettes.Count -gt $pipetteIdCells.Count -or
                                $instruments.Count -gt $instIdCells.Count) {
                                Gui-Log ("ℹ️ Infinity/GX: allt får inte plats i mallen (pipetter={0}, instrument={1})" -f $pipettes.Count, $instruments.Count) 'Info'
                            }
    
                        } else {
                            Gui-Log "ℹ️ Utrustning saknas – Infinity/GX lämnas som mall." 'Info'
                        }
                    }
                }
                finally {
                    if ($srcPkg) { $srcPkg.Dispose() }
                }
            } else {
                Gui-Log "ℹ️ Infinity/GX mall saknas: $($_.Exception.Message)" 'Info'
            }
        }
        catch {
            Gui-Log "⚠️ Kunde inte skapa 'Infinity/GX'-flik: $($_.Exception.Message)" 'Warn'
        }
    
        # ============================
        # === Control Material     ===
        # ============================
        try {
            if ($controlTab -and (Test-Path -LiteralPath $RawDataPath)) {
                $srcPkg = New-Object OfficeOpenXml.ExcelPackage (New-Object IO.FileInfo($RawDataPath))
                try { $srcPkg.Workbook.Calculate() } catch {}
                $candidates = if ($controlTab -match '\|') { $controlTab -split '\|' | ForEach-Object { $_.Trim() } | Where-Object { $_ } } else { @($controlTab) }
                $srcWs = $null
                foreach ($cand in $candidates) {
    
                    $srcWs = $srcPkg.Workbook.Worksheets | Where-Object { $_.Name -eq $cand } | Select-Object -First 1
                    if ($srcWs) { break }
                    $srcWs = $srcPkg.Workbook.Worksheets | Where-Object { $_.Name -like "*$cand*" } | Select-Object -First 1
                    if ($srcWs) { break }
                }
                if ($srcWs) {
                    $safeName = if ($srcWs.Name.Length -gt 31) { $srcWs.Name.Substring(0,31) } else { $srcWs.Name }
                    $destName = $safeName; $n=1
                    while ($pkgOut.Workbook.Worksheets[$destName]) { $base = if ($safeName.Length -gt 27) { $safeName.Substring(0,27) } else { $safeName }; $destName = "$base($n)"; $n++ }
                    $wsCM = $pkgOut.Workbook.Worksheets.Add($destName, $srcWs)
                    if ($wsCM.Dimension) {
                        foreach ($cell in $wsCM.Cells[$wsCM.Dimension.Address]) {
                            if ($cell.Formula -or $cell.FormulaR1C1) { $v=$cell.Value; $cell.Formula=$null; $cell.FormulaR1C1=$null; $cell.Value=$v }
                        }
                        try { $wsCM.Cells[$wsCM.Dimension.Address].AutoFitColumns() | Out-Null } catch {}
                    }
                    Gui-Log "✅ Control Material kopierad: '$($srcWs.Name)' → '$destName'" 'Info'
                } else { Gui-Log "ℹ️ Hittade inget blad i kontrollfilen som matchar '$controlTab'." 'Info' }
                $srcPkg.Dispose()
            } else { Gui-Log "ℹ️ Ingen Control-flik skapad (saknar mappning eller kontrollfil)." 'Info' }
        } catch { Gui-Log "⚠️ Control Material-fel: $($_.Exception.Message)" 'Warn' }
    
        # ============================
        # === SharePoint Info      ===
        # ============================
        try {
    
            if (-not $IncludeSharePoint) {
                Gui-Log "ℹ️ SharePoint Info ej valt – hoppar över." 'Info'
                try { $old = $pkgOut.Workbook.Worksheets["SharePoint Info"]; if ($old) { $pkgOut.Workbook.Worksheets.Delete($old) } } catch {}
            } else {
                $spOk = $false
                if ($global:SpConnected) { $spOk = $true }
                elseif (Get-Command Get-PnPConnection -ErrorAction SilentlyContinue) {
                    try { $null = Get-PnPConnection; $spOk = $true } catch { $spOk = $false }
                }
     
                if (-not $spOk) {
                    $errMsg = if ($global:SpError) { $global:SpError } else { 'Okänt fel' }
                    Gui-Log ("⚠️ SharePoint ej tillgängligt: $errMsg") 'Warn'
                }
    
                $batchInfo = Get-BatchLinkInfo -SealPosPath $selPos -SealNegPath $selNeg -Lsp $lsp
                $batch = $batchInfo.Batch
    
                if (-not $batch) {
                    Gui-Log "ℹ️ Inget Batch # i POS/NEG – skriver tom SharePoint Info." 'Info'
                    if (Get-Command Write-SPSheet-Safe -ErrorAction SilentlyContinue) {
                        [void](Write-SPSheet-Safe -Pkg $pkgOut -Rows @() -DesiredOrder @() -Batch '—')
                    } else {
                        $wsSp = $pkgOut.Workbook.Worksheets["SharePoint Info"]; if ($wsSp) { $pkgOut.Workbook.Worksheets.Delete($wsSp) }
                        $wsSp = $pkgOut.Workbook.Worksheets.Add("SharePoint Info")
                        $wsSp.Cells[1,1].Value = "Rubrik"; $wsSp.Cells[1,2].Value = "Värde"
                        $wsSp.Cells[2,1].Value = "Batch";  $wsSp.Cells[2,2].Value = "—"
                        try { $wsSp.Cells[$wsSp.Dimension.Address].AutoFitColumns() | Out-Null } catch {}
                    }
                } else {
                    Gui-Log "🔎 Batch hittad: $batch" 'Info'
    
                    $fields = @(
                        'Work_x0020_Center','Title','Batch_x0023_','SAP_x0020_Batch_x0023__x0020_2',
                        'LSP','Material','BBD_x002f_SLED','Actual_x0020_startdate_x002f__x0',
                        'PAL_x0020__x002d__x0020_Sample_x','Sample_x0020_Reagent_x0020_P_x00',
                        'Order_x0020_quantity','Total_x0020_good','ITP_x0020_Test_x0020_results',
                        'IPT_x0020__x002d__x0020_Testing_0','MES_x0020__x002d__x0020_Order_x0'
                    )
                    $renameMap = @{
                        'Work Center'            = 'Work Center'
                        'Title'                  = 'Order#'
                        'Batch#'                 = 'SAP Batch#'
                        'SAP Batch# 2'           = 'SAP Batch# 2'
                        'LSP'                    = 'LSP'
                        'Material'               = 'Material'
                        'BBD/SLED'               = 'BBD/SLED'
                        'Actual startdate/_x0'   = 'ROBAL - Actual start date/time'
                        'PAL - Sample_x'         = 'Sample Reagent use'
                        'Sample Reagent P'       = 'Sample Reagent P/N'
                        'Order quantity'         = 'Order quantity'
                        'Total good'             = 'ROBAL - Till Packning'
                        'IPT Test results'       = 'IPT Test results'
                        'IPT - Testing_0'        = 'IPT - Testing Finalized'
                        'MES - Order_x0'         = 'MES Order'
                    }
    
                    $desiredOrder = @(
                        'Work Center','Order#','SAP Batch#','SAP Batch# 2','LSP','Material','BBD/SLED',
                        'ROBAL - Actual start date/time','Sample Reagent use','Sample Reagent P/N',
                        'Order quantity','ROBAL - Till Packning','IPT Test results',
                        'IPT - Testing Finalized','MES Order'
                    )
    
                    $dateFields      = @('BBD/SLED','ROBAL - Actual start date/time','IPT - Testing Finalized')
                    $shortDateFields = @('BBD/SLED')
                    $rows = @()
                    if ($spOk) {
                        try {
                            $items = Get-PnPListItem -List "Cepheid | Production orders" -Fields $fields -PageSize 2000 -ErrorAction Stop
                            $match = $items | Where-Object {
                                $v1 = $_['Batch_x0023_']; $v2 = $_['SAP_x0020_Batch_x0023__x0020_2']
                                $s1 = if ($null -ne $v1) { ([string]$v1).Trim() } else { '' }
                                $s2 = if ($null -ne $v2) { ([string]$v2).Trim() } else { '' }
                                $s1 -eq $batch -or $s2 -eq $batch
                            } | Select-Object -First 1
                            if ($match) {
                                foreach ($f in $fields) {
                                    $val = $match[$f]
                                    $label = $f -replace '_x0020_', ' ' `
                                                 -replace '_x002d_', '-' `
                                                 -replace '_x0023_', '#' `
                                                 -replace '_x002f_', '/' `
                                                 -replace '_x2013_', '–' `
                                                 -replace '_x00',''
                                    $label = $label.Trim()
                                    if ($renameMap.ContainsKey($label)) { $label = $renameMap[$label] }
                                    if ($null -ne $val -and $val -ne '') {
                                        if ($val -eq $true) { $val = 'JA' }
                                        elseif ($val -eq $false) { $val = 'NEJ' }
    
                                        $dt = $null
                                        if ($val -is [datetime]) { $dt = [datetime]$val }
                                        else { try { $dt = [datetime]::Parse($val) } catch { $dt = $null } }
                                        if ($dt -ne $null -and ($dateFields -contains $label)) {
                                            $fmt = if ($shortDateFields -contains $label) { 'yyyy-MM-dd' } else { 'yyyy-MM-dd HH:mm' }
                                            $val = $dt.ToString($fmt)
                                        }
                                        $rows += [pscustomobject]@{ Rubrik = $label; 'Värde' = $val }
                                    }
                                }
     
                                if ($rows.Count -gt 0) {
                                    $ordered = @()
                                    foreach ($label in $desiredOrder) {
                                        $hit = $rows | Where-Object { $_.Rubrik -eq $label } | Select-Object -First 1
                                        if ($hit) { $ordered += $hit }
                                    }
                                    if ($ordered.Count -gt 0) { $rows = $ordered }
                                }
                                Gui-Log "📄 SharePoint-post hittad – skriver blad." 'Info'
                            } else {
                                Gui-Log "ℹ️ Ingen post i SharePoint för Batch=$batch." 'Info'
                            }
                        } catch {
                            Gui-Log "⚠️ SP: Get-PnPListItem misslyckades: $($_.Exception.Message)" 'Warn'
                        }
                    }
                    if (Get-Command Write-SPSheet-Safe -ErrorAction SilentlyContinue) {
                        [void](Write-SPSheet-Safe -Pkg $pkgOut -Rows $rows -DesiredOrder $desiredOrder -Batch $batch)
                    } else {
                        $wsSp = $pkgOut.Workbook.Worksheets["SharePoint Info"]; if ($wsSp) { $pkgOut.Workbook.Worksheets.Delete($wsSp) }
                        $wsSp = $pkgOut.Workbook.Worksheets.Add("SharePoint Info")
                        $wsSp.Cells[1,1].Value = "Rubrik"; $wsSp.Cells[1,2].Value = "Värde"
                        if ($rows.Count -gt 0) {
                            $r=2; foreach($rowObj in $rows) { $wsSp.Cells[$r,1].Value = $rowObj.Rubrik; $wsSp.Cells[$r,2].Value = $rowObj.'Värde'; $r++ }
                        } else {
                            $wsSp.Cells[2,1].Value = "Batch";  $wsSp.Cells[2,2].Value = $batch
                            $wsSp.Cells[3,1].Value = "Info";   $wsSp.Cells[3,2].Value = "No matching SharePoint row"
                        }
                        try { $wsSp.Cells[$wsSp.Dimension.Address].AutoFitColumns() | Out-Null } catch {}
                    }
                    try {
                        if ($SharePointLinkLabel -and $batch) {
                            $SharePointLinkLabel.Text = "SharePoint: $batch"
                            $SharePointLinkLabel.Tag  = $batchInfo.Url
                            $SharePointLinkLabel.Enabled = $true
                        }
                    } catch {}
                    try {
                        $wsSP = $pkgOut.Workbook.Worksheets['SharePoint Info']
                        if ($wsSP -and $wsSP.Dimension) {
                            $labelCol = 1; $valueCol = 2
                            for ($r = 1; $r -le $wsSP.Dimension.End.Row; $r++) {
                                if (($wsSP.Cells[$r,$labelCol].Text).Trim() -eq 'Sample Reagent use') {
                                    $wsSP.Cells[$r,$valueCol].Style.WrapText = $true
                                    $wsSP.Cells[$r,$valueCol].Style.VerticalAlignment = [OfficeOpenXml.Style.ExcelVerticalAlignment]::Top
                                    try { $wsSP.Column($valueCol).Width = 55 } catch {}
                                    $wsSP.Row($r).CustomHeight = $true
                                    break
                                }
                            }
                        }
                    } catch {
                        Gui-Log "⚠️ WrapText på 'Sample Reagent use' misslyckades: $($_.Exception.Message)" 'Warn'
                    }
                }
            }
        } catch {
            Gui-Log "⚠️ SP-blad: $($_.Exception.Message)" 'Warn'
        }
    
        # ============================
        # === Header watermark     ===
        # ============================
    
        try {
            foreach ($ws in $pkgOut.Workbook.Worksheets) {
                try {
                    $ws.HeaderFooter.OddHeader.CenteredText   = '&"Arial,Bold"&14 UNCONTROLLED'
                    $ws.HeaderFooter.EvenHeader.CenteredText  = '&"Arial,Bold"&14 UNCONTROLLED'
                    $ws.HeaderFooter.FirstHeader.CenteredText = '&"Arial,Bold"&14 UNCONTROLLED'
                } catch { Write-Warning "Kunde inte sätta header på blad: $($ws.Name)" }
            }
        } catch { Write-Warning "Fel vid vattenstämpling av rapporten." }
    
        # ============================
        # === Tab-färger (innan Save)
        # ============================
    
        try {
            $wsT = $pkgOut.Workbook.Worksheets['Information'];            if ($wsT) { $wsT.TabColor = [System.Drawing.Color]::FromArgb(255, 52, 152, 219) }
            $wsT = $pkgOut.Workbook.Worksheets['Infinity/GX'];            if ($wsT) { $wsT.TabColor = [System.Drawing.Color]::FromArgb(255, 33, 115, 70) }
            $wsT = $pkgOut.Workbook.Worksheets['SharePoint Info'];        if ($wsT) { $wsT.TabColor = [System.Drawing.Color]::FromArgb(255, 0, 120, 212) }
        } catch {
            Gui-Log "⚠️ Kunde inte sätta tab-färg: $($_.Exception.Message)" 'Warn'
        }
    
        # ============================
        # === Spara & Audit        ===
        # ============================
    
        $nowTs   = Get-Date -Format "yyyyMMdd_HHmmss"
        $baseName = "$($env:USERNAME)_output_${lsp}_$nowTs.xlsx"
        if ($SaveInLsp) {
            $saveDir = Split-Path -Parent $selNeg
            $SavePath = Join-Path $saveDir $baseName
            Gui-Log "💾 Sparläge: LSP-mapp → $saveDir"
        } else {
            $saveDir = $env:TEMP
            $SavePath = Join-Path $saveDir $baseName
            Gui-Log "💾 Sparläge: Temporärt → $SavePath"
        }
        try {
            $pkgOut.Workbook.View.ActiveTab = 0
            $wsInitial = $pkgOut.Workbook.Worksheets["Information"]
            if ($wsInitial) { $wsInitial.View.TabSelected = $true }
            $pkgOut.SaveAs($SavePath)
            Gui-Log "✅ Rapport sparad: $SavePath" 'Info'
            $global:LastReportPath = $SavePath
    
            try {
                $auditDir = Join-Path $PSScriptRoot 'audit'
                if (-not (Test-Path $auditDir)) { New-Item -ItemType Directory -Path $auditDir -Force | Out-Null }
                $auditObj = [pscustomobject]@{
                    DatumTid        = (Get-Date).ToString('yyyy-MM-dd HH:mm:ss')
                    Användare       = $env:USERNAME
                    LSP             = $lsp
                    ValdCSV         = if ($selCsv) { Split-Path $selCsv -Leaf } else { '' }
                    ValdSealNEG     = Split-Path $selNeg -Leaf
                    ValdSealPOS     = Split-Path $selPos -Leaf
                    SignaturSkriven = if ($SignSealTest) { 'Ja' } else { 'Nej' }
                    OverwroteSign   = if ($OverwriteSignature) { 'Ja' } else { 'Nej' }
                    SigMismatch     = if ($sigMismatch) { 'Ja' } else { 'Nej' }
                    MismatchSheets  = if ($mismatchSheets -and $mismatchSheets.Count -gt 0) { ($mismatchSheets -join ';') } else { '' }
                    ViolationsNEG   = $violationsNeg.Count
                    ViolationsPOS   = $violationsPos.Count
                    Violations      = ($violationsNeg.Count + $violationsPos.Count)
                    Sparläge        = if ($SaveInLsp) { 'LSP-mapp' } else { 'Temporärt' }
                    OutputFile      = $SavePath
                    Kommentar       = 'UNCONTROLLED rapport, ingen källfil ändrades automatiskt.'
                    ScriptVersion   = $ScriptVersion
                }
    
                $auditFile = Join-Path $auditDir ("$($env:USERNAME)_audit_${nowTs}.csv")
                $auditObj | Export-Csv -Path $auditFile -NoTypeInformation -Encoding UTF8
                try {
                    $statusText = 'OK'
                    if (($violationsNeg.Count + $violationsPos.Count) -gt 0 -or $sigMismatch -or ($mismatchSheets -and $mismatchSheets.Count -gt 0)) {
                        $statusText = 'Warnings'
                    }
                    $auditTests = $null
                    try { if ($csvStats) { $auditTests = $csvStats.TestCount } } catch {}
                    Add-AuditEntry -Lsp $lsp -Assay $runAssay -BatchNumber $batch -TestCount $auditTests -Status $statusText -ReportPath $SavePath
                } catch { Gui-Log "⚠️ Kunde inte skriva audit-CSV: $($_.Exception.Message)" 'Warn' }
            } catch { Gui-Log "⚠️ Kunde inte skriva revisionsfil: $($_.Exception.Message)" 'Warn' }
    
            try { Start-Process -FilePath "excel.exe" -ArgumentList "`"$SavePath`"" } catch {}
        }
        catch { Gui-Log "⚠️ Kunde inte spara/öppna: $($_.Exception.Message)" 'Warn' }
    } finally {
        try { if ($pkgNeg) { $pkgNeg.Dispose() } } catch {}
        try { if ($pkgPos) { $pkgPos.Dispose() } } catch {}
        try { if ($pkgOut) { $pkgOut.Dispose() } } catch {}
    }
    return $SavePath
}
