#requires -Version 5.1
[CmdletBinding()]
param(
    [string]$XlsxPath,
    [string]$WorksheetName = 'Slang till Assay',
    [int]$StartRow = 2,
    [string]$OutputPath
)

$here = Split-Path -Parent $MyInvocation.MyCommand.Path
$root = Split-Path -Parent $here
if (-not $OutputPath) { $OutputPath = Join-Path $root 'Modules\\AssayAliases.generated.psd1' }

if (-not $XlsxPath) {
    $cfgPath = Join-Path $root 'Modules\\Config.ps1'
    if (Test-Path -LiteralPath $cfgPath) {
        try { . $cfgPath -ScriptRoot $root } catch {}
        try { if (-not $XlsxPath -and $Config -and $Config.SlangAssayPath) { $XlsxPath = $Config.SlangAssayPath } } catch {}
    }
}

if (-not (Test-Path -LiteralPath $XlsxPath)) {
    Write-Error "Xlsx not found: $XlsxPath"
    return
}

try {
    $loader = Join-Path $root 'Modules\\EpplusLoader.ps1'
    if (Test-Path -LiteralPath $loader) { . $loader }
    if (Get-Command Initialize-EPPlus -ErrorAction SilentlyContinue) { Initialize-EPPlus | Out-Null }
} catch {}

try { [void][OfficeOpenXml.ExcelPackage] } catch {
    $dllPath = Join-Path $root 'Modules\\EPPlus.dll'
    if (Test-Path -LiteralPath $dllPath) {
        try { [Reflection.Assembly]::LoadFrom($dllPath) | Out-Null } catch {}
    }
}

$aliases = [ordered]@{}
try {
    $pkg = New-Object OfficeOpenXml.ExcelPackage (New-Object IO.FileInfo($XlsxPath))
    try {
        $ws = if ($WorksheetName) { $pkg.Workbook.Worksheets[$WorksheetName] } else { $null }
        if (-not $ws) { $ws = $pkg.Workbook.Worksheets[1] }
        if (-not $ws -or -not $ws.Dimension) { throw "Worksheet not found or empty" }

        $lastRow = $ws.Dimension.End.Row
        $lastCol = $ws.Dimension.End.Column
        for ($r = $StartRow; $r -le $lastRow; $r++) {
            $canonical = ($ws.Cells[$r,1].Text + '').Trim()
            if (-not $canonical) { continue }
            if (-not $aliases.Contains($canonical)) { $aliases[$canonical] = $canonical }
            for ($c = 2; $c -le $lastCol; $c++) {
                $val = ($ws.Cells[$r,$c].Text + '').Trim()
                if (-not $val) { continue }
                if (-not $aliases.Contains($val)) { $aliases[$val] = $canonical }
            }
        }
    } finally {
        if ($pkg) { $pkg.Dispose() }
    }
} catch {
    Write-Error "Failed to read Excel: $($_.Exception.Message)"
    return
}

$sb = New-Object System.Text.StringBuilder
[void]$sb.AppendLine('@{')
foreach ($k in $aliases.Keys) {
    $line = "  '{0}' = '{1}'" -f ($k -replace "'","''"), ($aliases[$k] -replace "'","''")
    [void]$sb.AppendLine($line)
}
[void]$sb.AppendLine('}')

$encoding = New-Object System.Text.UTF8Encoding($true)
[System.IO.File]::WriteAllText($OutputPath, $sb.ToString(), $encoding)
Write-Output ("Wrote alias map to {0} (entries={1})" -f $OutputPath, $aliases.Count)
