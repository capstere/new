#requires -Version 5.1
[CmdletBinding()]
param(
    [Parameter(Mandatory=$true)]
    [string]$CsvPath,
    [switch]$ShowAffected,
    [int]$Top = 10
)

$root = Split-Path -Parent $PSScriptRoot
$modules = Join-Path $root 'Modules'

. (Join-Path $modules 'Config.ps1') -ScriptRoot $root
. (Join-Path $modules 'EpplusLoader.ps1')
. (Join-Path $modules 'CsvBundle.ps1')
. (Join-Path $modules 'RuleBank.ps1')
. (Join-Path $modules 'RuleEngine.ps1')

if (-not (Test-Path -LiteralPath $CsvPath)) {
    Write-Error "CSV not found: $CsvPath"
    exit 1
}

$bundle = $null
try { $bundle = Get-TestsSummaryBundle -Path $CsvPath } catch { Write-Warning "Get-TestsSummaryBundle failed: $($_.Exception.Message)" }
$assayMap = $null
if ($Config -and $Config.AssayMap) { $assayMap = $Config.AssayMap }
elseif ($Config -and $Config.SlangAssay) { $assayMap = $Config.SlangAssay }

$ctx  = New-AssayRuleContext -Bundle $bundle -AssayMap $assayMap -RuleBank $global:RuleBank
$eval = Invoke-AssayRuleEngine -Context $ctx -RuleBank $global:RuleBank

Write-Host ("CSV: {0}" -f (Split-Path $CsvPath -Leaf))
Write-Host ("Assay: {0}  (Resolved: {1}, Source: {2})" -f $ctx.AssayRaw, $ctx.AssayCanonical, $ctx.MatchSource)
Write-Host ("Total tests: {0}" -f $ctx.TotalTests)
Write-Host ("Overall: {0}" -f $eval.OverallStatus)
Write-Host ""
Write-Host "Findings:"
$idx = 0
foreach ($f in ($eval.Findings | Select-Object -First $Top)) {
    $idx++
    Write-Host (" {0}. [{1}] {2} - {3} (Affected={4})" -f $idx, $f.Severity, $f.RuleId, $f.Message, $f.AffectedCount)
}

if ($ShowAffected) {
    Write-Host ""
    Write-Host "Affected tests (first 10):"
    $i = 0
    foreach ($row in ($eval.AffectedTests | Select-Object -First 10)) {
        $i++
        $fmt = " {0}. Row {1} Sample={2} Cartridge={3} Status={4} Tags={5}"
        Write-Host ($fmt -f $i, $row.RowIndex, $row.SampleID, $row.CartridgeSN, $row.Status, ($row.FailureTags -join ','))
    }
    if ($eval.AffectedTestsTruncated -gt 0) {
        Write-Host (" ... truncated: +{0} more" -f $eval.AffectedTestsTruncated)
    }
}
