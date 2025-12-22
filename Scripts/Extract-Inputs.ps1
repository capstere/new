#requires -Version 5.1
[CmdletBinding()]
param(
  [Parameter(Mandatory=$true)]
  [string[]]$ExcelPaths,
  [string]$OutDir = (Join-Path $PSScriptRoot '..\_debug')
)

. (Join-Path $PSScriptRoot '..\Modules\EpplusLoader.ps1')
Initialize-EPPlus

try { [OfficeOpenXml.ExcelPackage]::LicenseContext = [OfficeOpenXml.LicenseContext]::NonCommercial } catch {}

function Ensure-Dir([string]$p){ if(-not(Test-Path $p)){ New-Item -ItemType Directory -Path $p | Out-Null } }
function SafeName([string]$s){ if([string]::IsNullOrWhiteSpace($s)){return 'Sheet'}; return ($s -replace '[\\/:*?"<>|]+','_') }

Ensure-Dir $OutDir

foreach ($xp in $ExcelPaths) {
  if (-not (Test-Path $xp)) { throw "Missing: $xp" }
  $fi = Get-Item $xp
  $bookDir = Join-Path $OutDir (SafeName ([IO.Path]::GetFileNameWithoutExtension($fi.Name)))
  Ensure-Dir $bookDir

  $pkg = New-Object OfficeOpenXml.ExcelPackage($fi)
  try {
    foreach ($ws in $pkg.Workbook.Worksheets) {
      if (-not $ws.Dimension) { continue }
      $sheet = SafeName $ws.Name
      $dim = $ws.Dimension
      $maxRow = $dim.End.Row
      $maxCol = $dim.End.Column

      # Take row 1 as header (simple). If you need smarter, upgrade later.
      $headers = @()
      for ($c=1; $c -le $maxCol; $c++) { $headers += $ws.Cells[1,$c].Text }
      # ensure headers
      for ($i=0; $i -lt $headers.Count; $i++) {
        if ([string]::IsNullOrWhiteSpace($headers[$i])) { $headers[$i] = "Col$($i+1)" }
      }

      $rows = New-Object System.Collections.Generic.List[object]
      for ($r=2; $r -le $maxRow; $r++) {
        $obj = [ordered]@{}
        $any = $false
        for ($c=1; $c -le $maxCol; $c++) {
          $v = $ws.Cells[$r,$c].Text
          if (-not [string]::IsNullOrWhiteSpace($v)) { $any = $true }
          $obj[$headers[$c-1]] = $v
        }
        if ($any) { $rows.Add([pscustomobject]$obj) }
      }

      ($rows | ConvertTo-Json -Depth 12) | Out-File (Join-Path $bookDir "$sheet.json") -Encoding utf8

      # TSV
      $tsv = New-Object System.Text.StringBuilder
      [void]$tsv.AppendLine(($headers -join "`t"))
      foreach ($row in $rows) {
        $vals = @()
        foreach ($h in $headers) {
          $val = [string]$row.$h
          if ($null -eq $val) { $val = '' }
          $val = $val -replace "`r"," " -replace "`n"," " -replace "`t"," "
          $vals += $val
        }
        [void]$tsv.AppendLine(($vals -join "`t"))
      }
      $tsv.ToString() | Out-File (Join-Path $bookDir "$sheet.tsv") -Encoding utf8
    }
  } finally {
    $pkg.Dispose()
  }
}

Write-Host "[OK] Extracted to: $OutDir"
