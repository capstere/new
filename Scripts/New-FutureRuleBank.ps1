#requires -Version 5.1
[CmdletBinding()]
param(
  [string]$OutPath = (Join-Path $PSScriptRoot '..\Modules\RuleBank.Future.ps1')
)

$assays = @(
  'Xpert MTB-RIF Assay G4',
  'Xpert MTB-RIF JP IVD',
  'Xpert MTB-RIF Ultra',
  'Xpert MTB-XDR',
  'Xpress Flu IPT_EAT off',
  'Xpert Xpress Flu-RSV',
  'Xpert Xpress_SARS-CoV-2_Flu_RSV',
  'Xpress SARS-CoV-2_Flu_RSV plus',
  'Xpert Xpress CoV-2 plus',
  'Xpert CT_CE',
  'Xpert CT_NG',
  'Xpert C.difficile BT',
  'Xpert C.difficile G3',
  'Xpert Ebola EUA',
  'Xpert Ebola CE-IVD',
  'Xpert HPV v2 HR',
  'Xpert HPV HR',
  'Xpert HBV Viral Load',
  'Xpert HCV PQC',
  'Xpert_HCV Viral Load',
  'Xpert HCV VL Fingerstick',
  'Xpert_HIV-1 Viral Load',
  'Xpert HIV-1 Viral Load XC',
  'Xpert_HIV-1 Qual',
  'Xpert Xpress SARS-CoV-2 CE-IVD',
  'GXSACOMP-CE-10',
  'GXMRSA/SA-SSTI-CE',
  'GXSACOMP-CN-10',
  'Xpert SA Nasal Complete G3',
  'Xpert MRSA NxG',
  'Xpert Norovirus',
  'Xpert vanA vanB',
  'Xpert Xpress GBS',
  'Xpert GBS LB XC',
  'Xpert Xpress GBS US-IVD',
  'Xpert Xpress Strep A',
  'Xpert Xpress_Strep A',
  'Xpert Carba-R',
  'Xpert_Carba-R'
) | Select-Object -Unique

function Make-AssayKey([string]$name) {
  $t = $name.ToUpper()
  $t = ($t -replace '[^A-Z0-9]+','_').Trim('_')
  if ($t.Length -gt 50) { $t = $t.Substring(0,50) }
  return $t
}

$aliases = [ordered]@{}
foreach ($a in $assays) {
  $aliases[$a] = $a
  $aliases[($a -replace '\s+','_')] = $a
  $aliases[($a -replace '_',' ')] = $a
}

$today = (Get-Date).ToString('yyyy-MM-dd')
$sb = New-Object System.Text.StringBuilder

[void]$sb.AppendLine("﻿#requires -Version 5.1")
[void]$sb.AppendLine("<# RuleBank.Future.ps1 - generated skeleton (SchemaVersion=1). #>")
[void]$sb.AppendLine("")
[void]$sb.AppendLine("\$global:RuleBank = [ordered]@{")
[void]$sb.AppendLine("  Version       = '$today'")
[void]$sb.AppendLine("  SchemaVersion = 1")
[void]$sb.AppendLine("")
[void]$sb.AppendLine("  Baseline = [ordered]@{")
[void]$sb.AppendLine("    StandardBatch210 = [ordered]@{ BaselineTotal = 210 }")
[void]$sb.AppendLine("    SampleSizePolicy = [ordered]@{ MissingWarnAt=1; MissingErrorAt=15; ExtraWarnAt=1; ExtraErrorAt=15 }")
[void]$sb.AppendLine("  }")
[void]$sb.AppendLine("")
[void]$sb.AppendLine("  Global = [ordered]@{")
[void]$sb.AppendLine("    Identity   = [ordered]@{ Fields=@('Assay','Assay Version','Reagent Lot ID'); Severity='Error'; Message='Identity fields must be constant.' }")
[void]$sb.AppendLine("    Uniqueness = [ordered]@{ SampleId=[ordered]@{ Field='Sample ID'; Severity='Error' }; CartridgeSn=[ordered]@{ Field='Cartridge S/N'; Severity='Error' } }")
[void]$sb.AppendLine("    MaxPressure = [ordered]@{ Field='Max Pressure (PSI)'; MustBeLessThan=90; Severity='Error'; Message='Max Pressure must be < 90.' }")
[void]$sb.AppendLine("  }")
[void]$sb.AppendLine("")
[void]$sb.AppendLine("  ErrorBank = [ordered]@{")
[void]$sb.AppendLine("    ExtractRegex = '(?i)\\b(?:error|err)\\s*(?:code)?\\s*[:#]?\\s*(?<Code>\\d{4})\\b'")
[void]$sb.AppendLine("    Codes   = [ordered]@{}")
[void]$sb.AppendLine("    Special = @()")
[void]$sb.AppendLine("  }")
[void]$sb.AppendLine("")
[void]$sb.AppendLine("  AssayAliases = [ordered]@{")
foreach ($k in $aliases.Keys) {
  $kk = $k.Replace("'","''")
  $vv = $aliases[$k].Replace("'","''")
  [void]$sb.AppendLine("    '$kk' = '$vv'")
}
[void]$sb.AppendLine("  }")
[void]$sb.AppendLine("")
[void]$sb.AppendLine("  AssayProfiles = [ordered]@{")
foreach ($a in $assays) {
  $aa = $a.Replace("'","''")
  $key = (Make-AssayKey $a)
  [void]$sb.AppendLine("    '$aa' = [ordered]@{ AssayKey='$key'; DisplayName='$aa'; Mode='GlobalOnly'; Notes='TODO: add expected counts/results' }")
}
[void]$sb.AppendLine("    'default' = [ordered]@{ Mode='GlobalOnly'; Notes='Unknown assay -> global only' }")
[void]$sb.AppendLine("  }")
[void]$sb.AppendLine("}")

$sb.ToString() | Out-File -FilePath $OutPath -Encoding utf8
Write-Host "[OK] Wrote: $OutPath"
