<#
    Module: SignatureHelpers.ps1
    Purpose: Helpers for signature parsing, confirmation, and sign-off consistency across Seal Test files.
    Platform: PowerShell 5.1, EPPlus 4.5.3.3 (.NET 3.5)
    Notes:
      - Contains only signature-related helpers; data IO lives in DataHelpers.
#>

<#
    .SYNOPSIS
        Performs a lightweight syntax check of a signature string.

    .DESCRIPTION
        Splits the supplied text by commas to identify name, sign, and date parts.
        Used before writing signature cells to Seal Test sheets.

    .PARAMETER Text
        Signature string in format "Full Name, Sign, Date".

    .OUTPUTS
        PSCustomObject with Raw/Name/Sign/Date/Parts/DateOk/LooksOk.
#>
function Test-SignatureFormat {
    param([string]$Text)
    $raw = ($Text + '')
    $trim = $raw.Trim()
    $parts = $trim -split '\s*,\s*'
    $name = if ($parts.Count -ge 1) { $parts[0] } else { '' }
    $sign = if ($parts.Count -ge 2) { $parts[1] } else { '' }
    $date = if ($parts.Count -ge 3) { $parts[2] } else { '' }
    $dateOk = $false
    if ($date) { if ($date -match '^\d{4}-\d{2}-\d{2}$' -or $date -match '^\d{8}$') { $dateOk = $true } }
    [pscustomobject]@{ Raw=$raw; Name=$name; Sign=$sign; Date=$date; Parts=$parts.Count; DateOk=$dateOk; LooksOk=($name -ne '' -and $sign -ne '') }
}

<#
    .SYNOPSIS
        Asks the operator to confirm the parsed signature string.

    .DESCRIPTION
        Presents a WinForms MessageBox showing detected Name/Sign/Date segments to
        avoid accidental sign-off mistakes.

    .PARAMETER Text
        Signature string to parse and confirm.

    .OUTPUTS
        Boolean indicating whether the user confirmed.
#>
function Confirm-SignatureInput { param([string]$Text)
    $info = Test-SignatureFormat $Text
    $hint = @()
    if (-not $info.Name) { $hint += '• Namn verkar saknas' }
    if (-not $info.Sign) { $hint += '• Signatur verkar saknas' }
    $msg = "Har du skrivit korrekt 'Print Full Name, Sign, and Date'?
Text: $($info.Raw)
Tolkning:
  • Namn   : $($info.Name)
  • Sign   : $($info.Sign)
  • Datum  : $($info.Date)
" + ($(if ($hint.Count){ "Obs:`n  " + ($hint -join "`n  ") } else { "Ser bra ut." }))
    $res = [System.Windows.Forms.MessageBox]::Show($msg, "Bekräfta signatur", 'YesNo', 'Question')
    return ($res -eq 'Yes')
}

<#
    .SYNOPSIS
        Normalizes signature text for comparison (case/spacing insensitive).

    .PARAMETER s
        Raw signature string.

    .OUTPUTS
        Lower-cased, comma-trimmed signature representation.
#>
function Normalize-Signature {
    param([string]$s)
    if (-not $s) { return '' }
    $x = $s.Trim().ToLowerInvariant()
    $x = [regex]::Replace($x, '\s+', ' ')
    $x = $x -replace '\s*,\s*', ','
    return $x
}
 
<#
    .SYNOPSIS
        Builds a normalized signature set across all Seal Test data sheets.

    .DESCRIPTION
        Iterates the data sheets (skips "Worksheet Instructions") and collects
        signatures from B47 to detect mismatches between NEG and POS workbooks.

    .PARAMETER Pkg
        EPPlus package for a Seal Test workbook.

    .OUTPUTS
        PSCustomObject with RawFirst, NormSet (HashSet), Occurrence map, and RawByNorm.

    .NOTES
        Heavy operation: loops all data sheets to extract signatures.
#>
function Get-SignatureSetForDataSheets {
    param([OfficeOpenXml.ExcelPackage]$Pkg)
    $result = [pscustomobject]@{
        RawFirst  = $null
        NormSet   = New-Object 'System.Collections.Generic.HashSet[string]'
        Occ       = @{}
        RawByNorm = @{} 
    }
    if (-not $Pkg) { return $result }

    foreach ($ws in ($Pkg.Workbook.Worksheets | Where-Object { $_.Name -ne 'Worksheet Instructions' })) {
        $h3 = ($ws.Cells['H3'].Text + '').Trim()
        if ($h3 -match '^[0-9]') {
            $raw = ($ws.Cells['B47'].Text + '').Trim()
            if ($raw) {
                $norm = Normalize-Signature $raw
                [void]$result.NormSet.Add($norm)
                if (-not $result.RawFirst) { $result.RawFirst = $raw }
                if (-not $result.Occ.ContainsKey($norm)) {
                    $result.Occ[$norm] = New-Object 'System.Collections.Generic.List[string]'
                }
                if (-not $result.RawByNorm.ContainsKey($norm)) {
                    $result.RawByNorm[$norm] = $raw
                }
                [void]$result.Occ[$norm].Add($ws.Name)
            }
        } elseif ([string]::IsNullOrWhiteSpace($h3) -or $h3 -match '^(?i)(N\/?A|NA|Tomt( innehåll)?)$') {
            break
        }
    }
    return $result
}
