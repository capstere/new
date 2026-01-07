#requires -Version 5.1

function Read-AllLinesAuto {
  param([string]$Path)

  $fs = [IO.File]::Open($Path, [IO.FileMode]::Open, [IO.FileAccess]::Read, [IO.FileShare]::ReadWrite)
  try {
    $sr = New-Object IO.StreamReader($fs, [Text.Encoding]::UTF8, $true) # BOM aware
    $lines = New-Object System.Collections.Generic.List[string]
    while (-not $sr.EndOfStream) { $lines.Add($sr.ReadLine()) }
    $sr.Close()
    return ,$lines.ToArray()
  } finally {
    $fs.Dispose()
  }
}

function Normalize-HeaderName {
  param([string]$s)
  if ([string]::IsNullOrWhiteSpace($s)) { return '' }
  $t = $s.Trim().ToLowerInvariant()
  $t = $t.Replace('_',' ')
  $t = ($t -replace '[\(\)]',' ')
  $t = ($t -replace '\s+',' ').Trim()
  return $t
}

function Count-SeparatorsOutsideQuotes {
  param([string]$line, [char]$sep)
  $inQ = $false
  $cnt = 0
  for ($i=0; $i -lt $line.Length; $i++) {
    $ch = $line[$i]
    if ($ch -eq '"') {
      # handle escaped quote ""
      if ($inQ -and $i+1 -lt $line.Length -and $line[$i+1] -eq '"') { $i++; continue }
      $inQ = -not $inQ
      continue
    }
    if (-not $inQ -and $ch -eq $sep) { $cnt++ }
  }
  return $cnt
}

function Detect-Delimiter {
  param([string]$line)
  $cComma = Count-SeparatorsOutsideQuotes -line $line -sep ','
  $cSemi  = Count-SeparatorsOutsideQuotes -line $line -sep ';'
  $cTab   = Count-SeparatorsOutsideQuotes -line $line -sep "`t"

  # prefer the one with highest count
  if ($cSemi -gt $cComma -and $cSemi -ge $cTab) { return ';' }
  if ($cTab  -gt $cComma -and $cTab  -ge $cSemi) { return "`t" }
  return ','
}

function Parse-CsvLine {
  param([string]$line, [string]$delim)

  $vals = New-Object System.Collections.Generic.List[string]
  $sb = New-Object System.Text.StringBuilder
  $inQ = $false
  $d = [char]$delim

  for ($i=0; $i -lt $line.Length; $i++) {
    $ch = $line[$i]

    if ($ch -eq '"') {
      if ($inQ -and $i+1 -lt $line.Length -and $line[$i+1] -eq '"') {
        [void]$sb.Append('"'); $i++; continue
      }
      $inQ = -not $inQ
      continue
    }

    if (-not $inQ -and $ch -eq $d) {
      $vals.Add($sb.ToString()) | Out-Null
      $sb.Length = 0
      continue
    }

    [void]$sb.Append($ch)
  }

  $vals.Add($sb.ToString()) | Out-Null
  return ,$vals.ToArray()
}

function Find-HeaderRowIndex {
  param([string[]]$lines, [int]$maxScan = 80)

  $need = @(
    'assay',
    'assay version',
    'reagent lot id',
    'sample id',
    'cartridge s/n'
  )

  $max = [Math]::Min($lines.Count-1, $maxScan-1)
  for ($i=0; $i -le $max; $i++) {
    $line = $lines[$i]
    if ([string]::IsNullOrWhiteSpace($line)) { continue }
    $delim = Detect-Delimiter -line $line
    $cols = Parse-CsvLine -line $line -delim $delim

    $hits = 0
    foreach ($c in $cols) {
      $h = Normalize-HeaderName $c
      if ($need -contains $h) { $hits++ }
    }

    if ($hits -ge 3) { return @{ Index=$i; Delim=$delim } }
  }

  # fallback: first non-empty line
  for ($i=0; $i -lt $lines.Count; $i++) {
    if (-not [string]::IsNullOrWhiteSpace($lines[$i])) { return @{ Index=$i; Delim=(Detect-Delimiter -line $lines[$i]) } }
  }
  return @{ Index=0; Delim=',' }
}

function Find-DataStartRowIndex {
  param([string[]]$lines, [int]$headerIndex, [string]$delim, [string[]]$headers)

  $hCount = $headers.Count
  for ($i=$headerIndex+1; $i -lt $lines.Count; $i++) {
    $line = $lines[$i]
    if ([string]::IsNullOrWhiteSpace($line)) { continue }
    $cols = Parse-CsvLine -line $line -delim $delim

    # heuristic: must have many columns and contain something non-empty
    $nonEmpty = ($cols | Where-Object { -not [string]::IsNullOrWhiteSpace($_) }).Count
    if ($cols.Count -ge [Math]::Max(5, [int]($hCount*0.5)) -and $nonEmpty -ge 3) {
      return $i
    }
  }
  return ($headerIndex + 1)
}

function Get-TestsSummaryBundle_Core {
  [CmdletBinding()]
  param([Parameter(Mandatory=$true)][string]$Path)

  $lines = Read-AllLinesAuto -Path $Path
  if (-not $lines -or $lines.Count -eq 0) {
    return [ordered]@{
      Path              = $Path
      Lines             = $lines
      Delimiter         = ','
      HeaderRowIndex    = 0
      DataStartRowIndex = 0
      Headers           = @()
      HeaderIndex       = @{}
      RowCount          = 0
    }
  }
  $hdr = Find-HeaderRowIndex -lines $lines
  $headerRow = [int]$hdr.Index
  $delim = [string]$hdr.Delim

  $headers = Parse-CsvLine -line $lines[$headerRow] -delim $delim
  $headerIndex = @{}

  for ($i=0; $i -lt $headers.Count; $i++) {
    $norm = Normalize-HeaderName $headers[$i]
    if (-not [string]::IsNullOrWhiteSpace($norm) -and -not $headerIndex.ContainsKey($norm)) {
      $headerIndex[$norm] = $i
    }
  }

  $dataStart = Find-DataStartRowIndex -lines $lines -headerIndex $headerRow -delim $delim -headers $headers

  return [ordered]@{
    Path            = $Path
    Lines           = $lines
    Delimiter       = $delim
    HeaderRowIndex  = $headerRow
    DataStartRowIndex = $dataStart
    Headers         = $headers
    HeaderIndex     = $headerIndex
    RowCount        = ($lines.Count - $dataStart)
  }
}


# --- Bundle caching (read once per file stamp) -------------------------------
if (-not $script:CsvBundleCache) { $script:CsvBundleCache = @{} }

function Get-TestsSummaryBundle {
  [CmdletBinding()]
  param([Parameter(Mandatory=$true)][string]$Path)

  if (-not (Test-Path -LiteralPath $Path)) { throw ("CSV not found: {0}" -f $Path) }

  $fi = Get-Item -LiteralPath $Path -ErrorAction SilentlyContinue
  $key = $Path.ToLowerInvariant()
  $stamp = if ($fi) { $fi.LastWriteTimeUtc.Ticks } else { 0 }

  $cached = $script:CsvBundleCache[$key]
  if ($cached -and $cached.Stamp -eq $stamp -and $cached.Bundle) { return $cached.Bundle }

  $b = Get-TestsSummaryBundle_Core -Path $Path
  $script:CsvBundleCache[$key] = [pscustomobject]@{ Stamp = $stamp; Bundle = $b }
  return $b
}

function Get-AssayFromTestsSummaryBundle {
  [CmdletBinding()]
  param(
    [Parameter(Mandatory=$true)]$Bundle
  )

  if (-not $Bundle) { return $null }

  $lines = $Bundle.Lines
  if (-not $lines -or $lines.Count -eq 0) { return $null }

  $dataIdx = 0
  $delim = ','
  try { if ($Bundle.DataStartRowIndex -ge 0) { $dataIdx = [int]$Bundle.DataStartRowIndex } } catch { $dataIdx = 0 }
  try { if ($Bundle.Delimiter) { $delim = [string]$Bundle.Delimiter } } catch { $delim = ',' }

  $idxAssay = 0
  try {
    if ($Bundle.HeaderIndex -and $Bundle.HeaderIndex.ContainsKey('assay')) {
      $idxAssay = [int]$Bundle.HeaderIndex['assay']
    }
  } catch { $idxAssay = 0 }

  for ($i=$dataIdx; $i -lt $lines.Count; $i++) {
    $ln = $lines[$i]
    if ([string]::IsNullOrWhiteSpace($ln)) { continue }
    $cols = Parse-CsvLine -line $ln -delim $delim
    if (-not $cols -or $cols.Count -le $idxAssay) { continue }
    $a = ($cols[$idxAssay] + '').Trim().Trim('"')
    if ($a -and $a -notmatch '^(?i)assay$') { return $a }
  }

  return $null
}
