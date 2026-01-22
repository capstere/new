function Get-OutputBox {
    try { return $script:outputBox } catch { return $null }
}

function Set-LogOutputControl {
    param([System.Windows.Forms.TextBox]$Control)
    $script:outputBox = $Control
}

function Gui-Log {
    param(
        [string]$Text,
        [ValidateSet('Info', 'Warn', 'Error')]
        [string]$Severity = 'Info',
        [string]$Category = 'UI',
        [hashtable]$Data,
        [switch]$Immediate
    )

    # --- Förbered loggrad ---
    $prefix = switch ($Severity) {
        'Warn'  { '⚠️' }
        'Error' { '❌' }
        default { 'ℹ️' }
    }

    $timestamp = (Get-Date).ToString('HH:mm:ss')
    $line = "[$timestamp] $prefix $Text"

    # --- UI‑append scriptblock ---
    $append = {
        $outputBox.AppendText("$line`r`n")
        $outputBox.SelectionStart = $outputBox.TextLength
        $outputBox.ScrollToCaret()
        $outputBox.Refresh()
    }

    # --- Hämta OutputBox ---
    $outputBox = Get-OutputBox

    if ($outputBox -and ($outputBox.IsDisposed -or -not $outputBox.IsHandleCreated)) {
        $outputBox = $null
    }

    if ($outputBox) {
        try {
            if ($outputBox.InvokeRequired) {
                $null = $outputBox.BeginInvoke(
                    [System.Windows.Forms.MethodInvoker]$append
                )
            } else {
                & $append
            }

            if ($Immediate) {
                [System.Windows.Forms.Application]::DoEvents()
            }
        }
        catch {
            Write-Host $line
        }
    }
    else {
        Write-Host $line
    }

    # --- Patch: strukturerad logg (JSONL) ---
    try {
        Write-StructuredLog `
            -Message  $Text `
            -Severity $Severity `
            -Category $Category `
            -Data     $Data
    } catch {}

    # --- Klassisk filbaserad logg ---
    if ($global:LogPath) {
        try {
            Add-Content -Path $global:LogPath -Value $line -Encoding UTF8
        }
        catch {
            Write-Host "Loggning misslyckades: $($_.Exception.Message)"
        }
    }
}

function Add-AuditEntry {
    param(
        [string]$Lsp,
        [string]$Assay,
        [string]$BatchNumber,
        [int]$TestCount,
        [string]$Status,
        [string]$ReportPath,
        [string]$AuditDir

    )
    try {
        if (-not $AuditDir) {
            if ($PSScriptRoot) { $AuditDir = Join-Path $PSScriptRoot 'audit' }
            else { $AuditDir = Join-Path $env:TEMP 'audit' }
        }

        if (-not (Test-Path -LiteralPath $AuditDir)) { New-Item -ItemType Directory -Path $AuditDir -Force | Out-Null }

        $date = (Get-Date).ToString('yyyyMMdd')
        $safeLsp = if ($Lsp) { $Lsp } else { 'NA' }
        $file = Join-Path $AuditDir ("Audit_{0}_{1}_{2}.csv" -f $date, $env:USERNAME, $safeLsp)

        $row = [pscustomobject]@{
            Timestamp  = (Get-Date).ToString('yyyy-MM-dd HH:mm:ss')
            Username   = $env:USERNAME
            LSP        = $Lsp
            Assay      = $Assay
            Batch      = $BatchNumber
            TestCount  = $TestCount
            Status     = if ($Status) { $Status } else { 'OK' }
            ReportPath = $ReportPath
        }
        $exists = Test-Path -LiteralPath $file
        $row | Export-Csv -Path $file -NoTypeInformation -Append:($exists) -Encoding UTF8
    } catch {
        Gui-Log "⚠️ Kunde inte skriva audit-fil: $($_.Exception.Message)" 'Warn'
    }
}

function Write-StructuredLog {
    param(
        [string]$Message,
        [ValidateSet('Info','Warn','Error')][string]$Severity = 'Info',
        [string]$Category = 'General',
        [hashtable]$Context
    )
    if (-not $global:LogPath) { return }
    $jsonPath = "$($global:LogPath).jsonl"
    $entry = [ordered]@{
        Timestamp = (Get-Date).ToString('o')
        Severity  = $Severity
        Category  = $Category
        Message   = $Message
    }
    if ($Context) {
        foreach ($k in $Context.Keys) { $entry[$k] = $Context[$k] }
    }
    try {
        Add-Content -Path $jsonPath -Value ($entry | ConvertTo-Json -Compress) -Encoding UTF8
    } catch {
        # tyst
    }
}
