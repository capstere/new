#requires -Version 5.1
Set-StrictMode -Version Latest

<#
.SYNOPSIS
    Minimal, UI-safe logging + audit helpers for DocMerge (WinForms + EPPlus).
.DESCRIPTION
    - Keeps logging behavior consistent across modules.
    - Never throws for non-critical I/O (log files are "best effort").
    - Supports WinForms TextBox output when available.
    - Designed to minimize global state: only reads $global:LogPath when present.
.NOTES
    Why this exists:
      We want consistent audit trail and troubleshooting without sprinkling try/catch
      blocks everywhere. This module centralizes "best effort" logging so failures
      never break report generation.
#>

# ---- Internal module state (script scope, not global) ----
$script:LogControl = $null          # WinForms TextBox (multiline) or compatible
$script:MinLevel = 'Info'           # can be adjusted if needed

function Set-LogOutputControl {
    <#
    .SYNOPSIS
        Registers the WinForms TextBox (or similar control) where log lines should be appended.
    .PARAMETER Control
        A control supporting AppendText(). If InvokeRequired exists, thread-safe appending is used.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$false)]
        $Control
    )
    $script:LogControl = $Control
}

function Set-LogMinLevel {
    <#
    .SYNOPSIS
        Sets the minimum level that will be written to UI/file.
    .PARAMETER Level
        Info, Warn, Error
    #>
    [CmdletBinding()]
    param(
        [ValidateSet('Info','Warn','Error')]
        [string]$Level = 'Info'
    )
    $script:MinLevel = $Level
}

function Test-ShouldLog {
    param(
        [ValidateSet('Info','Warn','Error')]
        [string]$Level
    )
    $order = @{ Info = 1; Warn = 2; Error = 3 }
    return ($order[$Level] -ge $order[$script:MinLevel])
}

function Write-LogFileLine {
    param([string]$Line)
    # Best effort: never throw.
    try {
        if ($global:LogPath -and ($global:LogPath -is [string]) -and $global:LogPath.Trim()) {
            $dir = Split-Path -Parent $global:LogPath
            if ($dir -and -not (Test-Path -LiteralPath $dir)) {
                New-Item -ItemType Directory -Path $dir -Force | Out-Null
            }
            Add-Content -Path $global:LogPath -Value $Line -Encoding UTF8 -ErrorAction SilentlyContinue
        }
    } catch { }
}

function Append-ToUi {
    param([string]$Text)
    # Best effort: never throw.
    try {
        $ctl = $script:LogControl
        if (-not $ctl) { return }

        $append = {
            param($c,$t)
            try { $c.AppendText($t) } catch { }
        }

        # Thread-safe if control supports InvokeRequired
        try {
            if ($ctl.PSObject.Properties.Match('InvokeRequired').Count -gt 0 -and $ctl.InvokeRequired) {
                $null = $ctl.BeginInvoke($append, @($ctl,$Text))
            } else {
                & $append $ctl $Text
            }
        } catch {
            # Fallback
            try { $ctl.AppendText($Text) } catch { }
        }
    } catch { }
}

function Gui-Log {
    <#
    .SYNOPSIS
        Writes a timestamped line to UI log box (if registered) and to log file (if configured).
    .PARAMETER Message
        Log message.
    .PARAMETER Level
        Info/Warn/Error.
    .PARAMETER Immediate
        Calls Application.DoEvents() to flush UI (use sparingly).
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$Message,
        [ValidateSet('Info','Warn','Error')][string]$Level = 'Info',
        [switch]$Immediate
    )

    if (-not (Test-ShouldLog -Level $Level)) { return }

    $ts = Get-Date -Format 'HH:mm:ss'
    $line = "[$ts][$Level] $Message"
    Append-ToUi ($line + [Environment]::NewLine)
    Write-LogFileLine $line

    if ($Immediate) {
        try { [System.Windows.Forms.Application]::DoEvents() } catch { }
    }
}

function Add-AuditEntry {
    <#
    .SYNOPSIS
        Appends a compact audit record to the shared log file (if configured).
    .DESCRIPTION
        This is intentionally "best effort": failures never interrupt the workflow.
        Uses semicolon separated values to match common Swedish Excel parsing.
    #>
    [CmdletBinding()]
    param(
        [string]$Lsp,
        [string]$Assay,
        [string]$BatchNumber,
        [int]$TestCount,
        [string]$Status,
        [string]$ReportPath
    )

    try {
        $when = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
        $user = $env:USERNAME
        $line = ("{0};{1};{2};{3};{4};{5};{6};{7}" -f $when,$user,($Lsp -as [string]),($Assay -as [string]),($BatchNumber -as [string]),$TestCount,($Status -as [string]),($ReportPath -as [string]))
        Write-LogFileLine $line
    } catch { }
}

Export-ModuleMember -Function Set-LogOutputControl, Set-LogMinLevel, Gui-Log, Add-AuditEntry
