function Start-PerfTimer {
    param(
        [Parameter(Mandatory=$true)][string]$Label
    )
    $timer = [System.Diagnostics.Stopwatch]::StartNew()
    return [PSCustomObject]@{
        Timer = $timer
        Label = $Label
    }
}

function Stop-PerfTimer {
    param(
        [Parameter(Mandatory=$true)][pscustomobject]$Perf
    )
    try {
        if ($Perf -and $Perf.Timer) { $Perf.Timer.Stop() }
        $elapsed = if ($Perf -and $Perf.Timer) { $Perf.Timer.Elapsed } else { $null }
        if ($elapsed) {
            $seconds = [math]::Round($elapsed.TotalSeconds, 2)
            $msg = " $($Perf.Label) tog $seconds s"
            try { Gui-Log $msg 'Info' } catch { Write-Host $msg }
        }
    } catch {
        # Swallow any timing/logging errors to avoid affecting main code paths
    }
}

Export-ModuleMember -Function Start-PerfTimer, Stop-PerfTimer