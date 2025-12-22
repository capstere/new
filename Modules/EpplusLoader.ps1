function Write-UiLog {
    param([string]$Msg, [string]$Level = 'Info')
    if (Get-Command Gui-Log -ErrorAction SilentlyContinue) {
        try { Gui-Log $Msg $Level } catch { Write-Host "[$Level] $Msg" }
    } else {
        $prefix = switch ($Level) { 'Error' { '[ERROR]' } 'Warn' { '[WARN]' } default { '[INFO]' } }
        Write-Host "$prefix $Msg"
    }
}

function Ensure-EPPlus {
    [CmdletBinding()]
    param(
        [string]$Version = '4.5.3.3',
        # Primary (known-good) source path(s) on N:\ in your environment
        [string]$SourceDllPath = 'N:\QC\QC-1\IPT\Skiftspecifika dokument\PQC analyst\JESPER\Scripts\Modules\EPPlus\EPPlus.4.5.3.3\lib\net40\EPPlus.dll',
        [string]$HintPath = $null,
        [string]$CacheRoot = "$env:ProgramData\EPPlus\$Version"
    )

    $candidatePaths = New-Object System.Collections.Generic.List[string]

    # 0) Explicit hint (from Config.EpplusDllPath or caller)
    if (-not [string]::IsNullOrWhiteSpace($HintPath)) { $candidatePaths.Add([string]$HintPath) }

    # 1) Explicit source (N:\)
    if (-not [string]::IsNullOrWhiteSpace($SourceDllPath)) { $candidatePaths.Add([string]$SourceDllPath) }

    # 1b) Also try net35 variant (some installs use it)
    $net35 = 'N:\QC\QC-1\IPT\Skiftspecifika dokument\PQC analyst\JESPER\Scripts\Modules\EPPlus\EPPlus.4.5.3.3\.5.3.3\lib\net35\EPPlus.dll'
    if (-not [string]::IsNullOrWhiteSpace($net35)) { $candidatePaths.Add($net35) }

    # 2) Same folder as this project (repo root) / same folder as this file
    try {
        if ($PSScriptRoot) {
            $candidatePaths.Add((Join-Path $PSScriptRoot 'EPPlus.dll'))
        }
    } catch {}

    # 3) PowerShell module folders
    try {
        $userModRoot = Join-Path ([Environment]::GetFolderPath('MyDocuments')) 'WindowsPowerShell\Modules'
        if (Test-Path $userModRoot) {
            Get-ChildItem -Path (Join-Path $userModRoot 'EPPlus') -Directory -ErrorAction SilentlyContinue | ForEach-Object {
                $candidatePaths.Add((Join-Path $_.FullName 'lib\net45\EPPlus.dll'))
                $candidatePaths.Add((Join-Path $_.FullName 'lib\net40\EPPlus.dll'))
            }
        }

        $pf64 = [Environment]::GetFolderPath('ProgramFiles')
        $pf86 = [Environment]::GetFolderPath('ProgramFilesX86')
        foreach ($pf in @($pf64, $pf86)) {
            if ([string]::IsNullOrWhiteSpace($pf)) { continue }
            $systemModRoot = Join-Path $pf 'WindowsPowerShell\Modules'
            if (Test-Path $systemModRoot) {
                Get-ChildItem -Path (Join-Path $systemModRoot 'EPPlus') -Directory -ErrorAction SilentlyContinue | ForEach-Object {
                    $candidatePaths.Add((Join-Path $_.FullName 'lib\net45\EPPlus.dll'))
                    $candidatePaths.Add((Join-Path $_.FullName 'lib\net40\EPPlus.dll'))
                }
            }
        }
    } catch {}

    # 4) Cache
    $cacheDll = $null
    try {
        if (-not [string]::IsNullOrWhiteSpace($CacheRoot)) {
            $cacheDll = Join-Path $CacheRoot 'EPPlus.dll'
            $candidatePaths.Add($cacheDll)
        }
    } catch {}

    # Return first existing
    foreach ($cand in $candidatePaths) {
        try {
            if (-not [string]::IsNullOrWhiteSpace($cand) -and (Test-Path -LiteralPath $cand)) {
                return $cand
            }
        } catch {}
    }

    # 5) NuGet fallback -> cache
    try {
        try { [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12 } catch {}

        $nugetUrl = "https://www.nuget.org/api/v2/package/EPPlus/$Version"
        $guid     = [Guid]::NewGuid().ToString()
        $tempDir  = Join-Path $env:TEMP "EPPlus_$guid"
        New-Item -ItemType Directory -Path $tempDir -Force | Out-Null
        $nupkgPath = Join-Path $tempDir "EPPlus.$Version.nupkg"

        Write-UiLog "Hämtar EPPlus $Version från NuGet (fallback)…" 'Info'
        Invoke-WebRequest -Uri $nugetUrl -OutFile $nupkgPath -UseBasicParsing -Headers @{ 'User-Agent' = 'DocMerge/1.0' } -ErrorAction Stop | Out-Null

        $extractDir = Join-Path $tempDir 'extracted'
        Expand-Archive -Path $nupkgPath -DestinationPath $extractDir -Force

        $dllCandidates = @(
            Join-Path $extractDir 'lib\net45\EPPlus.dll',
            Join-Path $extractDir 'lib\net40\EPPlus.dll'
        ) | Where-Object { Test-Path $_ }

        if (-not $dllCandidates -or $dllCandidates.Count -eq 0) {
            throw "Kunde inte hitta EPPlus.dll i nupkg (lib\\net45/net40)."
        }

        if (-not [string]::IsNullOrWhiteSpace($CacheRoot)) {
            New-Item -ItemType Directory -Path $CacheRoot -Force | Out-Null
            Copy-Item -Path $dllCandidates[0] -Destination $cacheDll -Force
            try { Unblock-File -Path $cacheDll } catch {}
            Write-UiLog "EPPlus kopierad till cache: $cacheDll" 'Info'
            return $cacheDll
        }

        # no cache root -> return extracted file
        return $dllCandidates[0]

    } catch {
        Write-UiLog "⚠️ EPPlus: Kunde inte hämta EPPlus ($Version): $($_.Exception.Message)" 'Warn'
    }

    return $null
}

function Load-EPPlus {
    [CmdletBinding()]
    param(
        [string]$HintPath = $null,
        [string]$Version  = '4.5.3.3'
    )

    # Already loaded?
    try {
        if ([System.AppDomain]::CurrentDomain.GetAssemblies() | Where-Object { $_.GetName().Name -eq 'EPPlus' -or $_.GetName().Name -eq 'OfficeOpenXml' }) {
            return $true
        }
    } catch {}

    $dllPath = Ensure-EPPlus -Version $Version -HintPath $HintPath
    if (-not $dllPath) {
        return $false
    }

    try {
        try { Unblock-File -Path $dllPath -ErrorAction SilentlyContinue } catch {}
        $bytes = [System.IO.File]::ReadAllBytes($dllPath)
        [System.Reflection.Assembly]::Load($bytes) | Out-Null

        # Sanity check: ExcelPackage type
        try {
            $pkg = New-Object OfficeOpenXml.ExcelPackage
            $pkg.Dispose()
        } catch {}

        return $true
    } catch {
        Write-UiLog "❌ EPPlus-fel: $($_.Exception.Message)" 'Error'
        return $false
    }
}

function Initialize-EPPlus {
    [CmdletBinding()]
    param(
        [string]$HintPath = $null,
        [string]$Version  = '4.5.3.3'
    )

    if (Load-EPPlus -HintPath $HintPath -Version $Version) {
        return $true
    }

    # Build a helpful error (similar to your previous loader output)
    $attempts = @(
        $HintPath,
        'N:\\QC\\QC-1\\IPT\\Skiftspecifika dokument\\PQC analyst\\JESPER\\Scripts\\Modules\\EPPlus\\EPPlus.4.5.3.3\\lib\\net40\\EPPlus.dll',
        'N:\\QC\\QC-1\\IPT\\Skiftspecifika dokument\\PQC analyst\\JESPER\\Scripts\\Modules\\EPPlus\\EPPlus.4.5.3.3\\.5.3.3\\lib\\net35\\EPPlus.dll'
    ) | Where-Object { $_ -and ($_ + '').Trim() } | Select-Object -Unique

    $msg = @()
    $msg += 'EPPlus.dll could not be loaded.'
    $msg += 'Attempted (top):'
    $msg += ($attempts | ForEach-Object { " - $_" })
    $msg += ''
    $msg += 'Fix options:'
    $msg += ' - Verify the N:\\...\\EPPlus.dll path exists (net40/net35), or'
    $msg += ' - Put EPPlus.dll in .\\EPPlus.dll (project root), or'
    $msg += ' - Set Config.EpplusDllPath to the full path of EPPlus.dll.'

    throw ($msg -join "`r`n")
}