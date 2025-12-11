<# 
    Module: Config.ps1
    Purpose: Central configuration and constants for the Validate-Assay tool.
    Platform: PowerShell 5.1, EPPlus 4.5.3.3 (.NET 3.5)
    Notes:
      - Contains ONLY configuration/constant values and configuration validation.
      - Business logic lives in helper modules; this file is dot-sourced first.
#>

param(
    [string]$ScriptRoot = (Split-Path -Parent $MyInvocation.MyCommand.Path)
)

# === App metadata ===
$ScriptVersion = "v45.1.0"   # Shown in UI header; keep in sync with validation scope.

# === Default search roots for LSP scanning (Solna environment specific) ===
$RootPaths = @(
    'N:\QC\QC-1\IPT\Skiftspecifika dokument\PQC analyst\JESPER\Scripts\Tests',
    'N:\QC\QC-1\IPT\3. IPT - KLART FÖR SAMMANSTÄLLNING',
    'N:\QC\QC-1\IPT\4. IPT - KLART FÖR GRANSKNING'
)

# === Local asset/template paths ===
$ikonSokvag        = Join-Path $ScriptRoot "icon.png"
$TemplatePath      = Join-Path $ScriptRoot 'output_template-v4.xlsx'
$UtrustningListPath = "N:\QC\QC-1\IPT\Skiftspecifika dokument\PQC analyst\JESPER\Scripts\Click Less Project\Utrustninglista5.0.xlsx"
$RawDataPath        = "N:\QC\QC-1\IPT\KONTROLLPROVSFIL - Version 2.4.xlsm"
$SlangAssayPath     = "N:\QC\QC-1\IPT\Skiftspecifika dokument\PQC analyst\JESPER\Scripts\Click Less Project\Click Less Project\slangassay.xlsx"

# === Legacy script shortcuts (kept for compatibility/logging only) ===
$OtherScriptPath = ''
$Script1Path  = 'N:\QC\QC-1\IPT\Skiftspecifika dokument\PQC analyst\JESPER\Kontrollprovsfil 2025\Script Raw Data\Kontrollprovsfil_EPPlus_2025.ps1'
$Script2Path  = 'N:\QC\QC-1\IPT\Skiftspecifika dokument\PQC analyst\JESPER\Scripts\Click Less Project\rename-GUI.bat'
$Script3Path  = 'N:\QC\QC-1\IPT\Skiftspecifika dokument\PQC analyst\JESPER\Scripts\Click Less Project\rename-GUI.bat'

# === SharePoint connection defaults (PnP) ===
$env:PNPPOWERSHELL_UPDATECHECK = "Off"
$global:SP_ClientId   = "INSERT MYSELF"
$global:SP_Tenant     = "danaher.onmicrosoft.com"
$global:SP_CertBase64 = "INSERT MYSELF"
$global:SP_SiteUrl    = "https://danaher.sharepoint.com/sites/CEP-Sweden-Production-Management"
$SharePointBatchLinkTemplate = 'https://danaher.sharepoint.com/sites/CEP-Sweden-Production-Management/Lists/Cepheid%20%20Production%20orders/ROBAL.aspx?viewid=6c9e53c9-a377-40c1-a154-13a13866b52b&view=7&q={BatchNumber}'

# === Centraliserad konfiguration ===
$Config = [ordered]@{
    CsvPath        = ''                                 # Aktuellt vald CSV
    SealNegPath    = ''                                 # Aktuellt vald Seal Test NEG
    SealPosPath    = ''                                 # Aktuellt vald Seal Test POS
    WorksheetPath  = ''                                 # Aktuellt vald Worksheet.xlsx
    SiteUrl        = $global:SP_SiteUrl                 # SharePoint plats
    Tenant         = $global:SP_Tenant                  # SharePoint tenant
    ClientId       = $global:SP_ClientId                # App-ID för PnP
    Certificate    = $global:SP_CertBase64              # Base64-kodad cert
    EpplusDllPath  = (Join-Path $ScriptRoot 'EPPlus.dll') # Lokal EPPlus dll
    EpplusVersion  = '4.5.3.3'                          # EPPlus version att kräva
    TemplatePath   = $TemplatePath                      # Mallfil för rapporten
    UtrustningPath = $UtrustningListPath                # Kopieras till Infinity/GX-bladet
    RawDataPath    = $RawDataPath                       # Kontrollprovsfil (Control Material)
    SlangAssayPath = $SlangAssayPath                    # Slang→Assay-mappning (används vid behov)
}

# === Instrument-LSP mapping (legacy Infinity/GX summary) ===
$script:GXINF_Map = @{
    'Infinity-VI'   = '847922'
    'Infinity-VIII' = '803094'
    'GX5'           = '750210,750211,750212,750213'
    'GX6'           = '750246,750247,750248,750249'
    'GX1'           = '709863,709864,709865,709866'
    'GX2'           = '709951,709952,709953,709954'
    'GX3'           = '710084,710085,710086,710087'
    'GX7'           = '750170,750171,750172,750213'
    'Infinity-I'    = '802069'
    'Infinity-III'  = '807363'
    'Infinity-V'    = '839032'
}

# === Loggning ===
$DevLogDir = Join-Path $ScriptRoot 'Loggar'
if (-not (Test-Path $DevLogDir)) { New-Item -ItemType Directory -Path $DevLogDir -Force | Out-Null }
$global:LogPath = Join-Path $DevLogDir ("$($env:USERNAME)_$(Get-Date -Format 'yyyyMMdd_HHmmss').txt")

<#
    .SYNOPSIS
        Verifies that mandatory templates and directories exist and can be written to.

    .DESCRIPTION
        Performs lightweight checks to ensure the core Excel templates and logging
        folders are reachable before the GUI is shown. Warnings are non-fatal to keep
        the operator flow intact; errors block startup.

    .OUTPUTS
        PSCustomObject with Ok (bool), Errors (List), Warnings (List).

    .NOTES
        This function is intentionally lightweight; heavy IO is deferred until build time.
#>
function Test-Config {
    $result = [pscustomobject]@{
        Ok       = $true
        Errors   = New-Object System.Collections.Generic.List[object]
        Warnings = New-Object System.Collections.Generic.List[object]
    }

    try {
        $templatePath = $Config.TemplatePath
        if (-not (Test-Path -LiteralPath $templatePath)) {
            $null = $result.Errors.Add("Mallfil saknas: $templatePath")
        }
    } catch {
        $null = $result.Errors.Add("Test-Config (template): $($_.Exception.Message)")
    }
    try {
        if (-not (Test-Path -LiteralPath $UtrustningListPath)) {
            $null = $result.Warnings.Add("Utrustningslista saknas: $UtrustningListPath")
        }
    } catch {
        $null = $result.Warnings.Add("Test-Config (utrustning): $($_.Exception.Message)")
    }

    try {
        if (-not (Test-Path -LiteralPath $RawDataPath)) {
            $null = $result.Warnings.Add("Kontrollprovsfil saknas: $RawDataPath")
        }
    } catch {
        $null = $result.Warnings.Add("Test-Config (rawdata): $($_.Exception.Message)")
    }
    try {
        if (-not (Test-Path -LiteralPath $SlangAssayPath)) {
            $null = $result.Warnings.Add("Slang/assay-tabell saknas: $SlangAssayPath")
        }
    } catch {
        $null = $result.Warnings.Add("Test-Config (slang/assay): $($_.Exception.Message)")
    }
    try {
        if (-not (Test-Path -LiteralPath $DevLogDir)) {
            New-Item -ItemType Directory -Path $DevLogDir -Force | Out-Null
        }
        $probe = Join-Path $DevLogDir "write_probe.txt"
        Set-Content -Path $probe -Value 'probe' -Encoding UTF8 -Force
        Remove-Item -LiteralPath $probe -Force -ErrorAction SilentlyContinue
    } catch {
        $null = $result.Warnings.Add("Kunde inte verifiera skrivning till loggmapp: $($_.Exception.Message)")
    }
    if ($result.Errors.Count -gt 0) { $result.Ok = $false }
    return $result
}
