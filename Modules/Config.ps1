param(
    [string]$ScriptRoot = (Split-Path -Parent $MyInvocation.MyCommand.Path)
)
# === Inställningar ===
$ScriptVersion = "v45.1.0"
$RootPaths = @(
    'N:\QC\QC-1\IPT\Skiftspecifika dokument\PQC analyst\JESPER\Scripts\Tests',
    'N:\QC\QC-1\IPT\3. IPT - KLART FÖR SAMMANSTÄLLNING',
    'N:\QC\QC-1\IPT\4. IPT - KLART FÖR GRANSKNING'
)
$ikonSokvag = Join-Path $ScriptRoot "icon.png"
$UtrustningListPath = "N:\QC\QC-1\IPT\Skiftspecifika dokument\PQC analyst\JESPER\Scripts\Click Less Project v03\Utrustninglista5.0.xlsx"
$RawDataPath        = "N:\QC\QC-1\IPT\KONTROLLPROVSFIL - Version 2.5.xlsm"
$SlangAssayPath     = (
    @('N:\\QC\\QC-1\\IPT\\Skiftspecifika dokument\\PQC analyst\\JESPER\\Scripts\\Click Less Project v03\\Click Less Project\\slangassay.xlsx',
      (Join-Path $ScriptRoot 'Click Less Project\\slangassay.xlsx'),
      (Join-Path $ScriptRoot 'slangassay.xlsx')
    ) | Where-Object { $_ -and (Test-Path -LiteralPath $_) } | Select-Object -First 1
)
$OtherScriptPath = ''
$Script1Path  = 'N:\QC\QC-1\IPT\Skiftspecifika dokument\PQC analyst\JESPER\Kontrollprovsfil 2025\Script Raw Data\Kontrollprovsfil_EPPlus_2025.ps1'
$Script2Path  = 'N:\QC\QC-1\IPT\Skiftspecifika dokument\PQC analyst\JESPER\Scripts\Click Less Project\rename-GUI.bat'
$Script3Path  = 'N:\QC\QC-1\IPT\Skiftspecifika dokument\PQC analyst\JESPER\Scripts\Click Less Project\rename-GUI.bat'
$env:PNPPOWERSHELL_UPDATECHECK = "Off"
$global:SP_ClientId   = "23715695-a9a6-4f32-af7b-4cd164e0f1f9"
$global:SP_Tenant     = "danaher.onmicrosoft.com"
$global:SP_CertBase64 = "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"
$global:SP_SiteUrl    = "https://danaher.sharepoint.com/sites/CEP-Sweden-Production-Management"
# === Centraliserad konfiguration ===
$Config = [ordered]@{
    CsvPath       = ''           # Sökväg till CSV-fil
    SealNegPath   = ''           # Sökväg till Seal Test NEG
    SealPosPath   = ''           # Sökväg till Seal Test POS
    WorksheetPath = ''           # Sökväg till LSP worksheet (Worksheet.xlsx)
    SiteUrl      = $global:SP_SiteUrl
    Tenant       = $global:SP_Tenant
    ClientId     = $global:SP_ClientId
    Certificate  = $global:SP_CertBase64
    EpplusDllPath = (
        @('N:\QC\QC-1\IPT\Skiftspecifika dokument\PQC analyst\JESPER\Scripts\Modules\EPPlus\EPPlus.4.5.3.3\lib\net40\EPPlus.dll',
          'N:\QC\QC-1\IPT\Skiftspecifika dokument\PQC analyst\JESPER\Scripts\Modules\EPPlus\EPPlus.4.5.3.3\.5.3.3\lib\net35\EPPlus.dll',
          (Join-Path $ScriptRoot 'EPPlus.dll')
        ) | Where-Object { $_ -and (Test-Path -LiteralPath $_) } | Select-Object -First 1
    )
    EpplusVersion = '4.5.3.3'
}
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

$SharePointBatchLinkTemplate = 'https://danaher.sharepoint.com/sites/CEP-Sweden-Production-Management/Lists/Cepheid%20%20Production%20orders/ROBAL.aspx?viewid=6c9e53c9-a377-40c1-a154-13a13866b52b&view=7&q={BatchNumber}'
$DevLogDir = Join-Path $ScriptRoot 'Loggar'
if (-not (Test-Path $DevLogDir)) { New-Item -ItemType Directory -Path $DevLogDir -Force | Out-Null }
$global:LogPath = Join-Path $DevLogDir ("$($env:USERNAME)_$(Get-Date -Format 'yyyyMMdd_HHmmss').txt")
function Test-Config {
    $result = [pscustomobject]@{
        Ok       = $true
        Errors   = New-Object System.Collections.Generic.List[object]
        Warnings = New-Object System.Collections.Generic.List[object]
    }
    try {
        $templatePath = Join-Path $ScriptRoot 'output_template-v4.xlsx'
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
    return $result}