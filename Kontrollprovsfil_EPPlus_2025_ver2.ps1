param(
    [string]$LogExcelPath = "N:\QC\QC-1\IPT\8. IPT - WR + Rework\1. PQC - Kontrollprovsfil - RÖR EJ -\LogFile_Kontrollprovsfil.xlsx",
    [string]$RawDataPath  = "N:\QC\QC-1\IPT\8. IPT - WR + Rework\1. PQC - Kontrollprovsfil - RÖR EJ -\Script Raw Data\raw_data.xlsx",
    [string]$OutputDir    = "N:\QC\QC-1\IPT\8. IPT - WR + Rework\1. PQC - Kontrollprovsfil - RÖR EJ -\Inventeringsrapport"
)

# =====================[ Kolumnschema för rådata ]=====================
$InventoryColumns = @{
    PN              = 1
    Lot             = 2
    Exp             = 3
    Qty             = 4
    LastUpdate      = 5
    Signature       = 6
    ProductStartCol = 7
    ProductEndCol   = 13
    Description     = 14
    LabbDescription = 15
    LabbCode        = 16
}

# =====================[ EPPlus bootstrap (robust) ]=====================

function Ensure-EPPlus {
    param(
        [string] $Version       = "4.5.3.3",
        [string] $SourceDllPath = "N:\QC\QC-1\IPT\Skiftspecifika dokument\PQC analyst\JESPER\Scripts\Modules\EPPlus\EPPlus.4.5.3.3\lib\net35\EPPlus.dll"
    )

    $candidatePaths = @()

    # 1) Direktkänd sökväg (t.ex. från N:)
    if (-not [string]::IsNullOrWhiteSpace($SourceDllPath)) {
        $candidatePaths += $SourceDllPath
    }

    # 2) EPPlus.dll bredvid skriptet
    $localScriptDll = Join-Path $PSScriptRoot 'EPPlus.dll'
    $candidatePaths += $localScriptDll

    # 3) EPPlus-moduler under CurrentUser
    $userModRoot = Join-Path ([Environment]::GetFolderPath('MyDocuments')) 'WindowsPowerShell\Modules'
    if (Test-Path $userModRoot) {
        Get-ChildItem -Path (Join-Path $userModRoot 'EPPlus') -Directory -ErrorAction SilentlyContinue | ForEach-Object {
            $candidatePaths += (Join-Path $_.FullName 'lib\net45\EPPlus.dll')
            $candidatePaths += (Join-Path $_.FullName 'lib\net40\EPPlus.dll')
            $candidatePaths += (Join-Path $_.FullName 'lib\net35\EPPlus.dll')
        }
    }

    # 4) EPPlus-moduler under Program Files
    $systemModRoot = Join-Path $env:ProgramFiles 'WindowsPowerShell\Modules'
    if (Test-Path $systemModRoot) {
        Get-ChildItem -Path (Join-Path $systemModRoot 'EPPlus') -Directory -ErrorAction SilentlyContinue | ForEach-Object {
            $candidatePaths += (Join-Path $_.FullName 'lib\net45\EPPlus.dll')
            $candidatePaths += (Join-Path $_.FullName 'lib\net40\EPPlus.dll')
            $candidatePaths += (Join-Path $_.FullName 'lib\net35\EPPlus.dll')
        }
    }

    foreach ($cand in $candidatePaths) {
        if (-not [string]::IsNullOrWhiteSpace($cand) -and (Test-Path -LiteralPath $cand)) {
            return $cand
        }
    }

    # 5) Ladda ner från NuGet om inget hittas
    $nugetUrl = "https://www.nuget.org/api/v2/package/EPPlus/$Version"
    Write-Host "EPPlus hittades inte lokalt. Försöker hämta $Version från NuGet..." -ForegroundColor Yellow

    try {
        # TLS 1.2
        try {
            if (-not ([Net.ServicePointManager]::SecurityProtocol.ToString() -match 'Tls12')) {
                [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
            }
        } catch {}

        $guid    = [Guid]::NewGuid().ToString()
        $tempDir = Join-Path $env:TEMP "EPPlus_$guid"
        New-Item -ItemType Directory -Path $tempDir -Force | Out-Null
        $zipPath = Join-Path $tempDir 'EPPlus.zip'

        $reqParams = @{
            Uri             = $nugetUrl
            OutFile         = $zipPath
            UseBasicParsing = $true
            Headers         = @{ 'User-Agent' = 'Kontrollprovsfil/EPPlus-Loader' }
        }
        Invoke-WebRequest @reqParams -ErrorAction Stop | Out-Null

        if (-not ([System.AppDomain]::CurrentDomain.GetAssemblies() | Where-Object { $_.GetName().Name -eq 'System.IO.Compression.FileSystem' })) {
            Add-Type -AssemblyName 'System.IO.Compression.FileSystem' -ErrorAction SilentlyContinue
        }

        [System.IO.Compression.ZipFile]::ExtractToDirectory($zipPath, $tempDir)

        $libRoot = Join-Path $tempDir 'lib'
        $dllCandidates = @()
        if (Test-Path $libRoot) {
            $dllCandidates += Get-ChildItem -Path (Join-Path $libRoot 'net45') -Filter 'EPPlus.dll' -ErrorAction SilentlyContinue
            $dllCandidates += Get-ChildItem -Path (Join-Path $libRoot 'net40') -Filter 'EPPlus.dll' -ErrorAction SilentlyContinue
            $dllCandidates += Get-ChildItem -Path (Join-Path $libRoot 'net35') -Filter 'EPPlus.dll' -ErrorAction SilentlyContinue
        }

        $dll = $dllCandidates | Select-Object -First 1
        if ($dll) {
            try {
                if (-not (Test-Path -LiteralPath $localScriptDll)) {
                    Copy-Item -Path $dll.FullName -Destination $localScriptDll -Force -ErrorAction SilentlyContinue
                }
            } catch {}

            return $localScriptDll
        }
        else {
            Write-Warning "❌ EPPlus: Ingen EPPlus.dll hittades i det nedladdade paketet."
        }
    }
    catch {
        Write-Warning "❌ EPPlus: Kunde inte hämta EPPlus ($Version): $($_.Exception.Message)"
    }

    Write-Warning "❌ EPPlus.dll hittades inte. Installera EPPlus $Version manuellt eller placera EPPlus.dll bredvid skriptet."
    return $null
}

function Load-EPPlus {
    # Redan inläst?
    if ([System.AppDomain]::CurrentDomain.GetAssemblies() | Where-Object { $_.GetName().Name -eq 'EPPlus' }) {
        return $true
    }

    $dllPath = Ensure-EPPlus -Version '4.5.3.3'
    if (-not $dllPath) {
        Write-Host "EPPlus kunde inte lokaliseras." -ForegroundColor Red
        return $false
    }

    try {
        # Läsa som bytes för att minska fil-låsning
        $bytes = [System.IO.File]::ReadAllBytes($dllPath)
        [System.Reflection.Assembly]::Load($bytes) | Out-Null
        Write-Host "EPPlus.dll inläst från: $dllPath" -ForegroundColor Green
        return $true
    }
    catch {
        Write-Warning "❌ EPPlus-fel vid inläsning från '$dllPath': $($_.Exception.Message)"
        return $false
    }
}

# Lokalt bypass bara för denna process
Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass -Force -ErrorAction SilentlyContinue

if (-not (Load-EPPlus)) {
    Write-Host "Kritisk komponent (EPPlus) kunde inte laddas. Skriptet avbryts." -ForegroundColor Red
    exit 1
}

# Färger
Add-Type -AssemblyName System.Drawing -ErrorAction SilentlyContinue

# =====================[ Hjälpfunktioner ]=====================

function Backup-RawDataFile {
    param(
        [string]$FilePath,
        [string]$BackupDir = "$PSScriptRoot\Backups"
    )
    if (-not (Test-Path $BackupDir)) {
        New-Item -ItemType Directory -Path $BackupDir | Out-Null
    }
    $timestamp  = (Get-Date).ToString("yyyyMMdd_HHmmss")
    $filename   = Split-Path $FilePath -Leaf
    $backupPath = Join-Path $BackupDir "$($filename)_$timestamp.bak"
    Copy-Item $FilePath $backupPath -Force
    Write-Host "Säkerhetskopia skapad: $backupPath" -ForegroundColor Green
}

function Initialize-LogFile {
    param(
        [string]$LogExcelPath
    )

    # Se till att katalogen finns
    $logDir = Split-Path $LogExcelPath -Parent
    if (-not (Test-Path $logDir)) {
        New-Item -ItemType Directory -Path $logDir -Force | Out-Null
    }

    if (-not (Test-Path $LogExcelPath)) {
        $file    = New-Object System.IO.FileInfo($LogExcelPath)
        $package = New-Object OfficeOpenXml.ExcelPackage($file)
        try {
            $updateSheet = $package.Workbook.Worksheets.Add("UpdateLog")
            $updateSheet.Cells[1,1].Value = "Timestamp"
            $updateSheet.Cells[1,2].Value = "UserName"
            $updateSheet.Cells[1,3].Value = "Message"
            $updateSheet.Cells[1,4].Value = "PN"
            $updateSheet.Cells[1,5].Value = "Lot"
            $updateSheet.Cells[1,6].Value = "Exp"
            $updateSheet.Cells[1,7].Value = "Qty"

            $otherSheet = $package.Workbook.Worksheets.Add("OtherLogs")
            $otherSheet.Cells[1,1].Value = "Timestamp"
            $otherSheet.Cells[1,2].Value = "UserName"
            $otherSheet.Cells[1,3].Value = "Message"

            # Header-stil för UpdateLog
            $hdrUpd = $updateSheet.Cells["A1:G1"]
            $hdrUpd.Style.Font.Bold = $true
            $hdrUpd.Style.Fill.PatternType = [OfficeOpenXml.Style.ExcelFillStyle]::Solid
            $hdrUpd.Style.Fill.BackgroundColor.SetColor([System.Drawing.Color]::LightGray)
            $updateSheet.View.FreezePanes(2,1)

            # Header-stil för OtherLogs
            $hdrOther = $otherSheet.Cells["A1:C1"]
            $hdrOther.Style.Font.Bold = $true
            $hdrOther.Style.Fill.PatternType = [OfficeOpenXml.Style.ExcelFillStyle]::Solid
            $hdrOther.Style.Fill.BackgroundColor.SetColor([System.Drawing.Color]::LightGray)
            $otherSheet.View.FreezePanes(2,1)

            $package.Save()
        }
        catch {
            Write-Host "Fel vid skapande av loggfil: $_" -ForegroundColor Red
        }
        finally {
            $package.Dispose()
        }
    }
}

# Öppna ExcelPackage med retry (fil kan vara låst)
function Open-ExcelPackageWithRetry {
    param(
        [string]$Path,
        [int]   $MaxAttempts  = 5,
        [int]   $DelaySeconds = 2
    )

    $attempt = 0
    while ($attempt -lt $MaxAttempts) {
        try {
            $file    = New-Object System.IO.FileInfo($Path)
            $package = New-Object OfficeOpenXml.ExcelPackage($file)
            return $package
        }
        catch {
            Write-Host "Kunde inte öppna '$Path' (försök $($attempt + 1)/$MaxAttempts). Fel: $($_.Exception.Message)" -ForegroundColor Yellow
            Start-Sleep -Seconds $DelaySeconds
            $attempt++
        }
    }

    Write-Host "❌ Gav upp efter $MaxAttempts försök att öppna '$Path'." -ForegroundColor Red
    return $null
}

# Låsning för loggfil (för att minska risk för korruption)
function Get-LogLockPath {
    param(
        [string]$LogExcelPath
    )
    return ($LogExcelPath + ".lock")
}

function Acquire-LogLock {
    param(
        [string]$LogExcelPath,
        [int]$MaxWaitSeconds = 120
    )

    $lockPath = Get-LogLockPath -LogExcelPath $LogExcelPath
    $start    = Get-Date

    while ($true) {
        try {
            $fs = [System.IO.File]::Open(
                $lockPath,
                [System.IO.FileMode]::CreateNew,
                [System.IO.FileAccess]::Write,
                [System.IO.FileShare]::None
            )

            $info  = "LockedBy=$([Environment]::UserName);LockedAt=$(Get-Date -Format o);Target=$LogExcelPath"
            $bytes = [System.Text.Encoding]::UTF8.GetBytes($info)
            $fs.Write($bytes, 0, $bytes.Length)
            $fs.Close()

            return $lockPath
        }
        catch {
            if ((Get-Date) - $start -gt [TimeSpan]::FromSeconds($MaxWaitSeconds)) {
                throw "Kunde inte få exklusivt logg-lås på $lockPath inom $MaxWaitSeconds sekunder."
            }
            Start-Sleep -Milliseconds 500
        }
    }
}

function Release-LogLock {
    param(
        [string]$LockPath
    )

    if ($LockPath -and (Test-Path $LockPath)) {
        Remove-Item -Path $LockPath -Force -ErrorAction SilentlyContinue
    }
}


function Sanitize-ForExcel {
    param(
        [string]$Text
    )

    if ([string]::IsNullOrEmpty($Text)) {
        return ""
    }

    $sb = New-Object System.Text.StringBuilder

    foreach ($ch in $Text.ToCharArray()) {
        $code = [int][char]$ch

        # Tillåtna XML 1.0-tecken (som Excel klarar):
        # 0x09 (tab), 0x0A (LF), 0x0D (CR)
        # 0x20–0xD7FF, 0xE000–0xFFFD
        if ( $code -eq 0x9 -or $code -eq 0xA -or $code -eq 0xD -or
             ($code -ge 0x20   -and $code -le 0xD7FF) -or
             ($code -ge 0xE000 -and $code -le 0xFFFD) ) {

            [void]$sb.Append($ch)
        }
        else {
            # Släng ogiltiga kontrolltecken helt
        }
    }

    return $sb.ToString()
}

function Write-FallbackLog {
    param(
        [string]$LogExcelPath,
        [string]$WorksheetName,
        [string]$Message,
        [string]$PN  = "",
        [string]$Lot = "",
        [string]$Exp = "",
        [string]$Qty = ""
    )

    try {
        $fallbackDir = Join-Path $env:TEMP "Kontrollprovsfil_FallbackLogs"
        if (-not (Test-Path $fallbackDir)) { New-Item -ItemType Directory -Path $fallbackDir -Force | Out-Null }

        $safeName = ([IO.Path]::GetFileNameWithoutExtension($LogExcelPath) + "_fallback.log")
        $fallbackPath = Join-Path $fallbackDir $safeName

        $ts   = (Get-Date).ToString("yyyy-MM-dd HH:mm:ss")
        $user = [Environment]::UserName

        $Message = Sanitize-ForExcel $Message
        $PN      = Sanitize-ForExcel $PN
        $Lot     = Sanitize-ForExcel $Lot
        $Exp     = Sanitize-ForExcel $Exp
        $Qty     = Sanitize-ForExcel $Qty

        $line = "$ts`t$user`t$WorksheetName`t$Message`tPN=$PN`tLot=$Lot`tExp=$Exp`tQty=$Qty`tTarget=$LogExcelPath"
        Add-Content -Path $fallbackPath -Value $line -Encoding UTF8
    } catch {
        # sista utväg: ignorera
    }
}

function Save-ExcelPackageAtomically {
    param(
        [Parameter(Mandatory=$true)]
        [OfficeOpenXml.ExcelPackage]$Package,

        [Parameter(Mandatory=$true)]
        [string]$TargetPath
    )

    $localTemp = Join-Path $env:TEMP ("LogFile_Kontrollprovsfil_" + [guid]::NewGuid().ToString("N") + ".xlsx")
    $shareTmp  = $TargetPath + ".tmp_" + [guid]::NewGuid().ToString("N")

    try {
        # 1) Spara lokalt först
        $Package.SaveAs((New-Object System.IO.FileInfo($localTemp)))

        # 2) Kopiera upp till share som temporär fil
        Copy-Item -Path $localTemp -Destination $shareTmp -Force

        # 3) Atomisk ersättning via rename/move
        Move-Item -Path $shareTmp -Destination $TargetPath -Force
    }
    finally {
        if (Test-Path $localTemp) { Remove-Item $localTemp -Force -ErrorAction SilentlyContinue }
        if (Test-Path $shareTmp)  { Remove-Item $shareTmp  -Force -ErrorAction SilentlyContinue }
    }
}


# Färg per användare i loggen
function Get-UserColor {
    param(
        [string]$UserName
    )

    $u = $UserName.ToLower()
    switch ($u) {
        'jesper.fredriksson'  { return [System.Drawing.Color]::LightSkyBlue }
        'elin.sidstedt'       { return [System.Drawing.Color]::LightPink }
        'vivian.dao'          { return [System.Drawing.Color]::LightGreen }
        'afnan.vijitraphongs' { return [System.Drawing.Color]::Khaki }
        default               { return $null }
    }
}

function Log-Message {
    param(
        [ValidateSet("UpdateLog", "OtherLogs")]
        [string]$WorksheetName,
        [string]$Message,
        [string]$PN = "",
        [string]$Lot = "",
        [string]$Exp = "",
        [string]$Qty = "",
        [string]$LogExcelPath
    )


    # --- NYTT: sanera all text innan den hamnar i Excel-loggen ---
    $Message = Sanitize-ForExcel $Message
    $PN      = Sanitize-ForExcel $PN
    $Lot     = Sanitize-ForExcel $Lot
    $Exp     = Sanitize-ForExcel $Exp
    $Qty     = Sanitize-ForExcel $Qty


    if (-not (Test-Path $LogExcelPath)) {
        Write-Host "Loggfil saknas. Skapar ny med Initialize-LogFile..." -ForegroundColor Yellow
        Initialize-LogFile -LogExcelPath $LogExcelPath
    }

    $package  = $null
    $lockPath = $null
    try {
        # Exklusivt lås för loggning
        try {
            $lockPath = Acquire-LogLock -LogExcelPath $LogExcelPath
        } catch {
            Write-Host "⚠ Loggen är låst för länge. Skriver fallback-logg lokalt (Excel-filen lämnas orörd)." -ForegroundColor Yellow
            Write-FallbackLog -LogExcelPath $LogExcelPath -WorksheetName $WorksheetName -Message $Message -PN $PN -Lot $Lot -Exp $Exp -Qty $Qty
            return
        }
        $package = Open-ExcelPackageWithRetry -Path $LogExcelPath
        if (-not $package) {
            Write-Host "Kunde inte öppna loggfilen." -ForegroundColor Red
            return
        }

        $sheet = $package.Workbook.Worksheets[$WorksheetName]
        if (-not $sheet) {
            $sheet = $package.Workbook.Worksheets.Add($WorksheetName)
            if ($WorksheetName -eq "UpdateLog") {
                $sheet.Cells[1,1].Value = "Timestamp"
                $sheet.Cells[1,2].Value = "UserName"
                $sheet.Cells[1,3].Value = "Message"
                $sheet.Cells[1,4].Value = "PN"
                $sheet.Cells[1,5].Value = "Lot"
                $sheet.Cells[1,6].Value = "Exp"
                $sheet.Cells[1,7].Value = "Qty"

                $hdrUpd = $sheet.Cells["A1:G1"]
                $hdrUpd.Style.Font.Bold = $true
                $hdrUpd.Style.Fill.PatternType = [OfficeOpenXml.Style.ExcelFillStyle]::Solid
                $hdrUpd.Style.Fill.BackgroundColor.SetColor([System.Drawing.Color]::LightGray)
                $sheet.View.FreezePanes(2,1)
            }
            else {
                $sheet.Cells[1,1].Value = "Timestamp"
                $sheet.Cells[1,2].Value = "UserName"
                $sheet.Cells[1,3].Value = "Message"

                $hdrOther = $sheet.Cells["A1:C1"]
                $hdrOther.Style.Font.Bold = $true
                $hdrOther.Style.Fill.PatternType = [OfficeOpenXml.Style.ExcelFillStyle]::Solid
                $hdrOther.Style.Fill.BackgroundColor.SetColor([System.Drawing.Color]::LightGray)
                $sheet.View.FreezePanes(2,1)
            }
        }

        $lastRow = if ($sheet.Dimension) { $sheet.Dimension.End.Row } else { 1 }
        $newRow  = $lastRow + 1

        $timestamp = (Get-Date).ToString("yyyy-MM-dd HH:mm:ss")
        $userName  = [Environment]::UserName
        $sheet.Cells[$newRow,1].Value = $timestamp
        $sheet.Cells[$newRow,2].Value = $userName
        $sheet.Cells[$newRow,3].Value = $Message
        if ($WorksheetName -eq "UpdateLog") {
            $sheet.Cells[$newRow,4].Value = $PN
            $sheet.Cells[$newRow,5].Value = $Lot
            $sheet.Cells[$newRow,6].Value = $Exp
            $sheet.Cells[$newRow,7].Value = $Qty
        }

        # Färgkodning per användare
        $color = Get-UserColor -UserName $userName
        if ($color) {
            if ($WorksheetName -eq "UpdateLog") {
                $range = $sheet.Cells[$newRow,1,$newRow,7]
            }
            else {
                $range = $sheet.Cells[$newRow,1,$newRow,3]
            }
            $range.Style.Fill.PatternType = [OfficeOpenXml.Style.ExcelFillStyle]::Solid
            $range.Style.Fill.BackgroundColor.SetColor($color)
        }

        # AutoFit alla kolumner
        if ($sheet.Dimension) {
            $sheet.Cells[$sheet.Dimension.Address].AutoFitColumns()
        }

        Save-ExcelPackageAtomically -Package $package -TargetPath $LogExcelPath
        Write-Host "Logg sparad i '$WorksheetName'." -ForegroundColor Green
    }
    catch {
        Write-Host "❌ Fel vid loggning: $_" -ForegroundColor Red
    }
    finally {
        if ($package)  { $package.Dispose() }
        if ($lockPath) { Release-LogLock -LockPath $lockPath }
    }
}

function Log-Other {
    param(
        [string]$Message,
        [string]$LogExcelPath
    )
    Log-Message -WorksheetName "OtherLogs" -Message $Message -LogExcelPath $LogExcelPath
}

function Log-Update {
    param(
        [string]$Message,
        [string]$PN,
        [string]$Lot,
        [string]$Exp,
        [string]$Qty,
        [string]$LogExcelPath
    )
    Log-Message -WorksheetName "UpdateLog" -Message $Message `
        -PN $PN -Lot $Lot -Exp $Exp -Qty $Qty -LogExcelPath $LogExcelPath
}

# Om användaren är en av de specificerade, behövs inget lösenord.
function Request-Password {
    $autoUsers = @("vivian.dao", "elin.sidstedt", "jesper.fredriksson", "afnan.vijitraphongs")
    $currentUser = [Environment]::UserName.ToLower()
    if ($autoUsers -contains $currentUser) {
         Write-Host "Automatiskt godkänd för $currentUser utan lösenord." -ForegroundColor Green
         return $true
    }
    do {
         $securePwd = Read-Host "Ange lösenord eller tryck Enter för PQC-konto" -AsSecureString
         $bstr = [System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($securePwd)
         $unsecurePassword = [System.Runtime.InteropServices.Marshal]::PtrToStringAuto($bstr)
         [System.Runtime.InteropServices.Marshal]::ZeroFreeBSTR($bstr)
         if ($unsecurePassword -eq 'labbkontroll') {
              return $true
         }
         else {
              Write-Host "Fel lösenord. Tryck (1) för att återgå till huvudmenyn." -ForegroundColor Red -BackgroundColor DarkYellow
              $choice = Read-Host "Ange ditt val"
              if ($choice -eq '1') { return $false }
         }
    } while ($true)
}

function Get-UserSignature {
    param(
        [string]$Prompt = "Ange din signatur för att bekräfta ändringarna"
    )
   
    # Fördefinierade användarsignaturer
    $userSignature = @{
        "jesper.fredriksson"   = "JESP"
        "elin.sidstedt"        = "ELS"
        "vivian.dao"           = "vdao"
        "afnan.vijitraphoungs" = "AFVI"
    }

    $currentUser = [Environment]::UserName.ToLower()

    if ($userSignature.ContainsKey($currentUser)) {
        Write-Host "Automatisk signatur för $currentUser $($userSignature[$currentUser])" -ForegroundColor Green
        return $userSignature[$currentUser]
    }

    do {
        $signature = Read-Host $Prompt
        if ($signature.Length -ge 3 -and $signature.Length -le 4) {
            return $signature
        }
        Write-Host "Ogiltig signatur – måste vara exakt 3-4 tecken." -ForegroundColor Red
    } while ($true)
}

# Central datumtolkning
function Try-ParseInventoryDate {
    param(
        [string]$Text
    )

    if ([string]::IsNullOrWhiteSpace($Text) -or $Text -eq 'N/A') {
        return $null
    }

    try {
        # YYYY-MM-DD
        if ($Text -match '^\d{4}-\d{2}-\d{2}$') {
            return [datetime]$Text
        }
        # YYYY-MM eller YYYY MM
        elseif ($Text -match '^\d{4}[- ]\d{2}$') {
            $year  = [int]$Text.Substring(0,4)
            $month = [int]$Text.Substring(5,2)
            $first = New-Object datetime($year, $month, 1)
            return $first.AddMonths(1).AddDays(-1)  # sista dagen i månaden
        }
        else {
            return $null
        }
    }
    catch {
        return $null
    }
}

# Datum-input: tvinga YYYY-MM-DD eller YYYY-MM
function Read-ValidExpiryDate {
    param(
        [string]$Prompt = "Ange utgångsdatum (YYYY-MM-DD eller YYYY-MM)"
    )

    while ($true) {
        $input = Read-Host $Prompt

        if ([string]::IsNullOrWhiteSpace($input)) {
            Write-Host "Utgångsdatum får inte vara tomt." -ForegroundColor Red
            continue
        }

        if ($input -notmatch '^\d{4}-\d{2}-\d{2}$' -and
            $input -notmatch '^\d{4}-\d{2}$') {

            Write-Host "Ogiltigt format. Använd 'YYYY-MM-DD' eller 'YYYY-MM'." -ForegroundColor Red
            continue
        }

        $parsed = Try-ParseInventoryDate -Text $input
        if (-not $parsed) {
            Write-Host "Ogiltigt datum (t.ex. fel månad/dag). Försök igen." -ForegroundColor Red
            continue
        }

        return $input
    }
}

# Self-heal om loggfilen är korrupt
function Repair-LogFileIfCorrupt {
    param(
        [string]$LogExcelPath
    )

    if (-not (Test-Path $LogExcelPath)) {
        return
    }

    try {
        $file    = New-Object System.IO.FileInfo($LogExcelPath)
        $package = New-Object OfficeOpenXml.ExcelPackage($file)
        $package.Dispose()
    }
    catch {
        Write-Host "⚠ Loggfilen verkar korrupt. Tar backup och återskapar..." -ForegroundColor Yellow
        $timestamp  = (Get-Date).ToString("yyyyMMdd_HHmmss")
        $backupName = [System.IO.Path]::GetFileNameWithoutExtension($LogExcelPath) + "_corrupt_$timestamp.xlsx"
        $backupPath = Join-Path (Split-Path $LogExcelPath -Parent) $backupName
        try {
            Move-Item -Path $LogExcelPath -Destination $backupPath -Force
            Write-Host "Backup av korrupt loggfil skapad: $backupPath" -ForegroundColor Green
        } catch {
            Write-Host "Kunde inte döpa om korrupt loggfil: $($_.Exception.Message)" -ForegroundColor Red
        }

        Initialize-LogFile -LogExcelPath $LogExcelPath
    }
}

# =====================[ EPPlus-baserade funktioner ]=====================

function Update-ExcelInventory {
    param(
        [string]$FilePath,
        [string]$LogExcelPath
    )
    Backup-RawDataFile -FilePath $FilePath

    $package = $null
    try {
        $package = Open-ExcelPackageWithRetry -Path $FilePath
        if (-not $package) { return }

        $sheet = $package.Workbook.Worksheets[1]
        if (-not $sheet) {
            Write-Host "Blad 1 saknas i filen. Avbryter..." -ForegroundColor Red
            return
        }
        $rowCount = $sheet.Dimension.End.Row
        $continueUpdating = $true
        while ($continueUpdating) {
            Clear-Host
            Write-Host "================ Uppdatera Kontrollprovsfil =================" -ForegroundColor Cyan
            Write-Host "`nAnge P/N eller skriv [Q] för att avbryta"
            $pnToUpdate = Read-Host
            if ($pnToUpdate -eq 'Q') { return }
            if ([string]::IsNullOrWhiteSpace($pnToUpdate)) {
                Write-Host "P/N får inte vara tomt. Försök igen!" -ForegroundColor DarkYellow
                continue
            }

            $foundRows = @()
            for ($i = 2; $i -le $rowCount; $i++) {
                if ($sheet.Cells[$i,$InventoryColumns.PN].Text -eq $pnToUpdate) {
                    $foundRows += $i
                    if ($foundRows.Count -eq 4) { break }
                }
            }
            if ($foundRows.Count -eq 0) {
                Write-Host "Inget P/N hittades som matchar '$pnToUpdate'." -ForegroundColor Red
                continue
            }

            $index = 1
            foreach ($row in $foundRows) {
                $lotNum     = $sheet.Cells[$row,$InventoryColumns.Lot].Text
                $expiryDate = $sheet.Cells[$row,$InventoryColumns.Exp].Text
                $quantity   = $sheet.Cells[$row,$InventoryColumns.Qty].Text
                Write-Host "`n${index}: Lot: ${lotNum}, Exp: ${expiryDate}, Qty: ${quantity}" -ForegroundColor Green
                $index++
            }

            $updateAnother = $true
            while ($updateAnother) {
                Write-Host "`nVälj en rad att uppdatera (eller tryck 'B' för att backa)"
                $choice = Read-Host
                if ($choice -eq 'Q') { return }
                if ($choice -eq 'B') { break }
                if (($choice -as [int]) -and ($choice -gt 0) -and ($choice -le $foundRows.Count)) {
                    $selectedRow = $foundRows[$choice - 1]
                }
                else {
                    Write-Host "Ogiltigt val, försök igen." -ForegroundColor Red
                    continue
                }

                $selectedLotNum  = $sheet.Cells[$selectedRow,$InventoryColumns.Lot].Text
                $selectedExpDate = $sheet.Cells[$selectedRow,$InventoryColumns.Exp].Text
                $selectedQty     = $sheet.Cells[$selectedRow,$InventoryColumns.Qty].Text

                Write-Host "=============================================" -ForegroundColor Cyan
                Write-Host "`nVad vill du uppdatera?"
                Write-Host "`n1: Nytt Lotnummer, Exp och Qty" -ForegroundColor Magenta
                Write-Host "`n2: Endast kvantitet"          -ForegroundColor Magenta
                Write-Host "`n3: Ta bort post (markeras som N/A)" -ForegroundColor Magenta
                Write-Host "`nVälj ett alternativ (eller tryck 'B' för att backa)"
                $updateChoice = Read-Host
                Write-Host "=============================================" -ForegroundColor Cyan
                if ($updateChoice -eq 'B') { continue }

                $newLotNum = $selectedLotNum
                $newExp    = $selectedExpDate
                $newQty    = $selectedQty

                switch ($updateChoice) {
                    '1' {
                        $newLotNum = Read-Host "Ange Lotnummer"
                        $newExp    = Read-ValidExpiryDate -Prompt "Ange utgångsdatum (YYYY-MM-DD eller YYYY-MM)"
                        $newQty    = Read-Host "Ange antal"
                    }
                    '2' {
                        $newQty = Read-Host "Ange nytt antal"
                    }
                    '3' {
                        $newLotNum = "N/A"; $newExp = "N/A"; $newQty = "N/A"
                    }
                    default {
                        Write-Host "Ogiltigt val." -ForegroundColor Red
                        continue
                    }
                }

                if ($updateChoice -ne '3') {
                    Write-Host "`nBekräfta uppdateringen:"
                    Write-Host "`nLotnummer: $newLotNum" -ForegroundColor DarkYellow
                    Write-Host "`nUtgångsdatum: $newExp" -ForegroundColor DarkYellow
                    Write-Host "`nAntal: $newQty"        -ForegroundColor DarkYellow
                    $confirm = Read-Host "`nÄr detta korrekt? (1) Ja / (2) Nej"
                    Write-Host "=============================================" -ForegroundColor Cyan
                    if ($confirm -eq "1") {
                        $dagensDatum   = (Get-Date).ToString("yyyy-MM-dd")
                        $userSignature = Get-UserSignature

                        Log-Update -Message "'$userSignature' skrev över" `
                                   -PN $pnToUpdate -Lot $selectedLotNum -Exp $selectedExpDate -Qty $selectedQty -LogExcelPath $LogExcelPath

                        $sheet.Cells[$selectedRow,$InventoryColumns.Lot].Value        = $newLotNum
                        $sheet.Cells[$selectedRow,$InventoryColumns.Exp].Value        = $newExp
                        $sheet.Cells[$selectedRow,$InventoryColumns.Qty].Value        = $newQty
                        $sheet.Cells[$selectedRow,$InventoryColumns.LastUpdate].Value = $dagensDatum
                        $sheet.Cells[$selectedRow,$InventoryColumns.Signature].Value  = $userSignature

                        Log-Update -Message "'$userSignature' uppdaterade" `
                                   -PN $pnToUpdate -Lot $newLotNum -Exp $newExp -Qty $newQty -LogExcelPath $LogExcelPath

                        Write-Host "Uppdatering klar." -ForegroundColor Green
                    }
                    else {
                        Write-Host "Inga ändringar sparades."
                    }
                }
                else {
                    $sheet.Cells[$selectedRow,$InventoryColumns.Lot].Value        = "N/A"
                    $sheet.Cells[$selectedRow,$InventoryColumns.Exp].Value        = "N/A"
                    $sheet.Cells[$selectedRow,$InventoryColumns.Qty].Value        = "N/A"
                    $sheet.Cells[$selectedRow,$InventoryColumns.LastUpdate].Value = "N/A"
                    $sheet.Cells[$selectedRow,$InventoryColumns.Signature].Value  = "N/A"
                    $userSignature = Get-UserSignature "Ange signatur för borttagning"
                    Log-Update -Message "'$userSignature' tog bort" `
                               -PN $pnToUpdate -Lot $selectedLotNum -Exp $selectedExpDate -Qty $selectedQty -LogExcelPath $LogExcelPath
                    Write-Host "Post markerad som 'N/A'." -ForegroundColor Green

                    Write-Host "Notis: Omordning av N/A-rader hanteras i VBA-makrot."
                }

                $package.Save()
                Write-Host "=============================================" -ForegroundColor Cyan
                $nextAction = Read-Host "`nVill du uppdatera ett nytt P/N? (1) Ja / (2) Nej (tryck Enter för att fortsätta med samma P/N)"
                if ($nextAction -eq "1") {
                    $updateAnother = $false
                }
                elseif ($nextAction -eq "2") {
                    $updateAnother    = $false
                    $continueUpdating = $false
                }
                elseif ($nextAction -eq "") {
                    Clear-Host
                    Write-Host "================ Uppdatera Kontrollprovsfil =================" -ForegroundColor Cyan
                    Write-Host "`nP/N: $pnToUpdate"
                    $index = 1
                    foreach ($row in $foundRows) {
                        $lotNum     = $sheet.Cells[$row,$InventoryColumns.Lot].Text
                        $expiryDate = $sheet.Cells[$row,$InventoryColumns.Exp].Text
                        $quantity   = $sheet.Cells[$row,$InventoryColumns.Qty].Text
                        Write-Host "`n${index}: Lot: ${lotNum}, Exp: ${expiryDate}, Qty: ${quantity}" -ForegroundColor Green
                        $index++
                    }
                    $updateAnother = $true
                }
            }
        }
    }
    catch {
        Write-Host "Fel vid uppdatering: $_" -ForegroundColor Red
        Log-Other -Message "Fel vid uppdatering: $_" -LogExcelPath $LogExcelPath
    }
    finally {
        if ($package) { $package.Dispose() }
    }
}

function ShowProductInfo {
    param(
        [string]$FilePath,
        [string]$LogExcelPath
    )
    Clear-Host
    Write-Host "================ Sök på Produkt för material =================" -ForegroundColor Cyan
    $productName = Read-Host "`nAnge produktnamn"
    $package = $null
    try {
        $package = Open-ExcelPackageWithRetry -Path $FilePath
        if (-not $package) { return }

        $sheet = $package.Workbook.Worksheets[1]
        if (-not $sheet) {
            Write-Host "Inget blad hittades. Avbryter..." -ForegroundColor Red
            return
        }
        $rowCount = $sheet.Dimension.End.Row
        $infoFound = $false
        $productInfo = @()
        for ($i = 2; $i -le $rowCount; $i++) {
            for ($j = $InventoryColumns.ProductStartCol; $j -le $InventoryColumns.ProductEndCol; $j++) {
                if ($sheet.Cells[$i,$j].Text -eq $productName) {
                    $pn   = $sheet.Cells[$i,$InventoryColumns.PN].Text
                    $lot  = $sheet.Cells[$i,$InventoryColumns.Lot].Text
                    $exp  = $sheet.Cells[$i,$InventoryColumns.Exp].Text
                    $desc = $sheet.Cells[$i,$InventoryColumns.Description].Text
                    if ($lot -ne "N/A" -and $exp -ne "N/A") {
                        $productInfo += [PSCustomObject]@{
                            PN          = $pn
                            LotNr       = $lot
                            Exp         = $exp
                            Description = $desc
                        }
                        $infoFound = $true
                    }
                }
            }
        }
        if ($infoFound) {
            Write-Host "`nProduktinformation hittades:" -ForegroundColor Green
            $productInfo | Format-Table -AutoSize
            Log-Other -Message "Sökte produktinfo för: $productName" -LogExcelPath $LogExcelPath
        }
        else {
            Write-Host "`nIngen information hittades för '$productName'." -ForegroundColor Red
            Log-Other -Message "Ingen info hittades för: $productName" -LogExcelPath $LogExcelPath
        }
    }
    catch {
        Write-Host "Fel vid sökning: $_" -ForegroundColor Red
    }
    finally {
        if ($package) { $package.Dispose() }
    }
}

# Inventeringsrapport
function Generate-InventoryReport {
    param(
        [string]$FilePath,
        [string]$OutputDir,
        [string]$LogExcelPath
    )

    $package = $null
    try {
        $package = Open-ExcelPackageWithRetry -Path $FilePath
        if (-not $package) { return }

        $sheet = $package.Workbook.Worksheets[1]
        if (-not $sheet) {
            Write-Host "Blad 1 saknas, avbryter..." -ForegroundColor Red
            return
        }

        $rowCount = $sheet.Dimension.End.Row
        $data = @()
        for ($i = 2; $i -le $rowCount; $i++) {
            $pn       = $sheet.Cells[$i,$InventoryColumns.PN].Text
            $lotNr    = $sheet.Cells[$i,$InventoryColumns.Lot].Text
            $exp      = $sheet.Cells[$i,$InventoryColumns.Exp].Text
            $qty      = $sheet.Cells[$i,$InventoryColumns.Qty].Text
            $labbCode = $sheet.Cells[$i,$InventoryColumns.LabbCode].Text
            $labbDesc = $sheet.Cells[$i,$InventoryColumns.LabbDescription].Text

            $data += [PSCustomObject]@{
                PN              = $pn
                LotNr           = $lotNr
                Exp             = $exp
                Qty             = $qty
                LabbCode        = $labbCode
                LabbDescription = $labbDesc
            }
        }

        if ($data.Count -eq 0) {
            Write-Host "Inga data att rapportera."
            return
        }
        if (-not (Test-Path $OutputDir)) {
            New-Item -ItemType Directory -Path $OutputDir | Out-Null
        }

        $dateStamp = (Get-Date -Format "yyyyMMdd")
        $xlsxPath  = Join-Path $OutputDir ("InventoryReport_{0}.xlsx" -f $dateStamp)

        if (Test-Path $xlsxPath) { Remove-Item $xlsxPath -Force }
        $invPkg = New-Object OfficeOpenXml.ExcelPackage
        try {
            $ws = $invPkg.Workbook.Worksheets.Add("Inventory")

            # -------- Header via adress-strängar --------
            $ws.Cells["A1"].Value = "PN"
            $ws.Cells["B1"].Value = "LotNr"
            $ws.Cells["C1"].Value = "Exp"
            $ws.Cells["D1"].Value = "Qty"
            $ws.Cells["E1"].Value = "LabbCode"
            $ws.Cells["F1"].Value = "LabbDescription"

            # -------- Data-rader --------
            $row = 2
            foreach ($item in $data) {
                $ws.Cells["A$row"].Value = $item.PN
                $ws.Cells["B$row"].Value = $item.LotNr
                $ws.Cells["C$row"].Value = $item.Exp
                $ws.Cells["D$row"].Value = $item.Qty
                $ws.Cells["E$row"].Value = $item.LabbCode
                $ws.Cells["F$row"].Value = $item.LabbDescription
                $row++
            }

            $lastRow = $row - 1
            if ($lastRow -ge 2) {
                $usedRange = $ws.Cells["A1:F$lastRow"]

                # Header-stil
                $hdr = $ws.Cells["A1:F1"]
                $hdr.Style.Font.Bold = $true
                $hdr.Style.Fill.PatternType = [OfficeOpenXml.Style.ExcelFillStyle]::Solid
                $hdr.Style.Fill.BackgroundColor.SetColor([System.Drawing.Color]::Gainsboro)

                # Borders runt allt
                $usedRange.Style.Border.Top.Style    = [OfficeOpenXml.Style.ExcelBorderStyle]::Thin
                $usedRange.Style.Border.Bottom.Style = [OfficeOpenXml.Style.ExcelBorderStyle]::Thin
                $usedRange.Style.Border.Left.Style   = [OfficeOpenXml.Style.ExcelBorderStyle]::Thin
                $usedRange.Style.Border.Right.Style  = [OfficeOpenXml.Style.ExcelBorderStyle]::Thin

                # AutoFit, freeze panes
                $usedRange.AutoFitColumns()
                $ws.View.FreezePanes(2,1)
            }

            $fileInfo = New-Object System.IO.FileInfo($xlsxPath)
            $invPkg.SaveAs($fileInfo)

            Write-Host "XLSX-rapport genererad: $xlsxPath" -ForegroundColor Green
        }
        finally {
            if ($invPkg) { $invPkg.Dispose() }
        }

        Log-Other -Message "Inventeringsrapport genererad: $xlsxPath" -LogExcelPath $LogExcelPath
    }
    catch {
        Write-Host "Fel vid inventeringsrapport: $($_.Exception.Message)" -ForegroundColor Red
        Log-Other -Message "Fel vid inventeringsrapport: $($_.Exception.Message)" -LogExcelPath $LogExcelPath
    }
    finally {
        if ($package) { $package.Dispose() }
    }
}

# Kort-datum rapport – snygg .xlsx med färg per tidsfönster
function Generate-ExpiringPNsReport {
    param(
        [string]$FilePath,
        [string]$OutputDir = $OutputDir
    )

    $package = $null
    try {
        Clear-Host
        Write-Host "==== Genererar Materialrapport (Kort Datum) ====" -ForegroundColor Cyan
        Write-Host "`nSöker efter poster..." -ForegroundColor Green
        if (-not (Test-Path $OutputDir)) {
            New-Item -ItemType Directory -Path $OutputDir | Out-Null
        }

        $package = Open-ExcelPackageWithRetry -Path $FilePath
        if (-not $package) { return }

        $sheet = $package.Workbook.Worksheets[1]
        if (-not $sheet) {
            Write-Host "Blad saknas, avbryter..." -ForegroundColor Red
            return
        }

        $rowCount  = $sheet.Dimension.End.Row
        $now       = Get-Date
        $lowerDate = $now.AddDays(-15)
        $upperDate = $now.AddDays(30)

        $dateStamp = (Get-Date -Format "yyyyMMdd")
        $xlsxPath  = Join-Path $OutputDir ("ExpiringPNsReport_{0}.xlsx" -f $dateStamp)
        if (Test-Path $xlsxPath) { Remove-Item $xlsxPath -Force }

        $repPkg = New-Object OfficeOpenXml.ExcelPackage
        try {
            $ws = $repPkg.Workbook.Worksheets.Add("Expiring")

            # -------- Header via adress-strängar --------
            $ws.Cells["A1"].Value = "PN"
            $ws.Cells["B1"].Value = "Lotnummer"
            $ws.Cells["C1"].Value = "Exp. Date"
            $ws.Cells["D1"].Value = "Beskrivning"

            $outRow = 2
            for ($i = 2; $i -le $rowCount; $i++) {
                $expiryDateText = $sheet.Cells[$i,$InventoryColumns.Exp].Text
                $expiryDate     = Try-ParseInventoryDate -Text $expiryDateText

                if ($expiryDate -and $expiryDate -ge $lowerDate -and $expiryDate -le $upperDate) {
                    $pn          = $sheet.Cells[$i,$InventoryColumns.PN].Text
                    $lotNum      = $sheet.Cells[$i,$InventoryColumns.Lot].Text
                    $description = $sheet.Cells[$i,$InventoryColumns.Description].Text

                    $ws.Cells["A$outRow"].Value = $pn
                    $ws.Cells["B$outRow"].Value = $lotNum
                    $ws.Cells["C$outRow"].Value = $expiryDateText
                    $ws.Cells["D$outRow"].Value = $description

                    # Färg per tidsfönster
                    $rowRange = $ws.Cells["A$outRow:D$outRow"]
                    if ($expiryDate -lt $now) {
                        $rowRange.Style.Fill.PatternType = [OfficeOpenXml.Style.ExcelFillStyle]::Solid
                        $rowRange.Style.Fill.BackgroundColor.SetColor([System.Drawing.Color]::MistyRose)  # redan utgånget
                    }
                    elseif ($expiryDate -le $now.AddDays(7)) {
                        $rowRange.Style.Fill.PatternType = [OfficeOpenXml.Style.ExcelFillStyle]::Solid
                        $rowRange.Style.Fill.BackgroundColor.SetColor([System.Drawing.Color]::LightSalmon) # 0–7 dagar
                    }
                    elseif ($expiryDate -le $now.AddDays(30)) {
                        $rowRange.Style.Fill.PatternType = [OfficeOpenXml.Style.ExcelFillStyle]::Solid
                        $rowRange.Style.Fill.BackgroundColor.SetColor([System.Drawing.Color]::LightYellow) # 8–30 dagar
                    }

                    $outRow++
                }
            }

            if ($outRow -gt 2) {
                $lastRow   = $outRow - 1
                $usedRange = $ws.Cells["A1:D$lastRow"]

                # Header-stil
                $hdr = $ws.Cells["A1:D1"]
                $hdr.Style.Font.Bold = $true
                $hdr.Style.Fill.PatternType = [OfficeOpenXml.Style.ExcelFillStyle]::Solid
                $hdr.Style.Fill.BackgroundColor.SetColor([System.Drawing.Color]::Gainsboro)

                # Borders runt allt
                $usedRange.Style.Border.Top.Style    = [OfficeOpenXml.Style.ExcelBorderStyle]::Thin
                $usedRange.Style.Border.Bottom.Style = [OfficeOpenXml.Style.ExcelBorderStyle]::Thin
                $usedRange.Style.Border.Left.Style   = [OfficeOpenXml.Style.ExcelBorderStyle]::Thin
                $usedRange.Style.Border.Right.Style  = [OfficeOpenXml.Style.ExcelBorderStyle]::Thin

                $usedRange.AutoFitColumns()
                $ws.View.FreezePanes(2,1)

                $fileInfo = New-Object System.IO.FileInfo($xlsxPath)
                $repPkg.SaveAs($fileInfo)

                Write-Host "Rapport genererad: $xlsxPath" -ForegroundColor Green
                Invoke-Item $xlsxPath
            }
            else {
                Write-Host "`nInga poster hittades inom datumintervallet." -ForegroundColor Red
            }
        }
        finally {
            if ($repPkg) { $repPkg.Dispose() }
        }
    }
    catch {
        Write-Host "Fel vid generering av kort-datum-rapport: $($_.Exception.Message)" -ForegroundColor Red
    }
    finally {
        if ($package) { $package.Dispose() }
    }
}

function Collect-Feedback {
    param(
        [string]$LogExcelPath
    )
    Clear-Host
    Write-Host "========== Feedback ==========" -ForegroundColor Cyan
    do {
        $feedback = Read-Host "`nAnge din feedback"
        if ([string]::IsNullOrWhiteSpace($feedback)) {
            Write-Host "Feedback får inte vara tomt." -ForegroundColor Red
        }
    } until (-not [string]::IsNullOrWhiteSpace($feedback))
    Write-Host "`nDin feedback: $feedback" -ForegroundColor Green
    $confirm = Read-Host "`nVill du skicka feedback? (1) Ja / (2) Nej"
    if ($confirm -eq "1") {
        Log-Other -Message "Feedback: $feedback" -LogExcelPath $LogExcelPath
        Write-Host "`nTack för din feedback!" -ForegroundColor Green
    }
    else {
        Write-Host "`nFeedback avbruten." -ForegroundColor Red
    }
}

# =====================[ Huvudmeny ]=====================

function Main-Menu {
    do {
        Clear-Host
        $currentUser = [Environment]::UserName

        Write-Host "Inloggad användare : $currentUser" -ForegroundColor DarkGray
        Write-Host "Skriptväg         : $PSCommandPath" -ForegroundColor DarkGray
        Write-Host "Loggfil (Excel)   : $LogExcelPath"  -ForegroundColor DarkGray
        Write-Host "Rådatafil         : $RawDataPath"   -ForegroundColor DarkGray

        Write-Host "`n         Kontrollprovsfil - Version 2.5         " -ForegroundColor Cyan
        Write-Host "================================================================================" -ForegroundColor Cyan
        Write-Host "`n1: Uppdatera Kontrollprovsfil" -NoNewLine; Write-Host " (lösenord krävs)" -ForegroundColor Red
        Write-Host "`n2: Generera inventeringsrapport" -NoNewLine; Write-Host " (lösenord krävs)" -ForegroundColor Red
        Write-Host "`n3: Visa material med kort utgångsdatum"
        Write-Host "`n4: Visa materialinformation för produkt"
        Write-Host "`n5: Samla in feedback"
        Write-Host "`n6: Avsluta"
        Write-Host "================================================================================" -ForegroundColor Cyan
        $choice = Read-Host "`nVälj ett alternativ"
        switch ($choice) {
            '1' {
                if (Request-Password) {
                    Log-Other -Message "Valde att uppdatera kontrollprovsfilen" -LogExcelPath $LogExcelPath
                    Update-ExcelInventory -FilePath $RawDataPath -LogExcelPath $LogExcelPath
                }
            }
            '2' {
                if (Request-Password) {
                    Log-Other -Message "Valde att generera inventeringsrapport" -LogExcelPath $LogExcelPath
                    Generate-InventoryReport -FilePath $RawDataPath -OutputDir $OutputDir -LogExcelPath $LogExcelPath
                }
            }
            '3' {
                Log-Other -Message "Valde att generera materialrapport med kort datum" -LogExcelPath $LogExcelPath
                Generate-ExpiringPNsReport -FilePath $RawDataPath
            }
            '4' {
                Log-Other -Message "Valde att visa produktinformation" -LogExcelPath $LogExcelPath
                ShowProductInfo -FilePath $RawDataPath -LogExcelPath $LogExcelPath
            }
            '5' {
                Log-Other -Message "Valde att lämna feedback" -LogExcelPath $LogExcelPath
                Collect-Feedback -LogExcelPath $LogExcelPath
            }
            '6' {
                Log-Other -Message "Avslutar skriptet" -LogExcelPath $LogExcelPath
                Write-Host "Avslutar..." -ForegroundColor Cyan
                return
            }
            default {
                Write-Host "Ogiltigt val, försök igen." -ForegroundColor Red
            }
        }
        Write-Host "`nTryck på valfri tangent för att återgå till huvudmenyn..." -ForegroundColor Cyan
        $null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
    } while ($true)
}

# =====================[ Kör vid start ]=====================

if (-not (Test-Path -LiteralPath $RawDataPath)) {
    Write-Host "❌ Hittar inte rådatafilen:" -ForegroundColor Red
    Write-Host "   $RawDataPath" -ForegroundColor Yellow
    Write-Host "Kontrollera nätverksanslutning och sökväg innan du kör igen." -ForegroundColor Yellow
    exit 1
}

Repair-LogFileIfCorrupt -LogExcelPath $LogExcelPath
Initialize-LogFile      -LogExcelPath $LogExcelPath
Main-Menu


# =====================[ GOLDEN STANDARD: TEXT BACKEND LOG + XLSX VIEWER ]=====================
# Problem addressed: EPPlus/Excel .xlsx replacement on network shares can be corrupted or intermittently fail integrity checks.
# Solution: append-only JSONL backend as single source of truth + best-effort viewer .xlsx regenerated from backend.

# Viewer path remains $LogExcelPath (human-readable, can be open/locked)
# Backend path defaults to same folder with suffix _backend.jsonl (must NOT be opened in Excel)
if (-not (Get-Variable -Name LogBackendPath -Scope Script -ErrorAction SilentlyContinue)) {
    $script:LogBackendPath = ($LogExcelPath -replace '\.xlsx$','_backend.jsonl')
}

function Get-BackendLockPath {
    param([string]$BackendPath)
    return ($BackendPath + '.lock')
}

function Acquire-BackendLock {
    param(
        [string]$BackendPath,
        [int]$MaxWaitSeconds = 30,
        [int]$StaleMinutes = 180
    )

    $lockPath = Get-BackendLockPath -BackendPath $BackendPath

    # Stale lock recovery
    if (Test-Path -LiteralPath $lockPath) {
        try {
            $age = (Get-Date) - (Get-Item -LiteralPath $lockPath).LastWriteTime
            if ($age.TotalMinutes -ge $StaleMinutes) {
                Remove-Item -LiteralPath $lockPath -Force -ErrorAction SilentlyContinue
            }
        } catch {}
    }

    $start = Get-Date
    while ($true) {
        try {
            $fs = [System.IO.File]::Open(
                $lockPath,
                [System.IO.FileMode]::CreateNew,
                [System.IO.FileAccess]::Write,
                [System.IO.FileShare]::None
            )
            try {
                $info = "LockedBy=$([Environment]::UserName);LockedAt=$(Get-Date -Format o);Target=$BackendPath"
                $bytes = [System.Text.Encoding]::UTF8.GetBytes($info)
                $fs.Write($bytes, 0, $bytes.Length)
            } finally {
                $fs.Close()
            }
            return $lockPath
        } catch {
            if ((Get-Date) - $start -gt [TimeSpan]::FromSeconds($MaxWaitSeconds)) {
                throw "Kunde inte få backend-lås på $lockPath inom $MaxWaitSeconds sekunder."
            }
            Start-Sleep -Milliseconds 250
        }
    }
}

function Release-BackendLock {
    param([string]$LockPath)
    if ($LockPath -and (Test-Path -LiteralPath $LockPath)) {
        Remove-Item -LiteralPath $LockPath -Force -ErrorAction SilentlyContinue
    }
}

function Invoke-WithRetry {
    param(
        [scriptblock]$Action,
        [int]$MaxAttempts = 6,
        [int]$BaseDelayMs = 250
    )

    $attempt = 0
    while ($true) {
        try {
            return & $Action
        } catch {
            $attempt++
            if ($attempt -ge $MaxAttempts) { throw }
            $delay = [int]([Math]::Min(5000, ($BaseDelayMs * [Math]::Pow(2, $attempt-1))))
            # jitter
            $delay += (Get-Random -Minimum 0 -Maximum 120)
            Start-Sleep -Milliseconds $delay
        }
    }
}

function Ensure-BackendLog {
    param([string]$BackendPath)

    $dir = Split-Path $BackendPath -Parent
    if (-not (Test-Path -LiteralPath $dir)) {
        New-Item -ItemType Directory -Path $dir -Force | Out-Null
    }
    if (-not (Test-Path -LiteralPath $BackendPath)) {
        # Create empty file
        New-Item -ItemType File -Path $BackendPath -Force | Out-Null
    }
}

function Append-BackendLogJsonl {
    param(
        [string]$BackendPath,
        [ValidateSet('UpdateLog','OtherLogs')] [string]$WorksheetName,
        [string]$Message,
        [string]$PN  = '',
        [string]$Lot = '',
        [string]$Exp = '',
        [string]$Qty = ''
    )

    Ensure-BackendLog -BackendPath $BackendPath

    # Sanitize for Excel (viewer) + prevent huge control chars
    $Message = Sanitize-ForExcel $Message
    $PN      = Sanitize-ForExcel $PN
    $Lot     = Sanitize-ForExcel $Lot
    $Exp     = Sanitize-ForExcel $Exp
    $Qty     = Sanitize-ForExcel $Qty

    $obj = [pscustomobject]@{
        Timestamp     = (Get-Date).ToString('yyyy-MM-dd HH:mm:ss')
        UserName      = [Environment]::UserName
        WorksheetName = $WorksheetName
        Message       = $Message
        PN            = $PN
        Lot           = $Lot
        Exp           = $Exp
        Qty           = $Qty
    }

    $json = ($obj | ConvertTo-Json -Compress)

    $lockPath = $null
    try {
        $lockPath = Acquire-BackendLock -BackendPath $BackendPath
        Invoke-WithRetry -MaxAttempts 6 -Action {
            Add-Content -LiteralPath $BackendPath -Value $json -Encoding UTF8
        }
    } finally {
        if ($lockPath) { Release-BackendLock -LockPath $lockPath }
    }
}

function Read-BackendLogJsonl {
    param([string]$BackendPath)

    if (-not (Test-Path -LiteralPath $BackendPath)) { return @() }

    $lines = @()
    try {
        $lines = Get-Content -LiteralPath $BackendPath -Encoding UTF8 -ErrorAction Stop
    } catch {
        # If locked by scanner etc, retry a few times
        $lines = Invoke-WithRetry -MaxAttempts 5 -Action {
            Get-Content -LiteralPath $BackendPath -Encoding UTF8 -ErrorAction Stop
        }
    }

    $items = New-Object System.Collections.Generic.List[object]
    foreach ($line in $lines) {
        if ([string]::IsNullOrWhiteSpace($line)) { continue }
        try {
            $items.Add(($line | ConvertFrom-Json -ErrorAction Stop))
        } catch {
            # Skip malformed lines but keep running
        }
    }
    return $items
}

function Test-FileIsLocked {
    param([string]$Path)
    try {
        $fs = [System.IO.File]::Open($Path, [System.IO.FileMode]::OpenOrCreate, [System.IO.FileAccess]::ReadWrite, [System.IO.FileShare]::None)
        $fs.Close()
        return $false
    } catch {
        return $true
    }
}

function Publish-ViewerXlsxFromBackend {
    param(
        [string]$BackendPath,
        [string]$ViewerPath
    )

    $pendingPath = ($ViewerPath -replace '\.xlsx$','_PENDING.txt')

    # If viewer is open/locked, skip quietly but mark pending
    if (Test-Path -LiteralPath $ViewerPath) {
        if (Test-FileIsLocked -Path $ViewerPath) {
            try {
                Set-Content -LiteralPath $pendingPath -Value "Viewer kunde inte uppdateras eftersom filen är öppen/låst i Excel. Stäng filen och kör igen. $(Get-Date -Format o)" -Encoding UTF8
            } catch {}
            return $false
        }
    }

    $items = Read-BackendLogJsonl -BackendPath $BackendPath

    # Build workbook fresh each time (stable). Viewer can always be rebuilt.
    $tmpLocal = Join-Path $env:TEMP ("LogViewer_" + [guid]::NewGuid().ToString('N') + ".xlsx")

    $pkg = $null
    try {
        $file = New-Object System.IO.FileInfo($tmpLocal)
        $pkg  = New-Object OfficeOpenXml.ExcelPackage($file)

        $wsUpd   = $pkg.Workbook.Worksheets.Add('UpdateLog')
        $wsOther = $pkg.Workbook.Worksheets.Add('OtherLogs')

        # Headers
        $wsUpd.Cells[1,1].Value = 'Timestamp'
        $wsUpd.Cells[1,2].Value = 'UserName'
        $wsUpd.Cells[1,3].Value = 'Message'
        $wsUpd.Cells[1,4].Value = 'PN'
        $wsUpd.Cells[1,5].Value = 'Lot'
        $wsUpd.Cells[1,6].Value = 'Exp'
        $wsUpd.Cells[1,7].Value = 'Qty'

        $wsOther.Cells[1,1].Value = 'Timestamp'
        $wsOther.Cells[1,2].Value = 'UserName'
        $wsOther.Cells[1,3].Value = 'Message'

        # Style headers
        $hdr1 = $wsUpd.Cells['A1:G1']
        $hdr1.Style.Font.Bold = $true
        $hdr1.Style.Fill.PatternType = [OfficeOpenXml.Style.ExcelFillStyle]::Solid
        $hdr1.Style.Fill.BackgroundColor.SetColor([System.Drawing.Color]::LightGray)
        $wsUpd.View.FreezePanes(2,1)

        $hdr2 = $wsOther.Cells['A1:C1']
        $hdr2.Style.Font.Bold = $true
        $hdr2.Style.Fill.PatternType = [OfficeOpenXml.Style.ExcelFillStyle]::Solid
        $hdr2.Style.Fill.BackgroundColor.SetColor([System.Drawing.Color]::LightGray)
        $wsOther.View.FreezePanes(2,1)

        $rUpd = 2
        $rOth = 2

        foreach ($it in $items) {
            $ws = $it.WorksheetName
            if ($ws -eq 'UpdateLog') {
                $wsUpd.Cells[$rUpd,1].Value = $it.Timestamp
                $wsUpd.Cells[$rUpd,2].Value = $it.UserName
                $wsUpd.Cells[$rUpd,3].Value = $it.Message
                $wsUpd.Cells[$rUpd,4].Value = $it.PN
                $wsUpd.Cells[$rUpd,5].Value = $it.Lot
                $wsUpd.Cells[$rUpd,6].Value = $it.Exp
                $wsUpd.Cells[$rUpd,7].Value = $it.Qty

                $c = Get-UserColor -UserName $it.UserName
                if ($c) {
                    $rng = $wsUpd.Cells[$rUpd,1,$rUpd,7]
                    $rng.Style.Fill.PatternType = [OfficeOpenXml.Style.ExcelFillStyle]::Solid
                    $rng.Style.Fill.BackgroundColor.SetColor($c)
                }

                $rUpd++
            } else {
                $wsOther.Cells[$rOth,1].Value = $it.Timestamp
                $wsOther.Cells[$rOth,2].Value = $it.UserName
                $wsOther.Cells[$rOth,3].Value = $it.Message

                $c = Get-UserColor -UserName $it.UserName
                if ($c) {
                    $rng = $wsOther.Cells[$rOth,1,$rOth,3]
                    $rng.Style.Fill.PatternType = [OfficeOpenXml.Style.ExcelFillStyle]::Solid
                    $rng.Style.Fill.BackgroundColor.SetColor($c)
                }

                $rOth++
            }
        }

        # Autofit
        if ($wsUpd.Dimension) { $wsUpd.Cells[$wsUpd.Dimension.Address].AutoFitColumns() }
        if ($wsOther.Dimension) { $wsOther.Cells[$wsOther.Dimension.Address].AutoFitColumns() }

        $pkg.Save()
        $pkg.Dispose(); $pkg = $null

        # Atomically replace viewer (best effort)
        $dir = Split-Path $ViewerPath -Parent
        if (-not (Test-Path -LiteralPath $dir)) { New-Item -ItemType Directory -Path $dir -Force | Out-Null }

        $bak = $ViewerPath + '.bak'
        Invoke-WithRetry -MaxAttempts 6 -Action {
            if (Test-Path -LiteralPath $ViewerPath) {
                [System.IO.File]::Replace($tmpLocal, $ViewerPath, $bak, $true)
            } else {
                Move-Item -LiteralPath $tmpLocal -Destination $ViewerPath -Force
            }
        }

        # Clear pending if exists
        if (Test-Path -LiteralPath $pendingPath) {
            Remove-Item -LiteralPath $pendingPath -Force -ErrorAction SilentlyContinue
        }

        return $true
    } catch {
        try {
            # If publish fails, mark pending with reason
            Set-Content -LiteralPath ($ViewerPath -replace '\.xlsx$','_PENDING.txt') -Value ("Viewer publish misslyckades: " + $_.Exception.Message + " " + (Get-Date -Format o)) -Encoding UTF8
        } catch {}
        return $false
    } finally {
        if ($pkg) { $pkg.Dispose() }
        if (Test-Path -LiteralPath $tmpLocal) { Remove-Item -LiteralPath $tmpLocal -Force -ErrorAction SilentlyContinue }
    }
}

# OVERRIDE: Initialize-LogFile now initializes backend only.
function Initialize-LogFile {
    param([string]$LogExcelPath)

    # Ensure backend exists
    Ensure-BackendLog -BackendPath $script:LogBackendPath

    # Viewer is created lazily on first publish
}

# OVERRIDE: Repair becomes viewer rebuild (backend is append-only).
function Repair-LogFileIfCorrupt {
    param([string]$LogExcelPath)

    # If viewer corrupt/locked, it will be overwritten when it can.
    try {
        Publish-ViewerXlsxFromBackend -BackendPath $script:LogBackendPath -ViewerPath $LogExcelPath | Out-Null
    } catch {}
}

# OVERRIDE: Log-Message writes to JSONL backend, then best-effort publishes viewer.
function Log-Message {
    param(
        [ValidateSet('UpdateLog','OtherLogs')] [string]$WorksheetName,
        [string]$Message,
        [string]$PN = '',
        [string]$Lot = '',
        [string]$Exp = '',
        [string]$Qty = '',
        [string]$LogExcelPath
    )

    # Append backend (this is the only must-succeed step)
    try {
        Append-BackendLogJsonl -BackendPath $script:LogBackendPath -WorksheetName $WorksheetName -Message $Message -PN $PN -Lot $Lot -Exp $Exp -Qty $Qty
    } catch {
        # If even backend append fails, use existing fallback as last resort
        Write-Host "❌ Kritisk logg-fail (backend): $($_.Exception.Message). Skriver fallback i TEMP." -ForegroundColor Red
        Write-FallbackLog -LogExcelPath $LogExcelPath -WorksheetName $WorksheetName -Message $Message -PN $PN -Lot $Lot -Exp $Exp -Qty $Qty
        return
    }

    # Publish viewer best-effort (never blocks operations)
    $ok = $false
    try {
        $ok = Publish-ViewerXlsxFromBackend -BackendPath $script:LogBackendPath -ViewerPath $LogExcelPath
    } catch {
        $ok = $false
    }

    if ($ok) {
        Write-Host "Logg sparad (backend) + viewer uppdaterad." -ForegroundColor Green
    } else {
        Write-Host "Logg sparad (backend). Viewer kunde inte uppdateras (troligen öppen)." -ForegroundColor Yellow
    }
}

function Log-Other {
    param([string]$Message,[string]$LogExcelPath)
    Log-Message -WorksheetName 'OtherLogs' -Message $Message -LogExcelPath $LogExcelPath
}

function Log-Update {
    param([string]$Message,[string]$PN,[string]$Lot,[string]$Exp,[string]$Qty,[string]$LogExcelPath)
    Log-Message -WorksheetName 'UpdateLog' -Message $Message -PN $PN -Lot $Lot -Exp $Exp -Qty $Qty -LogExcelPath $LogExcelPath
}

# =====================[ END GOLDEN STANDARD LOG OVERRIDES ]=====================