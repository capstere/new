# ====================================================
# Validate-Assay.ps1
# ====================================================
[CmdletBinding()]
param (
    [string[]] $RootPaths     = @('N:\QC\QC-1\IPT\Skiftspecifika dokument\PQC analyst\JESPER\Allmänt\Powerpoint\02. sammanställ\Klar_samman_test', 'N:\QC\QC-1\IPT\3. IPT - KLART FÖR SAMMANSTÄLLNING'),
    [string]   $RawDataPath   = 'N:\QC\QC-1\IPT\Skiftspecifika dokument\PQC analyst\JESPER\Allmänt\Validate Assay\Inventory\raw_data.xlsx',
    [int]      $MaxPressurePSI= 90
)


# ====================================================
# STEG A) GLOBAL OUTPUT-MAPP
# ====================================================
$Global:ReportRoot = 'N:\QC\QC-1\IPT\Skiftspecifika dokument\PQC analyst\JESPER\01'
if (-not (Test-Path $Global:ReportRoot)) {
    New-Item -Path $Global:ReportRoot -ItemType Directory -Force | Out-Null
}

$excelBefore = @(Get-Process Excel -ErrorAction SilentlyContinue)

try {
    Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass -Force -ErrorAction Stop
}
catch {
    Write-Warning "Kunde inte ändra ExecutionPolicy: $_"
}

# ====================================================
# STEG B) GLOBAL KONFIGURATION OCH LOGGNING
# ====================================================
$ScriptVersion   = "1.0.0"
$ScriptBuildDate = (Get-Date).ToString('yyyy-MM-dd')

# Loggfil i Temp-report-root bara för enkelhetens skull
$LogFile = Join-Path $Global:ReportRoot ("Validate-Assay-Log_{0}.txt" -f (Get-Date -Format 'yyyyMMdd_HHmmssfff'))
function Write-Log {
    param ([ValidateSet('INFO','WARN','ERROR')] $Level='INFO',[string]$Message)
    $ts = (Get-Date -Format 'yyyy-MM-dd HH:mm:ss.fff')
    Add-Content -Path $LogFile -Value "[$ts][$Level] $Message"
}
Write-Log -Level INFO "=== START Validate-Assay v$ScriptVersion ==="


# ====================================================
# STEG B.1) PARAMETERVERIFIERING
# ====================================================
foreach ($path in $RootPaths) {
    if (-not (Test-Path $path)) {
        Write-Log -Level ERROR "Rotsökväg '$path' finns inte."
        Write-Error "Rotsökväg '$path' finns inte."
        exit 1
    }
}
if ($MaxPressurePSI -le 0) {
    Write-Log -Level ERROR "MaxPressurePSI ($MaxPressurePSI) måste vara > 0."
    Write-Error "MaxPressurePSI måste vara ett positivt heltal."
    exit 1
}
if (-not (Test-Path $RawDataPath)) {
    Write-Log -Level ERROR "Filen raw_data.xlsx hittades inte på $RawDataPath."
    Write-Error "❌ Filen raw_data.xlsx hittades inte på $RawDataPath."
    exit 1
}

# ====================================================
# STEG B.2) KONSTANTER OCH ASSAY-REGELVERK
# ====================================================
# Kolumner i CSV-filerna
$FixedColumns = @(
    'Assay',         'Assay Version',   'Sample ID',       'Cartridge S/N',
    'Reagent Lot ID','Test Type',       'Instrument S/N',   'Module S/N',
    'S/W Version',   'Start Time',      'Status',          'Test Result'
)

$InstrumentMap = @{
    '802069'    = @{ Name='Infinity-I';     CalibrationDue='Dec-25' }
    '807363'    = @{ Name='Infinity-III';   CalibrationDue='Oct-25' }
    '839032'    = @{ Name='Infinity-V';     CalibrationDue='Jul-25' }
    '847922'    = @{ Name='Infinity-VI';    CalibrationDue='Nov-25' }
    '803094'    = @{ Name='Infinity-VIII';  CalibrationDue='Jul-25' }
    '802598'    = @{ Name='GX1';            CalibrationDue='Sep-25' }
    '802624'    = @{ Name='GX2';            CalibrationDue='Sep-25' }
    '802685'    = @{ Name='GX3';            CalibrationDue='Sep-25' }
    '110012284' = @{ Name='GX5';            CalibrationDue='Oct-25' }
    '110012293' = @{ Name='GX6';            CalibrationDue='Oct-25' }
    '110012274' = @{ Name='GX7';            CalibrationDue='Sep-25' }
}

$AssayRules = @{
    'Xpert MTB-RIF Ultra' = @{
        SampleIDRegex   = '^([^_]+)_(\d{2})_([012])_(\d{2})(A{0,3})([X\+])(?:_D[^_]+(?:,[^_]+)*)?$'
        ValidResults    = @{
            '0' = @('MTB NOT DETECTED')
            '1' = @('MTB DETECTED LOW;RIF Resistance DETECTED', 'MTB DETECTED VERY LOW;RIF Resistance DETECTED')
            '2' = @('MTB DETECTED LOW;RIF Resistance NOT DETECTED')
        }
        TestTypeToCode  = @{
            'Negative Control 1' = '0'
            'Positive Control 1' = '1'
            'Positive Control 2' = '2'
        }
        StartTimeRegex  = '^\d{4}-\d{2}-\d{2}T\d{2}:\d{2}:\d{2}(?:\+|\-)\d{4}$'
        AllowedStatus   = @('Done','Aborted')
        ExtraColumns    = @{
            'Max Pressure (PSI)' = 'FY'
            'Error'              = 'GA'
        }
    }

    'Xpert_HIV-1 Viral Load' = @{
        SampleIDRegex   = '^([^_]+)_(\d{2})_([012])_(\d{2})(A{0,3})([X\+])(?:_D[^_]+(?:,[^_]+)*)?$'
        ValidResults    = @{
            '0' = @('HIV-1 NOT DETECTED')
            '1' = @('HIV-1 DETECTED *')
        }
        TestTypeToCode  = @{
            'Negative Control 1' = '0'
            'Positive Control 1' = '1'
            'Positive Control 2' = '1'
        }
        StartTimeRegex  = '^\d{4}-\d{2}-\d{2}T\d{2}:\d{2}:\d{2}(?:\+|\-)\d{4}$'
        AllowedStatus   = @('Done','Aborted')
        ExtraColumns    = @{
            'Max Pressure (PSI)' = 'CS'
            'Error'              = 'CU'
        }
    }

    'Xpert HPV v2 HR' = @{
        SampleIDRegex   = '^([^_]+)_(\d{2})_([012])_(\d{2})(A{0,3})([X\+])(?:_D[^_]+(?:,[^_]+)*)?$'
        ValidResults    = @{
            '0' = @('INVALID')
            '1' = @('HR HPV POS')
        }
        TestTypeToCode  = @{
            'Negative Control 1' = '0'
            'Positive Control 1' = '1'
            'Positive Control 2' = '1'
        }
        StartTimeRegex  = '^\d{4}-\d{2}-\d{2}T\d{2}:\d{2}:\d{2}(?:\+|\-)\d{4}$'
        AllowedStatus   = @('Done','Aborted')
        ExtraColumns    = @{
            'Max Pressure (PSI)' = 'CS'
            'Error'              = 'CU'
        }
    }

    'Xpert MRSA NxG' = @{
        SampleIDRegex   = '^([^_]+)_(\d{2})_([012])_(\d{2})(A{0,3})([X\+])(?:_D[^_]+(?:,[^_]+)*)?$'
        ValidResults    = @{
            '0' = @('MRSA NOT DETECTED')
            '1' = @('MRSA DETECTED')
        }
        TestTypeToCode  = @{
            'Negative Control 1' = '0'
            'Positive Control 1' = '1'
            'Positive Control 2' = '1'
        }
        StartTimeRegex  = '^\d{4}-\d{2}-\d{2}T\d{2}:\d{2}:\d{2}(?:\+|\-)\d{4}$'
        AllowedStatus   = @('Done','Aborted')
        ExtraColumns    = @{
            'Max Pressure (PSI)' = 'BC'
            'Error'              = 'BE'
        }
    }

    'Xpert CT_NG' = @{
        SampleIDRegex   = '^([^_]+)_(\d{2})_([012])_(\d{2})(A{0,3})([X\+])(?:_D[^_]+(?:,[^_]+)*)?$'
        ValidResults    = @{
            '0' = @('CT NOT DETECTED;NG NOT DETECTED')
            '1' = @('CT DETECTED;NG DETECTED')
        }
        TestTypeToCode  = @{ 'Specimen' = '0', '1' }
        StartTimeRegex  = '^\d{4}-\d{2}-\d{2}T\d{2}:\d{2}:\d{2}(?:\+|\-)\d{4}$'
        AllowedStatus   = @('Done','Aborted')
        ExtraColumns    = @{
            'Max Pressure (PSI)' = 'CE'
            'Error'              = 'CG'
        }
    }

    'Xpress SARS-CoV-2_Flu_RSV plus' = @{
        SampleIDRegex   = '^([^_]+)_(\d{2})_([012])_(\d{2})(A{0,3})([X\+])(?:_D[^_]+(?:,[^_]+)*)?$'
        ValidResults    = @{
            '0' = @('SARS-CoV-2 NEGATIVE;Flu A NEGATIVE;Flu B NEGATIVE;RSV NEGATIVE')
            '1' = @('SARS-CoV-2 POSITIVE;Flu A POSITIVE;Flu B POSITIVE;RSV POSITIVE')
        }
        TestTypeToCode  = @{
            'Negative Control 1' = '0'
            'Positive Control 1' = '1'
        }
        StartTimeRegex  = '^\d{4}-\d{2}-\d{2}T\d{2}:\d{2}:\d{2}(?:\+|\-)\d{4}$'
        AllowedStatus   = @('Done','Aborted')
        ExtraColumns    = @{
            'Max Pressure (PSI)' = 'CY'
            'Error'              = 'DA'
        }
    }

    'Xpert_Carba-R' = @{
        SampleIDRegex   = '^([^_]+)_(\d{2})_([012])_(\d{2})(A{0,3})([X\+])(?:_D[^_]+(?:,[^_]+)*)?$'
        ValidResults    = @{
            '0' = @('IMP1 NOT DETECTED;VIM NOT DETECTED;NDM NOT DETECTED;KPC NOT DETECTED;OXA48 NOT DETECTED')
            '1' = @('IMP1 DETECTED;VIM DETECTED;NDM DETECTED;KPC DETECTED;OXA48 DETECTED')
        }
        TestTypeToCode  = @{
            'Negative Control 1' = '0'
            'Positive Control 1' = '1'
            'Positive Control 2' = '1'
        }
        StartTimeRegex  = '^\d{4}-\d{2}-\d{2}T\d{2}:\d{2}:\d{2}(?:\+|\-)\d{4}$'
        AllowedStatus   = @('Done','Aborted')
        ExtraColumns    = @{
            'Max Pressure (PSI)' = 'CS'
            'Error'              = 'CU'
        }
    }
}

Write-Log -Level INFO "AssayRules inlästa: $($AssayRules.Keys -join ', ')"

$SlangToAssayMap = @{
    "FLUVID"     = "Xpress SARS-CoV-2_Flu_RSV plus"
    "CARBA"      = "Xpert_Carba-R"
    "CTNG"       = "Xpert CT_NG"
    "MRSA NXG"   = "Xpert MRSA NxG"
    "HPV"        = "Xpert HPV v2 HR"
    "HIV VL"     = "Xpert_HIV-1 Viral Load"
    "MTB U"      = "Xpert MTB-RIF Ultra"
    "MTB"        = "Xpert MTB-RIF Assay G4"
    "MTB JP"     = "Xpert MTB-RIF JP IVD"

}

Write-Log -Level INFO "SlangToAssayMap inläst: $($SlangToAssayMap.Keys -join ', ')"

# ====================================================
# STEG B.3) FUNKTION: Get-CSVFileForLot
# ====================================================
function Get-CSVFileForLot {
    [CmdletBinding()]
    param (
        [string] $LotID
    )
    Write-Verbose "Get-CSVFileForLot($LotID)"
    Write-Log -Level INFO "Försöker hitta CSV-filer med LotID '$LotID' under: $($RootPaths -join ', ')"

    $foundFiles = @()
    foreach ($root in $RootPaths) {
        try {
            $foundFiles += Get-ChildItem -Path $root -Filter "*$LotID*.csv" -File -Recurse -ErrorAction Stop
        }
        catch {
            Write-Log -Level WARN "Kunde inte söka i '$root': $($_.Exception.Message)"
        }
    }

    if ($foundFiles.Count -eq 0) {
        throw "❌ Hittade inga CSV-filer innehållande LotID '$LotID'."
    }
    elseif ($foundFiles.Count -eq 1) {
        Write-Log -Level INFO "Hittade en CSV-fil: $($foundFiles[0].FullName)"
        return $foundFiles[0]
    }
    else {
        Write-Host "`nFlera CSV-filer hittades för LotID '$LotID'. Vänligen välj via index:`n"
        for ($i = 0; $i -lt $foundFiles.Count; $i++) {
            Write-Host " [$i] $($foundFiles[$i].FullName)"
        }
        do {
            $ix = Read-Host "Ange index (0–$($foundFiles.Count - 1))"
            if ($ix -notmatch '^\d+$' -or [int]$ix -ge $foundFiles.Count) {
                Write-Warning "Ogiltigt val. Ange ett tal mellan 0 och $($foundFiles.Count - 1)."
            }
        } until ($ix -match '^\d+$' -and [int]$ix -lt $foundFiles.Count)
        Write-Log -Level INFO "Användaren valde CSV-fil: $($foundFiles[$ix].FullName)"
        return $foundFiles[[int]$ix]
    }
}

# ====================================================
# STEG B.4) FUNKTION: Import-CsvData
# ====================================================
function Import-CsvData {
    [CmdletBinding()]
    param (
        [string] $FilePath
    )

    Write-Verbose "Import-CsvData: Läser CSV '$FilePath'"
    Write-Log -Level INFO "Import-CsvData start: $FilePath"

    try {
        $rawLines = Get-Content -Path $FilePath -Encoding UTF8

        if ($rawLines.Count -lt 8) {
            throw "Filen innehåller färre än 8 rader."
        }

        $headerLine   = $rawLines[7] -split ',' | ForEach-Object { $_.Trim() }

        $pressureHeader = 'Max Pressure (PSI)'
        $errorHeader    = 'Error'
        $pressureIndex  = $headerLine.IndexOf($pressureHeader)
        $errorIndex     = $headerLine.IndexOf($errorHeader)

        if ($pressureIndex -eq -1 -or $errorIndex -eq -1) {
            throw "Kunde inte hitta kolumnerna '$pressureHeader' eller '$errorHeader' i header."
        }

        $dataLines = $rawLines | Select-Object -Skip 9
        if (-not $dataLines -or $dataLines.Count -eq 0) {
            throw "Inga datarader hittades på rad 10 och nedåt."
        }

        $columns = for ($i = 0; $i -lt $headerLine.Length; $i++) {
            if ($i -lt $FixedColumns.Count) {
                $FixedColumns[$i]
            }
            elseif ($i -eq $pressureIndex) {
                'MaxPressure'
            }
            elseif ($i -eq $errorIndex) {
                'ErrorText'
            }
            else {
                "Ignore$i"
            }
        }

        $data = $dataLines | ConvertFrom-Csv -Delimiter ',' -Header $columns

        $wanted   = $FixedColumns + 'MaxPressure','ErrorText'
        $filtered = $data | Select-Object $wanted

        Write-Log -Level INFO "Import-CsvData klar – antal datarader: $($filtered.Count)"
        return $filtered
    }
    catch {
        Write-Log -Level ERROR "Import-CsvData FEL: $($_.Exception.Message)"
        return $null
    }
}

# ====================================================
# STEG B.5) FUNKTION: Get-SuffixLogic
# ====================================================
function Get-SuffixLogic {
    [CmdletBinding()]
    param (
        [array]     $Rows,
        [hashtable] $AssayRules
    )

    Write-Verbose "Analyserar suffix-logik baserat on Sample ID och Cartridge S/N"

    [int] $xOdd     = 0
    [int] $xEven    = 0
    [int] $plusOdd  = 0
    [int] $plusEven = 0

    $totalRows = $Rows.Count
    $processed = 0

    foreach ($row in $Rows) {
        $processed++
        if ($processed % 100 -eq 0) {
            Write-Progress -Activity "Analyserar suffix-logik" `
                           -Status "$processed av $totalRows rader..." `
                           -PercentComplete ((($processed) / $totalRows) * 100)
        }

        $sampleID    = $row.'Sample ID'
        $assay       = $row.'Assay'
        $cartridgeSN = $row.'Cartridge S/N'
        if (-not $sampleID -or -not $cartridgeSN -or -not $assay) { continue }

        if (-not $AssayRules.ContainsKey($assay)) { continue }
        $regex = $AssayRules[$assay].SampleIDRegex

        if ($sampleID -match $regex) {
            $suffixFromID = $Matches[6]
            if ($cartridgeSN -match '(\d)$') {
                $lastDigit = [int] $Matches[1]
                $isEven    = (($lastDigit % 2) -eq 0)

                if ($suffixFromID -eq 'X') {
                    if ($isEven) { $xEven++ } else { $xOdd++ }
                }
                elseif ($suffixFromID -eq '+') {
                    if ($isEven) { $plusEven++ } else { $plusOdd++ }
                }
            }
        }
    }

    Write-Progress -Activity "Analyserar suffix-logik" -Completed

    $standardMatch = $xOdd + $plusEven
    $reversedMatch = $xEven + $plusOdd

    $swap = $false
    if ($reversedMatch -gt $standardMatch) {
        $swap = $true
    }

    Write-Verbose  "Suffix-logik bestämd: Swap=$swap (StandardMatch=$standardMatch, ReversedMatch=$reversedMatch)"
    Write-Verbose  "Detaljer: X_Odd=$xOdd, X_Even=$xEven, Plus_Odd=$plusOdd, Plus_Even=$plusEven"

    Write-Log -Level INFO "Suffix-analys: X_Odd=$xOdd, X_Even=$xEven, Plus_Odd=$plusOdd, Plus_Even=$plusEven"
    Write-Log -Level INFO "Suffix-logik: Swap=$swap (StandardMatch=$standardMatch, ReversedMatch=$reversedMatch)"

    return $swap
}

# ====================================================
# STEG B.6) FUNKTION: Validate-Data
# ====================================================
function Validate-Data {
    [CmdletBinding()]
    param (
        [array]     $Rows,
        [string]    $LotID,
        [int]       $MaxPressurePSI,
        [hashtable] $AssayRules,
        [hashtable] $InstrumentMap,
        [bool]      $SwapSuffixLogic
    )

    Write-Verbose "Validate-Data: Validerar $($Rows.Count) rader för LotID '$LotID'"
    Write-Log -Level INFO "Validate-Data start: $($Rows.Count) rader att validera"

    $deviations     = @()
    $systemErrors   = @()
    $instruments    = @{}
    $seenCartridges = [Collections.Generic.HashSet[string]]::new()
    $total          = 0

    foreach ($row in $Rows) {
        $total++

        # Extrahera värden
        $sampleID      = $row.'Sample ID'
        $assay         = $row.'Assay'
        $reagentLotID  = $row.'Reagent Lot ID'
        $status        = $row.'Status'
        $startTime     = $row.'Start Time'
        $cartridgeSN   = $row.'Cartridge S/N'
        $testType      = $row.'Test Type'
        $testResult    = $row.'Test Result'
        $maxPressure   = $row.MaxPressure
        $errorText     = $row.ErrorText
        $instSN        = $row.'Instrument S/N'

        try {
            # 1) Reagent Lot ID måste matcha LotID
            if ($reagentLotID -ne $LotID) {
                $deviations += [PSCustomObject]@{
                    Category       = 'Reagent Lot ID'
                    SampleID       = $sampleID
                    Field          = 'Reagent Lot ID'
                    Value          = $reagentLotID
                    Issue          = "Lot-ID mismatch (förväntade: $LotID)"
                    Assay          = $assay
                    'Cartridge S/N'= $cartridgeSN
                    'Start Time'   = $startTime
                    'Instrument S/N'= $instSN
                }
            }

            # 2) Är assay definierad?
            $rules = $null
            if (-not $AssayRules.ContainsKey($assay)) {
                $deviations += [PSCustomObject]@{
                    Category       = 'Assay'
                    SampleID       = $sampleID
                    Field          = 'Assay'
                    Value          = $assay
                    Issue          = 'Felaktig Assay'
                    Assay          = $assay
                    'Cartridge S/N'= $cartridgeSN
                    'Start Time'   = $startTime
                    'Instrument S/N'= $instSN
                }
            }
            else {
                $rules = $AssayRules[$assay]
            }

            # 3) Status (bara 'Done' eller 'Aborted')
            if ($rules) {
                if (-not ($rules.AllowedStatus -contains $status)) {
                    $deviations += [PSCustomObject]@{
                        Category       = 'Status'
                        SampleID       = $sampleID
                        Field          = 'Status'
                        Value          = $status
                        Issue          = "Felaktigt Status (förväntar: $($rules.AllowedStatus -join '/'))"
                        Assay          = $assay
                        'Cartridge S/N'= $cartridgeSN
                        'Start Time'   = $startTime
                        'Instrument S/N'= $instSN
                    }
                }
                elseif ($status -ieq 'Aborted') {
                    $deviations += [PSCustomObject]@{
                        Category       = 'Status'
                        SampleID       = $sampleID
                        Field          = 'Status'
                        Value          = $status
                        Issue          = 'Aborted – kontrollera ErrorText'
                        Assay          = $assay
                        'Cartridge S/N'= $cartridgeSN
                        'Start Time'   = $startTime
                        'Instrument S/N'= $instSN
                    }
                }
            }

            # 4) Start Time (ISO8601)
            if ($rules -and $rules.StartTimeRegex) {
                try {
                    [DateTime]::ParseExact($startTime, 'yyyy-MM-ddTHH:mm:ssK', $null) | Out-Null
                }
                catch {
                    $deviations += [PSCustomObject]@{
                        Category       = 'Start Time'
                        SampleID       = $sampleID
                        Field          = 'Start Time'
                        Value          = $startTime
                        Issue          = 'Felaktigt Start Time-format'
                        Assay          = $assay
                        'Cartridge S/N'= $cartridgeSN
                        'Start Time'   = $startTime
                        'Instrument S/N'= $instSN
                    }
                }
            }

            # 5) SampleID: regex, suffix + TestType→kod
            if ($rules -and $rules.SampleIDRegex) {
                if ($sampleID -match $rules.SampleIDRegex) {
                    $codeFromID   = $Matches[3]
                    $suffixFromID = $Matches[6]

                    # Hämta sista siffra i Cartridge S/N
                    if ($cartridgeSN -match '(\d)$') {
                        $lastDigit = [int] $Matches[1]
                        $isEven    = (($lastDigit % 2) -eq 0)

                        # Bestäm förväntat suffix (standard eller reverserat)
                        if (-not $SwapSuffixLogic) {
                            $expectedSuffix = if ($isEven) {'+'} else {'X'}
                        }
                        else {
                            $expectedSuffix = if ($isEven) {'X'} else {'+'}
                        }

                        if ($suffixFromID -ne $expectedSuffix) {
                            $deviations += [PSCustomObject]@{
                                Category       = 'Sample ID'
                                SampleID       = $sampleID
                                Field          = 'Sample ID'
                                Value          = $sampleID
                                Issue          = "Feldöpt X/+, rätt är: '$expectedSuffix'"
                                Assay          = $assay
                                'Cartridge S/N'= $cartridgeSN
                                'Start Time'   = $startTime
                                'Instrument S/N'= $instSN
                            }
                        }
                    }
                    else {
                        $deviations += [PSCustomObject]@{
                            Category       = 'Cartridge S/N'
                            SampleID       = $sampleID
                            Field          = 'Cartridge S/N'
                            Value          = $cartridgeSN
                            Issue          = 'Kan ej avgöra suffix – Cartridge S/N saknar siffra sist'
                            Assay          = $assay
                            'Cartridge S/N'= $cartridgeSN
                            'Start Time'   = $startTime
                            'Instrument S/N'= $instSN
                        }
                    }

                    # Kod-matchning: TestType → TestTypeToCode
                    if ($rules.TestTypeToCode.ContainsKey($testType)) {
                        $expectedCode = $rules.TestTypeToCode[$testType]
                        if ($codeFromID -ne $expectedCode) {
                            $deviations += [PSCustomObject]@{
                                Category       = 'Sample ID'
                                SampleID       = $sampleID
                                Field          = 'Sample ID'
                                Value          = $sampleID
                                Issue          = "Felvald Test Type, rätt är: '$expectedCode'"
                                Assay          = $assay
                                'Cartridge S/N'= $cartridgeSN
                                'Start Time'   = $startTime
                                'Instrument S/N'= $instSN
                            }
                        }
                    }
                    else {
                        $deviations += [PSCustomObject]@{
                            Category       = 'Test Type'
                            SampleID       = $sampleID
                            Field          = 'Test Type'
                            Value          = $testType
                            Issue          = 'Felaktig Test Type'
                            Assay          = $assay
                            'Cartridge S/N'= $cartridgeSN
                            'Start Time'   = $startTime
                            'Instrument S/N'= $instSN
                        }
                    }
                }
                else {
                    $deviations += [PSCustomObject]@{
                        Category       = 'Sample ID'
                        SampleID       = $sampleID
                        Field          = 'Sample ID'
                        Value          = $sampleID
                        Issue          = 'Felaktigt Sample ID-format'
                        Assay          = $assay
                        'Cartridge S/N'= $cartridgeSN
                        'Start Time'   = $startTime
                        'Instrument S/N'= $instSN
                    }
                }
            }

            # 6) Cartridge S/N: format + unik‐kontroll
            if ($cartridgeSN -notmatch '^\d{9,10}$') {
                $deviations += [PSCustomObject]@{
                    Category       = 'Cartridge S/N'
                    SampleID       = $sampleID
                    Field          = 'Cartridge S/N'
                    Value          = $cartridgeSN
                    Issue          = 'Felaktigt Cartridge S/N'
                    Assay          = $assay
                    'Cartridge S/N'= $cartridgeSN
                    'Start Time'   = $startTime
                    'Instrument S/N'= $instSN
                }
            }
            elseif ($seenCartridges.Contains($cartridgeSN)) {
                $deviations += [PSCustomObject]@{
                    Category       = 'Cartridge S/N'
                    SampleID       = $sampleID
                    Field          = 'Cartridge S/N'
                    Value          = $cartridgeSN
                    Issue          = 'Dubblett av Cartridge S/N'
                    Assay          = $assay
                    'Cartridge S/N'= $cartridgeSN
                    'Start Time'   = $startTime
                    'Instrument S/N'= $instSN
                }
            }
            else {
                $seenCartridges.Add($cartridgeSN) | Out-Null
            }

            # 7) TestType → TestResult
            if ($rules.TestTypeToCode.ContainsKey($testType)) {
                $codeForType = $rules.TestTypeToCode[$testType]
                $validArray  = $rules.ValidResults[$codeForType]
                if ($validArray.Count -gt 0) {
                    if ($testResult -notin $validArray) {
                        $deviations += [PSCustomObject]@{
                            Category       = 'Test Result'
                            SampleID       = $sampleID
                            Field          = 'Test Result'
                            Value          = $testResult
                            Issue          = "Ogiltigt Test Result för '$testType'"
                            Assay          = $assay
                            'Cartridge S/N'= $cartridgeSN
                            'Start Time'   = $startTime
                            'Instrument S/N'= $instSN
                        }
                    }
                }
            }

            # 8) Max Pressure (PSI)
            if ($maxPressure -ne '' -and $maxPressure -ne $null) {
                $parsed = $null
                if ([double]::TryParse($maxPressure, [ref]$parsed)) {
                    if ($parsed -gt $MaxPressurePSI) {
                        $deviations += [PSCustomObject]@{
                            Category       = 'Max Pressure'
                            SampleID       = $sampleID
                            Field          = 'Max Pressure'
                            Value          = $maxPressure
                            Issue          = "Max Pressure Failure ($parsed)"
                            Assay          = $assay
                            'Cartridge S/N'= $cartridgeSN
                            'Start Time'   = $startTime
                            'Instrument S/N'= $instSN
                        }
                    }
                }
                else {
                    $deviations += [PSCustomObject]@{
                        Category       = 'Max Pressure'
                        SampleID       = $sampleID
                        Field          = 'Max Pressure'
                        Value          = $maxPressure
                        Issue          = 'Ej numeriskt värde för Max Pressure'
                        Assay          = $assay
                        'Cartridge S/N'= $cartridgeSN
                        'Start Time'   = $startTime
                        'Instrument S/N'= $instSN
                    }
                }
            }

            # 9) ErrorText (instrumentfel)
            if ($errorText -and $errorText.Trim() -ne '') {
                $deviations += [PSCustomObject]@{
                    Category       = 'Error'
                    SampleID       = $sampleID
                    Field          = 'Error'
                    Value          = $errorText
                    Issue          = "Error: '$errorText'"
                    Assay          = $assay
                    'Cartridge S/N'= $cartridgeSN
                    'Start Time'   = $startTime
                    'Instrument S/N'= $instSN
                }
            }

            # 10) Instrument S/N – lookup + samla unika
            if ([string]::IsNullOrWhiteSpace($instSN)) {
                $deviations += [PSCustomObject]@{
                    Category       = 'Instrument'
                    SampleID       = $sampleID
                    Field          = 'Instrument S/N'
                    Value          = ''
                    Issue          = 'Saknar Instrument S/N'
                    Assay          = $assay
                    'Cartridge S/N'= $cartridgeSN
                    'Start Time'   = $startTime
                    'Instrument S/N'= $instSN
                }
            }
            else {
                if ($InstrumentMap.ContainsKey($instSN)) {
                    if (-not $instruments.ContainsKey($instSN)) {
                        $instruments[$instSN] = $InstrumentMap[$instSN]
                    }
                }
                else {
                    $deviations += [PSCustomObject]@{
                        Category       = 'Instrument'
                        SampleID       = $sampleID
                        Field          = 'Instrument S/N'
                        Value          = $instSN
                        Issue          = 'Okänt instrument'
                        Assay          = $assay
                        'Cartridge S/N'= $cartridgeSN
                        'Start Time'   = $startTime
                        'Instrument S/N'= $instSN
                    }
                }
            }
        }
        catch {
            $err = $_
            $systemErrors += [PSCustomObject]@{
                Category       = 'SystemError'
                SampleID       = ($sampleID -or "Rad$($total)")
                Field          = ''
                Value          = ''
                Issue          = "Systemfel: $($err.Exception.Message)"
                Assay          = $assay
                'Cartridge S/N'= $cartridgeSN
                'Start Time'   = $startTime
                'Instrument S/N'= $instSN
            }
            Write-Log -Level ERROR "Systemfel vid rad $total (SampleID: $sampleID) – $($err.Exception.Message)"
        }
    }

    $allDeviations = $deviations + $systemErrors
    Write-Log -Level INFO "Validate-Data klar: Totalt $total rader, $($deviations.Count) avvikelser, $($systemErrors.Count) systemfel, $($instruments.Count) unika instrument."
    return @{
        Deviations   = $allDeviations
        Instruments  = $instruments
        Total        = $total
    }
}

# ====================================================
# STEG B.7) FUNKTION: Get-AssayInventoryFromRawData (COM-baserad)
# ====================================================
function Get-AssayInventoryFromRawData {
    <#
      .SYNOPSIS
        Läser inventarie‐data från raw_data.xlsx via Excel COM, filtrerar på assay.
      .PARAMETER RawDataPath
        Sökväg till raw_data.xlsx
      .PARAMETER AssayName
        Fullständigt assay‐namn (t.ex. "Xpert MTB-RIF Ultra")
      .PARAMETER SlangMap
        Hashtable som mappar “slang”→“AssayName”
      .OUTPUTS
        En array av PSCustomObject med egenskaper:
          P/N, Lotnr., Utgångsdatum, Antal i lager, Senast inventering, SIGN, "Produkt (G-M)", Beskrivning, Förvaringsplats, Labb
    #>
    param(
        [string]     $RawDataPath,
        [string]     $AssayName,
        [hashtable]  $SlangMap
    )

    $result = @()

    try {
        $excel = New-Object -ComObject Excel.Application
        $excel.Visible       = $false
        $excel.DisplayAlerts = $false

        $wb   = $excel.Workbooks.Open($RawDataPath)
        $ws   = $wb.Worksheets.Item(1)
        $used = $ws.UsedRange
        $data2D = $used.Value2       # 2D array av alla cellvärden

        $rowCount = $used.Rows.Count
        $colCount = $used.Columns.Count

        # Hitta slang‐nyckel för AssayName
        $assaySlang = $SlangMap.GetEnumerator() |
                      Where-Object { $_.Value -eq $AssayName } |
                      Select-Object -First 1 -ExpandProperty Key

        if (-not $assaySlang) {
            Write-Host "⚠️ Kunde inte hitta någon nyckel för Assay '$AssayName' i SlangMap." -ForegroundColor Yellow
            $wb.Close($false)
            [void]($excel.Quit())
            [System.Runtime.Interopservices.Marshal]::ReleaseComObject($ws)    | Out-Null
            [System.Runtime.Interopservices.Marshal]::ReleaseComObject($wb)    | Out-Null
            [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
            return $result
        }

        # Första raden är header, går från A1 o.s.v. Nu loopar vi rader 2..$rowCount.
        for ($r = 2; $r -le $rowCount; $r++) {
            $match = $false
            $produkts = @()

            # Kolla kolumner G–M (7..13) för att se om de innehåller slang eller full assay-namn (match med wildcard).
            for ($c = 7; $c -le 13; $c++) {
                $cellVal = $data2D[$r, $c]
                if ($null -ne $cellVal) {
                    $cellText = ([string]$cellVal).Trim()
                    if ($cellText -like "*$assaySlang*" -or $cellText -like "*$AssayName*") {
                        $match = $true
                    }
                    if ($cellText -ne '') {
                        $produkts += $cellText
                    }
                }
            }

            if ($match) {
                # Läs av kolumnerna 1..6
                $pn         = $data2D[$r, 1]
                $lotnr      = $data2D[$r, 2]
                $utg        = $data2D[$r, 3]
                $antal      = $data2D[$r, 4]
                $senast     = $data2D[$r, 5]
                $sign       = $data2D[$r, 6]

                # Slå ihop alla produkter (G–M som vi lagt i $produkts)
                $gmProdukt  = ($produkts -join ", ").Trim(',')
                # Kolumnerna 14, 15, 16 är Beskrivning, Förvaringsplats, Labb
                $beskrivning     = $data2D[$r, 14]
                $forvaringsplats = $data2D[$r, 15]
                $labb            = $data2D[$r, 16]

                $result += [PSCustomObject]@{
                    "P/N"                 = $pn
                    "Lotnr."              = $lotnr
                    "Utgångsdatum"        = $utg
                    "Antal i lager"       = $antal
                    "Senast inventering"  = $senast
                    "SIGN"                = $sign
                    "Produkt (G-M)"       = $gmProdukt
                    "Beskrivning"         = $beskrivning
                    "Förvaringsplats"     = $forvaringsplats
                    "Labb"                = $labb
                }
            }
        }

        # Stäng COM-objekten
        $wb.Close($false)
        [void]($excel.Quit())
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($ws)    | Out-Null
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($wb)    | Out-Null
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
    }
    catch {
        Write-Host "❌ Fel vid inläsning av raw_data: $_" -ForegroundColor Red
    }

    return $result
}

# ====================================================
# STEG B.8) FUNKTION: Export-Report
# ====================================================
function Export-Report {
    param(
        [string]    $LotID,
        [array]     $Deviations,
        [hashtable] $Instruments,
        [int]       $Total,
        [array]     $InventoryRows
    )

    # 1) Skapa per-månad-mapp under ReportRoot
    $monthFolder = Join-Path $Global:ReportRoot (Get-Date -Format 'yyyy-MM')
    if (-not (Test-Path $monthFolder)) {
        New-Item -Path $monthFolder -ItemType Directory -Force | Out-Null
    }

    # 2) Bygg timestampat filnamn
    $timestamp  = Get-Date -Format 'yyyyMMdd_HHmmss'
    $reportFile = Join-Path $monthFolder "Report_${LotID}_$timestamp.xlsx"

    # Initiera COM-variabler
    $excel = $null
    $wbOut = $null

    try {
        # 3) Starta Excel COM synligt för felsökning
        $excel = New-Object -ComObject Excel.Application
        $excel.Visible       = $true
        $excel.DisplayAlerts = $true
        # 4) Skapa ny arbetsbok
        $wbOut = $excel.Workbooks.Add()

        # ---------- 2) Deviations-bladet ----------
        $wsDev = $wbOut.Worksheets.Item(1)
        $wsDev.Name = 'Deviations'
        for ($i = $wbOut.Worksheets.Count; $i -gt 1; $i--) {
            $wbOut.Worksheets.Item($i).Delete()
        }
        $headersDev = @('Sample ID','Assay','Cartridge S/N','Start Time','Instrument S/N','Deviations')
        for ($c = 0; $c -lt $headersDev.Count; $c++) {
            $wsDev.Cells.Item(1, $c+1).Value2 = [string]$headersDev[$c]
            $wsDev.Cells.Item(1, $c+1).Font.Bold = $true
        }
        $grouped     = $Deviations | Group-Object -Property SampleID
        $summaryRows = @()
        if ($grouped.Count -gt 0) {
            foreach ($g in $grouped) {
                $first = $g.Group[0]
                $summaryRows += [PSCustomObject]@{
                    'Sample ID'      = $first.SampleID
                    'Assay'          = $first.Assay
                    'Cartridge S/N'  = $first.'Cartridge S/N'
                    'Start Time'     = $first.'Start Time'
                    'Instrument S/N' = $first.'Instrument S/N'
                    'Deviations'     = ($g.Group | ForEach-Object { $_.Issue }) -join '; '
                }
            }
        } else {
            $summaryRows += [PSCustomObject]@{
                'Sample ID'      = $null
                'Assay'          = $null
                'Cartridge S/N'  = $null
                'Start Time'     = $null
                'Instrument S/N' = $null
                'Deviations'     = 'Inga avvikelser'
            }
        }
        for ($r = 0; $r -lt $summaryRows.Count; $r++) {
            $rowObj = $summaryRows[$r]
            $wsDev.Cells.Item($r+2, 1).Value2 = if ($null -eq $rowObj.'Sample ID')      { "" } else { [string]$rowObj.'Sample ID' }
            $wsDev.Cells.Item($r+2, 2).Value2 = if ($null -eq $rowObj.Assay)            { "" } else { [string]$rowObj.Assay }
            $wsDev.Cells.Item($r+2, 3).Value2 = if ($null -eq $rowObj.'Cartridge S/N')  { "" } else { [string]$rowObj.'Cartridge S/N' }
            $wsDev.Cells.Item($r+2, 4).Value2 = if ($null -eq $rowObj.'Start Time')     { "" } else { [string]$rowObj.'Start Time' }
            $wsDev.Cells.Item($r+2, 5).Value2 = if ($null -eq $rowObj.'Instrument S/N') { "" } else { [string]$rowObj.'Instrument S/N' }
            $wsDev.Cells.Item($r+2, 6).Value2 = if ($null -eq $rowObj.Deviations)       { "" } else { [string]$rowObj.Deviations }
        }
        $lastDevRow = $summaryRows.Count + 1
        $wsDev.Range("A1:F$lastDevRow").AutoFilter()
        $wsDev.Columns("A:F").AutoFit()
        if ($grouped.Count -gt 0) {
            for ($r = 2; $r -le $lastDevRow; $r++) {
                $devText = $wsDev.Cells.Item($r, 6).Text
                if ($devText -ne '' -and $devText -ne 'Inga avvikelser') {
                    for ($c = 1; $c -le 5; $c++) {
                        $wsDev.Cells.Item($r, $c).Interior.ColorIndex = 22
                    }
                }
            }
        }

        # ---------- 3) Instruments- och Inventory-blad ----------
        if ($Instruments.Count -gt 0) {
            $wsInstr = $wbOut.Worksheets.Add()
            $wsInstr.Name = 'Instruments'
            $headersInstr = @('Instrument S/N','Instrument Name','Calibration Due')
            for ($c = 0; $c -lt $headersInstr.Count; $c++) {
                $wsInstr.Cells.Item(1, $c+1).Value2 = [string]$headersInstr[$c]
                $wsInstr.Cells.Item(1, $c+1).Font.Bold = $true
            }
            $rowIndex = 2
            foreach ($kv in $Instruments.GetEnumerator()) {
                $wsInstr.Cells.Item($rowIndex, 1).Value2 = [string]$kv.Key
                $wsInstr.Cells.Item($rowIndex, 2).Value2 = [string]$kv.Value.Name
                $wsInstr.Cells.Item($rowIndex, 3).Value2 = [string]$kv.Value.CalibrationDue
                $rowIndex++
            }
            $lastRowInstr = $rowIndex - 1
            $wsInstr.Range("A1:C$lastRowInstr").AutoFilter()
            $wsInstr.Columns("A:C").AutoFit()

            # Inventory
            $wsInv = $wbOut.Worksheets.Add()
            $wsInv.Name = 'Inventory'

            if ($InventoryRows -and $InventoryRows.Count -gt 0) {
                $cleanInventory = $InventoryRows | Where-Object {
                    $val = $_.'Antal i lager'
                    -not [string]::IsNullOrWhiteSpace($val) -and $val.Trim().ToUpper() -ne 'N/A'
                }

                if ($cleanInventory.Count -gt 0) {
                    $props = $cleanInventory[0].PSObject.Properties.Name
                    for ($c = 0; $c -lt $props.Count; $c++) {
                        $wsInv.Cells.Item(1, $c+1).Value2 = [string]$props[$c]
                        $wsInv.Cells.Item(1, $c+1).Font.Bold = $true
                    }
                    for ($r = 0; $r -lt $cleanInventory.Count; $r++) {
                        for ($c = 0; $c -lt $props.Count; $c++) {
                            $val = $cleanInventory[$r].$($props[$c])
                            $wsInv.Cells.Item($r+2, $c+1).Value2 = if ($null -eq $val) { "" } else { [string]$val }
                        }
                    }
                    $wsInv.Range("A1").CurrentRegion.AutoFilter()
                    $wsInv.Columns.AutoFit()
                } else {
                    $wsInv.Cells.Item(1,1).Value2 = "Ingen inventory-information funnen för assay/LotID"
                    $wsInv.Cells.Item(1,1).Font.Bold = $true
                    $wsInv.Cells.Item(1,1).Font.Color = 255
                    $wsInv.Columns.Item(1).AutoFit()
                }
            } else {
                $wsInv.Cells.Item(1,1).Value2 = "Inga inventory-rader tillgängliga"
                $wsInv.Cells.Item(1,1).Font.Bold = $true
                $wsInv.Cells.Item(1,1).Font.Color = 255
                $wsInv.Columns.Item(1).AutoFit()
            }

        }

        # ---------- 4) Summary-bladet ----------
        $wsSum = $wbOut.Worksheets.Add()
        $wsSum.Name = 'Summary'
        $wsSum.Cells.Item(1,1).Value2    = [string]'Total Replicates'
        $wsSum.Cells.Item(1,1).Font.Bold = $true
        $wsSum.Cells.Item(2,1).Value2    = [string]$Total
        $wsSum.Columns.Item(1).AutoFit()

        # ---------- 5) CategorySummary-bladet ----------
        $catSummary = $Deviations |
                      Where-Object { $_.Category -ne 'SystemError' } |
                      Group-Object -Property Category |
                      Sort-Object Count -Descending |
                      ForEach-Object {
                          [PSCustomObject]@{ Category = $_.Name; Count = $_.Count }
                      }
        if ($catSummary.Count -gt 0) {
            $wsCat = $wbOut.Worksheets.Add()
            $wsCat.Name = 'CategorySummary'
            $wsCat.Cells.Item(1,1).Value2    = [string]'Category'
            $wsCat.Cells.Item(1,2).Value2    = [string]'Count'
            $wsCat.Cells.Item(1,1).Font.Bold = $true
            $wsCat.Cells.Item(1,2).Font.Bold = $true
            for ($r = 0; $r -lt $catSummary.Count; $r++) {
                $wsCat.Cells.Item($r+2,1).Value2 = [string]$catSummary[$r].Category
                $wsCat.Cells.Item($r+2,2).Value2 = [string]$catSummary[$r].Count
            }
            $lastRowCat = $catSummary.Count + 1
            $wsCat.Range("A1:B$lastRowCat").AutoFilter()
            $wsCat.Columns("A:B").AutoFit()
        }

        # ---------- 6) Metadata-bladet ----------
        $wsMeta = $wbOut.Worksheets.Add()
        $wsMeta.Name = 'Metadata'
        $metaProps = [ordered]@{
            'ExportDate'    = (Get-Date).ToString('yyyy-MM-dd HH:mm')
            'LotID'         = $LotID
            'ScriptVersion' = "Validate-Assay.ps1 v$ScriptVersion"
            'Computer'      = $env:COMPUTERNAME
            'TotalRows'     = $Total
            'Deviations'    = $Deviations.Count
            'UniqueInstr'   = $Instruments.Count
        }
        $rowM = 1
        foreach ($k in $metaProps.Keys) {
            $wsMeta.Cells.Item($rowM,1).Value2    = [string]$k
            $wsMeta.Cells.Item($rowM,1).Font.Bold = $true
            $wsMeta.Cells.Item($rowM,2).Value2    = [string]$metaProps[$k]
            $rowM++
        }
        $wsMeta.Columns("A:B").AutoFit()

        # ---------- 7) Spara & stäng Excel ----------
        try {
            Write-Host "Sparar rapport: $reportFile"
            $wbOut.SaveAs($reportFile, 51)
        }
        catch {
            Write-Warning "SaveAs misslyckades – försöker SaveCopyAs…"
            $wbOut.SaveCopyAs($reportFile)
        }
        if (-not (Test-Path $reportFile)) {
            throw "Kunde inte skapa Excel-rapport: $reportFile"
        }
        Write-Host "Rapport sparad: $reportFile"
    }
    catch {
        Write-Error "Export-Report FEL: $($_.Exception.Message)"
        throw
    }
    finally {
        if ($wbOut -ne $null) { $wbOut.Close($false) }
        if ($excel -ne $null) { $excel.Quit() }
        if ($wbOut -ne $null) { [System.Runtime.Interopservices.Marshal]::ReleaseComObject($wbOut) | Out-Null }
        if ($excel -ne $null) { [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null }
        [GC]::Collect(); [GC]::WaitForPendingFinalizers()
    }

    # 8) Döda nya Excel-processer
    $excelAfter = @(Get-Process Excel -ErrorAction SilentlyContinue)
    $newOnes    = $excelAfter | Where-Object { $excelBefore.ProcessId -notcontains $_.Id }
    foreach ($proc in $newOnes) { try { $proc.Kill() } catch { } }
}


# ====================================================
# STEG B.9) MAIN: HUVUDFLÖDE
# ====================================================
try {
    Write-Log -Level INFO "---- START NY SESSION ----"

    # 9.1) LotID
    do {
        $LotID = Read-Host "Ange Lot ID"
    } until ($LotID -match '^\d{5}$')
    Write-Log -Level INFO "LotID: $LotID"

    # 9.2) CSV-fil
    $csvFile = Get-CSVFileForLot -LotID $LotID
    Write-Host "CSV hittad: $($csvFile.FullName)"

    # 9.3) Importera + suffix-logik + validera
    $rows   = Import-CsvData -FilePath $csvFile.FullName
    $swap   = Get-SuffixLogic     -Rows $rows -AssayRules $AssayRules
    if ($swap) { Write-Host "X / + - logik" }
    $result = Validate-Data      -Rows $rows -LotID $LotID -MaxPressurePSI $MaxPressurePSI `
                -AssayRules $AssayRules -InstrumentMap $InstrumentMap -SwapSuffixLogic $swap

    # 9.4) Inventory
    $assayName     = $rows[0].Assay
    $inventoryRows = Get-AssayInventoryFromRawData -RawDataPath $RawDataPath -AssayName $assayName -SlangMap $SlangToAssayMap

    # 9.5) Anropa nya Export-Report
    Export-Report `
        -LotID         $LotID `
        -Deviations    $result.Deviations `
        -Instruments   $result.Instruments `
        -Total         $result.Total `
        -InventoryRows $inventoryRows

    Write-Log -Level INFO "=== SLUT Validate-Assay ==="
}
catch {
    Write-Error "❌ Fel uppstod: $($_.Exception.Message)"
    Write-Log -Level ERROR "FATAL: $($_.Exception.Message)"
    exit 1
}
