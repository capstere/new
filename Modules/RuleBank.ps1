#requires -Version 5.1

<#
  RuleBank.ps1
  DATA-ONLY root object used by RuleEngine.
  - No dynamic construction of rule data; values are deterministic and kept in-memory.
  - Optional override: Modules\AssayAliases.generated.(psd1|ps1) can supply additional aliases.
  - No Set-StrictMode here to keep dot-sourcing safe for the GUI/event handlers.
#>

$assayAliasesFromFile = & {
  param($paths)
  foreach ($p in $paths) {
    if (-not $p) { continue }
    if (Test-Path -LiteralPath $p) {
      try {
        $tmp = Import-PowerShellDataFile -Path $p -ErrorAction Stop
      } catch {
        $tmp = $null
        try { $tmp = . $p } catch {}
      }
      if ($tmp -is [System.Collections.IDictionary]) { return $tmp }
    }
  }
  return $null
} @(
  (Join-Path $PSScriptRoot 'AssayAliases.generated.psd1'),
  (Join-Path $PSScriptRoot 'AssayAliases.generated.ps1')
)

$baseline = [ordered]@{
  TotalTests = 210
  SampleSizePolicy = [ordered]@{
    Expected       = 210
    MinWarnCount   = 200
    AllowedMissing = 999
    AllowedExtra   = 999
    Severity       = 'Warn'
    RuleId         = 'BASELINE_MIN_COUNT'
  }
}

$globalRules = [ordered]@{
  Identity = [ordered]@{
    Fields = [ordered]@{
      Assay        = [ordered]@{ Severity='Error'; RuleId='IDENTITY_ASSAY' }
      AssayVersion = [ordered]@{ Severity='Error'; RuleId='IDENTITY_ASSAY_VERSION' }
      ReagentLotId = [ordered]@{ Severity='Error'; RuleId='IDENTITY_REAGENT_LOT_ID' }
    }
  }
  Uniqueness = [ordered]@{
    SampleId    = [ordered]@{ Severity='Error'; RuleId='DUP_SAMPLE_ID' }
    CartridgeSN = [ordered]@{ Severity='Error'; RuleId='DUP_CARTRIDGE_SN' }
  }
  SampleId = [ordered]@{
    Parse = [ordered]@{
      Pattern = '^(?<role>[A-Z0-9_]+)[-_](?<id>\d{3})[-_](?<idx>\d{2})$'
      Example = 'S1-001-01'
    }
    Roles = [ordered]@{
      CTRL  = 'Control'
      S1    = 'Specimen-SampleType1'
      S2    = 'Specimen-SampleType2'
      S3    = 'Specimen-SampleType3'
      S4    = 'Specimen-SampleType4'
      MIX_1 = 'Specimen-Mix1'
      MIX_2 = 'Specimen-Mix2'
      MIX_3 = 'Specimen-Mix3'
      MIX_4 = 'Specimen-Mix4'
    }
  }
  MaxPressure = [ordered]@{
    WarnThreshold = 85
    FailThreshold = 90
    Severity      = 'Error'
    RuleId        = 'MAX_PRESSURE'
  }
  Status = [ordered]@{ PassStatuses=@('Done'); FailStatuses=@('Error','Invalid','Aborted','Incomplete') }
}

$errorBank = [ordered]@{
  ExtractRegex = '(?i)(?<Code>[A-Z]{2}\d{2,4}|\d{3,5})'
  Codes = [ordered]@{
    'default' = [ordered]@{
      Name            = 'Unknown code'
      Group           = 'Unknown'
      Classification  = 'UNKNOWN'
      Description     = 'Error code not recognized in RuleBank.'
      GeneratesRetest = $false
    }
    'AB101' = [ordered]@{
      Name            = 'Alpha Beta'
      Group           = 'Hardware'
      Classification  = 'HARDWARE'
      Description     = 'Generic hardware error'
      GeneratesRetest = $true
    }
    'CD201' = [ordered]@{
      Name            = 'Gamma Delta'
      Group           = 'Functional'
      Classification  = 'FUNCTIONAL'
      Description     = 'Generic functional error'
      GeneratesRetest = $true
    }
    'PQ001' = [ordered]@{
      Name            = 'PQC'
      Group           = 'Workflow'
      Classification  = 'WORKFLOW'
      Description     = 'PQC workflow / data issue'
      GeneratesRetest = $false
    }
    '2037' = [ordered]@{
      Name            = '2037'
      Group           = 'Functional'
      Classification  = 'FUNCTIONAL'
      Description     = 'Define meaning for 2037 (site mapping).'
      GeneratesRetest = $false
    }
    '5015' = [ordered]@{
     Name            = '5015'
     Group           = 'Functional'
     Classification  = 'FUNCTIONAL'
     Description     = 'Define meaning for 5015 (site mapping).'
     GeneratesRetest = $false
      
      }
     }
    }

$builtinAssayAliases = [ordered]@{
  'GXMRSA/SA-SSTI-CE'                 = 'GXMRSA/SA-SSTI-CE'
  'GXSACOMP-CE-10'                    = 'GXSACOMP-CE-10'
  'GXSACOMP-CN-10'                    = 'GXSACOMP-CN-10'
  'Respiratory Panel IUO'             = 'Respiratory Panel IUO'
  'Xpert C.difficile BT'              = 'Xpert C.difficile BT'
  'Xpert C.difficile G3'              = 'Xpert C.difficile G3'
  'Xpert CT_CE'                       = 'Xpert CT_CE'
  'Xpert CT_NG'                       = 'Xpert CT_NG'
  'Xpert Carba-R'                     = 'Xpert Carba-R'
  'Xpert_Carba-R'                     = 'Xpert Carba-R'
  'Xpert Ebola CE-IVD'                = 'Xpert Ebola CE-IVD'
  'Xpert Ebola EUA'                   = 'Xpert Ebola EUA'
  'Xpert GBS LB XC'                   = 'Xpert GBS LB XC'
  'Xpert HBV Viral Load'              = 'Xpert HBV Viral Load'
  'Xpert HCV PQC'                     = 'Xpert HCV PQC'
  'Xpert HCV VL Fingerstick'          = 'Xpert HCV VL Fingerstick'
  'Xpert HCV Viral Load'              = 'Xpert HCV Viral Load'
  'Xpert_HCV Viral Load'              = 'Xpert HCV Viral Load'
  'Xpert HIV-1 Qual'                  = 'Xpert HIV-1 Qual'
  'Xpert_HIV-1 Qual'                  = 'Xpert HIV-1 Qual'
  'Xpert HIV-1 Qual XC PQC'           = 'Xpert HIV-1 Qual'
  'Xpert HIV-1 Viral Load'            = 'Xpert HIV-1 Viral Load'
  'Xpert_HIV-1 Viral Load'            = 'Xpert HIV-1 Viral Load'
  'Xpert HIV-1 Viral Load XC'         = 'Xpert HIV-1 Viral Load XC'
  'Xpert HPV HR'                      = 'Xpert HPV HR'
  'Xpert HPV v2 HR'                   = 'Xpert HPV v2 HR'
  'Xpert MRSA NxG'                    = 'Xpert MRSA NxG'
  'Xpert MRSA-SA SSTI G3'             = 'GXMRSA/SA-SSTI-CE'
  'Xpert MTB-RIF Assay G4'            = 'Xpert MTB-RIF Assay G4'
  'Xpert MTB-RIF JP IVD'              = 'Xpert MTB-RIF JP IVD'
  'Xpert MTB-RIF Ultra'               = 'Xpert MTB-RIF Ultra'
  'Xpert MTB-XDR'                     = 'Xpert MTB-XDR'
  'Xpert Norovirus'                   = 'Xpert Norovirus'
  'Xpert SA Nasal Complete G3'        = 'Xpert SA Nasal Complete G3'
  'Xpert Xpress CoV-2 plus'           = 'Xpert Xpress CoV-2 plus'
  'Xpert Xpress Flu-RSV'              = 'Xpert Xpress Flu-RSV'
  'Xpress Flu IPT_EAT off'            = 'Xpress Flu IPT_EAT off'
  'Xpert Xpress GBS'                  = 'Xpert Xpress GBS'
  'Xpert Xpress GBS US-IVD'           = 'Xpert Xpress GBS US-IVD'
  'Xpert Xpress SARS-CoV-2 CE-IVD'    = 'Xpert Xpress SARS-CoV-2 CE-IVD'
  'Xpert Xpress SARS-CoV-2'           = 'Xpert Xpress SARS-CoV-2 CE-IVD'
  'Xpert Xpress_SARS-CoV-2'           = 'Xpert Xpress SARS-CoV-2 CE-IVD'
  'Xpert Xpress SARS-CoV-2 Flu RSV'   = 'Xpert Xpress SARS-CoV-2 Flu RSV'
  'Xpert Xpress_SARS-CoV-2_Flu_RSV'   = 'Xpert Xpress SARS-CoV-2 Flu RSV'
  'Xpert Xpress SARS-CoV-2 Flu RSV plus' = 'Xpert Xpress SARS-CoV-2 Flu RSV plus'
  'Xpert Xpress_SARS-CoV-2_Flu_RSV plus' = 'Xpert Xpress SARS-CoV-2 Flu RSV plus'
  'Xpress SARS-CoV-2 Flu RSV'         = 'Xpert Xpress SARS-CoV-2 Flu RSV'
  'Xpress SARS-CoV-2_Flu_RSV'         = 'Xpert Xpress SARS-CoV-2 Flu RSV'
  'Xpress SARS-CoV-2 Flu RSV plus'    = 'Xpert Xpress SARS-CoV-2 Flu RSV plus'
  'Xpress SARS-CoV-2_Flu_RSV plus'    = 'Xpert Xpress SARS-CoV-2 Flu RSV plus'
  'Xpert Xpress Strep A'              = 'Xpert Xpress Strep A'
  'Xpert Xpress_Strep A'              = 'Xpert Xpress Strep A'
  'Xpert vanA vanB'                   = 'Xpert vanA vanB'
}

$assayAliases = if ($assayAliasesFromFile -is [System.Collections.IDictionary]) {
  $o = [ordered]@{}
  foreach ($k in $assayAliasesFromFile.Keys) { $o[$k] = $assayAliasesFromFile[$k] }
  $o
} else {
  $builtinAssayAliases
}

$assayProfiles = [ordered]@{
  '_DEFAULT'                           = [ordered]@{ Mode='GlobalOnly'; Description='Fallback when no specific profile exists.'; DisplayName='_DEFAULT' }
  'default'                            = [ordered]@{ Mode='GlobalOnly'; Description='Legacy default profile key.'; DisplayName='_DEFAULT' }
  'GXMRSA/SA-SSTI-CE'                  = [ordered]@{ Mode='GlobalOnly'; Description='Stub profile'; DisplayName='GXMRSA/SA-SSTI-CE' }
  'GXSACOMP-CE-10'                     = [ordered]@{ Mode='GlobalOnly'; Description='Stub profile'; DisplayName='GXSACOMP-CE-10' }
  'GXSACOMP-CN-10'                     = [ordered]@{ Mode='GlobalOnly'; Description='Stub profile'; DisplayName='GXSACOMP-CN-10' }
  'Respiratory Panel IUO'              = [ordered]@{ Mode='GlobalOnly'; Description='Stub profile'; DisplayName='Respiratory Panel IUO' }
  'Xpert C.difficile BT'               = [ordered]@{ Mode='GlobalOnly'; Description='Stub profile'; DisplayName='Xpert C.difficile BT' }
  'Xpert C.difficile G3'               = [ordered]@{ Mode='GlobalOnly'; Description='Stub profile'; DisplayName='Xpert C.difficile G3' }
  'Xpert CT_CE'                        = [ordered]@{ Mode='GlobalOnly'; Description='Stub profile'; DisplayName='Xpert CT_CE' }
  'Xpert CT_NG'                        = [ordered]@{ Mode='GlobalOnly'; Description='Stub profile'; DisplayName='Xpert CT_NG' }
  'Xpert Carba-R'                      = [ordered]@{ Mode='GlobalOnly'; Description='Stub profile'; DisplayName='Xpert Carba-R' }
  'Xpert Ebola CE-IVD'                 = [ordered]@{ Mode='GlobalOnly'; Description='Stub profile'; DisplayName='Xpert Ebola CE-IVD' }
  'Xpert Ebola EUA'                    = [ordered]@{ Mode='GlobalOnly'; Description='Stub profile'; DisplayName='Xpert Ebola EUA' }
  'Xpert GBS LB XC'                    = [ordered]@{ Mode='GlobalOnly'; Description='Stub profile'; DisplayName='Xpert GBS LB XC' }
  'Xpert HBV Viral Load'               = [ordered]@{ Mode='GlobalOnly'; Description='Stub profile'; DisplayName='Xpert HBV Viral Load' }
  'Xpert HCV PQC'                      = [ordered]@{ Mode='GlobalOnly'; Description='Stub profile'; DisplayName='Xpert HCV PQC' }
  'Xpert HCV VL Fingerstick'           = [ordered]@{ Mode='GlobalOnly'; Description='Stub profile'; DisplayName='Xpert HCV VL Fingerstick' }
  'Xpert HCV Viral Load'               = [ordered]@{ Mode='GlobalOnly'; Description='Stub profile'; DisplayName='Xpert HCV Viral Load' }
  'Xpert HIV-1 Qual'                   = [ordered]@{ Mode='GlobalOnly'; Description='Stub profile'; DisplayName='Xpert HIV-1 Qual' }
  'Xpert HIV-1 Viral Load'             = [ordered]@{ Mode='GlobalOnly'; Description='Stub profile'; DisplayName='Xpert HIV-1 Viral Load' }
  'Xpert HIV-1 Viral Load XC'          = [ordered]@{ Mode='GlobalOnly'; Description='Stub profile'; DisplayName='Xpert HIV-1 Viral Load XC' }
  'Xpert HPV HR'                       = [ordered]@{ Mode='GlobalOnly'; Description='Stub profile'; DisplayName='Xpert HPV HR' }
  'Xpert HPV v2 HR'                    = [ordered]@{ Mode='GlobalOnly'; Description='Stub profile'; DisplayName='Xpert HPV v2 HR' }
  'Xpert MRSA NxG'                     = [ordered]@{ Mode='GlobalOnly'; Description='Stub profile'; DisplayName='Xpert MRSA NxG' }
  'Xpert MTB-RIF Assay G4'             = [ordered]@{ Mode='GlobalOnly'; Description='Stub profile'; DisplayName='Xpert MTB-RIF Assay G4' }
  'Xpert MTB-RIF JP IVD'               = [ordered]@{ Mode='GlobalOnly'; Description='Stub profile'; DisplayName='Xpert MTB-RIF JP IVD' }
  'Xpert MTB-RIF Ultra'                = [ordered]@{ Mode='GlobalOnly'; Description='Stub profile'; DisplayName='Xpert MTB-RIF Ultra' }
  'Xpert MTB-XDR'                      = [ordered]@{ Mode='GlobalOnly'; Description='Stub profile'; DisplayName='Xpert MTB-XDR' }
  'Xpert Norovirus'                    = [ordered]@{ Mode='GlobalOnly'; Description='Stub profile'; DisplayName='Xpert Norovirus' }
  'Xpert SA Nasal Complete G3'         = [ordered]@{ Mode='GlobalOnly'; Description='Stub profile'; DisplayName='Xpert SA Nasal Complete G3' }
  'Xpert Xpress CoV-2 plus'            = [ordered]@{ Mode='GlobalOnly'; Description='Stub profile'; DisplayName='Xpert Xpress CoV-2 plus' }
  'Xpert Xpress Flu-RSV'               = [ordered]@{ Mode='GlobalOnly'; Description='Stub profile'; DisplayName='Xpert Xpress Flu-RSV' }
  'Xpert Xpress GBS'                   = [ordered]@{ Mode='GlobalOnly'; Description='Stub profile'; DisplayName='Xpert Xpress GBS' }
  'Xpert Xpress GBS US-IVD'            = [ordered]@{ Mode='GlobalOnly'; Description='Stub profile'; DisplayName='Xpert Xpress GBS US-IVD' }
  'Xpert Xpress SARS-CoV-2 CE-IVD'     = [ordered]@{ Mode='GlobalOnly'; Description='Stub profile'; DisplayName='Xpert Xpress SARS-CoV-2 CE-IVD' }
  'Xpert Xpress SARS-CoV-2 Flu RSV'    = [ordered]@{ Mode='GlobalOnly'; Description='Stub profile'; DisplayName='Xpert Xpress SARS-CoV-2 Flu RSV' }
  'Xpert Xpress SARS-CoV-2 Flu RSV plus' = [ordered]@{ Mode='GlobalOnly'; Description='Stub profile'; DisplayName='Xpert Xpress SARS-CoV-2 Flu RSV plus' }
  'Xpert Xpress Strep A'               = [ordered]@{ Mode='GlobalOnly'; Description='Stub profile'; DisplayName='Xpert Xpress Strep A' }
  'Xpert vanA vanB'                    = [ordered]@{ Mode='GlobalOnly'; Description='Stub profile'; DisplayName='Xpert vanA vanB' }
  'Xpress Flu IPT_EAT off'             = [ordered]@{ Mode='GlobalOnly'; Description='Stub profile'; DisplayName='Xpress Flu IPT_EAT off' }
}

$global:RuleBank = [ordered]@{
  Version       = '2025-01-05'
  SchemaVersion = 3
  Baseline      = $baseline
  Global        = $globalRules
  ErrorBank     = $errorBank
  AssayAliases  = $assayAliases
  AssayProfiles = $assayProfiles
}

# --- cleanup: avoid leaking helper variables into caller scope when dot-sourced ---
try {
  Remove-Variable assayAliasesFromFile,builtinAssayAliases,assayAliases,baseline,globalRules,errorBank,assayProfiles -ErrorAction SilentlyContinue
} catch {}

