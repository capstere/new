#requires -Version 5.1
Set-StrictMode -Version Latest

<#
  RuleBank.ps1
  - DATA ONLY: one root object `$RuleBank` (global) containing all assay profiles and global rules.
  - Generated from real 'Tests Summary' exports + ErrorBank.xlsx (Solna reality, 210 baseline).
  - X/+ parity logic intentionally skipped for now (future majority-/paritetsregel).

  Structure:
    $RuleBank = @{
      Version
      SchemaVersion
      Baseline
      Global
      ErrorBank
      AssayProfiles
    }
#>

$global:RuleBank = [ordered]@{
  'Version'       = '2025-12-14'
  'SchemaVersion' = 1

  # ---------------------------------------------------------------------------
  # BASELINE – 210 tester per batch, toleranser för missing/extra mm
  # ---------------------------------------------------------------------------
  'Baseline' = [ordered]@{
    'StandardBatch210' = [ordered]@{
      'BaselineTotal' = 210
      'BagLayout'     = @(
        [ordered]@{
          'BagRange'     = '00'
          'ExpectedTotal' = 10
        }
        [ordered]@{
          'BagRange'     = '01-10'
          'ExpectedTotal' = 20
        }
      )
      'Notes' = '210 is the 100% baseline sample size per batch. Real datasets may be <210 (visual fail) or >210 (replacements).'
    }
    'SampleSizePolicy' = [ordered]@{
      'CountToleranceInfo' = 1
      'CountToleranceWarn' = 2
      'MissingWarnAt'      = 1
      'MissingErrorAt'     = 15
      'ExtraWarnAt'        = 1
      'ExtraErrorAt'       = 15
      'Notes'              = 'Majority is truth: 79 usually means expected 80. Deviations are findings, not new truth.'
    }
  }

  # ---------------------------------------------------------------------------
  # GLOBAL RULES – gäller alla assays
  # ---------------------------------------------------------------------------
  'Global' = [ordered]@{
    'Identity' = [ordered]@{
      'Fields'   = @('Assay', 'Assay Version', 'Reagent Lot ID')
      'Severity' = 'Error'
      'Message'  = 'Assay / Assay Version / Reagent Lot ID must be constant within the dataset.'
    }

    'Uniqueness' = [ordered]@{
      'SampleId' = [ordered]@{
        'Field'    = 'Sample ID'
        'Severity' = 'Error'
        'Message'  = 'Duplicate Sample ID detected.'
      }
      'CartridgeSn' = [ordered]@{
        'Field'    = 'Cartridge S/N'
        'Severity' = 'Error'
        'Message'  = 'Duplicate Cartridge S/N detected.'
      }
    }

    'SampleId' = [ordered]@{
      # PREFIX_BAG_IDX_TestNo + ev. suffix (A/AA/AAA, Dxx osv)
      'Regex' = '^(?<Prefix>[A-Za-z0-9\-]+)_(?<BAG>\d{2})_(?<IDX>[0-5])_(?<TestNo>\d{2})(?<Suffix>.*)$'
      'Constraints' = [ordered]@{
        'BagMin'    = 0
        'BagMax'    = 10
        'IdxMin'    = 0
        'IdxMax'    = 5
        'TestNoMin' = 1
        'TestNoMax' = 20
      }
      'ReplacementSuffixRegex' = '(?i)A{1,3}'
      'DelamSuffixRegex'       = '(?i)D(?:0[1-9]|1[0-1]|2A|2B)'
    }

    'WorkCenter' = [ordered]@{
      'ParseRule' = [ordered]@{
        # Cartridge S/N som börjar med 10/11/12 → tvåsiffrig lina.
        'TwoDigitIfStartsWith' = @('10', '11', '12')
      }
      'Map' = [ordered]@{
        '2'  = 'R2'
        '6'  = 'R6'
        '8'  = 'R8'
        '9'  = 'R9'
        '10' = 'R10'
        '11' = 'R11'
        '12' = 'R12'
      }
      'Severity' = 'Info'
    }

    # -----------------------------------------------------------------------
    # ROLLABELS – hur vi vill visa Specimen-index-grupper i rapporten
    # -----------------------------------------------------------------------
    'RoleMeta' = [ordered]@{
      'NEG'   = 'Specimen-SampleType0'
      'NEG_2' = 'Specimen-SampleType1'
      'POS'   = 'Specimen-SampleType2'
      'MIX_1' = 'Specimen-SampleType3'
      'MIX_2' = 'Specimen-SampleType4'
    }

    # -----------------------------------------------------------------------
    # TEST TYPE LABELS – kontroller i ByTestType-assays
    # -----------------------------------------------------------------------
    'TestTypeMeta' = [ordered]@{
      'Negative Control 1' = 'Negative Control 1'
      'Positive Control 1' = 'Positive Control 1'
      'Positive Control 2' = 'Positive Control 2'
    }

    'StatusResultError' = [ordered]@{
      'AllowedStatus'         = @('Done', 'Aborted', 'Incomplete', 'Invalid', 'Error', 'In Progress')
      'MinorFunctionalResults' = @('INVALID', 'ERROR', 'NO RESULT')
      'MajorFunctionalPolicy' = [ordered]@{
        'Classification' = 'FUNCTIONAL_MAJOR_FALSE_CALL'
        'Message'        = 'Incorrect call (False Negative / False Positive) is a Major Functional Failure.'
      }
      'SpecialAllowedInvalid' = @(
        [ordered]@{
          'AssayRegex'    = '(?i)HPV'
          'TestTypeRegex' = '(?i)^Negative Control 1$'
        }
      )
    }

    'MaxPressure' = [ordered]@{
      'Field'          = 'Max Pressure (PSI)'
      'MustBeLessThan' = 90
      'Severity'       = 'Error'
      'Message'        = 'Max Pressure (PSI) must be < 90 for all products.'
    }

    'HardwareSignals' = [ordered]@{
      'ModuleSnRepeatErrorThreshold' = 3
      'Severity' = 'Warn'
      'Message'  = 'Same Module S/N recurring with errors may indicate a module-related issue.'
    }
  }

  # ---------------------------------------------------------------------------
  # ERROR BANK – Error-koder och specialregler
  # ---------------------------------------------------------------------------
  'ErrorBank' = [ordered]@{
    'ExtractRegex' = '(?i)\bError\s+(?<Code>\d{4})\b'
    'Codes' = [ordered]@{
      5001 = [ordered]@{
        'Name'           = 'Curve Fit Error'
        'Group'          = 'CurveFit'
        'Classification' = 'FUNCTIONAL_MINOR_CURVEFIT'
        'Description'    = 'Curve fit error – true functional failure, no re-test'
        'GeneratesRetest' = $false
      }
      5002 = [ordered]@{
        'Name'           = 'Curve Fit Error'
        'Group'          = 'CurveFit'
        'Classification' = 'FUNCTIONAL_MINOR_CURVEFIT'
        'Description'    = 'Curve fit error – true functional failure, no re-test'
        'GeneratesRetest' = $false
      }
      5003 = [ordered]@{
        'Name'           = 'Curve Fit Error'
        'Group'          = 'CurveFit'
        'Classification' = 'FUNCTIONAL_MINOR_CURVEFIT'
        'Description'    = 'Curve fit error – true functional failure, no re-test'
        'GeneratesRetest' = $false
      }
      5004 = [ordered]@{
        'Name'           = 'Curve Fit Error'
        'Group'          = 'CurveFit'
        'Classification' = 'FUNCTIONAL_MINOR_CURVEFIT'
        'Description'    = 'Curve fit error – true functional failure, no re-test'
        'GeneratesRetest' = $false
      }
      5005 = [ordered]@{
        'Name'           = 'Curve Fit Error'
        'Group'          = 'CurveFit'
        'Classification' = 'FUNCTIONAL_MINOR_CURVEFIT'
        'Description'    = 'Curve fit error – true functional failure, no re-test'
        'GeneratesRetest' = $false
      }
      5015 = [ordered]@{
        'Name'           = 'Curve Fit Error'
        'Group'          = 'CurveFit'
        'Classification' = 'FUNCTIONAL_MINOR_CURVEFIT'
        'Description'    = 'Curve fit error – true functional failure, no re-test'
        'GeneratesRetest' = $false
      }
      5016 = [ordered]@{
        'Name'           = 'Curve Fit Error'
        'Group'          = 'CurveFit'
        'Classification' = 'FUNCTIONAL_MINOR_CURVEFIT'
        'Description'    = 'Curve fit error – true functional failure, no re-test'
        'GeneratesRetest' = $false
      }
      5007 = [ordered]@{
        'Name'           = 'Probe Check Low'
        'Group'          = 'ProbeCheckLow'
        'Classification' = 'FUNCTIONAL_MINOR_PROBECHECK_LOW'
        'Description'    = 'Probe Check Low – true functional failure, no re-test'
        'GeneratesRetest' = $false
      }
      5017 = [ordered]@{
        'Name'           = 'Probe Check Low'
        'Group'          = 'ProbeCheckLow'
        'Classification' = 'FUNCTIONAL_MINOR_PROBECHECK_LOW'
        'Description'    = 'Probe Check Low – true functional failure, no re-test'
        'GeneratesRetest' = $false
      }
      5019 = [ordered]@{
        'Name'           = 'Probe Check Low'
        'Group'          = 'ProbeCheckLow'
        'Classification' = 'FUNCTIONAL_MINOR_PROBECHECK_LOW'
        'Description'    = 'Probe Check Low – true functional failure, no re-test'
        'GeneratesRetest' = $false
      }
      5006 = [ordered]@{
        'Name'           = 'Probe Check High'
        'Group'          = 'ProbeCheckHigh'
        'Classification' = 'FUNCTIONAL_MINOR_PROBECHECK_HIGH'
        'Description'    = 'Probe Check High – true functional failure, no re-test'
        'GeneratesRetest' = $false
      }
      5018 = [ordered]@{
        'Name'           = 'Probe Check High'
        'Group'          = 'ProbeCheckHigh'
        'Classification' = 'FUNCTIONAL_MINOR_PROBECHECK_HIGH'
        'Description'    = 'Probe Check High – true functional failure, no re-test'
        'GeneratesRetest' = $false
      }
      2037 = [ordered]@{
        'Name'           = 'CIT'
        'Group'          = 'CIT/SLD'
        'Classification' = 'FUNCTIONAL_MINOR_CIT_SLD'
        'Description'    = 'CIT – true functional failure, no re-test'
        'GeneratesRetest' = $false
      }
      5011 = [ordered]@{
        'Name'           = 'SLD'
        'Group'          = 'CIT/SLD'
        'Classification' = 'FUNCTIONAL_MINOR_CIT_SLD'
        'Description'    = 'SLD – true functional failure, no re-test'
        'GeneratesRetest' = $false
      }
      2008 = [ordered]@{
        'Name'           = 'Pressure Abort'
        'Group'          = 'Pressure'
        'Classification' = 'FUNCTIONAL_MINOR_PRESSURE'
        'Description'    = 'Pressure abort; protocol limit exceeded (also if Max Pressure > 90 PSI)'
        'GeneratesRetest' = $false
      }
      2096 = [ordered]@{
        'Name'           = 'Sample Volume Adequacy'
        'Group'          = 'SVA'
        'Classification' = 'FUNCTIONAL_MINOR_SVA'
        'Description'    = 'Sample Volume Adequacy failure – true functional failure, no re-test'
        'GeneratesRetest' = $false
      }
      2097 = [ordered]@{
        'Name'           = 'Sample Volume Adequacy'
        'Group'          = 'SVA'
        'Classification' = 'FUNCTIONAL_MINOR_SVA'
        'Description'    = 'Sample Volume Adequacy failure – true functional failure, no re-test'
        'GeneratesRetest' = $false
      }
      2125 = [ordered]@{
        'Name'           = 'Sample Volume Adequacy'
        'Group'          = 'SVA'
        'Classification' = 'FUNCTIONAL_MINOR_SVA'
        'Description'    = 'Sample Volume Adequacy failure – true functional failure, no re-test'
        'GeneratesRetest' = $false
      }
    }
    'Special' = @(
      [ordered]@{
        'Id'             = 'MAX_PRESSURE_GTE_90'
        'Classification' = 'FUNCTIONAL_MINOR_PRESSURE'
        'GeneratesRetest' = $false
        'When' = [ordered]@{
          'MaxPressureGte' = 90
        }
        'Message' = 'Max Pressure (PSI) >= 90 is a functional failure.'
      }
      [ordered]@{
        'Id'             = 'INVALID_NO_ERRORCODE'
        'Classification' = 'FUNCTIONAL_MINOR_INVALID'
        'GeneratesRetest' = $false
        'When' = [ordered]@{
          'StatusIn'        = @('Done')
          'TestResultEquals' = 'INVALID'
          'ErrorBlank'       = $true
        }
        'Message' = 'INVALID without error code is a minor functional failure (except allowed HPV NC case).'
      }
      [ordered]@{
        'Id'             = 'DELAM_SUFFIX_WITH_ERROR'
        'Classification' = 'FUNCTIONAL_MINOR_DELAM'
        'GeneratesRetest' = $false
        'When' = [ordered]@{
          'SampleIdDelamSuffix' = $true
          'ErrorPresent'        = $true
        }
        'Message' = 'Any error combined with delamination suffix is classified as delamination.'
      }
      [ordered]@{
        'Id'             = 'OTHER_ERROR_CODES'
        'Classification' = 'INSTRUMENT_ERROR'
        'GeneratesRetest' = $true
        'When' = [ordered]@{
          'ErrorPresent'        = $true
          'ErrorCodeNotInBank'  = $true
        }
        'Message' = 'Unknown error codes are classified as INSTRUMENT_ERROR by default.'
      }
    )
  }

  # ---------------------------------------------------------------------------
  # ASSAY-PROFILER – per assay: mode + idx-testtyp + förväntat resultat
  # ---------------------------------------------------------------------------
  'AssayProfiles' = [ordered]@{

    # -----------------------------------------------------------------------
    # Respiratory Panel IUO – SpecimenOnly_IdxDriven med NEG + MIX_1-4
    # -----------------------------------------------------------------------
    'Respiratory Panel IUO' = [ordered]@{
      'AssayKey'    = 'RESP_PANEL_IUO'
      'DisplayName' = 'Respiratory Panel IUO'
      'Mode'        = 'SpecimenOnly_IdxDriven'

      'IdentityHints' = [ordered]@{
        'AssayVersionExpected' = 6
      }

      # IDX → rollabel (knyts vidare till Specimen-SampleTypeX via Global.RoleMeta)
      'IdxToRole' = [ordered]@{
        0 = 'NEG'    # NEG_ALL
        1 = 'MIX_1'
        2 = 'MIX_2'
        3 = 'MIX_3'
        4 = 'MIX_4'
      }

      'ExpectedCountsByIdx' = [ordered]@{
        0 = [ordered]@{
          'Target'   = 110
          'Severity' = 'Warn'
        }
        1 = [ordered]@{
          'Target'   = 40
          'Severity' = 'Warn'
        }
        2 = [ordered]@{
          'Target'   = 20
          'Severity' = 'Warn'
        }
        3 = [ordered]@{
          'Target'   = 20
          'Severity' = 'Warn'
        }
        4 = [ordered]@{
          'Target'   = 20
          'Severity' = 'Warn'
        }
      }

      'ExpectedLiteralByRole' = [ordered]@{
        'NEG' = 'Adenovirus NEGATIVE;Coronavirus NEGATIVE;SARS-CoV-2 NEGATIVE;Flu A NEGATIVE;Flu B NEGATIVE;hMPV NEGATIVE;Parainfluenza NEGATIVE;Rhino-Enterovirus NEGATIVE;RSV NEGATIVE;B. parapertussis NEGATIVE;B. pertussis NEGATIVE;Chlam. pneumoniae NEGATIVE;Myco. pneumoniae NEGATIVE'
        'MIX_1' = 'Adenovirus NEGATIVE;Coronavirus POSITIVE;SARS-CoV-2 NEGATIVE;Flu A POSITIVE;Flu B NEGATIVE;hMPV NEGATIVE;Parainfluenza POSITIVE;Rhino-Enterovirus NEGATIVE;RSV POSITIVE;B. parapertussis NEGATIVE;B. pertussis NEGATIVE;Chlam. pneumoniae POSITIVE;Myco. pneumoniae NEGATIVE'
        'MIX_2' = 'Adenovirus NEGATIVE;Coronavirus POSITIVE;SARS-CoV-2 POSITIVE;Flu A NEGATIVE;Flu B POSITIVE;hMPV NEGATIVE;Parainfluenza POSITIVE;Rhino-Enterovirus NEGATIVE;RSV NEGATIVE;B. parapertussis POSITIVE;B. pertussis NEGATIVE;Chlam. pneumoniae NEGATIVE;Myco. pneumoniae NEGATIVE'
        'MIX_3' = 'Adenovirus POSITIVE;Coronavirus POSITIVE;SARS-CoV-2 NEGATIVE;Flu A POSITIVE;Flu B NEGATIVE;hMPV POSITIVE;Parainfluenza POSITIVE;Rhino-Enterovirus POSITIVE;RSV NEGATIVE;B. parapertussis NEGATIVE;B. pertussis NEGATIVE;Chlam. pneumoniae NEGATIVE;Myco. pneumoniae NEGATIVE'
        'MIX_4' = 'Adenovirus NEGATIVE;Coronavirus POSITIVE;SARS-CoV-2 NEGATIVE;Flu A NEGATIVE;Flu B NEGATIVE;hMPV NEGATIVE;Parainfluenza POSITIVE;Rhino-Enterovirus NEGATIVE;RSV POSITIVE;B. parapertussis NEGATIVE;B. pertussis POSITIVE;Chlam. pneumoniae NEGATIVE;Myco. pneumoniae POSITIVE'
      }

      'Notes' = 'Respiratory Panel IUO med NEG + 4 POS-mixar. Roller mappas vidare till Specimen-SampleType0/3/4 via Global.RoleMeta.'
    }

    # -----------------------------------------------------------------------
    # Xpert CT/NG – SpecimenOnly_IdxDriven (NEG/POS)
    # -----------------------------------------------------------------------
    'Xpert CT_NG' = [ordered]@{
      'AssayKey'    = 'CTNG'
      'DisplayName' = 'Xpert CT/NG'
      'Mode'        = 'SpecimenOnly_IdxDriven'

      'IdentityHints' = [ordered]@{
        'AssayVersionExpected' = 3
      }

      # Specimen-index → roll
      'IdxToRole' = [ordered]@{
        0 = 'NEG'  # Specimen-SampleType0
        1 = 'POS'  # Specimen-SampleType2
      }

      'ExpectedCountsByIdx' = [ordered]@{
        0 = [ordered]@{
          'Target'   = 110
          'Severity' = 'Warn'
        }
        1 = [ordered]@{
          'Target'   = 100
          'Severity' = 'Warn'
        }
      }

      'ExpectedResultByRole' = [ordered]@{
        'NEG' = [ordered]@{
          'WhenStatusIn' = @('Done')
          'Severity'     = 'Error'
          'ResultRegex'  = @(
            '(?i)^CT\s+NOT\s+DETECTED;NG\s+NOT\s+DETECTED$',
            '(?i)^INVALID$'
          )
        }
        'POS' = [ordered]@{
          'WhenStatusIn' = @('Done')
          'Severity'     = 'Error'
          'ResultRegex'  = @(
            '(?i)^CT\s+DETECTED;NG\s+DETECTED$'
          )
        }
      }

      'Notes' = 'Specimen-only CT/NG. IDX=0 är NEG-panel, IDX=1 POS-panel. Roller → Specimen-SampleType0/2 i rapport via Global.RoleMeta.'
    }

    # -----------------------------------------------------------------------
    # Xpert Ebola EUA – ByTestType
    # (oförändrad logik, bara följer din struktur)
    # -----------------------------------------------------------------------
    'Xpert Ebola EUA' = [ordered]@{
      'Mode' = 'ByTestType'
      'IdentityHints' = [ordered]@{
        'AssayVersionExpected' = 2
      }
      'ExpectedCounts' = @(
        [ordered]@{
          'TestType' = 'Negative Control 1'
          'Target'   = 110
          'Severity' = 'Warn'
        }
        [ordered]@{
          'TestType' = 'Positive Control 1'
          'Target'   = 80
          'Severity' = 'Warn'
        }
        [ordered]@{
          'TestType' = 'Positive Control 2'
          'Target'   = 20
          'Severity' = 'Warn'
        }
      )
      'ExpectedResultByTestType' = [ordered]@{
        'Negative Control 1' = [ordered]@{
          'WhenStatusIn' = @('Done')
          'Severity'     = 'Error'
          'ResultRegex'  = @('(?i)^Ebola\s+GP\s+NOT\s+DETECTED;Ebola\s+NP\s+NOT\s+DETECTED$')
        }
        'Positive Control 1' = [ordered]@{
          'WhenStatusIn' = @('Done')
          'Severity'     = 'Error'
          'ResultRegex'  = @('(?i)^Ebola\s+GP\s+DETECTED;Ebola\s+NP\s+DETECTED$')
        }
        'Positive Control 2' = [ordered]@{
          'WhenStatusIn' = @('Done')
          'Severity'     = 'Error'
          'ResultRegex'  = @('(?i)^Ebola\s+GP\s+DETECTED;Ebola\s+NP\s+DETECTED$')
        }
      }
    }

    # -----------------------------------------------------------------------
    # Xpert GBS LB XC – SpecimenOnly_IdxDriven (som du hade)
    # -----------------------------------------------------------------------
    'Xpert GBS LB XC' = [ordered]@{
      'Mode' = 'SpecimenOnly_IdxDriven'
      'IdentityHints' = [ordered]@{
        'AssayVersionExpected' = 2
      }
      'IdxToRole' = [ordered]@{
        0 = 'ROLE_1'
        1 = 'POS'
      }
      'ExpectedCountsByIdx' = [ordered]@{
        0 = [ordered]@{
          'Target'   = 100
          'Severity' = 'Warn'
        }
        1 = [ordered]@{
          'Target'   = 110
          'Severity' = 'Warn'
        }
      }
      'ExpectedResultByRole' = [ordered]@{
        'ROLE_1' = [ordered]@{
          'WhenStatusIn' = @('Done')
          'Severity'     = 'Error'
          'ResultRegex'  = @('(?i)^GBS\s+NEGATIVE$')
        }
        'POS' = [ordered]@{
          'WhenStatusIn' = @('Done')
          'Severity'     = 'Error'
          'ResultRegex'  = @('(?i)^GBS\s+POSITIVE$')
        }
      }
    }

    # -----------------------------------------------------------------------
    # Xpert HBV Viral Load – ByTestType
    # -----------------------------------------------------------------------
    'Xpert HBV Viral Load' = [ordered]@{
      'Mode' = 'ByTestType'
      'IdentityHints' = [ordered]@{
        'AssayVersionExpected' = 2
      }
      'ExpectedCounts' = @(
        [ordered]@{
          'TestType' = 'Negative Control 1'
          'Target'   = 100
          'Severity' = 'Warn'
        }
        [ordered]@{
          'TestType' = 'Positive Control 1'
          'Target'   = 90
          'Severity' = 'Warn'
        }
        [ordered]@{
          'TestType' = 'Positive Control 2'
          'Target'   = 20
          'Severity' = 'Warn'
        }
      )
      'ExpectedResultByTestType' = [ordered]@{
        'Negative Control 1' = [ordered]@{
          'WhenStatusIn' = @('Done')
          'Severity'     = 'Error'
          'ResultRegex'  = @('(?i)^HBV\s+NOT\s+DETECTED$', '(?i)^INVALID$')
        }
        'Positive Control 1' = [ordered]@{
          'WhenStatusIn' = @('Done')
          'Severity'     = 'Error'
          'ResultRegex'  = @('(?i)^HBV\s+DETECTED\s+\d+\s+IU/mL\s+\(log\s+\d+\.\d+\)$')
        }
        'Positive Control 2' = [ordered]@{
          'WhenStatusIn' = @('Done')
          'Severity'     = 'Error'
          'ResultRegex'  = @('(?i)^HBV\s+DETECTED\s+\d+(?:\.\d+)?E\d+\s+IU/mL\s+\(log\s+\d+\.\d+\)$')
        }
      }
    }

    # -----------------------------------------------------------------------
    # Xpert HCV PQC – ByTestType (oförändrad)
    # -----------------------------------------------------------------------
    'Xpert HCV PQC' = [ordered]@{
      'Mode' = 'ByTestType'
      'IdentityHints' = [ordered]@{
        'AssayVersionExpected' = 1
      }
      'ExpectedCounts' = @(
        [ordered]@{
          'TestType' = 'Negative Control 1'
          'Target'   = 100
          'Severity' = 'Warn'
        }
        [ordered]@{
          'TestType' = 'Positive Control 1'
          'Target'   = 110
          'Severity' = 'Warn'
        }
      )
      'ExpectedResultByTestType' = [ordered]@{
        'Negative Control 1' = [ordered]@{
          'WhenStatusIn' = @('Done')
          'Severity'     = 'Error'
          'ResultRegex'  = @('(?i)^HCV\s+NOT\s+DETECTED$')
        }
        'Positive Control 1' = [ordered]@{
          'WhenStatusIn' = @('Done')
          'Severity'     = 'Error'
          'ResultRegex'  = @('(?i)^HCV\s+DETECTED$')
        }
      }
    }

    # -----------------------------------------------------------------------
    # Xpert HCV VL Fingerstick – ByTestType
    # -----------------------------------------------------------------------
    'Xpert HCV VL Fingerstick' = [ordered]@{
      'Mode' = 'ByTestType'
      'IdentityHints' = [ordered]@{
        'AssayVersionExpected' = 1
      }
      'ExpectedCounts' = @(
        [ordered]@{
          'TestType' = 'Negative Control 1'
          'Target'   = 100
          'Severity' = 'Warn'
        }
        [ordered]@{
          'TestType' = 'Positive Control 1'
          'Target'   = 90
          'Severity' = 'Warn'
        }
        [ordered]@{
          'TestType' = 'Positive Control 2'
          'Target'   = 20
          'Severity' = 'Warn'
        }
      )
      'ExpectedResultByTestType' = [ordered]@{
        'Negative Control 1' = [ordered]@{
          'WhenStatusIn' = @('Done')
          'Severity'     = 'Error'
          'ResultRegex'  = @('(?i)^HCV\s+NOT\s+DETECTED$')
        }
        'Positive Control 1' = [ordered]@{
          'WhenStatusIn' = @('Done')
          'Severity'     = 'Error'
          'ResultRegex'  = @('(?i)^HCV\s+DETECTED\s+\d+\s+IU/mL\s+\(log\s+\d+\.\d+\)$')
        }
        'Positive Control 2' = [ordered]@{
          'WhenStatusIn' = @('Done')
          'Severity'     = 'Error'
          'ResultRegex'  = @('(?i)^HCV\s+DETECTED\s+\d+(?:\.\d+)?E\d+\s+IU/mL\s+\(log\s+\d+\.\d+\)$')
        }
      }
    }

    # -----------------------------------------------------------------------
    # Xpert HIV-1 Qual XC PQC – SpecimenOnly_IdxDriven
    # -----------------------------------------------------------------------
    'Xpert HIV-1 Qual XC PQC' = [ordered]@{
      'Mode' = 'SpecimenOnly_IdxDriven'
      'IdentityHints' = [ordered]@{
        'AssayVersionExpected' = 1
      }
      'IdxToRole' = [ordered]@{
        0 = 'NEG'
        1 = 'POS'
      }
      'ExpectedCountsByIdx' = [ordered]@{
        0 = [ordered]@{
          'Target'   = 100
          'Severity' = 'Warn'
        }
        1 = [ordered]@{
          'Target'   = 110
          'Severity' = 'Warn'
        }
      }
      'ExpectedResultByRole' = [ordered]@{
        'NEG' = [ordered]@{
          'WhenStatusIn' = @('Done')
          'Severity'     = 'Error'
          'ResultRegex'  = @('(?i)^HIV\-\d+\s+NOT\s+DETECTED$')
        }
        'POS' = [ordered]@{
          'WhenStatusIn' = @('Done')
          'Severity'     = 'Error'
          'ResultRegex'  = @('(?i)^HIV\-\d+\s+DETECTED$')
        }
      }
    }

    # -----------------------------------------------------------------------
    # Xpert HIV-1 Viral Load XC – ByTestType
    # -----------------------------------------------------------------------
    'Xpert HIV-1 Viral Load XC' = [ordered]@{
      'Mode' = 'ByTestType'
      'IdentityHints' = [ordered]@{
        'AssayVersionExpected' = 3
      }
      'ExpectedCounts' = @(
        [ordered]@{
          'TestType' = 'Negative Control 1'
          'Target'   = 100
          'Severity' = 'Warn'
        }
        [ordered]@{
          'TestType' = 'Positive Control 1'
          'Target'   = 90
          'Severity' = 'Warn'
        }
        [ordered]@{
          'TestType' = 'Positive Control 2'
          'Target'   = 20
          'Severity' = 'Warn'
        }
      )
      'ExpectedResultByTestType' = [ordered]@{
        'Negative Control 1' = [ordered]@{
          'WhenStatusIn' = @('Done')
          'Severity'     = 'Error'
          'ResultRegex'  = @('(?i)^HIV\-\d+\s+NOT\s+DETECTED$')
        }
        'Positive Control 1' = [ordered]@{
          'WhenStatusIn' = @('Done')
          'Severity'     = 'Error'
          'ResultRegex'  = @('(?i)^HIV\-\d+\s+DETECTED\s+\d+\s+copies/mL\s+\(log\s+\d+\.\d+\)$')
        }
        'Positive Control 2' = [ordered]@{
          'WhenStatusIn' = @('Done')
          'Severity'     = 'Error'
          'ResultRegex'  = @('(?i)^HIV\-\d+\s+DETECTED\s+\d+(?:\.\d+)?E\d+\s+copies/mL\s+\(log\s+\d+\.\d+\)$')
        }
      }
    }

    # -----------------------------------------------------------------------
    # Xpert HPV v2 HR – ByTestType
    # -----------------------------------------------------------------------
    'Xpert HPV v2 HR' = [ordered]@{
      'Mode' = 'ByTestType'
      'IdentityHints' = [ordered]@{
        'AssayVersionExpected' = 1
      }
      'ExpectedCounts' = @(
        [ordered]@{
          'TestType' = 'Negative Control 1'
          'Target'   = 60
          'Severity' = 'Warn'
        }
        [ordered]@{
          'TestType' = 'Positive Control 1'
          'Target'   = 130
          'Severity' = 'Warn'
        }
        [ordered]@{
          'TestType' = 'Positive Control 2'
          'Target'   = 20
          'Severity' = 'Warn'
        }
      )
      'ExpectedResultByTestType' = [ordered]@{
        'Negative Control 1' = [ordered]@{
          'WhenStatusIn' = @('Done')
          'Severity'     = 'Error'
          'ResultRegex'  = @('(?i)^INVALID$')
        }
        'Positive Control 1' = [ordered]@{
          'WhenStatusIn' = @('Done')
          'Severity'     = 'Error'
          'ResultRegex'  = @('(?i)^HR\s+HPV\s+POS$')
        }
        'Positive Control 2' = [ordered]@{
          'WhenStatusIn' = @('Done')
          'Severity'     = 'Error'
          'ResultRegex'  = @('(?i)^HR\s+HPV\s+POS$')
        }
      }
    }

    # -----------------------------------------------------------------------
    # Xpert MRSA NxG – ByTestType
    # -----------------------------------------------------------------------
    'Xpert MRSA NxG' = [ordered]@{
      'Mode' = 'ByTestType'
      'IdentityHints' = [ordered]@{
        'AssayVersionExpected' = 2
      }
      'ExpectedCounts' = @(
        [ordered]@{
          'TestType' = 'Negative Control 1'
          'Target'   = 110
          'Severity' = 'Warn'
        }
        [ordered]@{
          'TestType' = 'Positive Control 1'
          'Target'   = 80
          'Severity' = 'Warn'
        }
        [ordered]@{
          'TestType' = 'Positive Control 2'
          'Target'   = 20
          'Severity' = 'Warn'
        }
      )
      'ExpectedResultByTestType' = [ordered]@{
        'Negative Control 1' = [ordered]@{
          'WhenStatusIn' = @('Done')
          'Severity'     = 'Error'
          'ResultRegex'  = @('(?i)^MRSA\s+NOT\s+DETECTED$')
        }
        'Positive Control 1' = [ordered]@{
          'WhenStatusIn' = @('Done')
          'Severity'     = 'Error'
          'ResultRegex'  = @('(?i)^MRSA\s+DETECTED$')
        }
        'Positive Control 2' = [ordered]@{
          'WhenStatusIn' = @('Done')
          'Severity'     = 'Error'
          'ResultRegex'  = @('(?i)^MRSA\s+DETECTED$')
        }
      }
    }

    # -----------------------------------------------------------------------
    # Xpert MTB-RIF Assay G4 – SpecimenOnly_IdxDriven (NEG/NEG_2)
    # -----------------------------------------------------------------------
    'Xpert MTB-RIF Assay G4' = [ordered]@{
      'AssayKey'    = 'MTB_RIF_G4'
      'DisplayName' = 'Xpert MTB-RIF Assay G4'
      'Mode'        = 'SpecimenOnly_IdxDriven'

      'IdentityHints' = [ordered]@{
        'AssayVersionExpected' = 6
      }

      'IdxToRole' = [ordered]@{
        0 = 'NEG'
        1 = 'NEG_2'
      }

      'ExpectedCountsByIdx' = [ordered]@{
        0 = [ordered]@{
          'Target'   = 110
          'Severity' = 'Warn'
        }
        1 = [ordered]@{
          'Target'   = 100
          'Severity' = 'Warn'
        }
      }

      'ExpectedResultByRole' = [ordered]@{
        'NEG' = [ordered]@{
          'WhenStatusIn' = @('Done')
          'Severity'     = 'Error'
          'ResultRegex'  = @(
            '(?i)^MTB\s+NOT\s+DETECTED$',
            '(?i)^ERROR$'
          )
        }
        'NEG_2' = [ordered]@{
          'WhenStatusIn' = @('Done')
          'Severity'     = 'Error'
          'ResultRegex'  = @(
            '(?i)^MTB\s+DETECTED\s+MEDIUM;Rif\s+Resistance\s+NOT\s+DETECTED$',
            '(?i)^MTB\s+DETECTED\s+LOW;Rif\s+Resistance\s+NOT\s+DETECTED$'
          )
        }
      }

      'Notes' = 'Specimen-only assay. IDX=0/1 styr NEG/NEG_2-panel; mappas vidare till Specimen-SampleType0/1 i rapport via Global.RoleMeta.'
    }

    # -----------------------------------------------------------------------
    # Xpert MTB-RIF JP IVD – SpecimenOnly_IdxDriven (NEG/NEG_2/POS)
    # -----------------------------------------------------------------------
    'Xpert MTB-RIF JP IVD' = [ordered]@{
      'Mode' = 'SpecimenOnly_IdxDriven'
      'IdentityHints' = [ordered]@{
        'AssayVersionExpected' = 1
      }
      'IdxToRole' = [ordered]@{
        0 = 'NEG'
        1 = 'NEG_2'
        2 = 'POS'
        3 = 'POS'
        4 = 'POS'
        5 = 'POS'
      }
      'ExpectedCountsByIdx' = [ordered]@{
        0 = [ordered]@{
          'Target'   = 110
          'Severity' = 'Warn'
        }
        1 = [ordered]@{
          'Target'   = 60
          'Severity' = 'Warn'
        }
        2 = [ordered]@{
          'Target'   = 10
          'Severity' = 'Warn'
        }
        3 = [ordered]@{
          'Target'   = 10
          'Severity' = 'Warn'
        }
        4 = [ordered]@{
          'Target'   = 10
          'Severity' = 'Warn'
        }
        5 = [ordered]@{
          'Target'   = 10
          'Severity' = 'Warn'
        }
      }
      'ExpectedResultByRole' = [ordered]@{
        'NEG' = [ordered]@{
          'WhenStatusIn' = @('Done')
          'Severity'     = 'Error'
          'ResultRegex'  = @('(?i)^MTB\s+NOT\s+DETECTED$')
        }
        'NEG_2' = [ordered]@{
          'WhenStatusIn' = @('Done')
          'Severity'     = 'Error'
          'ResultRegex'  = @('(?i)^MTB\s+DETECTED;Rif\s+Resistance\s+NOT\s+DETECTED$')
        }
        'POS' = [ordered]@{
          'WhenStatusIn' = @('Done')
          'Severity'     = 'Error'
          'ResultRegex'  = @('(?i)^MTB\s+DETECTED;Rif\s+Resistance\s+DETECTED$')
        }
      }
    }

    # -----------------------------------------------------------------------
    # Xpert MTB-RIF Ultra – ByTestType
    # -----------------------------------------------------------------------
    'Xpert MTB-RIF Ultra' = [ordered]@{
      'Mode' = 'ByTestType'
      'IdentityHints' = [ordered]@{
        'AssayVersionExpected' = 4
      }
      'ExpectedCounts' = @(
        [ordered]@{
          'TestType' = 'Negative Control 1'
          'Target'   = 110
          'Severity' = 'Warn'
        }
        [ordered]@{
          'TestType' = 'Positive Control 1'
          'Target'   = 80
          'Severity' = 'Warn'
        }
        [ordered]@{
          'TestType' = 'Positive Control 2'
          'Target'   = 20
          'Severity' = 'Warn'
        }
      )
      'ExpectedResultByTestType' = [ordered]@{
        'Negative Control 1' = [ordered]@{
          'WhenStatusIn' = @('Done')
          'Severity'     = 'Error'
          'ResultRegex'  = @('(?i)^MTB\s+NOT\s+DETECTED$')
        }
        'Positive Control 1' = [ordered]@{
          'WhenStatusIn' = @('Done')
          'Severity'     = 'Error'
          'ResultRegex'  = @(
            '(?i)^MTB\s+DETECTED\s+LOW;\s*RIF\s+Resistance\s+DETECTED$',
            '(?i)^MTB\s+DETECTED\s+VERY\s+LOW;\s*RIF\s+Resistance\s+DETECTED$',
            '(?i)^MTB\s+DETECTED\s+MEDIUM;\s*RIF\s+Resistance\s+DETECTED$',
            '(?i)^MTB\s+DETECTED\s+HIGH;\s*RIF\s+Resistance\s+DETECTED$'
          )
        }
        'Positive Control 2' = [ordered]@{
          'WhenStatusIn' = @('Done')
          'Severity'     = 'Error'
          'ResultRegex'  = @(
            '(?i)^MTB\s+DETECTED\s+LOW;\s*RIF\s+Resistance\s+NOT\s+DETECTED$',
            '(?i)^MTB\s+DETECTED\s+VERY\s+LOW;\s*RIF\s+Resistance\s+NOT\s+DETECTED$',
            '(?i)^MTB\s+DETECTED\s+MEDIUM;\s*RIF\s+Resistance\s+NOT\s+DETECTED$',
            '(?i)^MTB\s+DETECTED\s+HIGH;\s*RIF\s+Resistance\s+NOT\s+DETECTED$'
          )
        }
      }
    }

    # -----------------------------------------------------------------------
    # Xpert MTB-XDR – ByTestType
    # -----------------------------------------------------------------------
    'Xpert MTB-XDR' = [ordered]@{
      'Mode' = 'ByTestType'
      'IdentityHints' = [ordered]@{
        'AssayVersionExpected' = 1
      }
      'ExpectedCounts' = @(
        [ordered]@{
          'TestType' = 'Negative Control 1'
          'Target'   = 110
          'Severity' = 'Warn'
        }
        [ordered]@{
          'TestType' = 'Positive Control 1'
          'Target'   = 80
          'Severity' = 'Warn'
        }
        [ordered]@{
          'TestType' = 'Positive Control 2'
          'Target'   = 20
          'Severity' = 'Warn'
        }
      )
      'ExpectedResultByTestType' = [ordered]@{
        'Negative Control 1' = [ordered]@{
          'WhenStatusIn' = @('Done')
          'Severity'     = 'Error'
          'ResultRegex'  = @('(?i)^MTB\s+NOT\s+DETECTED$')
        }
        'Positive Control 1' = [ordered]@{
          'WhenStatusIn' = @('Done')
          'Severity'     = 'Error'
          'ResultRegex'  = @('(?i)^MTB\s+DETECTED;INH\s+Resistance\s+NOT\s+DETECTED;FLQ\s+Resistance\s+NOT\s+DETECTED;AMK\s+Resistance\s+NOT\s+DETECTED;KAN\s+Resistance\s+NOT\s+DETECTED;CAP\s+Resistance\s+NOT\s+DETECTED;ETH\s+Resistance\s+NOT\s+DETECTED$')
        }
        'Positive Control 2' = [ordered]@{
          'WhenStatusIn' = @('Done')
          'Severity'     = 'Error'
          'ResultRegex'  = @('(?i)^MTB\s+DETECTED;INH\s+Resistance\s+DETECTED;FLQ\s+Resistance\s+DETECTED;AMK\s+Resistance\s+DETECTED;KAN\s+Resistance\s+DETECTED;CAP\s+Resistance\s+DETECTED;ETH\s+Resistance\s+DETECTED$')
        }
      }
    }

    # -----------------------------------------------------------------------
    # Xpert Norovirus – ByTestType
    # -----------------------------------------------------------------------
    'Xpert Norovirus' = [ordered]@{
      'Mode' = 'ByTestType'
      'IdentityHints' = [ordered]@{
        'AssayVersionExpected' = 1
      }
      'ExpectedCounts' = @(
        [ordered]@{
          'TestType' = 'Negative Control 1'
          'Target'   = 110
          'Severity' = 'Warn'
        }
        [ordered]@{
          'TestType' = 'Positive Control 1'
          'Target'   = 80
          'Severity' = 'Warn'
        }
        [ordered]@{
          'TestType' = 'Positive Control 2'
          'Target'   = 20
          'Severity' = 'Warn'
        }
      )
      'ExpectedResultByTestType' = [ordered]@{
        'Negative Control 1' = [ordered]@{
          'WhenStatusIn' = @('Done')
          'Severity'     = 'Error'
          'ResultRegex'  = @('(?i)^NORO\s+GI\s+NOT\s+DETECTED;NORO\s+GII\s+NOT\s+DETECTED$')
        }
        'Positive Control 1' = [ordered]@{
          'WhenStatusIn' = @('Done')
          'Severity'     = 'Error'
          'ResultRegex'  = @('(?i)^NORO\s+GI\s+DETECTED;NORO\s+GII\s+DETECTED$')
        }
        'Positive Control 2' = [ordered]@{
          'WhenStatusIn' = @('Done')
          'Severity'     = 'Error'
          'ResultRegex'  = @('(?i)^NORO\s+GI\s+DETECTED;NORO\s+GII\s+DETECTED$')
        }
      }
    }

    # -----------------------------------------------------------------------
    # Xpert Xpress CoV-2 plus IVD – ByTestType
    # -----------------------------------------------------------------------
    'Xpert Xpress CoV-2 plus IVD' = [ordered]@{
      'Mode' = 'ByTestType'
      'IdentityHints' = [ordered]@{
        'AssayVersionExpected' = 2
      }
      'ExpectedCounts' = @(
        [ordered]@{
          'TestType' = 'Negative Control 1'
          'Target'   = 110
          'Severity' = 'Warn'
        }
        [ordered]@{
          'TestType' = 'Positive Control 1'
          'Target'   = 100
          'Severity' = 'Warn'
        }
      )
      'ExpectedResultByTestType' = [ordered]@{
        'Negative Control 1' = [ordered]@{
          'WhenStatusIn' = @('Done')
          'Severity'     = 'Error'
          'ResultRegex'  = @('(?i)^SARS\-CoV\-\d+\s+NEGATIVE$')
        }
        'Positive Control 1' = [ordered]@{
          'WhenStatusIn' = @('Done')
          'Severity'     = 'Error'
          'ResultRegex'  = @('(?i)^SARS\-CoV\-\d+\s+POSITIVE$')
        }
      }
    }

    # -----------------------------------------------------------------------
    # Xpert Xpress Flu-RSV – ByTestType
    # -----------------------------------------------------------------------
    'Xpert Xpress Flu-RSV' = [ordered]@{
      'Mode' = 'ByTestType'
      'IdentityHints' = [ordered]@{
        'AssayVersionExpected' = 6
      }
      'ExpectedCounts' = @(
        [ordered]@{
          'TestType' = 'Negative Control 1'
          'Target'   = 110
          'Severity' = 'Warn'
        }
        [ordered]@{
          'TestType' = 'Positive Control 1'
          'Target'   = 80
          'Severity' = 'Warn'
        }
        [ordered]@{
          'TestType' = 'Positive Control 2'
          'Target'   = 20
          'Severity' = 'Warn'
        }
      )
      'ExpectedResultByTestType' = [ordered]@{
        'Negative Control 1' = [ordered]@{
          'WhenStatusIn' = @('Done')
          'Severity'     = 'Error'
          'ResultRegex'  = @('(?i)^Flu\s+A\s+NEGATIVE;Flu\s+B\s+NEGATIVE;RSV\s+NEGATIVE$', '(?i)^INVALID$')
        }
        'Positive Control 1' = [ordered]@{
          'WhenStatusIn' = @('Done')
          'Severity'     = 'Error'
          'ResultRegex'  = @('(?i)^Flu\s+A\s+POSITIVE;Flu\s+B\s+POSITIVE;RSV\s+POSITIVE$')
        }
        'Positive Control 2' = [ordered]@{
          'WhenStatusIn' = @('Done')
          'Severity'     = 'Error'
          'ResultRegex'  = @('(?i)^Flu\s+A\s+POSITIVE;Flu\s+B\s+POSITIVE;RSV\s+POSITIVE$')
        }
      }
    }

    # -----------------------------------------------------------------------
    # Xpert Xpress GBS – SpecimenOnly_IdxDriven (NEG/POS)
    # -----------------------------------------------------------------------
    'Xpert Xpress GBS' = [ordered]@{
      'AssayKey'    = 'XPRESS_GBS'
      'DisplayName' = 'Xpert Xpress GBS'
      'Mode'        = 'SpecimenOnly_IdxDriven'

      'IdentityHints' = [ordered]@{
        'AssayVersionExpected' = 1
      }

      'IdxToRole' = [ordered]@{
        0 = 'NEG'
        1 = 'POS'
      }

      'ExpectedCountsByIdx' = [ordered]@{
        0 = [ordered]@{
          'Target'   = 100
          'Severity' = 'Warn'
        }
        1 = [ordered]@{
          'Target'   = 110
          'Severity' = 'Warn'
        }
      }

      'ExpectedResultByRole' = [ordered]@{
        'NEG' = [ordered]@{
          'WhenStatusIn' = @('Done')
          'Severity'     = 'Error'
          'ResultRegex'  = @('(?i)^GBS\s+NEGATIVE$')
        }
        'POS' = [ordered]@{
          'WhenStatusIn' = @('Done')
          'Severity'     = 'Error'
          'ResultRegex'  = @('(?i)^GBS\s+POSITIVE$')
        }
      }

      'Notes' = 'Specimen-only GBS. IDX=0 NEG, IDX=1 POS. Roller mappas till Specimen-SampleType0/2 i rapport via Global.RoleMeta.'
    }

    # -----------------------------------------------------------------------
    # Xpert Xpress GBS US-IVD – SpecimenOnly_IdxDriven
    # -----------------------------------------------------------------------
    'Xpert Xpress GBS US-IVD' = [ordered]@{
      'Mode' = 'SpecimenOnly_IdxDriven'
      'IdentityHints' = [ordered]@{
        'AssayVersionExpected' = 2
      }
      'IdxToRole' = [ordered]@{
        0 = 'ROLE_1'
        1 = 'POS'
      }
      'ExpectedCountsByIdx' = [ordered]@{
        0 = [ordered]@{
          'Target'   = 100
          'Severity' = 'Warn'
        }
        1 = [ordered]@{
          'Target'   = 110
          'Severity' = 'Warn'
        }
      }
      'ExpectedResultByRole' = [ordered]@{
        'ROLE_1' = [ordered]@{
          'WhenStatusIn' = @('Done')
          'Severity'     = 'Error'
          'ResultRegex'  = @('(?i)^GBS\s+PRESUMPTIVE\s+NEGATIVE$')
        }
        'POS' = [ordered]@{
          'WhenStatusIn' = @('Done')
          'Severity'     = 'Error'
          'ResultRegex'  = @('(?i)^GBS\s+POSITIVE$')
        }
      }
    }

    # -----------------------------------------------------------------------
    # Xpert Xpress SARS-CoV-2 CE-IVD – ByTestType
    # -----------------------------------------------------------------------
    'Xpert Xpress SARS-CoV-2 CE-IVD' = [ordered]@{
      'Mode' = 'ByTestType'
      'IdentityHints' = [ordered]@{
        'AssayVersionExpected' = 1
      }
      'ExpectedCounts' = @(
        [ordered]@{
          'TestType' = 'Negative Control 1'
          'Target'   = 110
          'Severity' = 'Warn'
        }
        [ordered]@{
          'TestType' = 'Positive Control 1'
          'Target'   = 100
          'Severity' = 'Warn'
        }
      )
      'ExpectedResultByTestType' = [ordered]@{
        'Negative Control 1' = [ordered]@{
          'WhenStatusIn' = @('Done')
          'Severity'     = 'Error'
          'ResultRegex'  = @('(?i)^SARS\-CoV\-\d+\s+NEGATIVE$')
        }
        'Positive Control 1' = [ordered]@{
          'WhenStatusIn' = @('Done')
          'Severity'     = 'Error'
          'ResultRegex'  = @('(?i)^SARS\-CoV\-\d+\s+POSITIVE$')
        }
      }
    }

    # -----------------------------------------------------------------------
    # Xpert Xpress Strep A – ByTestType
    # -----------------------------------------------------------------------
    'Xpert Xpress Strep A' = [ordered]@{
      'Mode' = 'ByTestType'
      'IdentityHints' = [ordered]@{
        'AssayVersionExpected' = 2
      }
      'ExpectedCounts' = @(
        [ordered]@{
          'TestType' = 'Negative Control 1'
          'Target'   = 110
          'Severity' = 'Warn'
        }
        [ordered]@{
          'TestType' = 'Positive Control 1'
          'Target'   = 80
          'Severity' = 'Warn'
        }
        [ordered]@{
          'TestType' = 'Positive Control 2'
          'Target'   = 20
          'Severity' = 'Warn'
        }
      )
      'ExpectedResultByTestType' = [ordered]@{
        'Negative Control 1' = [ordered]@{
          'WhenStatusIn' = @('Done')
          'Severity'     = 'Error'
          'ResultRegex'  = @('(?i)^Strep\s+A\s+NOT\s+DETECTED$')
        }
        'Positive Control 1' = [ordered]@{
          'WhenStatusIn' = @('Done')
          'Severity'     = 'Error'
          'ResultRegex'  = @('(?i)^Strep\s+A\s+DETECTED$')
        }
        'Positive Control 2' = [ordered]@{
          'WhenStatusIn' = @('Done')
          'Severity'     = 'Error'
          'ResultRegex'  = @('(?i)^Strep\s+A\s+DETECTED$')
        }
      }
    }

    # -----------------------------------------------------------------------
    # Xpert vanA vanB – ByTestType
    # -----------------------------------------------------------------------
    'Xpert vanA vanB' = [ordered]@{
      'Mode' = 'ByTestType'
      'IdentityHints' = [ordered]@{
        'AssayVersionExpected' = 20
      }
      'ExpectedCounts' = @(
        [ordered]@{
          'TestType' = 'Negative Control 1'
          'Target'   = 110
          'Severity' = 'Warn'
        }
        [ordered]@{
          'TestType' = 'Positive Control 1'
          'Target'   = 100
          'Severity' = 'Warn'
        }
      )
      'ExpectedResultByTestType' = [ordered]@{
        'Negative Control 1' = [ordered]@{
          'WhenStatusIn' = @('Done')
          'Severity'     = 'Error'
          'ResultRegex'  = @('(?i)^vanA\s+NEGATIVE;vanB\s+NEGATIVE$')
        }
        'Positive Control 1' = [ordered]@{
          'WhenStatusIn' = @('Done')
          'Severity'     = 'Error'
          'ResultRegex'  = @('(?i)^vanA\s+POSITIVE;vanB\s+POSITIVE$')
        }
      }
    }

    # -----------------------------------------------------------------------
    # Xpert_Carba-R – ByTestType
    # -----------------------------------------------------------------------
    'Xpert_Carba-R' = [ordered]@{
      'Mode' = 'ByTestType'
      'IdentityHints' = [ordered]@{
        'AssayVersionExpected' = 1
      }
      'ExpectedCounts' = @(
        [ordered]@{
          'TestType' = 'Negative Control 1'
          'Target'   = 150
          'Severity' = 'Warn'
        }
        [ordered]@{
          'TestType' = 'Positive Control 1'
          'Target'   = 40
          'Severity' = 'Warn'
        }
        [ordered]@{
          'TestType' = 'Positive Control 2'
          'Target'   = 20
          'Severity' = 'Warn'
        }
      )
      'ExpectedResultByTestType' = [ordered]@{
        'Negative Control 1' = [ordered]@{
          'WhenStatusIn' = @('Done')
          'Severity'     = 'Error'
          'ResultRegex'  = @('(?i)^IMP\s+NOT\s+DETECTED;VIM\s+NOT\s+DETECTED;NDM\s+NOT\s+DETECTED;KPC\s+NOT\s+DETECTED;OXA\d+\s+NOT\s+DETECTED$')
        }
        'Positive Control 1' = [ordered]@{
          'WhenStatusIn' = @('Done')
          'Severity'     = 'Error'
          'ResultRegex'  = @('(?i)^IMP\s+DETECTED;VIM\s+DETECTED;NDM\s+DETECTED;KPC\s+DETECTED;OXA\d+\s+DETECTED$')
        }
        'Positive Control 2' = [ordered]@{
          'WhenStatusIn' = @('Done')
          'Severity'     = 'Error'
          'ResultRegex'  = @('(?i)^IMP\s+DETECTED;VIM\s+DETECTED;NDM\s+DETECTED;KPC\s+DETECTED;OXA\d+\s+DETECTED$')
        }
      }
    }

    # -----------------------------------------------------------------------
    # Xpert_HCV Viral Load – ByTestType (duplikat av HCV VL men annan nyckel)
    # -----------------------------------------------------------------------
    'Xpert_HCV Viral Load' = [ordered]@{
      'Mode' = 'ByTestType'
      'IdentityHints' = [ordered]@{
        'AssayVersionExpected' = 1
      }
      'ExpectedCounts' = @(
        [ordered]@{
          'TestType' = 'Negative Control 1'
          'Target'   = 100
          'Severity' = 'Warn'
        }
        [ordered]@{
          'TestType' = 'Positive Control 1'
          'Target'   = 90
          'Severity' = 'Warn'
        }
        [ordered]@{
          'TestType' = 'Positive Control 2'
          'Target'   = 20
          'Severity' = 'Warn'
        }
      )
      'ExpectedResultByTestType' = [ordered]@{
        'Negative Control 1' = [ordered]@{
          'WhenStatusIn' = @('Done')
          'Severity'     = 'Error'
          'ResultRegex'  = @('(?i)^HCV\s+NOT\s+DETECTED$')
        }
        'Positive Control 1' = [ordered]@{
          'WhenStatusIn' = @('Done')
          'Severity'     = 'Error'
          'ResultRegex'  = @('(?i)^HCV\s+DETECTED\s+\d+\s+IU/mL\s+\(log\s+\d+\.\d+\)$')
        }
        'Positive Control 2' = [ordered]@{
          'WhenStatusIn' = @('Done')
          'Severity'     = 'Error'
          'ResultRegex'  = @('(?i)^HCV\s+DETECTED\s+\d+(?:\.\d+)?E\d+\s+IU/mL\s+\(log\s+\d+\.\d+\)$')
        }
      }
    }

    # -----------------------------------------------------------------------
    # Xpert_HIV-1 Viral Load – ByTestType
    # -----------------------------------------------------------------------
    'Xpert_HIV-1 Viral Load' = [ordered]@{
      'Mode' = 'ByTestType'
      'IdentityHints' = [ordered]@{
        'AssayVersionExpected' = 2
      }
      'ExpectedCounts' = @(
        [ordered]@{
          'TestType' = 'Negative Control 1'
          'Target'   = 100
          'Severity' = 'Warn'
        }
        [ordered]@{
          'TestType' = 'Positive Control 1'
          'Target'   = 90
          'Severity' = 'Warn'
        }
        [ordered]@{
          'TestType' = 'Positive Control 2'
          'Target'   = 20
          'Severity' = 'Warn'
        }
      )
      'ExpectedResultByTestType' = [ordered]@{
        'Negative Control 1' = [ordered]@{
          'WhenStatusIn' = @('Done')
          'Severity'     = 'Error'
          'ResultRegex'  = @('(?i)^HIV\-\d+\s+NOT\s+DETECTED$')
        }
        'Positive Control 1' = [ordered]@{
          'WhenStatusIn' = @('Done')
          'Severity'     = 'Error'
          'ResultRegex'  = @('(?i)^HIV\-\d+\s+DETECTED\s+\d+\s+copies/mL\s+\(log\s+\d+\.\d+\)$')
        }
        'Positive Control 2' = [ordered]@{
          'WhenStatusIn' = @('Done')
          'Severity'     = 'Error'
          'ResultRegex'  = @('(?i)^HIV\-\d+\s+DETECTED\s+\d+(?:\.\d+)?E\d+\s+copies/mL\s+\(log\s+\d+\.\d+\)$')
        }
      }
    }

    # -----------------------------------------------------------------------
    # Xpress Flu IPT_EAT off – ByTestType
    # -----------------------------------------------------------------------
    'Xpress Flu IPT_EAT off' = [ordered]@{
      'Mode' = 'ByTestType'
      'IdentityHints' = [ordered]@{
        'AssayVersionExpected' = 3
      }
      'ExpectedCounts' = @(
        [ordered]@{
          'TestType' = 'Negative Control 1'
          'Target'   = 110
          'Severity' = 'Warn'
        }
        [ordered]@{
          'TestType' = 'Positive Control 1'
          'Target'   = 80
          'Severity' = 'Warn'
        }
        [ordered]@{
          'TestType' = 'Positive Control 2'
          'Target'   = 20
          'Severity' = 'Warn'
        }
      )
      'ExpectedResultByTestType' = [ordered]@{
        'Negative Control 1' = [ordered]@{
          'WhenStatusIn' = @('Done')
          'Severity'     = 'Error'
          'ResultRegex'  = @('(?i)^Flu\s+A\s+NEGATIVE;Flu\s+B\s+NEGATIVE$')
        }
        'Positive Control 1' = [ordered]@{
          'WhenStatusIn' = @('Done')
          'Severity'     = 'Error'
          'ResultRegex'  = @('(?i)^Flu\s+A\s+POSITIVE;Flu\s+B\s+POSITIVE$')
        }
        'Positive Control 2' = [ordered]@{
          'WhenStatusIn' = @('Done')
          'Severity'     = 'Error'
          'ResultRegex'  = @('(?i)^Flu\s+A\s+POSITIVE;Flu\s+B\s+POSITIVE$')
        }
      }
    }

    # -----------------------------------------------------------------------
    # Xpress SARS-CoV-2_Flu_RSV plus – ByTestType
    # -----------------------------------------------------------------------
    'Xpress SARS-CoV-2_Flu_RSV plus' = [ordered]@{
      'Mode' = 'ByTestType'
      'IdentityHints' = [ordered]@{
        'AssayVersionExpected' = 2
      }
      'ExpectedCounts' = @(
        [ordered]@{
          'TestType' = 'Negative Control 1'
          'Target'   = 110
          'Severity' = 'Warn'
        }
        [ordered]@{
          'TestType' = 'Positive Control 1'
          'Target'   = 100
          'Severity' = 'Warn'
        }
      )
      'ExpectedResultByTestType' = [ordered]@{
        'Negative Control 1' = [ordered]@{
          'WhenStatusIn' = @('Done')
          'Severity'     = 'Error'
          'ResultRegex'  = @('(?i)^SARS\-CoV\-\d+\s+NEGATIVE;Flu\s+A\s+NEGATIVE;Flu\s+B\s+NEGATIVE;RSV\s+NEGATIVE$')
        }
        'Positive Control 1' = [ordered]@{
          'WhenStatusIn' = @('Done')
          'Severity'     = 'Error'
          'ResultRegex'  = @('(?i)^SARS\-CoV\-\d+\s+POSITIVE;Flu\s+A\s+POSITIVE;Flu\s+B\s+POSITIVE;RSV\s+POSITIVE$')
        }
      }
    }

    # -----------------------------------------------------------------------
    # DEFAULT-PROFIL – om assay inte känns igen
    # -----------------------------------------------------------------------
    'default' = [ordered]@{
      'Mode'  = 'GlobalOnly'
      'Notes' = 'Unknown assay -> apply only RuleBank.Global + RuleBank.ErrorBank.'
    }
  }
}