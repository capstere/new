@{
    # SchemaVersion for RuleBank template
    SchemaVersion = '1.0'

    # Default settings applicable to all assays
    Defaults = @{
        # Canonical field names and their possible aliases in incoming CSV files
        FieldAliases = @{
            'SampleID'    = @('Sample ID', 'ID')
            'CartridgeSN' = @('Cartridge SN', 'Cartridge Serial')
            # Add additional canonical fields and their aliases here
        }
        # Fields that must exist; missing values will generate Error severity
        RequiredFields  = @('SampleID','CartridgeSN')
        # Fields that should exist; missing values will generate a Warning severity
        PreferredFields = @('ModuleSN', 'TestResult')
        # Fields that are nice to have; missing values have no impact on severity
        OptionalFields  = @('ProbeCheck')
        # Validation policy defines default severities for various validation issues
        ValidationPolicy = @{
            MissingRequiredSeverity  = 'Error'
            MissingPreferredSeverity = 'Warn'
            DuplicateSeverity        = 'Warn'
            # Additional policy settings can be added here
        }
    }

    # Global rule definitions used by the RuleEngine
    RuleDefinitions = @{
        'R001' = @{
            Title           = 'Dubblett av SampleID'
            Description     = 'SampleID förekommer mer än en gång.'
            DefaultSeverity = 'Error'
            Category        = 'Duplicate'
            EvidenceFields  = @('SampleID','Row#')
            SuggestedAction = 'Kontrollera filen och ta bort dubbletter.'
        }
        # Add additional rule definitions here
    }

    # Assay-specific profiles can override defaults for match names or enabled/disabled rules
    AssayProfiles = @(
        # Example profile (empty for template)
        #@{
        #    MatchNames   = @('AssayA','AssayA v2')
        #    EnabledRules = @('R001')
        #    DisabledRules= @()
        #    Overrides    = @{
        #        'R001' = @{ DefaultSeverity = 'Warn' }
        #    }
        #}
    )
}