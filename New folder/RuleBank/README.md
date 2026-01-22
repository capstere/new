# RuleBank Configuration Guide

## Filhierarki
- 01_ResultCallPatterns.csv: Tolkar "Test Result" → Call (POS/NEG/ERROR)
- 02_SampleExpectationRules.csv: Förväntat call baserat på Sample ID
- 03_ErrorCodes.csv: Felkodsmappning (inkl. intervall)
- 04_MissingSamplesConfig.csv: Konfig för saknade samples (template)
- 05_SampleIdMarkers.csv: Markers/tokenindex för Sample ID
- 06_ParityCheckConfig.csv: Parity/suffix-logik per assay
- 07_SampleNumberRules.csv: Sample-nummerregler (regex, min/max, padding)

## Lägga till nytt assay (checklista)
1. Lägg till patterns i 01_ResultCallPatterns.csv
2. Lägg till expectations i 02_SampleExpectationRules.csv
3. Lägg till markers i 05_SampleIdMarkers.csv (om behövs)
4. Lägg till parity-config i 06_ParityCheckConfig.csv (om behövs)
5. Lägg till sample-number-regler i 07_SampleNumberRules.csv
