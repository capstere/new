ValidateAssay GoldenPatch v5
- Added Modules/RuleBank.ps1: ONE root object $global:RuleBank with Global + ErrorBank + AssayProfiles (auto-built from your exports)
- Main.ps1 now dot-sources RuleBank.ps1 right after Config.ps1.

Next step (engine):
- Use $RuleBank.AssayProfiles[<assay>] + $RuleBank.Global + $RuleBank.ErrorBank to produce Findings and the final report output.
