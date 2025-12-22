# Click Less Project – struktur & best practise (v3)

## Var händer vad?

### 1) Entry point
- `Main.ps1` – GUI + workflow: Sök LSP → Välj filer → Skapa rapport.

### 2) Data + regler
- `Modules/RuleBank.ps1` – **data only** (assayprofiler, felkoder, aliases).
  - Nytt: `AssayDisplayNames` = hur vi “pratar” (korta, mänskliga namn).
  - Nytt: `ColumnDisplayNames` = mänskliga kolumnnamn (för Missing columns osv).

### 3) Rule engine (logik)
- `Modules/Rules/RuleEngine.Core.ps1` – **pure logic**
  - Bygger context från CSV (`New-AssayRuleContext`)
  - Kör regler (`Invoke-AssayRuleEngine`)
  - Nytt: sätter `Context.AssayDisplayName` + `Context.ReagentLotDisplay`.

### 4) Writers (rapport-layout)
- `Modules/Writers/Information2Writer.ps1` – **EPPlus writer only**
  - Skriver ett blad: `Information2`
  - Human-friendly sammanfattning överst, sen kompakta tabeller.
  - Capping: Findings max 50, Error codes max 60, Affected tests max 250.

### 5) Facade för backward compatibility
- `Modules/RuleEngine.ps1` – dot-sourcar Core + Writer (Main ändras inte).

## Prestanda
- `Modules/Config.ps1` har nu:
  - `$Config.Performance.AutoFitColumns` (default **false**)
  - `$Config.Performance.AutoFitMaxRows` (default 300)

AutoFit är ofta den tyngsta EPPlus-operationen. Slå på vid behov.

## Nästa steg (regler)
1) Lägg regler i `RuleBank.ps1` under `$Rules` (eller skapa AssayProfiles per assay).
2) Skriv “människotexter” i `Message` + håll `RuleId` stabil.
3) Om du vill visa mer “hur vi pratar”:
   - fyll på `AssayDisplayNames` (canonical → human)
   - fyll på `ColumnDisplayNames`
