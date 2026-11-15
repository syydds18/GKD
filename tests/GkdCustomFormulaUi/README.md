# Native custom formula manager UI

Entry point: `internal sealed GkdCustomFormulaForm : Form`, parameterless constructor. `Controller.Execute("customFormula")` opens an Excel-owned modal manager before requiring an active workbook, prevents duplicate manager windows, and `Controller.Close()` calls `CloseForShutdown()` to cancel an active design and close its editor.

## Core contract consumed

- `GkdCustomFormulaStore.Load()` → `List<GkdCustomFormula>`
- `GkdCustomFormulaStore.Save(List<GkdCustomFormula>)` → void; persist atomically and throw on failure, without changing the original file.
- Existing real `GkdCustomFormulaJson.Import(string)` → `List<GkdCustomFormula>` and `Export(IEnumerable<GkdCustomFormula>)` → string. Import is validation-only, never persistence.
- Existing real `GkdCustomFormula` read/write fields: `Id`, `ExcelName`, `DisplayName`, `Description`, `Kind`, `Inputs`, `FormulaExpression`, `PromptTemplate`, `OutputUnit`, `Version`, `Enabled`; `Validate()`.
- `GkdCustomFormulaKind.Deterministic` / `.Ai`.
- Existing `GkdAiStore.Load()`, `GkdAiProfile.Clone()/Validate()`, and `GkdAiClient.ExecuteAsync(profile, prompt, null, "单值", cancellationToken)` returning reply `.Value` as JSON text.

UI owns deep draft copying; no additional model methods are required. It uses versioned `SchemaVersion: 1` / `Formulas` JSON and string enum names. New/Edit/Import require nonblank meaning, a read-only preview, an unchecked reviewed checkbox, and explicit confirmation before writing. Delete requires a separate confirmation. Imports append only and reject duplicate IDs or Excel names; errors leave the saved manager state unchanged. Export includes all saved definitions.

AI design previews the exact design prompt and model name before allowing a network request. It passes no worksheet data, validates the returned single-definition JSON, and only fills an unsaved draft. A subsequent save still requires review. The UI never evaluates definitions, compiles generated code, calls Excel Evaluate, or inserts workbook formulas. Closing/cancelling an AI request prevents late publication to disposed controls. Corrupt local stores disable mutation rather than silently replacing the file.

## Exact tests

Run from the repository root (`dotnet` 8 SDK; on this machine use `/Users/tmsxhc/.hermes/cache/scratch/dotnet-sdk/dotnet`):

```sh
dotnet run --project tests/GkdCustomFormulaUi/GkdCustomFormulaUi.csproj -c Release
python3 tests/GkdCustomFormulaUi/test_integration.py
python3 tests/GkdAiIntegration/test_integration.py
dotnet run --project tests/GkdAiSettingsUi/GkdAiSettingsUi.csproj -c Release
```

Verified results: **26 custom formula UI checks**, **2 custom formula Ribbon/safety tests**, **6 existing AI integration tests**, **129 existing AI settings checks**.

The custom formula test compiles the complete form against net48 WinForms references, actual shared `Ui`, real model, real JSON importer and real AI protocol. Store/client compile seams throw if executed. Reflection-safe behavioral seams execute the production copy, meaning-preview, required-meaning validation, prompt builder, manager save and editor save methods. Cases cover isolated input lists, complete explanation, valid AI schema round-trip, no-data protocol acceptance, blank/oversized AI requirements, cancelled confirmation, failed writes preserving manager state and draft, and confirmed persistence. Source checks guard unchecked confirmation, prompt approval, no implicit AI save, cancellation/disposal guards, Ribbon icon/group/action and pre-workbook dispatch.

Full add-in build was attempted with `dotnet build src/Outline.Addin/Outline.Addin.csproj -c Release --nologo`; it currently requires the parent-owned `GkdCustomFormulaStore` contract above. It reported only four missing-store errors when attempted. Native UI type compatibility passes independently.

## Windows / Excel acceptance still required

- Open Ribbon 自定义公式 with/without workbook; inspect DPI 100%/150%/200%, keyboard navigation, read-only selectable meaning/algorithm preview and editor layout.
- Create/edit/reopen definitions in both kinds; cancel confirmation, leave reviewed unchecked, retry failed saves, delete and verify cancellation.
- Import/export round-trip, malformed/oversized JSON and duplicate conflicts; ensure cancelled import does not change the local library.
- Use a locally configured model: review prompt, send/cancel design, review filled draft, confirm save separately; close during design and unload the add-in without late control access.

No live AI request, native WinForms interaction, Excel formula execution, or real user formula library was exercised by these tests.
