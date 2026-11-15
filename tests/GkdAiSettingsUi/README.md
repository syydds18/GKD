# GKD AI settings UI checks

Run from the repository root:

```sh
dotnet run --project tests/GkdAiSettingsUi/GkdAiSettingsUi.csproj -c Release
```

Requires .NET 8 SDK and the restored `Microsoft.NETFramework.ReferenceAssemblies.net48` 1.0.3 package. Uses Roslyn shipped with the SDK, with no additional test framework dependency.

## What runs here

- Compile and execute the actual `GkdAiSettingsEditor` extracted from the production form source with the real `GkdAiProfile.cs`: clone isolation, multiple saved profiles, unsaved drafts, defaults, deletion, invalid saves and backend-compatible IDs.
- Compile and execute the actual reply validation: reject scalar/HTTP status for a table, wrong shape, missing title, nonblank separator, incorrect header, wrong quantities and duplicated material rows; accept the complete 5×2 layout and arrays with nonzero lower bounds.
- Compile the complete form against net48 WinForms references, the existing `Ui` implementation and the actual profile class. Store/client compile seams throw if executed; this proves type compatibility with the agreed API, not backend execution.
- Execute the production key visibility method against lightweight control seams: hidden by default, reveal/hide captions, unchanged existing key and no dirty state. Parse native button wiring, profile/close masking resets and the flexible key-row layout.
- Parse source contracts for clickable actions, all editable mappings, masked keys, async cancellation, disabled conflicting controls, no automatic save in tests, disposal guards, timestamp/total latency and visible formula guidance.

TDD sequence was exercised on missing editor, missing probe/form and subsequent new-layout/backend compatibility changes. These checks are not native WinForms interaction tests and do not call a live service.

## API consumed by the form

Entry point: `internal sealed GkdAiSettingsForm : Form`, parameterless internal constructor. Parent opens `new GkdAiSettingsForm().ShowDialog(owner)`.

- `Ui.Base(Form, Size, string)` and `Ui.Button(string, bool = false)`.
- `GkdAiProfile`: public read/write `Id`, `Name`, `Provider`, `Endpoint`, `Model`, `ApiKey`, `PromptInput`, `RouteInput`, `AnswerOutput`, `TimeoutSeconds`; `Clone()` and `Validate()`.
- `GkdAiSettings`: public read/write `List<GkdAiProfile> Profiles` and `string DefaultProfileId`.
- `GkdAiStore.Load()` and `GkdAiStore.Save(GkdAiSettings)`.
- `await GkdAiClient.ExecuteAsync(GkdAiProfile, string prompt, object[,] data, string mode, CancellationToken)` returning a reply with `object Value`. Modes are exactly `单值` and `表格`. The form measures its own total elapsed time; it does not consume the backend elapsed-time field.
- Table reply `Value` is 5×2 for the capability fixture: row 0 nonempty title in column 0, remainder blank; row 1 blank; row 2 `名称 / 数量`; row 3 `钢筋 / 2`; row 4 `水泥 / 3`. This uses the client's `{title,data}` model contract; UI does not parse JSON or apply Excel formatting.

The actual backend currently requires D-format GUID IDs, a timeout of 1–600 seconds and at least one saved profile; the UI reflects those constraints. No additional backend methods are required.

## Required Windows / Excel acceptance (not executed on macOS)

1. Open from the integrated Ribbon action with an owner; verify native controls, keyboard navigation, 100% / 150% / 200% scaling, scrolling and visible footer/test controls.
2. Save two custom profiles, reopen, switch/default/delete and verify the right model and mappings are retained. Test discarding edits and cancelled discard. Verify last-profile deletion explains why it is refused.
3. Use valid service credentials locally. Click **连接 / 鉴权测试**: one tiny single-turn request, nonempty actual result, total milliseconds and local timestamp. Repeat with invalid key / endpoint / model; ensure failure and no credential/response leakage.
4. Modify a saved endpoint or model without saving, click either test, and reopen the form: test must use edited settings, while stored configuration stays unchanged. Editing settings invalidates the old displayed result.
5. Click **表格能力测试** against a capable service: it must validate title + blank separator + headers + both material rows, and display 5×2 success. A scalar, malformed table or wrong content must fail.
6. During a slow request, verify editing and duplicate tests are disabled while Cancel/Close work. Cancel then run again; close during a request and confirm no disposed-control exception.
7. In Excel separately verify `=GKD("总结",A2)` and `=GKD("整理为名称、数量两列",A2,1)` with an empty spill area and a dynamic-array host. Title styling, fonts and borders belong to parent integration and are not implemented by this settings form.
8. With an existing saved key, verify **明文** reveals it and changes to **隐藏**, and **隐藏** masks it again without changing the key or marking settings dirty. Switch profiles, create a new profile and close/reopen: the key must start hidden. At 100% / 150% / 200% scaling verify the input stretches and the native button remains visible within the existing field width.

No real credentials or configuration files are included in this test directory.
