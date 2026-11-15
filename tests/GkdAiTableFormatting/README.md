# Deferred GKD native formatting

Production implementation: `src/Outline.Addin/GkdAiTableFormatting.cs` (C# 7.3/net48).

## Integration

On the Excel UDF/main thread, after `ExcelAsyncUtil.RunTaskWithCancellation` returns a final result, call:

```csharp
GkdAiTableFormatting.Schedule(caller, finalArray);
GkdAiTableFormatting.ScheduleScalar(caller, finalString);
```

Call only the overload matching the final result. Exclude waiting/error/cancellation states; scalar strings starting `GKD：` are additionally refused by the helper. Do not call either API from the background request task. Call `Start()` in AutoOpen and `Stop()` in AutoClose on Excel's main thread. Stop disposes retry/notice timers and invalidates queued macros.

The table must already include a title at `[0,0]`, blank remaining title cells, a fully blank second row, then nonblank column headers and optional subsequent data. The helper rejects malformed layout or more than 10000 total output cells; it never inserts rows or writes values.

Schedule clones the array, captures `GET.CELL(41, caller)` through the C API, and enqueues a macro. Every attempt rechecks that exact formula, resolves the original workbook/sheet through the caller's sheet id, and refuses constants or wrapped/non-direct GKD calls. Tables additionally require the actual spill parent, same worksheet, same top-left and exact dimensions, no merges, and all expected values including the title. Pending/different old spill dimensions and values retry at most 16 macro attempts, with 250ms WinForms UI timers; ticks only enqueue macros. At most 256 retained request entries provide pending/result deduplication. Macros never use active workbook/sheet/selection.

Title: 黑体 → Times New Roman, 12, normal weight, center across selection (`7`); second row onwards: 宋体 → Times New Roman, 10.5. The sequence is deliberate; actual Chinese face retention requires native/version verification below. Top two rows have no borders. Header and data have continuous thin black grid borders with diagonals disabled. Scalar output changes only the anchor font name/size. No value/formula, merge/unmerge, fill, selection, column width or row-height writes occur. Formatting outside the current result footprint is not cleared when a spill shrinks.

## Mixed-script limitation (unresolved for live formula results)

A single formula cell containing `中文 ABC123` cannot be given arbitrary `Characters` rich-text runs. This is separate from Excel's historically reported sequential script-font retention. Per the user's requested strategy, the shared font helper sets 黑体 (title) or 宋体 (body/scalar) first, then uniformly sets Times New Roman, then the requested size. The portable tests verify this setter order, not the host's resulting script-specific fonts.

The Microsoft Q&A reports this sequence working in Excel 2013/2016 but reproduces the last-selected face replacing both scripts in Excel 2019: https://learn.microsoft.com/en-us/answers/questions/5022204/excel-2019-mixed-font-styles-for-chinese-english-p . Therefore mixed-script font compliance remains unverified and version-dependent until tested in native Windows Excel. A nonmodal status-bar notice explains the version variability, expires after five seconds via a queued macro, and restores the previous status only if the notice still owns it.

Evidence checked:

- Microsoft Excel Font object lists `Name`, not `NameFarEast`: https://learn.microsoft.com/en-us/office/vba/api/excel.font(object)
- `NameFarEast` belongs to Office Font2 text objects, not Excel worksheet Font: https://learn.microsoft.com/en-us/office/vba/api/office.font2.namefareast
- Microsoft's Characters example uses a constant string: https://learn.microsoft.com/en-us/office/vba/api/excel.range.characters
- Reported Excel formula restriction: Characters cannot format substrings of formula results: https://stackoverflow.com/questions/71448257/excel-vba-range-characters-properties-not-working
- OOXML phonetic font ids apply to pronunciation annotations, not base text: https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.spreadsheet.phoneticproperties?view=openxml-3.0.1

Sequentially setting `Font.Name` to a Chinese face then Times New Roman is not treated as proof of separate script slots. Theme mutation, phonetic text, workbook ZIP rewrites, or unsupported NameFarEast dispatch are not used. Per-character `Characters` formatting can be considered after an explicitly confirmed conversion to constant values, as separate work; the helper never replaces formulas to get rich text.

## Portable tests

```sh
dotnet run --project tests/GkdAiTableFormatting -c Release
dotnet build src/Outline.Addin -c Release
```

The executable compiles the real helper against minimal Excel-DNA/WinForms stubs and dynamic Excel fakes. Tests exercise actual formatting/guard/scheduling logic, array lower bounds, copied snapshots, stale formula/result rejection, pending-spill retries, bounded exhaustion, scalar legacy unsupported-member behavior, busy COM handling, unload, temporary notice ownership and no data/merge/fill writes. They do not prove native Excel COM semantics or font rendering.

## Windows Excel acceptance still required

1. Load the Windows XLL, run table and scalar GKD formulas, verify title/spacer/header/data values and spill dimensions, no merged cells and correct top-row center-across alignment.
2. Inspect Chinese title 黑体12; blank/header/data 宋体10.5; independent Latin/numeric cells Times New Roman. Mixed-script formula cells must show the limitation notice; do not mark their exact fonts compliant.
3. Verify borders only from the third relative row through the current spill. Recalculate into a new size, then test #SPILL!, edited/deleted formula, renamed sheet, changed active workbook, closed workbook and add-in unload while awaiting a result.
4. Confirm scalar formatting does not add title/border/alignment styling; verify a legacy host without HasSpill still formats scalar text.
5. Ensure repeated unchanged recalculation does not enqueue or restyle indefinitely and that status notices restore/preserve user status.
