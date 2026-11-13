# Live PrintArea reader

## Integration

```csharp
// Both the initial ReadChoices population and an explicit refresh button:
choice.PrintRange = HydroLivePrintArea.Read(app, choice.SourceSheet);
// During initial sheet enumeration, app can be workbook.Application.
```

API: `internal static string HydroLivePrintArea.Read(dynamic app, dynamic sheet)`
in namespace `Outline.Addin`. Call on Excel's UI/COM thread. The result is
local A1 address text without `$`, preserving comma-separated areas. Existing
extraction validation must continue rejecting non-contiguous/oversized ranges.
An empty result means all readable sources failed; report this rather than
silently exporting an empty range. No state is retained between calls.

Only new helper/test files were added. `HydropowerExtraction.cs` and the runner
are not modified here; parent integration must replace the old inline read.

## Selection and state contract

1. If available, read `Application.PrintCommunication`.
2. If false, temporarily set true, committing cached PageSetup commands.
3. Read a fresh `sheet.PageSetup.PrintArea`.
4. A valid nonempty PageSetup result is authoritative; do not consult stale names.
5. A successful empty result is also authoritative (area cleared): use UsedRange.
6. Only when PageSetup could not be read, freshly resolve the sheet-local
   `sheet.Names.Item("Print_Area").RefersToRange`; require the same sheet parent.
7. Otherwise use current UsedRange; if unavailable, return empty.
8. Restore original false PrintCommunication in finally, even after failures.

The helper never calls Save/SaveAs, never sets Workbook.Saved, never changes
PrintArea, never activates/selects a sheet, and never resets/moves page breaks.
It deliberately does not set Saved=true to hide existing or newly committed
workbook changes. Enabling PrintCommunication commits **application-wide** cached
PageSetup commands, so this is not a strictly side-effect-free getter. Excel may
mark a workbook dirty when pending settings are committed. The helper cannot
guarantee restoration if Excel rejects the restoration setter; this best-effort
failure is swallowed so it does not replace the read result.

Quoted worksheet qualifiers are parsed without splitting a comma inside the
sheet name; escaped apostrophes and `!` inside the name are supported. An explicit
different-sheet/external-workbook qualifier is rejected, not stripped into a
local address. The parser checks reference syntax, not Excel row/column limits;
the existing `sheet.Range[...]` export validation remains necessary.

## Evidence and limits

Microsoft documents:

- [PrintCommunication](https://learn.microsoft.com/en-us/office/vba/api/excel.application.printcommunication):
  setting true commits all cached PageSetup commands.
- [PrintArea](https://learn.microsoft.com/en-us/office/vba/api/excel.pagesetup.printarea):
  a read/write A1 address string; empty clears the explicit print area.
- [Name.RefersToRange](https://learn.microsoft.com/en-us/office/vba/api/excel.name.referstorange):
  includes a sheet-specific Print_Area example; non-range names can fail.
- [Worksheet.Names](https://learn.microsoft.com/en-us/office/vba/api/excel.worksheet.names):
  the collection contains worksheet-specific names.
- [Page breaks](https://support.microsoft.com/en-us/excel/insert-move-or-delete-page-breaks-in-a-worksheet):
  page breaks divide printing into pages. Moving an **inner** page divider is
  not the same operation as changing the **outer** printable extent/PrintArea.

These docs do NOT establish that every page-break-preview drag remains stale
until Save. That reported host bug could not be reproduced: native Windows Excel
is unavailable on this macOS host. This is a defensively tested refresh helper,
not a claim that the reported native UI failure has been proved fixed.

Portable C# 7.3 dynamic fakes cover pending/committed order, repeated fresh reads,
preserving initially true/false communication state, read/enable/partially
successful enable failures, same-sheet current names, cleared-area authority,
conflicting names, tricky qualifiers, no save/Saved writes, unsupported hosts,
and total fallback failure. They simulate the documented COM contract; they do
not test actual printer drivers, native Excel COM/UI, or add-in installation.

Run from the repository root:

```sh
/Users/tmsxhc/.hermes/cache/scratch/dotnet-sdk/dotnet run --project tests/HydroLivePrintArea/HydroLivePrintArea.csproj
/Users/tmsxhc/.hermes/cache/scratch/dotnet-sdk/dotnet build src/Outline.Addin/Outline.Addin.csproj --no-restore
```

Verified locally: helper tests PASS; existing HydropowerRange and
HydropowerPageSetup tests PASS; net48 add-in build succeeds with zero warnings
and errors.

Required Windows acceptance check: unsaved outer-boundary drag then refresh;
unsaved inner-divider drag then refresh (PrintArea may correctly stay unchanged);
explicit Set/Clear Print Area then refresh; initially disabled communication;
verify source file timestamp/content not saved and existing unsaved edits remain.