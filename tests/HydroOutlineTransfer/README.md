# HydroOutlineTransfer portable tests

Run from any directory:

```sh
/Users/tmsxhc/.hermes/cache/scratch/dotnet-sdk/dotnet run --project "/Users/tmsxhc/Desktop/Excel分级校验加载项/tests/HydroOutlineTransfer/HydroOutlineTransfer.csproj" -c Release
```

Links the real C# 7.3 production helper into a .NET 8 console executable; no test framework or Excel installation required. Fakes expose the Excel dynamic range members used by the helper and model mixed properties as null/DBNull, Group() depth increments and ShowDetail visibility side effects.

Integration: `HydroOutlineTransfer.Copy(sourceRange, destinationSheet)` (`internal static void`, dynamic arguments). Pass a single rectangular source Range and a **different, fresh Worksheet**, not a destination Range. Source origin maps to destination A1. Call after value/format copy and all AutoFit operations. The helper does not copy values or cell formatting, clear an existing outline, change global Excel settings, or write anything on the source. Invalid extents, source levels and preexisting destination outline in the mapped extent are rejected before destination writes.

Coverage includes clipped nested groups in both axes, exact hidden segments, all levels 1–8, outer-to-inner Group order, both summary directions, best-effort ShowDetail, unavailable/clipped summaries, source immutability, Excel extent limits, multi-area/same-sheet rejection, fresh-target validation, full-sheet uniform batching and varied singleton fallback. Uniform axes use one source range read per property; two axes means two OutlineLevel reads and two Hidden reads. Reads split only mixed segments, not cells. Individual summary reads are once per distinct in-clip group boundary.

Limitations: these are portable dynamic-fake tests, not Windows Excel COM tests. Native Excel outline button rendering, save/reopen persistence, filter metadata and adjacent groups with indistinguishable levels require Excel integration testing. Summary ShowDetail is best effort (only valid in-clip source summaries are used; COM failures are ignored for this optional property); exact final hidden visibility wins over ShowDetail side effects. Clipping can remove summaries, and levels alone cannot recover distinct adjacent groups lacking a lower-level separator. Operational Group/Hidden COM errors propagate, so callers should discard a partially constructed export sheet on failure.
