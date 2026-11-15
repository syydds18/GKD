# GKD selection action verification

## Portable behavioral tests

```sh
dotnet run --project tests/GkdAiActions/GkdAiActions.csproj -c Release
dotnet build tests/GkdAiActions/Compatibility/Compatibility.csproj -c Release
```

The executable links the actual production `GkdAiActions.cs` and uses dynamic fake Excel ranges. It exercises complete direct GKD formula recognition; spill-child anchoring and duplicate selection; per-cell revisions and native Dirty/Calculate dispatch; cancel without selection; selection and aggregate spill limits; full-spill confirmation and decline; ignoring input references/unrelated formulas; refusal of errors, reserved failure/cancel markers, protected sheets, merged ranges and legacy arrays; restoring events after paste failure; and rechecking every target after modal confirmation. The compatibility project compiles the production source against real Excel-DNA 1.9.0 and net48 references (only the integration scheduler/UDF contracts are stubs).

## Windows Excel checks still required

- Load the XLL and create a scalar `=GKD(...)`, a table-mode GKD spill, `GKD.PROMPT`, an unrelated formula, and input cells containing formulas/data. Select a spill child and disjoint GKD anchors; refresh must touch each GKD anchor once and leave all formula text unchanged.
- Refresh a still-running request; verify the earlier reply cannot replace the refreshed revision. Cancel queued/running requests; verify no automatic rerun until explicitly refreshed.
- Select one spill child and use Results to Values. Confirm the dialog lists the complete spill address/count. Decline once, then accept; verify the full spill becomes constants and input references/unrelated cells remain unchanged. Include strings starting with `=` and mixed scalar result types to check native clipboard semantics.
- Verify `#N/A`, other Excel errors (including a spill obstruction), reserved `GKD：` failure/cancel messages, protected sheets, merged ranges, legacy arrays, and a full-spill aggregate over 10000 cells all refuse before a write.
- Test multiple selected spills and both original EnableEvents states. Native Copy/PasteSpecial may fail in an Excel version or clipboard environment; the action restores events and stops on that error, but cannot roll back already converted earlier ranges. The macro does not create an undo transaction.

Native Excel COM is not executable by these portable tests. In particular, a fake PasteSpecial call is not evidence that a real Excel build supports same-range values paste over a dynamic spill.
