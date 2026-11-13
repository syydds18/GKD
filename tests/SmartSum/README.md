# SmartSum pure-core tests

Run from the repository root (requires .NET SDK 8):

```sh
dotnet run --project tests/SmartSum/SmartSum.csproj -c Release
dotnet build tests/SmartSum/Compatibility/Compatibility.csproj -c Release
```

The console tests reference an isolated **netstandard2.0 / C# 7.3** build of the actual `src/Outline.Core/*.cs` sources. No UI, COM, Excel process, third-party test packages or workbook writes. All generated build outputs remain in `tests/SmartSum/`; production Core project/output files need not be touched.

To run one case, append `-- case_name`. Nonzero exit indicates failure (including an unknown case name).

## Contract

`SmartSumPlanner.Build(sheet, detection, columns, useExistingOutline: false)` returns a preview only. `SumPlan.Proposals/Warnings` and `SumProposal.Row/Col/Level/Address/RowName/Formula/OriginalText/ChildRows/HasExistingFormula` are public properties.

- Default hierarchy is Detection.Levels; outline is an explicit alternative, never an automatic fallback.
- Caller must preflight **summary rows above details** before selecting existing outline. SheetData has no SummaryRow metadata; core cannot verify the Excel setting.
- Only recognized codes with nonempty names are structural. A name beginning with explicit `第一部分`-style part numbering is an explicit banner edge, not an uncoded total. Detection's family and dotted relative levels are preserved below that edge.
- Ambiguous rows, missing levels, duplicate sibling codes, dotted prefix mismatches and level gaps block the affected ancestor scope; subsequent nodes in that scope are also conservative. This may intentionally generate fewer formulas than a human could infer.
- Direct children only, sorted/deduplicated; same-column downward references. Contiguous direct children become ranges; grandchildren are never folded into a parent's range.
- At least two direct children. Empty leaves are valid references. A blank child group without an existing formula must have a valid proposal before its parent can be proposed. Proposals are bottom-up. UI must enforce selected blank-group proposal dependencies before applying.
- Columns are checked independently. Error/non-numeric/Boolean parent values are rejected. Numeric text is permitted as a replaceable parent target, but rejected as a direct child reference because SUM ignores cell text. Existing parent formulas are offered for explicit selection; original formula/text is retained.
- Invalid Excel columns/dimensions throw `ArgumentOutOfRangeException`; null arguments throw `ArgumentNullException`. Missing hierarchy/header or empty selection produce warnings and no proposals.
- More than 255 noncontiguous SUM argument segments are skipped with warning (no chunking). At 255 segments, maximum A1-range formula length is 5615 characters, below 8192. More than 5000 proposals throws `InvalidOperationException`; no partial plan is returned.

## TDD evidence

Each new implementation behavior below was run as a failing console assertion first, then the full suite was run green after the corresponding minimal implementation. These are actual terminal outcomes, not a fabricated test log:

| Case | Observed RED assertion |
|---|---|
| direct_children_three_levels | SmartSumPlanner API missing |
| multiple_columns_existing_formula | existing formula should be offered for explicit preview selection |
| single_child_warning | single child warning missing |
| ambiguous_rows_block_ancestors | must not cross ambiguous row 小计 |
| typed_amount_safety | unsafe child must block blank ancestors: 12 |
| explicit_outline_source | explicit flat outline must not fall back to detected |
| input_and_excel_bounds | Expected ArgumentOutOfRangeException: illegal column 0 |
| missing_or_invalid_levels_are_barriers | invalid level must not bridge |
| part_banner_is_structural | Sequence contains no matching element |
| sum_argument_limit_255 | 256 segments must skip with warning |
| candidate_cap_is_atomic | Expected InvalidOperationException: candidate overflow must fail whole plan, not truncate |
| dotted_prefix_boundary | must not attach 1.3 to group 2 |
| duplicate_sibling_codes_are_ambiguous | duplicate numbered siblings must not be guessed |
| hierarchy_gaps_are_ambiguous | missing intermediate level must block root |
| ambiguous_scope_does_not_spawn_new_groups | no new group proposals inside ambiguous range |
| nonempty_unselected_cells_are_not_blank_rows | nonempty note-only row is not a blank gap |
| preview_preserves_original_text | preview must retain original cell text exactly |

Three additional regression/boundary tests exercise already-green behavior without requiring implementation changes: blank_leaf_and_group_dependencies, bottom_up_order_and_formula_limits, maximal_range_formula_stays_below_8192.

Final verified result: **20 tests, 20 passed, 0 failed**; netstandard2.0 build: **0 warnings, 0 errors**.

## Not verified here

Excel formula recalculation, real workbook layout/merged-cell extraction, live COM error marshaling, native preview selections/dependency checks, applying formulas, rollback, and Windows installation. Core does not evaluate existing formulas or audit their dependencies. This harness tests planning only.
