# AcadDATA — Kehittäjämuistiinpanot / Developer Notes

> **Luokitus / Classification:** `[ACTIVE]` — Tekninen viitedokumentti kehittäjille
> **Päivitetty / Updated:** 8.11.2025 (64-bit refaktoroinnin jälkeen / post-64-bit refactoring)
> **Kohderyhmä / Audience:** Kehittäjät / Developers only
>
> *Tämä dokumentti on tekninen kehittäjäopas. Jos haluat vain käyttää AcadDATA-työkalua, katso pääohje: `README.md`.*
> *This is a technical developer reference. For end-user instructions, see `README.md`.*

## 2025-10-30 – Non-functional cleanup

- Introduced a small helper `BuildTypeFilter(includeTexts, FilterType(), FilterData())` to construct the DXF entity-type filter arrays for selection (`INSERT` and optional `TEXT`/`MTEXT`). This removes duplicated code blocks and avoids repeated `ReDim Preserve` operations by sizing arrays exactly.
- Removed a couple of unused locals and improved inline comments. No functional changes.
- Behavior remains the same:
  - Selection still enforces Model Space and uses type filters first.
  - When specific names are provided, effective-name pruning remains in place to include dynamic blocks.
  - Coordinate extraction and Excel calculation-mode handling are unchanged.

Rationale: Keeping the selection filter construction DRY improves readability and slightly reduces overhead while keeping the late-binding behavior stable.

This document explains the AutoCAD import/export implementation in Excel (AcadDATA), focusing on the selection pipeline, dynamic blocks, DXF filter usage, tracing, performance, and troubleshooting.

## Selection pipeline overview

1. Build DXF filters by entity type for performance
   - Always include `INSERT`
   - Optionally include `TEXT` and `MTEXT` when user selects "Blokit ja tekstit"
2. If specific block names are provided (not just `*`)
   - Add a code-2 OR-group with those names
3. Execute selection (previous selection or all)
4. If name-filtered selection is empty (common with dynamic blocks)
   - Reselect using only entity-type filters
   - Prune in VBA by `EffectiveName` to keep only requested blocks
5. Extra safety: When names are specified, remove non-matching BlockReferences from the selection set before processing (pre-filter removal)

Rationale: Dynamic blocks have anonymous DXF names (e.g., `*U###`), so a strict code-2 filter may match nothing. Falling back to type-only selection and pruning by `EffectiveName` ensures correct results without pulling unrelated entities.

## Dynamic blocks: Name vs EffectiveName

- `Name`/DXF code 2 may be anonymous for dynamic blocks, typically `*U...`
- `EffectiveName` resolves the base block name used for matching
- The import code compares requested names against `EffectiveName` to include the right dynamic blocks

## DXF filter usage details

- Selection API: `SelectionSet.Select mode, pt1, pt2, FilterType(), FilterData()`
- CRITICAL: `FilterType` must be an Integer array
  - Passing `Long()` causes late-bound COM to error: "Invalid argument FilterType in Select"
- Code groups used:
  - 0 → Entity type (e.g., `INSERT`, `TEXT`, `MTEXT`)
  - 2 → Block name (for OR-group of requested names)
  - -4 → Boolean operators: "<or" and "or>" to delimit OR-groups
- Layer filtering was intentionally removed (simplifies logic and avoids over-restriction)

Example (conceptual):

```text
Filter: <or 0=INSERT 0=TEXT 0=MTEXT or>
If names: <or 2=NAME_A 2=NAME_B or>
```

## Pre-filter removal and pruning

- After selection, when specific names were requested, we build a small removal list of items whose `EffectiveName` does not match any requested name and call `SelectionSet.RemoveItems` once.
- `RemoveItems` expects a Variant array of COM objects; populate as `Variant` and assign objects directly.

## Tracing (developer mode)

- Controlled in `Excel/Moduulit/AcadDATA/Koodit.bas` by `Public Const DEBUG_TRACE As Boolean`
- When `True`, logs to the Immediate Window (Ctrl+G):
  - Step breadcrumbs (sheet select, parsing names)
  - Filter construction decisions (name OR-group, fallback)
  - Selection counts before and after fallback
  - Per-document rows written and grand total
- No functional changes—only verbosity

## Late binding and robustness

- All AutoCAD COM references are late-bound (`As Object`), enabling compatibility without setting references
- Prefer `TypeName` and then fallback to `EntityName`/`ObjectName` via `CallByName` to identify entity type safely
- Access potentially missing members under `On Error Resume Next` with immediate error clearing; then restore `On Error GoTo ErrHandler`

## Buffered writes for performance

- Rows are buffered into a 2D Variant array and flushed to the sheet in a single assignment per drawing
- Column capacity expands when new attribute tags appear
- Adds cell comments to headers with the source block name (optional)

## Zoom helpers (navigation)

- Double-click uses `ZoomWindow` to the entity bounding box, then a safe, relative zoom helper to avoid enum mismatches in late binding
- The helper tries enum values (1 then 3) and swallows benign errors to keep UX smooth

## Error handling conventions

- `StepMsg` tracks the step for error dialogs
- All exit paths go through a `Cleanup` label that restores Excel/AutoCAD state and releases COM objects
- When nothing matches, an informational message is shown

## Testing checklist

- Import with a single static block name → verify count and attribute mapping
- Import with a dynamic block name → verify fallback triggers and results are correct
- Import with `*` → include all INSERT (and TEXT/MTEXT if chosen)
- Export changed attributes → verify values update in DWG by handle
- Double-click navigation → verify zoom and activation

## Troubleshooting

- "Invalid argument FilterType in Select" → ensure `FilterType()` is `Integer()`, not `Long()`
- "No blocks found" with dynamic blocks → confirm fallback path executes and pruning by `EffectiveName` is active
- Excess selection counts → verify no wildcard like `*U*` is included in name filters; rely on fallback + pruning
- Zoom errors under late binding → use `SafeZoomScaled` helper instead of direct enums

## Implementation locations

- Import logic: `Excel/Moduulit/AcadDATA/Koodit.bas`
- Worksheet navigation: `Excel/Moduulit/AcadDATA/DATA.bas`
- Zoom helper: `Excel/Moduulit/AcadDATA/AcadHelpers.bas`
