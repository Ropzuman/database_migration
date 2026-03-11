# AcadDATA Cleanup – 2025-10-30

Scope: Non-functional cleanup and micro-optimizations in the Excel AcadDATA integration.

Changes

- Refactored DXF entity-type filter construction into helper `BuildTypeFilter(includeTexts, FilterType(), FilterData())` to remove duplicated code and avoid repeated ReDim Preserve calls.
- Removed unused locals and clarified comments; no logic changes.
- Preserved selection behavior, dynamic-block handling (EffectiveName pruning), coordinate extraction, and Excel calculation-mode handling.

Verification

- Manual smoke run paths unchanged: import all, import with a single static name, import with a dynamic block name, and previous-selection mode.
- Checked that selection counts and written columns match prior behavior.

Notes

- Helper is late-binding friendly and only constructs common type filters (`INSERT`, optional `TEXT`/`MTEXT`).
- Any future extension (e.g., adding LINE, CIRCLE) can leverage the same helper or an overload.
