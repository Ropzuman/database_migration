# Changelog: 64-bit compatibility and performance fixes

## Overview

This document lists the automated and manual changes applied to the open VBA modules to make them compatible with 64-bit Office (VBA7) and to apply low-risk performance improvements and safety fixes.

## Files modified

- kanta.vba
- kanta1.vba
- kanta2.vba
- kanta3.vba

## Summary of changes

### 1) API Declare compatibility

- Added conditional compilation blocks (`#If VBA7 Then / #Else / #End If`) around Declare statements to support both 32-bit and 64-bit Office.
- Added `PtrSafe` to VBA7 branch declares where appropriate.
- Used `LongPtr` for pointer/handle fields inside UDTs in the VBA7 branch (for example `OPENFILENAME.hwndOwner`, `hInstance`, `lCustData`, `lpfnHook`).
- Kept `Long` for BOOL/DWORD return values and buffer length parameters where APIs expect 32-bit types (for example `GetUserName`, `GetComputerName`, `GetCursorPos`).
- Rationale: use `LongPtr` only for pointer-sized values (handles/callbacks). Using `Long` for BOOL/DWORD is correct and avoids incorrect sizing on 64-bit.

### 2) DAO / ADO clarification and resource cleanup

- Qualified object types explicitly where ambiguous: `Database -> DAO.Database`, `Recordset -> DAO.Recordset`.
- Ensured recordsets are closed and set to `Nothing` using a safe close pattern (`On Error Resume Next / Close / Set Nothing`) to avoid leaks.
- Replaced an incorrect `Recordset.State` check (ADO property) with a safe close pattern for DAO recordsets.

### 3) Recordset insert optimization (`SniffUser`)

- Replaced `dbOpenTable` with `dbOpenDynaset` and wrapped the insert in a transaction (`BeginTrans / CommitTrans / Rollback` on error).
- Benefits: improved portability across database engines, better behavior with multi-user environments, and faster grouped inserts.
- Note: after initial changes a compile-time error was encountered when calling `BeginTrans` on the `Database` object in some DAO versions. To fix this, the implementation was updated to call `DBEngine.BeginTrans / DBEngine.CommitTrans / DBEngine.Rollback` (see **Fixes**).

### 4) Typing, input validation, and Excel macro safety

- Explicitly typed previously untyped module-level variables (`Global last_criteria As Variant`, `last_used As Variant`).
- Provided explicit parameter and return types for helper functions (`Set_last`, `Show_last`, `Show_last_criteria`) to avoid implicit-typing bugs.
- Hardened input validation in `CustomMessage`:
  - check `IsNumeric` before numeric comparisons
  - convert to `Long` via `CLng` and re-check range
  - prevents runtime errors due to string/number comparisons
- In Excel macro `HaeData`, added a guard so `ORDER BY` is only appended if not already present, preventing duplicate clauses and SQL syntax errors.
- In Excel macro `Checkout`, all counters and ByRef arguments are now `Long` for 64-bit compatibility and to avoid ByRef type mismatches.
- Fixed duplicate variable declarations in Module1 (`HaeData`, `Checkout`) and Module2 (`VaihdaInfo`, `EtsiOts`).
- Added ODBC error handling in `HaeData` with database file existence check and detailed error messages for runtime error 1004.
- Fixed `HaeData` error handler logic so success message appears only on successful completion.
- Removed `SQLSuffix` logic from `HaeData` because it was commented out and caused issues.
- `HaeData` now skips Excel-based queries (`_qryForExcel`) to prevent ODBC errors for Excel-file-backed document property queries.
- Fixed `EtsiOts`: removed erroneous `EndFastMode2` call and added max column iteration safety check (`16384`).
- Fixed `VaihdaInfo`: removed orphaned `BeginFastMode2` call in `revid` case that had no matching `EndFastMode2`.
- **CRITICAL FIX**: Removed `VaihdaInfo("Revisions")` call from `Checkout` to prevent freeze caused by nested loops and comment processing.
- Changed numeric variables to `Long` throughout for 64-bit consistency (`RMAX`, `Valinta`, loop counters).
- Fixed `GenPrintout` subscript-out-of-range by activating macro workbook before accessing `DB1`.
- **CRITICAL PERFORMANCE FIX**: Created `PopulateRevisionsSimple` to replace heavy `VaihdaInfo("Revisions")` workflow:
  - one-pass comment-marked column detection
  - direct writes from `DIRevArr`
  - reduced complexity from roughly O(comments × revisions) to O(n)
  - safety checks (`IsArray`, bounds checks, guarded split operations)
- Kept `VaihdaInfo` optimization for Info sheet only.
- `PopulateRevisionsSimple` is called in `GenPrintout` after Revisions sheet copy.
- Revisions source: `DB2` sheet column `rev` (split by `Chr(10)`).
- Info source: `DB2` sheet document metadata.

### 5) Column width optimization in `GenPrintout`

- Fixed issue where printout contained extra columns beyond TEMPLATE definition.
- Changed `dataCols` calculation from `DB1` last column to use `Sarakkeita` (TEMPLATE column count).
- Added cleanup step to delete columns beyond `Sarakkeita` after data population.
- Result: printout width now matches TEMPLATE exactly.
- Reverted `Checkout` implementation to original logic with only 64-bit typing changes (`i`, `j`, `Apu` as `Long`).
- Fixed global types for 64-bit (`RMAX` and `Valinta` from `Integer` to `Long`).
- Fixed missing `End If` at end of `GenPrintout` (`If CheckOK = False`).
- Enhanced `Checkout` missing-sheet handling to avoid runtime error 9.
- Reverted `Checkout` to lightweight template marker lookup (`Cells.Find`) and kept error logging to ERRORS sheet.
- Removed `Cells.ClearComments` from `Checkout` because it could freeze Excel.
- Added lightweight sheet existence checks in `HaeDocTiedot`, `VaihdaInfo`, and `EtsiOts`.
- Optimized `DIRevArr` access in `VaihdaInfo` with consolidated error handling.

## Why these changes were made

- 64-bit compatibility: `PtrSafe` and `LongPtr` are required for pointer-sized values on VBA7/64-bit.
- Stability: explicit DAO qualification and safe recordset closing reduce runtime conflicts and leaks.
- Portability and performance: `dbOpenDynaset` + transactions are more robust and often faster for inserts.
- Maintainability: explicit typing reduces subtle bugs and improves auditability.

## Testing instructions (quick smoke tests)

1. Open the database in 64-bit Access (import modified modules if needed).
2. In VBA editor, run **Debug -> Compile**.
3. Run smoke tests:
   - call `SniffUser` and verify inserted `UsysUsers` row
   - call `Set_last` and `Show_last` to verify state behavior
   - run `CustomMessage` with non-numeric, cancel, and out-of-range inputs
4. Check Immediate window for errors.

## Notes and manual review items

- Manual review is required for any Declare patterns using `VarPtr`, `StrPtr`, `ObjPtr`, `CopyMemory`, or `GetWindowLong/SetWindowLong`.
- If other modules/forms exist in the database, run same compatibility scan and patch them as needed.

## Fixes

- Transaction methods corrected for DAO compatibility:
  - Problem: calling `BeginTrans/CommitTrans/Rollback` on `DAO.Database` caused compile-time errors in some Access/DAO configurations.
  - Change: replaced `db.BeginTrans/db.CommitTrans/db.Rollback` with `DBEngine.BeginTrans/DBEngine.CommitTrans/DBEngine.Rollback` in `SniffUser`.
  - Effect: compile error removed, using supported DAO transaction API.

## Files backed up

- For each file edited, a `.bak` copy was created in the same directory (for example `kanta.vba.bak`).

## Next steps

- Complete validation and testing (full checklist can be prepared).
- Optional performance work:
  - convert repetitive `DLookup` patterns into joins or preloaded dictionaries
  - reduce UI updates during bulk operations (`Application.Echo False / True`)
  - provide consolidated patch across all modules/forms

Prepared by: (automated edits)
Date: $(Get-Date)

## Access migration addendum (2026-03-13)

- Consolidated summary: `Access/ACCESS_64bit_DECLARE_MIGRATION_SUMMARY.md`
- Per-file changelogs:
  - `Access/DOCUMENTS/GlobalVBAs_changelog.md`
  - `Access/FunctionDiagrams/Form_LisääKuviin_ACAD_changelog.md`
  - `Access/Function_descriptions_html/USysCheck_changelog.md`
  - `Access/instru3/USysCheck_changelog.md`
  - `Access/MAINEQ/For ACAD Utility_changelog.md`
  - `Access/LoopCircuit/USysCheck_changelog.md`
  - `Access/PIPE/Koodit_changelog.md`
  - `Access/LoopCircuit/General_changelog.md`
  - `Access/MAINEQ/Form_Revisiointi_changelog.md`
  - `Access/LoopCircuit/Form_Tee Kuvat_changelog.md`
  - `Access/Lukituskaavio/APIKoodit_changelog.md`
  - `Access/MAINEQ/USysCheck_changelog.md`
  - `Access/LoopCircuit/For ACAD Utility_changelog.md`
