# Database Migration: Access/Excel/AutoCAD (32-bit → 64-bit Office)

Purpose
-------
This project migrates an MS Access–based, database-driven design system from 32-bit Office to 64-bit Office (VBA7). The system integrates AutoCAD, Excel, and Access via macros and scripts. The goal is to update declarations, types, and data-access patterns so the same code runs under 64-bit Office as well as update the codebase to run more efficiently through optimization and modernization.

Scope and components
--------------------
- Access VBA modules: `Access/kanta.vba`, `Access/kanta1.vba`, `Access/kanta2.vba`, `Access/kanta3.vba`
- Excel artifacts: `Excel/` (supporting spreadsheets and macros)
- AutoCAD VBA macros:
- Changelog: `CHANGELOG_64bit_and_perf.md` (documenting compatibility and performance fixes)

Key migration changes (from the changelog)
-----------------------------------------
1) Win32 API declares updated for VBA7/64-bit
	- Conditional compilation blocks: `#If VBA7 Then ... #Else ... #End If`
	- `PtrSafe` applied to declares in the VBA7 branch.
	- `LongPtr` used for pointer/handle-sized fields and callbacks in UDTs (e.g., `OPENFILENAME.hwndOwner`, `hInstance`, `lCustData`, `lpfnHook`).
	- `Long` retained for BOOL/DWORD parameters and returns where the underlying API expects 32-bit values (e.g., `GetUserName`, `GetComputerName`, `GetCursorPos`).

2) DAO object model clarity and safe cleanup
	- Explicit DAO qualifications: `DAO.Database`, `DAO.Recordset` to avoid ambiguity.
	- Safe close pattern for recordsets: `On Error Resume Next` → `Close` → `Set Nothing`.
	- ADO-only properties (like `Recordset.State`) removed; DAO-safe patterns used instead.

3) Insert behavior and transactions (example: `SniffUser`)
	- Use `dbOpenDynaset` instead of `dbOpenTable` for portability and multi-user robustness.
	- Wrap inserts in DAO transactions via `DBEngine.BeginTrans / CommitTrans / Rollback` (fixes compile issues with calling `BeginTrans` on `DAO.Database`).

4) Typing and input validation improvements
	- Explicit module-level typing for globals (e.g., `last_criteria As Variant`, `last_used As Variant`).
	- Helper functions (`Set_last`, `Show_last`, `Show_last_criteria`) given explicit param/return types.
	- Hardened input validation in `CustomMessage`: numeric checks (`IsNumeric`), conversion via `CLng`, and range guards to prevent runtime errors.

Quick validation (smoke tests)
------------------------------
1. Open the database in 64-bit Access. Import or confirm the updated modules.
2. In VBA editor: Debug → Compile to verify no compile-time errors.
3. Run these:
	- `SniffUser`: Confirm a new record in `UsysUsers` with username, Access user, computer name, and timestamp.
	- `Set_last` / `Show_last` / `Show_last_criteria`: Verify values are stored and returned.
	- `CustomMessage`: Try non-numeric input, cancel, and out-of-range values to verify validation and messaging.
4. Watch the Immediate window for errors or warnings.

Usage notes
-----------
- Ensure Office references match your installation (32-bit vs 64-bit) and that macros are enabled.
- AutoCAD integration: If macros/scripts interact with AutoCAD, confirm any declares or COM references are compatible with 64-bit AutoCAD; apply the same conditional compilation and `PtrSafe/LongPtr` rules.
- Excel macros: Validate that any API calls and file dialogs compile and run under 64-bit Excel.

Repository contents
-------------------
- `Access/` — core VBA modules; primary entry is `kanta3.vba`.
- `Excel/` — supporting Excel files used with the VBA code.
- `CHANGELOG_64bit_and_perf.md` — 64-bit compatibility and performance notes.
- `README.md` — project overview and validation steps.

Contributing and next steps
---------------------------
- Add a LICENSE and CONTRIBUTING guide to streamline collaboration.
- If other modules/forms exist, run the same compatibility pass and include them here; open an issue for assistance.
- Optional performance work: reduce UI updates during bulk operations, optimize repeated lookups, and consolidate database interactions.

Contact
-------
For help migrating, testing, or documenting additional modules (including AutoCAD/Excel integrations), open an issue or contact the repository owner.
