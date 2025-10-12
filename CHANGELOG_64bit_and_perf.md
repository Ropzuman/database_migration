Changelog: 64-bit compatibility and performance fixes

Overview
--------
This document lists the automated and manual changes applied to the open VBA modules to make them compatible with 64-bit Office (VBA7) and to apply low-risk performance improvements and safety fixes.

Files modified
--------------
- kanta.vba
- kanta1.vba
- kanta2.vba
- kanta3.vba

Summary of changes
------------------
1) API Declare compatibility
   - Added conditional compilation blocks (#If VBA7 Then / #Else / #End If) around Declare statements to support both 32-bit and 64-bit Office.
   - Added PtrSafe to VBA7 branch declares where appropriate.
   - Used LongPtr for pointer/handle fields inside UDTs in the VBA7 branch (e.g., OPENFILENAME.hwndOwner, hInstance, lCustData, lpfnHook).
   - Kept Long for BOOL/DWORD return values and buffer length parameters where APIs expect 32-bit types (e.g., GetUserName, GetComputerName, GetCursorPos).
   - Rationale: Use LongPtr only for pointer-sized values (handles/callbacks). Using Long for BOOL/DWORD is correct and avoids incorrect sizing on 64-bit.

2) DAO / ADO clarification and resource cleanup
   - Qualified object types explicitly where ambiguous: Database -> DAO.Database, Recordset -> DAO.Recordset.
   - Ensured recordsets are closed and set to Nothing using a safe close pattern (On Error Resume Next / Close / Set Nothing) to avoid leaks.
   - Replaced an incorrect Recordset.State check (ADO property) with a safe Close pattern for DAO recordsets.

3) Recordset insert optimization (SniffUser)
   - Replaced dbOpenTable with dbOpenDynaset and wrapped the insert in a transaction:
     - BeginTrans / CommitTrans / Rollback on error.
   - Benefits: Improved portability across database engines, better behavior with multi-user environments, and faster grouped inserts.
   - Note: After initial changes a compile-time error was encountered when calling BeginTrans on the Database object in some DAO versions. To fix this, the implementation was updated to call DBEngine.BeginTrans / DBEngine.CommitTrans / DBEngine.Rollback (see "Fixes" below).


4) Typing, input validation, and Excel macro safety
    - Explicitly typed previously untyped module-level variables (Global last_criteria As Variant, last_used As Variant).
    - Provided explicit parameter and return types for small helper functions (Set_last, Show_last, Show_last_criteria) for clarity and to avoid implicit-typing bugs.
    - Hardened input validation in CustomMessage:
       - Check IsNumeric before numeric comparisons.
       - Convert to Long via CLng and re-check range.
       - This prevents runtime errors due to string/number comparisons and improves user input safety.
    - In Excel macro `HaeData`, added a guard so `ORDER BY` is only appended if not already present in the SQL, preventing duplicate clauses and SQL syntax errors.
    - In Excel macro `Checkout`, all counters and ByRef arguments are now `Long` for 64-bit compatibility and to avoid VBA ByRef argument type mismatches.
    - Fixed duplicate variable declarations in Module1 (HaeData: changed `i` from Integer to Long; Checkout: removed duplicate `foundRange` declaration) and Module2 (VaihdaInfo: removed duplicate `i` and `Row` declarations; EtsiOts: moved Dim statements to function top).
    - Added ODBC error handling in HaeData with database file existence check and detailed error messages to diagnose runtime error 1004.
    - Fixed HaeData error handler logic: moved success message outside error handler so it only shows on successful completion, not after errors.
    - Removed SQLSuffix logic from HaeData: was commented out in original code and caused issues. Now uses ORDER BY only if specified in faceplate query.
    - HaeData now skips Excel-based queries (_qryForExcel): prevents ODBC errors when document property queries reference Excel files instead of Access database.
    - Fixed EtsiOts function: removed erroneous EndFastMode2 call that was causing Excel to freeze during Checkout validation.
    - Fixed VaihdaInfo function: removed orphaned BeginFastMode2 call in "revid" case that had no matching EndFastMode2, causing Excel to freeze when Checkout validated Revisions sheet.
    - Reverted Checkout function to original implementation: restored Cells.ClearComments, added missing Sheets("TEMPLATE").Select call, and fixed function parameter to numeric 1 instead of 1& (Long literal). Kept 64-bit compatibility (Long instead of Integer) but otherwise matches original function that worked reliably.
    - Enhanced Checkout function to handle missing sheets gracefully: VaihdaInfo now checks if sheet exists before accessing it (fixes runtime error 9 "subscript out of range").
    - Reverted Checkout to lightweight template marker lookup: removed heavy error checking with MsgBox for each marker. Now uses direct Cells.Find() calls like original code - errors logged to ERRORS sheet as designed.
    - Removed Cells.ClearComments from Checkout: this operation on entire sheet was causing Excel to freeze. Comments are now only cleared when needed in GenPrintout.
    - Added lightweight sheet existence validation in HaeDocTiedot, VaihdaInfo, and EtsiOts: silently exits if required sheets are missing instead of showing multiple message boxes.
    - Optimized DIRevArr array access in VaihdaInfo: consolidated error handling with single On Error Resume Next block per case instead of nested checks, improving performance.

Why these changes were made
--------------------------
- 64-bit compatibility: PtrSafe and LongPtr are required for pointer-sized values on VBA7/64-bit. Conditional compilation ensures the code still runs under 32-bit VBA.
- Stability: Explicit DAO qualifications and recordset closing avoid runtime conflicts and resource leaks.
- Portability and performance: Using dbOpenDynaset and transactions is more robust than dbOpenTable and can improve performance for inserts, especially in multi-user environments.
- Maintainability: Explicit typing reduces subtle bugs and makes the code easier to understand and audit.

Testing instructions (quick smoke tests)
--------------------------------------
1) Open the database in 64-bit Access. Import the modified modules if necessary.
2) Open the VBA editor and choose Debug -> Compile to catch compile-time errors.
3) Run these smoke tests:
   - Call SniffUser and verify that a new record appears in UsysUsers with the expected fields.
   - Call Set_last and Show_last functions to confirm they preserve and return values.
   - Run CustomMessage and attempt entering non-numeric input, Cancel, and out-of-range numbers to verify validation behavior.
4) Check the Immediate window for errors during tests.

Notes and manual review items
-----------------------------
- Manual review is required for any Declare that uses VarPtr, StrPtr, ObjPtr, CopyMemory, or GetWindowLong/SetWindowLong patterns; none were found in these open modules.
- If other modules/forms exist in the database, run the same compatibility scan on them (see script in earlier messages) and send them if you want me to patch them too.

Fixes
-----
- Transaction methods corrected for DAO compatibility:
   - Problem: Calling BeginTrans/CommitTrans/Rollback on a DAO.Database object caused a compile-time error in some Access/DAO configurations: "Function or interface marked as restricted, or the function uses an Automation type not supported in Visual Basic." This occurs because the Database object does not expose those transaction methods in all versions; the DBEngine object provides the supported transaction API.
   - Change: Replaced uses of db.BeginTrans/db.CommitTrans/db.Rollback with DBEngine.BeginTrans/DBEngine.CommitTrans/DBEngine.Rollback in `SniffUser`.
   - Effect: Fixes the compile error and uses the supported DAO transaction API.

Files backed up
---------------
- For each file edited, a .bak copy was created in the same directory (e.g., kanta.vba.bak).

Next steps
----------
- Complete Validation and testing (I can produce a full test checklist and help iterate on any compile/runtime errors you find).
- If you want further performance improvements, I can:
  - Convert repetitive DLookup patterns (if present in other modules) into joins or preloaded dictionaries.
  - Reduce UI updates during bulk operations (Application.Echo False / True), and remove redundant SetFocus calls.
  - Provide a consolidated patch across all modules/forms in the database.

Prepared by: (automated edits)
Date: $(Get-Date)
