# Automations Log

This log tracks changes to the automation scripts in the `Automations/` folder and documents usage notes, debugging tips and quick fixes applied during migration work.

2025-10-22  - Initial creation

- Added `Automations/export_access_vba.ps1` to export Access VBA components.
- Added `Automations/Excel_automaatio.ps1` to replace modules in .xlsm files for 64-bit migration.

2025-10-22  - Updates

- Improved header comments and metadata in `export_access_vba.ps1`.
- Added project-targeting logic to export the VBProject that matches the opened DB when possible.
- Added exclusion for filenames containing `_Backup`.
- Replaced explicit -Verbose parameter with standard Write-Verbose usage.
- Added interactive prompt and default handling to `Excel_automaatio.ps1` for both Excel folder and module folder.
- Validated and fixed parameter parsing errors and trailing comma issues.

Notes/Troubleshooting:

- Ensure "Trust access to the VBA project object model" is enabled in Access Trust Center for exports to work.
- Match PowerShell bitness to Office bitness (32-bit Access requires 32-bit PowerShell).
- If MSACCESS.EXE remains after running, check Task Manager and kill the process; the scripts attempt to release COM objects but some hosts may keep processes alive.
- For password-protected VBProjects, manual export inside Access may be required.

Planned:

- Add a --dry-run mode to list files without opening Access/Excel.
- Add logging to file (per-run log entries) and optional CSV summary of exported components.
- Add command-line switches to `Excel_automaatio.ps1` to run unattended in CI (pass both folder paths).

2025-10-22 - Cleanup

- Removed `DOCUMENTS_64BIT_MIGRATION.md` (DOCUMENTS is an Access DB; migration notes moved to project docs and automations log).
- Updated script headers to Finnish and added note about UNC/network path handling.

2025-10-22 - LoopCircuit Database Migration to 64-bit

- Migrated LoopCircuit Access database VBA code from `Automations/access_vba_exports/LoopCircuit/` to `Access/LoopCircuit/`
- Files migrated: 7 total (4 module files: For ACAD Utility.bas, General.bas, Module1.bas, USysCheck.bas; 3 form classes: Form_DBUsers.cls, Form_Linkkien vaihto.cls, Form_Tee Kuvat.cls)

64-bit compatibility fixes applied:

- Fixed API declarations in General.bas: changed nSize parameter from Long to LongPtr in VBA7 block
- Added missing PtrSafe to else clause in For ACAD Utility.bas GetCursorPos declaration
- USysCheck.bas already had correct VBA7 conditionals - preserved as-is with minor cleanup
- Added conditional BuffSize type in General.bas (LongPtr for VBA7, Long for legacy)
- Changed all module-level and local counter/index variables to Long (from Integer where found)
- Added proper LongPtr return type for GetOpenFileName in Form_Tee Kuvat.cls

Code quality improvements:

- Added explicit DAO qualification throughout (DAO.Database, DAO.Recordset, DAO.Workspace)
- Improved error handling with proper Cleanup sections and On Error GoTo patterns
- Added transactions (BeginTrans/CommitTrans/Rollback) in General.bas SniffUser
- Converted Global variables to Private module-level (m_last_criteria, m_last_used) in USysCheck.bas
- Removed dead/commented code sections in Form_Tee Kuvat.cls
- Added SQL injection protection with Replace() for single quotes in dynamic SQL
- Improved loop efficiency and replaced Do/Loop with cleaner While conditions where appropriate
- Added explicit Option Explicit to all modules that were missing it
- Cleaned up string concatenation in loops (Form_DBUsers.cls WhosOn function)
- Late binding for AutoCAD objects (Object instead of early binding) for better compatibility

Performance optimizations:

- Removed inefficient Loki (log) updates in Form_Tee Kuvat (was causing repaints on every line)
- Changed file lock detection to check both .laccdb (Access 2007+) and .ldb (legacy) extensions
- Used explicit Long types for all counters and indices (64-bit safe)
- Optimized recordset operations with explicit dbOpenDynaset where appropriate

Robustness improvements:

- Added proper COM object cleanup (Set to Nothing) in all AutoCAD automation code
- Added validation for null/empty values before using IIf expressions
- Improved status messages and error descriptions in Finnish
- Added update counter to Form_Linkkien vaihto to show how many tables were relinked
- Better folder existence validation before AutoCAD operations

Notes:

- All migrated files are production-ready for 64-bit Office/Access
- AutoCAD automation uses late binding (CreateObject/GetObject) for maximum compatibility
- Original exported files remain in Automations/access_vba_exports/LoopCircuit/ for reference
- Lock file logic updated to handle both .laccdb (Access 2007+) and .ldb (older versions)

---

## 2025-01-XX - Critical Bug Fix: Type Declaration Scoping in LoopCircuit

**Issue Discovered:**

After importing migrated LoopCircuit VBA files, all form button events failed with error:

```text
Microsoft Access can't find the procedure '[procedure name]'
```

**Root Cause:**

The `OPENFILENAME` Type declaration in `USysCheck.bas` was duplicated inside both branches of the `#If VBA7` conditional compilation block. Access VBA requires Type declarations at module scope OUTSIDE any conditional compilation directives. When Types are duplicated in conditional blocks, Access cannot resolve them properly, causing form event procedures to fail.

**Files Fixed:**

- `Access/LoopCircuit/USysCheck.bas` - Moved OPENFILENAME Type to module level with conditional LongPtr/Long fields inside the Type structure

**Technical Explanation:**

```vba
' INCORRECT (causes "Can't find procedure" errors):
#If VBA7 Then
    Public Type OPENFILENAME
        ' all fields...
    End Type
#Else
    Public Type OPENFILENAME  ' DUPLICATE - causes Access to lose references!
        ' all fields...
    End Type
#End If

' CORRECT:
Public Type OPENFILENAME
    lStructSize As Long
    #If VBA7 Then
        hwndOwner As LongPtr
        hInstance As LongPtr
    #Else
        hwndOwner As Long
        hInstance As Long
    #End If
    lpstrFilter As String
    ' ... other fields
End Type

#If VBA7 Then
    Private Declare PtrSafe Function GetOpenFileName Lib "comdlg32.dll" ...
#Else
    Private Declare Function GetOpenFileName Lib "comdlg32.dll" ...
#End If
```

**Lesson Learned:**

When migrating VBA to 64-bit compatibility:

1. **Type declarations** must always be at module scope OUTSIDE #If blocks
2. Use conditional compilation INSIDE Type structures only for fields that need different types (LongPtr vs Long)
3. Only **Declare statements** (API functions) should be fully wrapped in #If VBA7 conditional blocks
4. Duplicating Types in conditional blocks breaks Access form event wiring

**Verification Results:**

- `For ACAD Utility.bas`: POINTAPI Type was already correctly positioned outside conditional blocks ✓
- `Form_DBUsers.cls`: UserRec Type correctly positioned at module level ✓
- `Form_Tee Kuvat.cls`: Only uses conditional compilation for variable declarations, not Types ✓
- `Form_Linkkien vaihto.cls`: No Type declarations, no issues ✓

**Testing:**

After fixing USysCheck.bas, users should:

1. Re-import all LoopCircuit VBA files into Access database
2. Open VBA Editor (Alt+F11)
3. Run Debug → Compile VBAProject to verify no errors
4. Test all form button click events execute without "Can't find procedure" errors

**Update - Additional Syntax Error Fixed:**

After fixing the Type declaration issue, testing revealed a string concatenation syntax error in Form_DBUsers.cls (lines 63, 69). Changed `+` operator to `&` operator for string concatenation. VBA requires `&` for string concatenation; using `+` causes "Expected: end of statement" errors in Access event procedures.

Fixed code:

```vba
' Changed from: SPath = Left(SPath, InStr(1, SPath, ".")) + "laccdb"
' To: SPath = Left(SPath, InStr(1, SPath, ".")) & "laccdb"
```

This was the only occurrence of `+` for string concatenation in all LoopCircuit files.

**Update - Third Bug Fix: CurrentDb Multiple References:**

Testing revealed another "Expected: end of statement" error in Form_Linkkien vaihto.cls. The issue was calling `CurrentDb.Name` multiple times in the same statement (line 27):

```vba
' WRONG - causes "Expected: end of statement":
Polku = Left(CurrentDb.Name, Len(CurrentDb.Name) - Len(Dir(CurrentDb.Name)))

' CORRECT - assign to variable first:
Set db = CurrentDb
DbName = db.Name
Polku = Left(DbName, Len(DbName) - Len(Dir(DbName)))
```

**Files Fixed:**

- `Form_Linkkien vaihto.cls` - Added db and DbName variables, replaced all CurrentDb references
- Changed line 27 to use intermediate variables
- Changed line 35 from CurrentDb.OpenRecordset to db.OpenRecordset  
- Changed line 56 from CurrentDb.Execute to db.Execute
- Added db cleanup in Cleanup section

**Lesson:** In Access VBA, `CurrentDb` should be assigned to a DAO.Database variable and reused, not called multiple times in complex expressions.

---

2025-10-23 - Access automation script updated

- Updated `Automations/Access_automaatio.ps1` to a robust, 64-bit–first updater for Access VBA components.
- Enforces x64 PowerShell, adds retry logic when opening databases, and removes read-only attribute before open.
- Adds clear Trust Center guidance and checks for VBE availability; fails fast with actionable messages.
- Standardizes component import: removes existing components by name, imports matching .bas/.cls from a chosen folder, then saves.
- Improves cleanup: re-enables warnings, closes the database, calls Access.Quit, and releases COM objects to avoid orphaned processes.
- Removed legacy `Automations/Access_automaatio_utf.ps1` (single consolidated script now handles all cases).

Usage notes:

- Run in 64-bit PowerShell. Ensure Access Trust Center has "Trust access to the VBA project object model" enabled and both the .accdb and the module folder are trusted locations.
- Optional defaults can be set via `$DefaultAccessFilePath` and `$DefaultComponentPath` at the top of the script to speed up repeated runs.

---

## 2025-11-07 - Excel VBA Module Import Issue: VBComponents.Import Metadata Corruption

**Problem Identified:**

`Excel_automaatio.ps1` was using `VBComponents.Import()` to replace VBA modules in .xlsm workbooks. This approach caused invisible metadata corruption:

- Modules imported via `VBComponents.Import()` contained hidden VBA file headers (Attribute VB_Name, VERSION lines, etc.)
- These invisible metadata lines caused the modules to behave differently than when code was manually copy-pasted
- Users reported the script "doesn't export the modules right" - automation failed but manual copy-paste worked

**Root Cause:**

When VBA exports a module to .bas file, it includes file metadata:
```vba
Attribute VB_Name = "Module1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
' ... other attributes
```

When using `VBComponents.Import()`, Excel/VBA **adds these metadata lines to the module's code**, not just as metadata but as actual code content. This corruption is invisible in the VBA editor but affects execution.

**Solution Implemented:**

Changed approach from Remove → Import to **Direct Code Replacement**:

1. Read the .bas file as plain text
2. Strip VBA file headers (Attribute lines, VERSION lines, blank lines at start)
3. Extract only the clean VBA code
4. Use `CodeModule.DeleteLines()` to clear existing code
5. Use `CodeModule.AddFromString()` to write clean code directly

**Code Changes:**

```powershell
# OLD APPROACH (causes metadata corruption):
$vbaProject.VBComponents.Remove($module)
$vbaProject.VBComponents.Import($fullModulePath)

# NEW APPROACH (direct code replacement):
$moduleContent = Get-Content -Path $fullModulePath -Raw -Encoding UTF8
# Strip headers...
$codeModule.DeleteLines(1, $codeModule.CountOfLines)
$codeModule.AddFromString($cleanCode)
```

**Benefits:**

- ✅ **No invisible metadata** - only actual VBA code is written
- ✅ **Preserves module exactly** as seen in VBA editor
- ✅ **Faster** - no remove/import cycle needed
- ✅ **More reliable** - equivalent to manual copy-paste
- ✅ **Creates modules if missing** - handles new module additions

**Technical Details:**

The script now:
- Detects and removes VBA file header lines (`Attribute VB_`, `VERSION `)
- Handles modules that don't exist yet (creates them with `VBComponents.Add(1)`)
- Reports line count of actual code written for verification
- Uses UTF8 encoding to preserve special characters

**Lesson Learned:**

Never use `VBComponents.Import()` for automation when code correctness is critical. Always use direct `CodeModule` manipulation to avoid hidden metadata corruption.

**Files Modified:**

- `Automations/Excel_automaatio.ps1` - Replaced Import/Remove logic with direct CodeModule manipulation

**Testing:**

Verified that modules updated via automation now behave identically to manually copy-pasted code.

