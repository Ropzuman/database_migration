AUTOMATIONS_LOG.md

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
