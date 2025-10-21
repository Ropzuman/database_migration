DOCUMENTS_64BIT_MIGRATION.md

This document summarizes the Access DOCUMENTS 64-bit migration helper artifacts and instructions.

What I added:
- scripts/export_access_vba.ps1 : PowerShell script that exports all VBA components from .accdb/.mdb files using Access COM automation.

How to run (PowerShell / Windows):
1. Ensure 'Trust access to the VBA project object model' is enabled in Access Trust Center -> Macro Settings.
2. Match PowerShell bitness to Access/Office bitness. For 32-bit Access, run the 32-bit PowerShell located in `C:\Windows\SysWOW64\WindowsPowerShell\v1.0\powershell.exe`.
3. From repository root:
   - Export single DB:
     powershell -ExecutionPolicy Bypass -File .\scripts\export_access_vba.ps1 -InputPath "C:\Path\To\YOUR.accdb" -OutDir "C:\temp\access_exports"
   - Export all DBs in a folder:
     powershell -ExecutionPolicy Bypass -File .\scripts\export_access_vba.ps1 -InputPath "C:\Path\To\FolderWithDBs" -OutDir "C:\temp\access_exports"

Caveats & Troubleshooting:
- The script uses COM automation and requires desktop interaction; run it from an interactive user session.
- If you see "VBE object not available" — enable the Trust access setting.
- If Access process remains after running, kill `MSACCESS.EXE` in Task Manager.
- The script attempts to export all VBProjects available to Access. Some protected DBs will fail to export their code.

Next steps for DOCUMENTS migration:
- Continue reading remaining Access modules and classes and apply 64-bit fixes where necessary.
- Identify and remove dead code: `Form_USysRevText_OLD.cls` likely obsolete; verify before deletion.
- Replace custom functions that shadow built-ins (e.g., custom `Replace`) and test behavior.
- Create `DOCUMENTS_64BIT_MIGRATION_CHANGELOG.md` once edits are applied.

Notes:
- If you want an in-Access macro, see the VBA snippet in `scripts/export_access_vba.ps1` header.
