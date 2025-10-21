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
