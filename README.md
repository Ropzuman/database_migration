# Database Migration: 32-bit to 64-bit Office VBA

## Purpose

Migration of an MS Access/Excel/AutoCAD design system from 32-bit to 64-bit Office (VBA7). Updates API declarations, data access patterns, and ODBC compatibility for 64-bit operation with performance optimization.

## Components

[This section will be updated with new information once the placeholder is filled.]

### Access VBA

- **API Compatibility:** Conditional compilation (`#If VBA7 Then`) with `PtrSafe` and `LongPtr` for pointers/handles
- **DAO:** Explicit qualification (`DAO.Database`, `DAO.Recordset`), safe cleanup patterns
- **Transactions:** `DBEngine.BeginTrans/CommitTrans/Rollback` for robust inserts
- **Validation:** Hardened input checks in `CustomMessage`, explicit typing for globals

### Excel Macros

- **OLE DB Migration:** Switched from ODBC to OLE DB (ACE.OLEDB) for better 64-bit compatibility
- **Provider Fallback:** Automatic fallback: 16.0 → 15.0 → 12.0 for Office version compatibility
- **DB2 Query Fix:** Fixed saved query `_qryForExcel` to work with OLE DB (removed `Nz()` function)
- **Column Mapping:** Complete mapping from DOCUMENTS table to Info sheet (see `COLUMN_MAPPING_COMPLETE.md`)
- **Performance:** Consistent BeginFastMode/EndFastMode usage, optimized screen updating
- **Safety:** Infinite loop protection, sheet existence validation, database file checks
- **Type Safety:** `Long` for all counters/indices (64-bit compatibility)

## Excel Workflows

### 1. Get Data (HaeData)

- Fetches data from Access database via OLE DB (with automatic provider fallback)
- Populates DB1 (Circuit_Diagrams_IO_Terminals) and DB2 (DOCUMENTS) sheets
- Validates database file existence, handles connection errors gracefully

### 2. Run Check (Checkout)

- Validates TEMPLATE headers against DB1 data
- Extracts document metadata from DB2 (via HaeDocTiedot)
- Populates Info sheet using comment-based templating (via VaihdaInfo)
- Reports missing headers to ERRORS sheet

### 3. Generate Printout (GenPrintout)

- Creates new workbook from TEMPLATE
- Populates with DB1 data and document metadata
- Generates print-ready output

## Testing

1. **Access:** Debug → Compile in VBA editor, test `SniffUser`, `CustomMessage`
2. **Excel:** Click "Get Data" → verify DB1/DB2 populate without errors
3. **Excel:** Click "Run Check" → verify Info sheet populates, no ERRORS
4. **Excel:** Click "Generate Printout" → verify new workbook created

## Excel: AutoCAD Import/Export (AcadDATA)

This workbook includes an AutoCAD integration to read block attributes (and optional text entities) from DWG and write changes back.

### Import (TuoDATA)

- Run via Macros: TuoDATA_All (all) or TuoDATA_Selected (previous selection in AutoCAD).
- Inputs on the Start sheet:
  - D7: Block names, comma-separated. Use `*` to import all blocks.
  - D5: Entity scope: "Blokit" or "Blokit ja tekstit" (includes TEXT/MTEXT).
- Output columns: PATH, DWG, BLOCK, HANDLE, XCord, YCord, Layer, then one column per attribute tag (created as needed).

Selection behavior

- No layer filter (by design for simplicity and correctness).
- The selection uses DXF filters for performance:
  - Always filters by entity type: `INSERT` (and `TEXT`/`MTEXT` if chosen).
  - If specific block names are given (not just `*`), a code-2 OR-group is added for those names.
  - If that yields zero entities (typical for dynamic blocks with anonymous names), the code falls back to type-only selection and prunes in VBA by `EffectiveName` to keep only the requested blocks.
- Text entities (when enabled) bypass the block-name filter and are included as-is.

Dynamic blocks

- Matching is performed against `EffectiveName`, so dynamic blocks are handled correctly.
- The fallback path ensures dynamic blocks are included even if their anonymous DXF names don’t match the code-2 name filter.

Dev tracing

- Controlled in `Excel/Moduulit/AcadDATA/Koodit.bas` by `Public Const DEBUG_TRACE As Boolean`.
- When `True`, detailed timestamps and steps are printed to the Immediate Window (Ctrl+G) during import: filter setup, selection counts, fallback triggers, and per-document row totals.
- Functional behavior is unchanged; tracing only affects verbosity.

### Export (VieDATA)

More details for developers: see `Logs/ACADDATA_DEVELOPER_NOTES.md`.

- Writes edited attribute values back to blocks by HANDLE.
- Respects AutoCAD SingleDocumentMode and performs safe COM cleanup.

### Double-click navigation

- On the data sheet, double-click a row to locate the entity in AutoCAD and zoom to it.
- Uses a robust zoom sequence: ZoomWindow to the entity’s bounding box and a safe scaled zoom helper to avoid enum mismatches under late binding.

## File Organization

- `Access/` - Access VBA modules
- `Excel/Kytkentälista/` - Excel macro modules
- `Excel/Kytkentälista/Debugging/` - Diagnostic tools (DiagnosticTest, ColumnMappingDiagnostic, FixInfoSheetComments)
- `Logs/` - Changelogs, fixes, and analysis documentation
- `COLUMN_MAPPING_COMPLETE.md` - DB2→Info sheet mapping reference
- `Automations/` - Automation scripts used to migrate and extract VBA components (see `Automations/export_access_vba.ps1`, `Automations/Excel_automaatio.ps1`, and `Automations/Access_automaatio.ps1`)
- `Logs/AUTOMATIONS_LOG.md` - Log and changelog for automation scripts

## Automations: Access updater

Use `Automations/Access_automaatio.ps1` to replace Access VBA modules (.bas) and class modules (.cls) inside an .accdb safely on 64-bit Office.

Prerequisites:

- Run in 64-bit PowerShell (x64). The script enforces this and exits if not.
- Microsoft Access installed (same bitness as PowerShell).
- Trust Center: enable "Trust access to the VBA project object model" and ensure the database and module folder are in trusted locations.

What it does:

- Prompts for the database file path and the folder containing exported VBA components.
- Opens the database with retry logic and removes read-only attribute if set.
- Removes listed components if present, then imports the corresponding .bas/.cls files.
- Saves the database, re-enables warnings, and performs robust COM cleanup to avoid orphaned MSACCESS.EXE.

Notes:

- You can pre-fill default paths by editing `$DefaultAccessFilePath` and `$DefaultComponentPath` at the top of the script.

## Automations: Excel updater

Use `Automations/Excel_automaatio.ps1` to replace VBA modules in .xlsm workbooks across multiple locations.

Prerequisites:

- Run in PowerShell (any bitness, but must match Excel if opening workbooks).
- Microsoft Excel installed.
- Trust Center: enable "Trust access to the VBA project object model".

What it does:

- Prompts for the folder containing .xlsm files and the folder with updated .bas module files.
- For each .xlsm workbook:
  - Opens with retry logic (handles OneDrive locks).
  - **Replaces module code directly** (reads .bas content, strips headers, writes clean code via CodeModule API).
  - Saves to temporary file, deletes original, renames (atomic replacement pattern).
- Performs robust COM cleanup to avoid orphaned EXCEL.EXE processes.

**Important:** The script uses **direct code replacement** instead of `VBComponents.Import()` to avoid invisible metadata corruption that can cause modules to behave incorrectly. This approach is equivalent to manually copy-pasting code into the VBA editor.

Notes:

- Pre-fill default paths by editing `$DefaultExcelFilesPath` and `$DefaultModulePath` at the top.
- Module names to update are defined in `$moduleNames` array (default: Module1, Module2, Module3).
- See `Logs/AUTOMATIONS_LOG.md` for technical details on the VBComponents.Import issue.

## Version History

See `Logs/CHANGELOG_64bit_and_perf.md` for detailed change history.

## Maintenance: VBA cleanup (2025-10-30)

- Non-functional cleanup in `Excel/Moduulit/AcadDATA/Koodit.bas`:
  - Extracted a tiny helper to build DXF type filters (reduces duplication and ReDim Preserve calls).
  - Removed a couple of unused variables and clarified comments.
  - No behavior changes; import/export flows and selection logic are identical.
- Developer notes updated: see `Logs/ACADDATA_DEVELOPER_NOTES.md`.
- Changelog entry: `Logs/ACADDATA_CLEANUP_2025-10-30.md`.

## Excel: Kytkentälista (DB fetch + printout)

The Kytkentälista tool uses two SQL inputs on the faceplate:

- DB1 (body data): populates the main rows for the printout.
- DB2 (document info): fills Info/Revisions and drives the default Save As location.

Inputs

- SQL can target tables or saved Access queries (e.g., `_qryForExcel`) via ODBC.
- Prefer ANSI-92 wildcards (`%`) in SQL for portability if filtering with LIKE.

Save As defaults

- Path comes from DB2 WorkPath (header is case-insensitive; common synonyms like `workpath`, `path`, `work_path`, `listpath`, `lists_path`, `savepath`, `targetpath`, `outputpath` are accepted). The path is normalized and ends with `\`.
- File name comes from DB2 File. If missing, it falls back to the faceplate Body Sheet Name.
- If the name has no extension, `.xlsx` is appended automatically.

Template population

- GenPrintout uses template-driven population: copies TEMPLATE blocks per data group (RMAX rows), then maps values from LINKING sheet via comment-based markers.
- This preserves the template's layout, formatting, and cell linking logic.
- Each comment marker in the TEMPLATE (created during Checkout) links to a DB1 column; VaihdaLinkit reads these markers and populates cells with corresponding data.

Diagnostics

- After Get Data, the status bar briefly shows row counts per sheet (DB1/DB2). The Immediate Window also prints a summary (view via Ctrl+G).

Recent updates (2025-11-03)

- Fixed critical bug in VaihdaLinkit that caused incorrect template population
- Optimized code: removed dead variables, added constants, improved comments
- See `Logs/KYTKENTALISTA_OPTIMIZATION_2025-11-03.md` for details
