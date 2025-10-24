# Database Migration: 32-bit to 64-bit Office VBA

## Purpose

Migration of an MS Access/Excel/AutoCAD design system from 32-bit to 64-bit Office (VBA7). Updates API declarations, data access patterns, and ODBC compatibility for 64-bit operation with performance optimization.

## Components

- **Access VBA:** `Access/kanta.vba`, `kanta1.vba`, `kanta2.vba`, `kanta3.vba`
- **Excel Macros:** `Excel/Kytkentälista/Module1.vba`, `Module2.vba`, `Module3.vba`
- **Diagnostics:** `Excel/Kytkentälista/Debugging/` (diagnostic and testing tools)
- **Documentation:** `COLUMN_MAPPING_COMPLETE.md`, `Logs/CHANGELOG_64bit_and_perf.md`

## Key Changes

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


## Version History

See `Logs/CHANGELOG_64bit_and_perf.md` for detailed change history.
