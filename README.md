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
- **64-bit ODBC:** Auto-brackets Access query names: `FROM _qryForExcel` → `FROM [_qryForExcel]`
- **DB2 Query Fix:** Changed from saved query to direct DOCUMENTS table query to avoid ODBC WHERE clause limitations
- **Column Mapping:** Complete mapping from DOCUMENTS table to Info sheet (see `COLUMN_MAPPING_COMPLETE.md`)
- **Performance:** Consistent BeginFastMode/EndFastMode usage, optimized screen updating
- **Safety:** Infinite loop protection, sheet existence validation, database file checks
- **Type Safety:** `Long` for all counters/indices (64-bit compatibility)

## Excel Workflows

### 1. Get Data (HaeData)
- Fetches data from Access database via ODBC
- Populates DB1 (Circuit_Diagrams_IO_Terminals) and DB2 (DOCUMENTS) sheets
- Validates database file existence, handles ODBC errors

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
- `Automations/` - Automation scripts used to migrate and extract VBA components (see `Automations/export_access_vba.ps1` and `Automations/Excel_automaatio.ps1`)
- `Logs/AUTOMATIONS_LOG.md` - Log and changelog for automation scripts

## Version History
See `Logs/CHANGELOG_64bit_and_perf.md` for detailed change history.
