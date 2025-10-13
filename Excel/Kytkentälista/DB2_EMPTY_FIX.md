# Info Sheet & DB2 Empty - Root Cause Fixed

## Issue Summary
- **Info sheet:** Completely empty after Checkout
- **DB2 sheet:** Completely empty after Get Data
- **DB1 sheet:** Has data (working correctly)

## Root Cause Identified ✅

### Problem: `_qryForExcel` Query was Being Skipped

The optimized code had this logic in `HaeData()`:

```vba
' Skip Excel-based queries (_qryForExcel) - only process Access database queries
If sSQL(i) <> "" And InStr(1, sSQL(i), "_qryForExcel", vbTextCompare) = 0 Then
```

This was **skipping any SQL query containing `_qryForExcel`** in the name!

But according to the faceplate, the SQL for DB2 (document metadata) is:
```sql
SELECT * FROM _qryForExcel WHERE DocName3 like 'Kytkentälista'
```

**Result:** DB2 query was never executed → DB2 stayed empty → Info sheet had no data to populate!

### Secondary Issue: Missing Revisions Call

The original code calls `VaihdaInfo` **twice** in Checkout:

```vba
HaeDocTiedot 'Hakee dokumentin tiedot DB2-sheetiltä
VaihdaInfo   'Vaihtaa dokumentin tiedot info sheetille
VaihdaInfo ("Revisions") 'Tiedot revisions sheetille
```

The "working version" was missing the second call to populate the Revisions sheet.

## Fixes Applied

### Fix #1: Removed Query Skip Logic (Module1.vba, HaeData)

**Before:**
```vba
' Skip Excel-based queries (_qryForExcel) - only process Access database queries
If sSQL(i) <> "" And InStr(1, sSQL(i), "_qryForExcel", vbTextCompare) = 0 Then
  Set TAULUKKO = ws.QueryTables.Add(Connection:=Yhteys, Destination:=ws.Range("A1"))
  ' ... query execution ...
End If
```

**After:**
```vba
' Run query if SQL is not empty
If sSQL(i) <> "" Then
  Set TAULUKKO = ws.QueryTables.Add(Connection:=Yhteys, Destination:=ws.Range("A1"))
  ' ... query execution ...
End If
```

Now **all queries are executed**, including the `_qryForExcel` query for DB2.

### Fix #2: Added Revisions Sheet Population (Module1.vba, Checkout)

**Before:**
```vba
HaeDocTiedot
VaihdaInfo   'Populate document info to Info sheet only (not Revisions during checkout)
```

**After:**
```vba
HaeDocTiedot
VaihdaInfo   'Populate document info to Info sheet
VaihdaInfo ("Revisions") 'Populate revisions to Revisions sheet
```

Now **both Info and Revisions sheets are populated** during Checkout.

## Why the `_qryForExcel` Skip Was Added (Historical Context)

Looking at previous documentation, the skip logic was added because someone thought `_qryForExcel` queries were "Excel-based" and shouldn't be run through ODBC.

**This was incorrect!** `_qryForExcel` is simply a **naming convention** in the Access database. It's still an Access query that needs to be executed via ODBC to populate DB2.

The query `SELECT * FROM _qryForExcel WHERE...` works perfectly fine through ODBC - it's just querying a table/view named `_qryForExcel` in the Access database.

## Testing Steps

### Test 1: Verify DB2 Gets Populated
1. Click **"Get Data"** button
2. Check DB2 sheet
3. **Expected:** Row 1 has headers (rev, revid, revdate, docno, project, manager, etc.)
4. **Expected:** Row 2 has document metadata values

### Test 2: Verify Info Sheet Gets Populated
1. After Get Data succeeds
2. Click **"Run Check"** button
3. Check Info sheet
4. **Expected:** Cells with comments should now display document info (project name, manager, document number, etc.)

### Test 3: Verify Revisions Sheet Gets Populated
1. After Checkout completes
2. Check Revisions sheet
3. **Expected:** Revision history table should be populated with revision IDs, dates, designers, checkers, approvers, descriptions

## Technical Details

### Database Query Structure

**DB1 Query (Circuit Diagrams):**
```sql
SELECT * FROM Circuit_Diagrams_IO_Terminals 
WHERE [Control Place]='CBA10'
```
- Populates: Circuit/terminal data
- Used by: Template linking (£ markers)

**DB2 Query (Document Metadata):**
```sql
SELECT * FROM _qryForExcel 
WHERE DocName3 like 'Kytkentälista'
```
- Populates: Document properties (project, manager, revision info, etc.)
- Used by: Info sheet, Revisions sheet, page footers

### Info Sheet Population Flow

```
Get Data Button
  ↓
HaeData() executes both SQL queries
  ↓
DB1 populated (circuit data)
DB2 populated (document metadata)
  ↓
Run Check Button
  ↓
HaeDocTiedot() reads DB2, stores in global variables (DIProject, DIManager, etc.)
  ↓
VaihdaInfo("Info") reads comments in Info sheet, fills cells with global variable values
  ↓
VaihdaInfo("Revisions") reads comments in Revisions sheet, fills revision table
```

## Compile Status
✅ Module1.vba: No errors  
✅ Module2.vba: No errors

## Summary of Changes

| File | Function | Change | Reason |
|------|----------|--------|--------|
| Module1.vba | HaeData | Removed `_qryForExcel` skip logic | DB2 query contains `_qryForExcel` and needs to execute |
| Module1.vba | Checkout | Added `VaihdaInfo("Revisions")` call | Original code populates both Info and Revisions sheets |

## Expected Behavior After Fix

1. **Get Data** → DB1 and DB2 both populated ✅
2. **Run Check** → Info and Revisions sheets both populated ✅
3. **Generate Printout** → Document metadata appears in footers ✅

## Why This Wasn't Caught Earlier

The `_qryForExcel` skip logic was added during an optimization pass with the assumption that:
- Queries with "Excel" in the name were Excel-based formulas
- They shouldn't be executed through ODBC

This assumption was wrong. `_qryForExcel` is just a query name in the Access database, and it executes normally through ODBC like any other query.

The fix restores the original behavior: **execute all non-empty SQL queries, regardless of name**.
