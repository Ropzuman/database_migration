# Fixed SQL for _qryForExcel - Faceplate Update Required

## Problem Identified ✅

The `_qryForExcel` is a complex SELECT query with:
- Multiple JOINs (DOCUMENTS, ProjInfo, Status tables)
- Calculated fields (Nz, IIf, Left, Mid, InStr functions)
- Complex nested expressions

**Current Faceplate SQL (WRONG):**
```sql
SELECT * FROM _qryForExcel WHERE DocName3 like 'HVAC kytkentälista'
```

This tries to **SELECT from a saved query**, which ODBC struggles with for complex queries.

## Solution: Use the Query Definition Directly

Instead of selecting FROM the saved query, use the query's SQL definition directly in the faceplate.

### Updated SQL for Faceplate

Replace the SQL in the faceplate with this:

**For Row 1 (Kytkentälista):**
```sql
SELECT DOCUMENTS.Pick1, Nz([ClientNo], [MetsoDocNo]) AS DocNo, IIf([DOCUMENTS]![DocName1] Like "Motor*",[DOCUMENTS]![DocName1] & " " & [DOCUMENTS]![DocName2],[DOCUMENTS]![DocName1]) AS DocName, DOCUMENTS.DocName1, DOCUMENTS.DocName2, DOCUMENTS.DocName3, [Status]![Description] AS Status, Left([Rev],InStr([Rev],"/")-1) AS RevNum, DOCUMENTS.Area AS Department, DOCUMENTS.WorkPath, DOCUMENTS.File, Left([RevApu],InStr([RevAPU],"/")-1) AS RevDate, Mid([rev],InStr([rev]," ")+1) AS RevAPU, ProjInfo.ProjNo, ProjInfo.ContractNo, ProjInfo.Mill, ProjInfo.Customer, ProjInfo.Manager, ProjInfo.Name AS Project, IIf(IsNull([rev])," /",Mid([rev],InStr([rev]," ")+1)) AS APU, IIf(IsNull([rev]),"",Left([rev],InStr([rev]," ")-1)) AS RevID, Left([Apu],InStr([apu],"/")-1) AS [Date], DOCUMENTS.Area, DOCUMENTS.ClientNo, DOCUMENTS.MetsoDocNo, DOCUMENTS.Note, DOCUMENTS.Rev FROM (DOCUMENTS LEFT JOIN ProjInfo ON DOCUMENTS.pick1 = ProjInfo.Dept) INNER JOIN Status ON DOCUMENTS.Status = Status.Status WHERE DOCUMENTS.DocName3 LIKE 'Kytkentälista'
```

**For Row 2 (HVAC kytkentälista):**
```sql
SELECT DOCUMENTS.Pick1, Nz([ClientNo], [MetsoDocNo]) AS DocNo, IIf([DOCUMENTS]![DocName1] Like "Motor*",[DOCUMENTS]![DocName1] & " " & [DOCUMENTS]![DocName2],[DOCUMENTS]![DocName1]) AS DocName, DOCUMENTS.DocName1, DOCUMENTS.DocName2, DOCUMENTS.DocName3, [Status]![Description] AS Status, Left([Rev],InStr([Rev],"/")-1) AS RevNum, DOCUMENTS.Area AS Department, DOCUMENTS.WorkPath, DOCUMENTS.File, Left([RevApu],InStr([RevAPU],"/")-1) AS RevDate, Mid([rev],InStr([rev]," ")+1) AS RevAPU, ProjInfo.ProjNo, ProjInfo.ContractNo, ProjInfo.Mill, ProjInfo.Customer, ProjInfo.Manager, ProjInfo.Name AS Project, IIf(IsNull([rev])," /",Mid([rev],InStr([rev]," ")+1)) AS APU, IIf(IsNull([rev]),"",Left([rev],InStr([rev]," ")-1)) AS RevID, Left([Apu],InStr([apu],"/")-1) AS [Date], DOCUMENTS.Area, DOCUMENTS.ClientNo, DOCUMENTS.MetsoDocNo, DOCUMENTS.Note, DOCUMENTS.Rev FROM (DOCUMENTS LEFT JOIN ProjInfo ON DOCUMENTS.pick1 = ProjInfo.Dept) INNER JOIN Status ON DOCUMENTS.Status = Status.Status WHERE DOCUMENTS.DocName3 LIKE 'HVAC kytkentälista'
```

**For Row 3 (Kytkentälista - duplicate):**
Same as Row 1.

## Why This Works

1. **Original approach:** `SELECT * FROM _qryForExcel WHERE...`
   - ODBC tries to query a saved query (nested query)
   - Fails because ODBC doesn't handle complex Access queries as objects

2. **New approach:** Uses the query's SQL definition directly
   - ODBC executes the query against the actual tables
   - Works because ODBC can query tables and join them

## How to Update Faceplate

### Step 1: Open Excel Workbook
Open the Kytkentälista workbook

### Step 2: Go to Main Sheet
Click on the "Main" tab at the bottom

### Step 3: Update SQL Queries
In the "SQL For This Document (Revision, Name etc.):" section:

**Cell C13 (Row 1):** Clear current content and paste the SQL for "Kytkentälista"

**Cell C14 (Row 2):** Clear current content and paste the SQL for "HVAC kytkentälista"  

**Cell C15 (Row 3):** Clear current content and paste the SQL for "Kytkentälista" (or leave as-is if same as Row 1)

### Step 4: Test
1. Click "Get Data" button
2. Should now work without ODBC error ✅
3. DB2 sheet should populate with document metadata ✅

## Field Mapping

The query returns these fields (which map to the global variables):

| Query Field | Maps To Variable | Purpose |
|-------------|------------------|---------|
| DocNo | DIDocNo | Document number (client or Metso) |
| DocName | DIDocName | Full document name |
| DocName1 | DIDocName1 | Document name line 1 |
| DocName2 | DIDocName2 | Document name line 2 |
| DocName3 | DIDocName3 | Document name line 3 (filter field) |
| Status | DIStatus | Document status |
| RevNum | (not used) | Revision number |
| Department | DIDepartName | Department/Area |
| WorkPath | DIPath | File path |
| File | DIFile | Filename |
| RevDate | DIRevDate | Revision date |
| ProjNo | DIProjNo | Project number |
| ContractNo | DIContract | Contract number |
| Mill | DIMill | Mill name |
| Customer | DICustomer | Customer name |
| Manager | DIManager | Project manager |
| Project | DIProjName | Project name |
| RevID | DIRevID | Revision ID |
| Date | DIDate | Document date |
| Rev | DIRev | Full revision string |

## Testing Checklist

After updating the SQL in the faceplate:

- [ ] Click "Get Data" button
- [ ] Verify no ODBC error appears
- [ ] Check DB1 sheet - should have circuit data
- [ ] Check DB2 sheet - should have ONE row with document metadata
- [ ] Check DB2 columns match field names above
- [ ] Click "Run Check" button
- [ ] Check Info sheet - should now populate with project info, manager, etc.
- [ ] Check Revisions sheet - should populate with revision history

## If Still Getting Errors

### Issue: Field Names Don't Match

If `HaeDocTiedot` doesn't find the fields, check the column headers in DB2:

The query uses aliases like:
- `Nz([ClientNo], [MetsoDocNo]) AS DocNo` → Column name will be "DocNo"
- `[Status]![Description] AS Status` → Column name will be "Status"

Make sure these match what `HaeDocTiedot` is looking for (see Module2.vba, Case statements).

### Issue: Too Many Rows Returned

If the WHERE clause doesn't filter properly:

```sql
WHERE DOCUMENTS.DocName3 LIKE 'Kytkentälista'
```

Change to exact match if needed:
```sql
WHERE DOCUMENTS.DocName3 = 'Kytkentälista'
```

Or use wildcard:
```sql
WHERE DOCUMENTS.DocName3 LIKE '*Kytkentälista*'
```

### Issue: No Rows Returned

Check the database to see what values exist in DocName3:
```sql
SELECT DISTINCT DocName3 FROM DOCUMENTS
```

Then adjust the WHERE clause accordingly.

## Alternative: Simpler Approach

If the complex query causes issues, you can simplify to just get what's needed:

```sql
SELECT 
  DOCUMENTS.ClientNo AS MetsoDocNo,
  DOCUMENTS.DocName1, 
  DOCUMENTS.DocName2, 
  DOCUMENTS.DocName3,
  DOCUMENTS.WorkPath,
  DOCUMENTS.File,
  DOCUMENTS.Rev,
  ProjInfo.ProjNo,
  ProjInfo.ContractNo,
  ProjInfo.Mill,
  ProjInfo.Customer,
  ProjInfo.Manager,
  ProjInfo.Name AS Project
FROM DOCUMENTS 
LEFT JOIN ProjInfo ON DOCUMENTS.pick1 = ProjInfo.Dept
WHERE DOCUMENTS.DocName3 LIKE 'Kytkentälista'
```

This removes the Status table join and complex calculated fields, which may not all be needed.

## Summary

✅ **Root cause:** ODBC can't handle `SELECT * FROM _qryForExcel` for complex saved queries  
✅ **Solution:** Use the query's SQL definition directly in the faceplate  
✅ **Action:** Update the 3 SQL cells in Main sheet with the full query SQL  
✅ **Expected result:** DB2 will populate, Info sheet will populate, no ODBC errors

Copy the SQL above into the faceplate and test!
