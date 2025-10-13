# ODBC Error Fix - Missing `_qryForExcel` Table

## Error Encountered

**ODBC Error:**
```
Database: C:\Data\Opinnäytetyö\Mallikanta\24PRO229 Fortum Nuojala\2\DBLoopCircuit.accdb
SQL Query 2: SELECT * FROM _qryForExcel WHERE DocName3 like 'HVAC kytkentälista'
```

**Error Message:** "Table `_qryForExcel` doesn't exist in this database"

## Root Cause

The faceplate has SQL queries configured to fetch document metadata from a table called `_qryForExcel`:

```sql
SELECT * FROM _qryForExcel WHERE DocName3 like 'Kytkentälista'
```

**However:** Not all Access databases have this table/query!

- **DB1 query:** Works ✅ (Circuit_Diagrams_IO_Terminals table exists)
- **DB2 query:** Fails ❌ (_qryForExcel table doesn't exist in this particular database)

When DB2 query failed, the entire `HaeData()` function stopped with an error dialog, preventing DB1 from being used.

## Solution: Graceful Error Handling

Changed `HaeData()` to **continue execution even if individual queries fail**:

### Before (Stop on First Error)
```vba
On Error GoTo ErrorHandler

For i = 1 To 2
  ' Run query
  .Refresh
Next i

ErrorHandler:
  MsgBox "ODBC Error: " & Err.Description
  ' Stops entire function!
End Sub
```

### After (Skip Failed Queries, Continue)
```vba
For i = 1 To 2
  On Error Resume Next
  ' Run query
  .Refresh
  If Err.Number <> 0 Then
    ' Log error but continue with next query
    queryFailed = True
    errorMsg = errorMsg & "DB" & i & " Query Failed: " & Err.Description
    Err.Clear
  End If
  On Error GoTo 0
Next i

' Show appropriate message
If queryFailed Then
  MsgBox "Data partially loaded! Some queries failed..."
Else
  MsgBox "Data brought successfully!"
End If
```

## Behavior After Fix

### Scenario 1: Both Queries Succeed ✅
- DB1 populated with circuit data
- DB2 populated with document metadata
- Message: "Data brought successfully!"
- Info sheet will populate correctly

### Scenario 2: DB2 Query Fails (Current Situation) ⚠️
- DB1 populated with circuit data ✅
- DB2 remains empty (query failed)
- Message: "Data partially loaded! Some queries failed..."
- **Impact:** Info sheet will be empty, but template validation and printout generation can still work with DB1 data

### Scenario 3: DB1 Query Fails ❌
- DB1 remains empty (query failed)
- DB2 may or may not populate
- Message: "Data partially loaded! Some queries failed..."
- **Impact:** Template validation will fail (no headers to match)

## Why This Database Doesn't Have `_qryForExcel`

Looking at the faceplate SQL queries, `_qryForExcel` appears to be a **project-specific query/table** that exists in some Access databases but not others.

**Possible reasons:**
1. Different database versions (older databases don't have this query)
2. Database was created without the metadata structure
3. Query was renamed or deleted
4. This is a different type of project that doesn't use the standard metadata query

## Recommended Solutions

### Option A: Check if Database Has Document Metadata Table (Recommended)

Add logic to detect which metadata table exists in the database:

```vba
' Try _qryForExcel first
sSQL(2) = "SELECT * FROM _qryForExcel WHERE DocName3 like 'Kytkentälista'"

' If that fails, try alternative table names
If DB2 query fails Then
  sSQL(2) = "SELECT * FROM DocumentInfo WHERE DocName3 like 'Kytkentälista'"
  ' Or try other common table names
End If
```

### Option B: Make DB2 Query Optional

Modify the faceplate to have a checkbox "Use Document Metadata Query" that can be unchecked for databases without `_qryForExcel`.

### Option C: Skip DB2 Query if Table Doesn't Exist

Add a pre-check before running the query:

```vba
' Check if table exists before querying
If TableExists("_qryForExcel", Kanta) Then
  ' Run DB2 query
Else
  ' Skip DB2 query, leave metadata fields empty
End If
```

### Option D: Use Empty Defaults (Current Behavior)

With the current fix, if DB2 fails:
- User sees warning message
- Info sheet stays empty
- User can still proceed with printout generation (just without document metadata in footers)

## Testing Results Expected

### Test 1: Click "Get Data"
**Expected Result:**
- Message box: "Data partially loaded! Some queries failed..."
- Shows: "DB2 Query Failed: [Table '_qryForExcel' does not exist]"
- DB1 sheet: Contains circuit data ✅
- DB2 sheet: Empty (query failed)

### Test 2: Click "Run Check"
**Expected Result:**
- Template validation should work (uses DB1 only)
- Info sheet: Empty (no DB2 data)
- Warning in ERRORS sheet: "No document metadata found in DB2"

### Test 3: Click "Generate Printout"
**Expected Result:**
- Printout generates successfully using DB1 data
- Page footers show empty fields for project/document info (no DB2 data)
- Circuit/terminal linking works correctly

## Workarounds for Missing Document Metadata

If you need document metadata but the database doesn't have `_qryForExcel`, you can:

1. **Manually fill Info sheet** - After Checkout, manually type in project name, manager, etc.
2. **Modify SQL query** - Change the DB2 query in faceplate to use a different table that exists in this database
3. **Add `_qryForExcel` to database** - Open the Access database and create the missing query/table
4. **Use a different database** - If this is the wrong database, select one that has the metadata structure

## File Modified

**File:** `Module1.vba`  
**Function:** `HaeData()`  
**Change:** Added per-query error handling with `On Error Resume Next`

## Compile Status

✅ Module1.vba: No errors  
✅ Code will now handle missing tables gracefully

## Summary

The ODBC error occurred because the Access database doesn't have a `_qryForExcel` table. The fix allows the code to **continue execution** even when individual queries fail, so:

- ✅ DB1 data loads successfully
- ⚠️ DB2 fails gracefully with warning message
- ✅ User can still use DB1 data for template validation and printouts
- ⚠️ Info sheet will be empty (expected with no DB2 data)

The error is now **handled gracefully** rather than stopping execution!
