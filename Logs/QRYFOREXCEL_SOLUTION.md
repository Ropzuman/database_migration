# ODBC Error with _qryForExcel Saved Query - Solution

## Issue
ODBC error when trying to query `_qryForExcel`:
```
Table '_qryForExcel' does not exist
```

**But:** You confirmed that `_qryForExcel` DOES exist as a saved query in Access that opens as a datasheet.

## Root Cause: ODBC Can't See Saved Queries Directly

The issue is that **ODBC drivers treat Access saved queries differently than tables**:

- ✅ **Tables:** ODBC can query directly with `SELECT * FROM TableName`
- ⚠️ **Saved Queries:** ODBC may not be able to query them the same way
- ❌ **Parameter Queries:** ODBC definitely can't query these without special handling

### Why This Happens

When you use ODBC to connect to Access:
1. ODBC sees **tables** as queryable objects
2. ODBC may or may not see **saved queries** as queryable objects (depends on query type)
3. If the saved query is a **parameter query** or has special Access features, ODBC can't execute it

## Solutions Attempted in Code

The updated `HaeData()` function now tries multiple approaches:

### Attempt 1: Original SQL
```sql
SELECT * FROM _qryForExcel WHERE DocName3 like 'HVAC kytkentälista'
```

### Attempt 2: With Brackets (if Attempt 1 fails)
```sql
SELECT * FROM [_qryForExcel] WHERE DocName3 like 'HVAC kytkentälista'
```

### Attempt 3: Graceful Failure
If both attempts fail, shows warning but continues with DB1 data.

## Recommended Solutions

### Option A: Convert Query to Table (Easiest) ✅

**In Access:**
1. Open the database
2. Run the `_qryForExcel` query to see results
3. Select all results (Ctrl+A)
4. Copy (Ctrl+C)
5. Create new table: Right-click → Paste → "Yes" to create new table
6. Name it `_qryForExcel_Table`
7. Update faceplate SQL to use `_qryForExcel_Table` instead

**Advantage:** ODBC can definitely query tables

### Option B: Check if _qryForExcel is a Parameter Query

**In Access:**
1. Right-click `_qryForExcel` → Design View
2. Check if there are any parameters defined (Parameters dialog)
3. If YES: This is why ODBC fails - parameter queries need special handling

**Solution if Parameter Query:**
- Remove parameters and hard-code the filtering
- OR create a new query without parameters

### Option C: Use Make-Table Query (Automated)

**In Access:**
1. Create a new Make-Table query:
```sql
SELECT * INTO _qryForExcel_Export
FROM _qryForExcel
WHERE DocName3 like 'Kytkentälista'
```
2. Run this query before using the Excel tool
3. Update faceplate to query `_qryForExcel_Export` instead

### Option D: Export Query Definition

**Check what type of query `_qryForExcel` is:**
1. In Access, right-click `_qryForExcel` → Design View
2. Switch to SQL View
3. Copy the SQL and paste it directly into the faceplate instead of `SELECT * FROM _qryForExcel`

**Example:** If _qryForExcel SQL is:
```sql
SELECT DocName3, Project, Manager FROM Documents
```

Then in faceplate, use:
```sql
SELECT DocName3, Project, Manager FROM Documents WHERE DocName3 like 'Kytkentälista'
```

### Option E: Use ADO Instead of ODBC (Code Change)

Replace ODBC QueryTables with ADO connection:

```vba
' Instead of QueryTables, use ADO
Dim cn As Object
Dim rs As Object
Set cn = CreateObject("ADODB.Connection")
cn.Open "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & Kanta
Set rs = cn.Execute(sSQL(2))

' Copy recordset to worksheet
ws.Range("A1").CopyFromRecordset rs
rs.Close
cn.Close
```

**Advantage:** ADO can execute saved queries that ODBC can't

## Current Code Behavior

With the updated code:

1. **Tries original SQL**
   - If succeeds: DB2 populates ✅
   - If fails: Goes to step 2

2. **Tries SQL with brackets**
   - If succeeds: DB2 populates ✅
   - If fails: Goes to step 3

3. **Shows warning, continues**
   - DB1 still works ✅
   - DB2 empty ⚠️
   - User can proceed with DB1 data

## Diagnostic Steps

### Step 1: Check Query Type in Access
1. Open Access database
2. Find `_qryForExcel` in queries list
3. Right-click → Design View
4. Check for:
   - Parameters (View → Parameters)
   - Crosstab queries
   - Action queries (UPDATE, DELETE, APPEND)
   - Special functions that ODBC doesn't support

### Step 2: Test Query in Access
1. Open `_qryForExcel` in datasheet view
2. Does it prompt for parameters? ❌ This is the problem
3. Does it show data immediately? ✅ Should work with ODBC

### Step 3: Get Query SQL
1. Open `_qryForExcel` in Design View
2. View → SQL View
3. Copy the SQL text
4. Share it so we can convert it to ODBC-compatible SQL

## Quick Fix You Can Try Now

### Test with Direct Table Access

If `_qryForExcel` is based on a table, try querying that table directly:

1. Open Access, open `_qryForExcel` in Design View
2. Look at what table(s) it queries (shown in top section)
3. Try changing the faceplate SQL to query that table directly:

**Instead of:**
```sql
SELECT * FROM _qryForExcel WHERE DocName3 like 'HVAC kytkentälista'
```

**Try:**
```sql
SELECT * FROM [ActualTableName] WHERE DocName3 like 'HVAC kytkentälista'
```

(Replace `ActualTableName` with whatever table _qryForExcel queries)

## Files Modified

- **Module1.vba** - Added retry logic with brackets and graceful error handling

## Next Steps

To permanently fix this, I need to know:
1. What type of query is `_qryForExcel`? (SELECT query, parameter query, crosstab, etc.)
2. What is the SQL definition of `_qryForExcel`? (View → SQL View in Access)
3. What table(s) does it query?

With this information, I can provide the exact SQL to use in the faceplate that will work with ODBC.
