# Nz() Function Error Fix

Date: 2025-10-15  
Branch: OBDC-kokeilu  
Error: "Undefined function 'Nz' in expression"

## The Problem

**Error Message:**
```
Database Connection Failed for DB2

Tried OLE DB providers: 16.0, 15.0, 12.0
Last error: Undefined function 'Nz' in expression.

Query: SELECT * FROM _qryForExcel WHERE DocName3 like '%Kytkentälista%'
```

### Root Cause

The `Nz()` function is an **Access VBA function** that:
- Works **only inside Access** (in queries, forms, reports)
- Does **NOT work** in external SQL queries (via ODBC or OLE DB)
- Is used in the `_qryForExcel` saved query

When you query `_qryForExcel` from Excel, Access tries to execute the saved query's SQL, which contains `Nz()`, but the external connection doesn't support it.

## Solution: Query DOCUMENTS Table Directly

Instead of querying the Access saved query `_qryForExcel`, we should query the **DOCUMENTS table directly** with standard SQL.

### Replace Nz() with Standard SQL

**Access Nz() function:**
```sql
Nz([FieldName], "default")  -- Access-only function
```

**Standard SQL alternatives:**
```sql
-- Option 1: ISNULL (SQL Server, Access SQL)
ISNULL([FieldName], 'default')

-- Option 2: IIF with ISNULL (Access SQL)
IIF(ISNULL([FieldName]), 'default', [FieldName])

-- Option 3: Handle Nulls in Excel/VBA instead
-- Just let Null values come through and handle them in Module2
```

## Recommended Query for DB2

Replace your current DB2 query with this:

### Simple Version (No JOINs)
```sql
SELECT * FROM DOCUMENTS WHERE DocName3 LIKE '%Kytkentälista%'
```

### With Project Info (If you need ProjInfo table data)
```sql
SELECT 
  d.*,
  p.Name, 
  p.Manager, 
  p.Customer, 
  p.Mill,
  p.Contract,
  p.MetsoUnitName,
  p.DepartName
FROM DOCUMENTS d 
LEFT JOIN ProjInfo p ON d.ProjNo = p.ProjNo
WHERE d.DocName3 LIKE '%Kytkentälista%'
```

### Handling Null Values

If you need to replace Null values, use Access SQL's `IIF` function:

```sql
SELECT 
  d.DocNo,
  d.ProjNo,
  IIF(ISNULL(d.DocName1), '', d.DocName1) AS DocName1,
  IIF(ISNULL(d.DocName2), '', d.DocName2) AS DocName2,
  IIF(ISNULL(d.DocName3), '', d.DocName3) AS DocName3,
  d.Rev,
  d.Status,
  d.Date,
  d.WorkPath,
  IIF(ISNULL(p.Customer), '', p.Customer) AS Customer,
  IIF(ISNULL(p.Mill), '', p.Mill) AS Mill
FROM DOCUMENTS d 
LEFT JOIN ProjInfo p ON d.ProjNo = p.ProjNo
WHERE d.DocName3 LIKE '%Kytkentälista%'
```

## How to Update the Query

### Method 1: Update in Excel Main Sheet

1. Open your Excel workbook
2. Go to the **Main** sheet
3. Find the DB2 query cell (probably row 14, column C)
4. Replace the query with the new SQL above

**Current query:**
```sql
SELECT * FROM _qryForExcel WHERE DocName3 LIKE '%Kytkentälista%'
```

**New query:**
```sql
SELECT * FROM DOCUMENTS WHERE DocName3 LIKE '%Kytkentälista%'
```

### Method 2: Fix the Saved Query in Access

If you want to keep using `_qryForExcel`:

1. Open your Access database
2. Find the `_qryForExcel` query
3. Switch to SQL View
4. Replace all `Nz(...)` with `IIF(ISNULL(...), 'default', ...)`
5. Save the query

**Example conversion:**
```sql
-- Before:
SELECT Nz([DocName1], '') AS DocName1

-- After:
SELECT IIF(ISNULL([DocName1]), '', [DocName1]) AS DocName1
```

## Why Nz() Doesn't Work Externally

| Context | Nz() Support | Alternative |
|---------|--------------|-------------|
| **Access Forms/Reports** | ✅ Yes | Use Nz() freely |
| **Access Queries (internal)** | ✅ Yes | Use Nz() freely |
| **ODBC Queries from Excel** | ❌ No | Use IIF(ISNULL()) |
| **OLE DB Queries from Excel** | ❌ No | Use IIF(ISNULL()) |

The Nz() function is part of the **Access VBA runtime**, not part of **Access SQL Engine**. When you query externally, only the SQL engine is available.

## Module2 Already Handles Nulls

Good news: Your Module2.vba already handles Null values safely!

**Line 54-57 in Module2:**
```vba
DIRev = wsDB2.Cells(2, i).Value
If Not IsNull(DIRev) And DIRev <> "" Then
  DIRevArr = Split(DIRev, Chr(10))
End If
```

**Line 105-114 in Module2:**
```vba
pathStr = wsDB2.Cells(2, i).Value & ""  ' Converts Null to empty string
If pathStr <> "" Then
  ' Process path
End If
```

So you **don't need** to replace Nulls in the SQL query—Module2 will handle them!

## Recommended Action

**Update your DB2 query in the Main sheet:**

Change from:
```sql
SELECT * FROM _qryForExcel WHERE DocName3 LIKE '%Kytkentälista%'
```

To:
```sql
SELECT * FROM DOCUMENTS WHERE DocName3 LIKE '%Kytkentälista%'
```

This will:
- ✅ Avoid the Nz() error
- ✅ Work with OLE DB
- ✅ Return the same data (Module2 handles Nulls)
- ✅ Be simpler and faster

## Testing

After updating the query:

1. Click **"Get Data"** button
2. Check DB2 sheet has data
3. Click **"Run Check"** button  
4. Verify Info sheet populates correctly

## Summary

The error isn't actually about OLE DB providers—**all three providers worked!** The error is that your saved query `_qryForExcel` uses the `Nz()` function, which doesn't work in external queries.

**Fix:** Query the DOCUMENTS table directly instead of using the saved query.
