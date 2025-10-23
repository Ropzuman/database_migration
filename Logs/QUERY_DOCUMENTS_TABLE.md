# Solution: Query DOCUMENTS Table Directly

## Problem
The saved query `_qryForExcel` cannot be filtered through ODBC with WHERE clauses.

## Solution
Query the DOCUMENTS table directly with proper JOINs, just like DB1 queries Circuit_Diagrams_IO_Terminals.

## Recommended Query for DB2

Based on typical Access database structures for document management systems, try this query:

```sql
SELECT d.*, p.Name as ProjName, p.Manager, p.Customer, p.Mill, p.Contract as ContractNo, s.StatusName
FROM DOCUMENTS d 
LEFT JOIN ProjInfo p ON d.ProjNo = p.ProjNo
LEFT JOIN Status s ON d.Status = s.StatusID
WHERE d.DocName3 LIKE '%Kytkentälista%'
```

## Alternative Simpler Query (If Above Fails)

If you're not sure about the table structure, start simple:

```sql
SELECT * FROM DOCUMENTS WHERE DocName3 LIKE '%Kytkentälista%'
```

## How to Find the Correct Table Structure

### Option 1: Check in Access
1. Open your Access database
2. Look at the `_qryForExcel` query in Design View
3. See which tables it uses
4. Note the JOIN conditions
5. Copy that structure for the ODBC query

### Option 2: Query Information Schema
Temporarily use this to see all tables:
```sql
SELECT Name FROM MSysObjects WHERE Type=1 AND Flags=0
```

### Option 3: Test Each Table
Try these queries one by one to find the right table:

**Test 1: DOCUMENTS table**
```sql
SELECT * FROM DOCUMENTS
```

**Test 2: Check for DocName3 column**
```sql
SELECT DocNo, DocName, DocName1, DocName2, DocName3, Status, ProjNo 
FROM DOCUMENTS 
WHERE DocName3 LIKE '%Kytk%'
```

**Test 3: With ProjInfo JOIN**
```sql
SELECT d.*, p.* 
FROM DOCUMENTS d 
LEFT JOIN ProjInfo p ON d.ProjNo = p.ProjNo
WHERE d.DocName3 LIKE '%Kytk%'
```

## Columns Needed for Info Sheet

Based on your template, the query needs to return these columns:
- **Customer** (from ProjInfo table)
- **Mill** (from ProjInfo table)  
- **ContractNo** (from ProjInfo table)
- **Name** / ProjName (from ProjInfo table)
- **ProjNo** (from DOCUMENTS table)
- **DocName** / DocName1 / DocName2 / DocName3 (from DOCUMENTS table)
- **DocNo** (from DOCUMENTS table)
- **Date** (from DOCUMENTS table)
- **Status** (from Status table or DOCUMENTS table)
- **Rev** / RevID (from DOCUMENTS table)
- **RevDate** (from DOCUMENTS table)
- **Manager** (from ProjInfo table)

## Recommended Steps

1. **First, test the simple query:**
   ```sql
   SELECT * FROM DOCUMENTS WHERE DocName3 LIKE '%Kytkentälista%'
   ```

2. **If that works and returns data, check what columns you get**

3. **Then add the JOINs to get project info:**
   ```sql
   SELECT d.*, p.* 
   FROM DOCUMENTS d 
   LEFT JOIN ProjInfo p ON d.ProjNo = p.ProjNo
   WHERE d.DocName3 LIKE '%Kytkentälista%'
   ```

4. **Test in your workbook**

## Expected Column Names in DOCUMENTS Table

Common column names (check your actual database):
- `DocNo` or `MetsoDocNo`
- `DocName`, `DocName1`, `DocName2`, `DocName3`
- `ProjNo`
- `Status` or `StatusID`
- `Rev` or `RevID`
- `RevDate` or `Date`
- `File` or `FileName`

## Update Your Main Sheet

Replace the current DB2 query in Main sheet with:

```sql
SELECT d.*, p.Name as ProjName, p.Manager, p.Customer, p.Mill, p.Contract as ContractNo
FROM DOCUMENTS d 
LEFT JOIN ProjInfo p ON d.ProjNo = p.ProjNo
WHERE d.DocName3 LIKE '%Kytkentälista%'
```

Or start with the simple version first to test:

```sql
SELECT * FROM DOCUMENTS WHERE DocName3 LIKE '%Kytkentälista%'
```
