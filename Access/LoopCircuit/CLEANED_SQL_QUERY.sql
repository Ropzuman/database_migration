# Cleaned SQL Query for OLE DB

## Original _qryForExcel Query (With Problems)

```sql
SELECT 
  DOCUMENTS.Pick1, 
  Nz([ClientNo], [MetsoDocNo]) AS DocNo,  -- ❌ Nz() doesn't work in OLE DB
  IIf([DOCUMENTS]![DocName1] Like "Motor*",[DOCUMENTS]![DocName1] & " " & [DOCUMENTS]![DocName2],[DOCUMENTS]![DocName1]) AS DocName, 
  DOCUMENTS.DocName1, 
  DOCUMENTS.DocName2, 
  DOCUMENTS.DocName3, 
  [Status]![Description] AS Status, 
  Left([Rev],InStr([Rev],"/")-1) AS RevNum, 
  DOCUMENTS.Area AS Department, 
  DOCUMENTS.WorkPath, 
  DOCUMENTS.File, 
  Left([RevApu],InStr([RevAPU],"/")-1) AS RevDate, 
  Mid([rev],InStr([rev]," ")+1) AS RevAPU, 
  ProjInfo.ProjNo, 
  ProjInfo.ContractNo, 
  ProjInfo.Mill, 
  ProjInfo.Customer, 
  ProjInfo.Manager, 
  ProjInfo.Name AS Project, 
  IIf(IsNull([rev])," /",Mid([rev],InStr([rev]," ")+1)) AS APU, 
  IIf(IsNull([rev]),"",Left([rev],InStr([rev]," ")-1)) AS RevID, 
  Left([Apu],InStr([apu],"/")-1) AS [Date], 
  DOCUMENTS.Area, 
  DOCUMENTS.ClientNo, 
  DOCUMENTS.MetsoDocNo, 
  DOCUMENTS.Note, 
  DOCUMENTS.Rev
FROM (DOCUMENTS LEFT JOIN ProjInfo ON DOCUMENTS.pick1 = ProjInfo.Dept) 
INNER JOIN Status ON DOCUMENTS.Status = Status.Status;
```

## Fixed Query for OLE DB (Replace Nz with IIF)

```sql
SELECT 
  DOCUMENTS.Pick1, 
  IIF(ISNULL([ClientNo]), [MetsoDocNo], [ClientNo]) AS DocNo,
  IIF([DOCUMENTS].[DocName1] LIKE 'Motor*', [DOCUMENTS].[DocName1] & ' ' & [DOCUMENTS].[DocName2], [DOCUMENTS].[DocName1]) AS DocName, 
  DOCUMENTS.DocName1, 
  DOCUMENTS.DocName2, 
  DOCUMENTS.DocName3, 
  Status.Description AS Status, 
  LEFT([Rev], INSTR([Rev], '/') - 1) AS RevNum, 
  DOCUMENTS.Area AS Department, 
  DOCUMENTS.WorkPath, 
  DOCUMENTS.File, 
  LEFT([RevApu], INSTR([RevAPU], '/') - 1) AS RevDate, 
  MID([rev], INSTR([rev], ' ') + 1) AS RevAPU, 
  ProjInfo.ProjNo, 
  ProjInfo.ContractNo, 
  ProjInfo.Mill, 
  ProjInfo.Customer, 
  ProjInfo.Manager, 
  ProjInfo.Name AS Project, 
  IIF(ISNULL([rev]), ' /', MID([rev], INSTR([rev], ' ') + 1)) AS APU, 
  IIF(ISNULL([rev]), '', LEFT([rev], INSTR([rev], ' ') - 1)) AS RevID, 
  LEFT([Apu], INSTR([apu], '/') - 1) AS [Date], 
  DOCUMENTS.Area, 
  DOCUMENTS.ClientNo, 
  DOCUMENTS.MetsoDocNo, 
  DOCUMENTS.Note, 
  DOCUMENTS.Rev
FROM (DOCUMENTS LEFT JOIN ProjInfo ON DOCUMENTS.pick1 = ProjInfo.Dept) 
INNER JOIN Status ON DOCUMENTS.Status = Status.Status
WHERE DOCUMENTS.DocName3 LIKE '%Kytkentälista%';
```

## Key Changes Made

1. **Replaced `Nz()` with `IIF(ISNULL())`:**
   ```sql
   -- Before:
   Nz([ClientNo], [MetsoDocNo])
   
   -- After:
   IIF(ISNULL([ClientNo]), [MetsoDocNo], [ClientNo])
   ```

2. **Changed `!` to `.` notation:**
   ```sql
   -- Before: [DOCUMENTS]![DocName1]
   -- After:  [DOCUMENTS].[DocName1]
   ```
   The `!` operator is VBA/Access UI syntax, `.` is standard SQL

3. **Changed double quotes to single quotes:**
   ```sql
   -- Before: "Motor*"
   -- After:  'Motor*'
   ```
   Single quotes are standard SQL for string literals

4. **Added WHERE clause at the end:**
   ```sql
   WHERE DOCUMENTS.DocName3 LIKE '%Kytkentälista%'
   ```

## Simpler Alternative (RECOMMENDED)

If you don't need all the calculated fields right away, use this simpler version:

```sql
SELECT 
  DOCUMENTS.Pick1,
  DOCUMENTS.ClientNo,
  DOCUMENTS.MetsoDocNo,
  DOCUMENTS.DocName1, 
  DOCUMENTS.DocName2, 
  DOCUMENTS.DocName3,
  DOCUMENTS.Rev,
  DOCUMENTS.Area,
  DOCUMENTS.WorkPath, 
  DOCUMENTS.File,
  DOCUMENTS.Note,
  ProjInfo.ProjNo, 
  ProjInfo.ContractNo, 
  ProjInfo.Mill, 
  ProjInfo.Customer, 
  ProjInfo.Manager, 
  ProjInfo.Name,
  Status.Description AS Status
FROM (DOCUMENTS LEFT JOIN ProjInfo ON DOCUMENTS.Pick1 = ProjInfo.Dept) 
INNER JOIN Status ON DOCUMENTS.Status = Status.Status
WHERE DOCUMENTS.DocName3 LIKE '%Kytkentälista%';
```

**Why simpler is better:**
- ✅ All raw data is retrieved
- ✅ Module2.vba can do the calculations in VBA (more reliable)
- ✅ Fewer places for errors
- ✅ Easier to debug

Module2 already has logic to:
- Handle Null values (`& ""` converts Null to empty string)
- Parse revision data
- Extract dates and IDs
- Build composite fields

## How to Update Your Query

### Option 1: Update in Excel Main Sheet (EASIEST)

1. Open your Excel workbook
2. Go to **Main** sheet
3. Find the DB2 query cell (row 14, column C - check the pattern)
4. Paste the **simpler query** above
5. Save and click "Get Data"

### Option 2: Fix the Saved Query in Access

1. Open your Access database
2. Right-click `_qryForExcel` → Design View
3. Switch to SQL View
4. Replace with the **fixed query** above
5. Save the query
6. Keep your Excel query as: `SELECT * FROM _qryForExcel WHERE DocName3 LIKE '%Kytkentälista%'`

## Testing the New Query

After updating:

1. Click **"Get Data"**
2. Check **DB2 sheet** - should have columns:
   - Pick1, ClientNo, MetsoDocNo
   - DocName1, DocName2, DocName3
   - Rev, Area, WorkPath, File
   - ProjNo, ContractNo, Mill, Customer, Manager, Name
   - Status
3. Click **"Run Check"**
4. Verify **Info sheet** populates correctly

## Column Mapping Reference

Module2 expects these column names (case-insensitive):

| Column Name | Used For |
|-------------|----------|
| `rev` | Revision history |
| `clientno` | Document number |
| `metsodocno` | Metso document number |
| `docname1`, `docname2`, `docname3` | Document names |
| `status` | Document status |
| `workpath` | File path |
| `file` | Filename |
| `projno` | Project number |
| `customer` | Customer name |
| `mill` | Mill name |
| `manager` | Project manager |
| `name` | Project name |

All these columns are in the simpler query, so Module2 will work perfectly!

## Summary

**Problem:** `Nz()` function in saved query doesn't work with OLE DB

**Solution:** Use the simpler query that gets raw data and lets Module2 do the processing

**Benefit:** More reliable, easier to maintain, already tested in Module2
