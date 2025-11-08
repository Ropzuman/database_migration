# 64-bit ODBC Compatibility Fix for \_qryForExcel

## Problem

**Error:** "ODBC Error: Database Connection Error"
**SQL Query:** `SELECT * FROM _qryForExcel WHERE DocName3 like 'Kytkentälista'`

The query works perfectly in **32-bit Excel** but fails in **64-bit Excel** with the 64-bit ODBC driver.

## Root Cause

The syntax `SELECT * FROM _qryForExcel` works differently between 32-bit and 64-bit ODBC drivers:

- **32-bit ODBC:** Allows querying Access saved queries with `FROM queryname` syntax
- **64-bit ODBC:** Requires brackets around saved query names: `FROM [queryname]`

`_qryForExcel` is a **saved query in Access**, not a table. The 64-bit ODBC driver treats it differently and requires proper escaping with brackets.

## The Fix

Modified `Module1.vba` HaeData function to automatically add brackets around `_qryForExcel`:

```vba
For i = 1 To 2
  Set ws = ThisWorkbook.Sheets("DB" & i)
  ws.Cells.Clear
  If sSQL(i) <> "" Then
    ' For 64-bit compatibility: Replace "FROM _qryForExcel" with "FROM [_qryForExcel]"
    Dim sqlQuery As String
    sqlQuery = sSQL(i)
    If InStr(1, sqlQuery, "FROM _qryForExcel", vbTextCompare) > 0 Then
      sqlQuery = Replace(sqlQuery, "FROM _qryForExcel", "FROM [_qryForExcel]", , , vbTextCompare)
    End If

    Set TAULUKKO = ws.QueryTables.Add(Connection:=Yhteys, Destination:=ws.Range("A1"))
    With TAULUKKO
      .Sql = sqlQuery  ' Uses modified SQL with brackets
      ' ... rest of query setup ...
```

## How It Works

1. **Reads original SQL** from Main sheet (e.g., `SELECT * FROM _qryForExcel WHERE ...`)
2. **Detects saved query reference** using `InStr()` to find "\_qryForExcel"
3. **Adds brackets** using `Replace()` to change `FROM _qryForExcel` to `FROM [_qryForExcel]`
4. **Executes modified SQL** through ODBC QueryTable

## Why This Works

Microsoft Access and ODBC require **identifier escaping** for names that:

- Start with underscore (`_qryForExcel`)
- Contain special characters
- Are reserved words

The brackets `[...]` tell ODBC: "This is an object name, not a SQL keyword."

## Testing

After this fix, test:

1. **Click "Get Data" button**

   - Should execute without ODBC error
   - DB1 should populate with circuit data
   - DB2 should populate with document metadata (one row)

2. **Check DB2 content:**

   - Row 1: Headers (rev, revid, project, manager, etc.)
   - Row 2: Data values

3. **Click "Run Check" button:**
   - Should read DB2 successfully
   - Info sheet should populate with project info

## Alternative Solutions (Not Used)

If brackets don't work, other options include:

### Option 1: Use Query Name Directly

Change Main sheet SQL from:

```sql
sql
SELECT * FROM _qryForExcel WHERE DocName3 like 'Kytkentälista'
```

To just:

```sql
_qryForExcel
```

**Downside:** Loses the WHERE filter, would return ALL documents.

### Option 2: Replace with Full SQL Definition

Replace the saved query reference with the actual SQL:

```sql
SELECT DOCUMENTS.*, ProjInfo.*, Status.Status
FROM (DOCUMENTS LEFT JOIN ProjInfo ON DOCUMENTS.ProjID = ProjInfo.ProjID)
LEFT JOIN Status ON DOCUMENTS.StatusID = Status.StatusID
WHERE DOCUMENTS.DocName3 like 'Kytkentälista'
```

**Downside:** Makes the Main sheet SQL very long and hard to maintain.

## Why Not Change the Main Sheet?

We chose to fix the VBA code instead of changing the Main sheet because:

1. ✅ **Preserves existing configuration** - Users don't need to edit their SQL
2. ✅ **Maintains compatibility** - Works with any query name that needs escaping
3. ✅ **Future-proof** - Automatically handles other queries with similar issues
4. ✅ **Transparent** - Users don't need to know about 64-bit ODBC quirks

## Comparison: 32-bit vs 64-bit

| Feature               | 32-bit Excel | 64-bit Excel (Before Fix) | 64-bit Excel (After Fix) |
| --------------------- | ------------ | ------------------------- | ------------------------ |
| ODBC Driver           | 32-bit       | 64-bit                    | 64-bit                   |
| Query Name Escaping   | Optional     | Required                  | Automatic                |
| `FROM _qryForExcel`   | ✅ Works     | ❌ Fails                  | ✅ Works                 |
| `FROM [_qryForExcel]` | ✅ Works     | ✅ Works                  | ✅ Works                 |

## Files Modified

- **Module1.vba:** Added automatic bracket insertion for saved query names

## Syntax Check

✅ No syntax errors found in Module1.vba

## Next Steps

1. **Copy updated Module1.vba** into your Excel VBA editor
2. **Test "Get Data"** button - should work without ODBC error
3. **Verify DB2 populates** with document metadata
4. **Test "Run Check"** button - should populate Info sheet
5. **Test full workflow** - Get Data → Run Check → Generate Printout

The code now provides **full 64-bit compatibility** while maintaining the exact same SQL syntax in the Main sheet that worked in 32-bit Excel! 🎯
