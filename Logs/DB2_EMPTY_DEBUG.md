# DB2 Empty Issue - Debugging Steps

## Problem
DB2 sheet has headers but no data in row 2, causing Info sheet to remain unpopulated.

## Symptoms
1. DB2 shows headers in row 1
2. DB2 has NO data in row 2  
3. Info sheet shows all red asterisks (not populated)
4. User reports: "Running the code also seems to completely empty DB2 sheet"

## Possible Causes

### 1. DB2 SQL Query Returns No Rows
**Check:** Look at Main sheet, cell corresponding to DB2 SQL query (probably C13 or C14)
- Query might be too restrictive (wrong WHERE clause)
- Query might reference non-existent data
- Query syntax might be invalid

### 2. DB2 Gets Cleared During Checkout
**Evidence:** User says "Running the code completely empties DB2"
- Checkout calls HaeDocTiedot which reads DB2
- But doesn't clear it
- Need to check if there's accidental clearing happening

### 3. Worksheet Selection Issue
**Fixed in latest code:**
- Changed from `Worksheets("DB2").Select` + unqualified `Cells()`
- To: `Set wsDB2 = Worksheets("DB2")` + qualified `wsDB2.Cells()`
- This prevents reading from wrong sheet if selection changes

## Debugging Steps

### Step 1: Check if DB2 has data after "Get Data"
1. Click "Get Data" button
2. Immediately check DB2 sheet
3. **Question:** Does DB2 have a row 2 with data? Or just headers?

### Step 2: Check DB2 SQL Query
1. Open Main sheet
2. Look at the cell containing DB2 SQL query (likely C13 or C14)
3. **Check the query - what does it say?**

Expected query should be something like:
```sql
SELECT * FROM [_qryForExcel] WHERE DocName3 LIKE 'Kytkentälista'
```

### Step 3: Test if Checkout clears DB2
1. Click "Get Data" - verify DB2 has data
2. **DO NOT** click "Run Check" yet
3. Manually check DB2 - does it still have data?
4. Now click "Run Check"  
5. Check DB2 again - did it get cleared?

### Step 4: Check Immediate Window
After clicking "Run Check", press Ctrl+G in VBA editor to open Immediate Window.
Look for debug messages:
- `"HaeDocTiedot: WARNING - DB2 sheet has no data in row 2!"`
- `"Checkout: No data found in DB2 - Info sheet will be empty"`

## Fixes Applied

### Module2.vba - HaeDocTiedot Function
1. **Added worksheet object reference:**
   ```vba
   Dim wsDB2 As Worksheet
   Set wsDB2 = Worksheets("DB2")
   ```

2. **Changed all unqualified Cells() to wsDB2.Cells():**
   - Ensures we're always reading from DB2, not ActiveSheet
   
3. **Added debug output:**
   ```vba
   If wsDB2.Cells(2, 1).Value = "" Then
     Debug.Print "HaeDocTiedot: WARNING - DB2 sheet has no data in row 2!"
   End If
   ```

### Module2.vba - VaihdaInfo Function
- Error handling already added in previous fix

## Next Steps

**CRITICAL: We need to know:**
1. What is the SQL query in Main sheet for DB2?
2. Does DB2 have row 2 data immediately after clicking "Get Data"?
3. When exactly does DB2 get emptied?

Please run the debugging steps above and report findings.
