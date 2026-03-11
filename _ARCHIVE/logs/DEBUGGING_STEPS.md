# STEP-BY-STEP DEBUGGING INSTRUCTIONS

## Problem
Info and Revisions sheets not populating - DB2 appears to have headers but no data.

## Diagnostic Process

### STEP 1: Import Diagnostic Module
1. Open your Excel file with macros
2. Press `Alt+F11` to open VBA Editor
3. Go to File → Import File
4. Import the file: `DiagnosticTest.vba`
5. You should now see a new module with diagnostic procedures

### STEP 2: Test DB2 After "Get Data"
1. **In Excel:** Click the "Get Data" button
2. **Wait for it to complete**
3. **In VBA Editor:** Press `Ctrl+G` to open Immediate Window (if not already open)
4. **In VBA Editor:** Press `F5` or go to Run → Run Sub/UserForm
5. **Select:** `TestDB2Contents` and click Run
6. **Check the Immediate Window** - it will show:
   - Whether row 2 has data
   - All column headers and their values
   - Last row number

**REPORT FINDING:** Does DB2 have data in row 2? Copy the Immediate Window output.

### STEP 3: Check What SQL Queries Are Configured
1. Go to the "Main" sheet in Excel
2. Look for cells that contain SQL queries (probably around C8-C14)
3. **Find the query for DB2** (likely the second query)
4. **Copy the exact SQL text**

**REPORT FINDING:** What is the SQL query for DB2?

### STEP 4: Test Complete Workflow
1. Make sure DB2 has data (click "Get Data" if needed)
2. **In VBA Editor:** Run the macro `TestFullWorkflow`
3. This will:
   - Show DB2 contents
   - Run HaeDocTiedot
   - Show what variables were populated
   - Run VaihdaInfo
4. **Check Immediate Window** for detailed output

**REPORT FINDING:** 
- Are the DI variables populated (DIProject, DIManager, etc.)?
- Does Info sheet get populated?

### STEP 5: Manual Inspection
1. After clicking "Get Data", **immediately go to DB2 sheet**
2. **Before clicking anything else**, check:
   - Does row 1 have headers?
   - Does row 2 have actual data values?
3. Take a screenshot if possible

### STEP 6: Check for DB2 Clearing
1. Click "Get Data" - verify DB2 has row 2 data
2. **WITHOUT clicking "Run Check"**, just switch sheets
3. Go back to DB2 - is the data still there?
4. Now click "Run Check"
5. Check DB2 again - is the data gone?

**REPORT FINDING:** When does DB2 data disappear?

## Possible Scenarios

### Scenario A: DB2 Never Gets Data
**Symptom:** After "Get Data", DB2 has only headers, no row 2
**Cause:** SQL query returns no rows or has error
**Solution:** Fix the SQL query in Main sheet

### Scenario B: DB2 Gets Cleared During Checkout
**Symptom:** DB2 has data after "Get Data", but disappears when clicking "Run Check"
**Cause:** Something in Checkout is clearing DB2
**Solution:** Need to prevent clearing or call HaeData again

### Scenario C: Reading Wrong Sheet
**Symptom:** DB2 has data but HaeDocTiedot reads empty values
**Cause:** Worksheet reference issue (FIXED in latest code)
**Solution:** Already fixed with wsDB2 object reference

## What To Report Back

Please run the diagnostic tests above and report:

1. **Output from TestDB2Contents** (copy from Immediate Window)
2. **SQL query text** from Main sheet for DB2
3. **When does DB2 data disappear?** (After Get Data? After Run Check? Never had data?)
4. **Output from TestFullWorkflow** (copy from Immediate Window)

This will help identify the exact problem!
