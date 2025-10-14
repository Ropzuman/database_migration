# Info Sheet Population - Testing and Verification

## Status: DB2 is Now Populated ✅

Great! The DOCUMENTS table query is working.

## Next Step: Populate Info Sheet

The code to populate the Info sheet already exists in `Module2.vba`:
- `HaeDocTiedot()` - Reads data from DB2 into global variables
- `VaihdaInfo("Info")` - Writes variables to Info sheet based on cell comments

## How It Should Work

1. **Click "Run Check"** (Checkout button)
2. Checkout calls `HaeDocTiedot()` which reads DB2
3. Checkout calls `VaihdaInfo("Info")` which populates Info sheet
4. Info sheet cells with comments get filled with corresponding data

## Quick Test Procedure

### Test 1: Run the Diagnostic
1. Open VBA Editor (Alt+F11)
2. Press F5 to run a macro
3. Select `TestFullWorkflow`
4. Check the results

### Test 2: Manual Test
1. Make sure DB2 has data (click "Get Data" if needed)
2. Open VBA Editor (Alt+F11)
3. Press F5
4. Select `TestDB2Contents` - verify DB2 has data in row 2
5. Then run `TestHaeDocTiedotVariables` - see what variables were populated
6. Go back to Excel and check Info sheet

### Test 3: Normal Workflow
1. Click "Get Data" button
2. Verify DB2 has data
3. Click "Run Check" button  
4. Check Info sheet - should be populated!

## Troubleshooting

### If Info Sheet is Still Empty

**Possible Issue 1: Column Names Don't Match**

The `HaeDocTiedot()` function looks for specific column names in DB2:
- "customer"
- "mill"
- "project"
- "manager"
- "docname", "docname1", "docname2", "docname3"
- "status"
- "rev", "revid", "revdate"
- etc.

**Check:** Open DB2 sheet and look at row 1 (headers). Do the column names match what HaeDocTiedot expects?

**Solution:** If column names are different, we need to update the case statements in HaeDocTiedot.

### Possible Issue 2: Comment Names Don't Match

Info sheet cells need comments with specific text:
- "customer"
- "mill"
- "project"
- "manager"
- etc.

**Check:** Click on the cells in Info sheet - do they have comments? What do the comments say?

### Possible Issue 3: Variables Are Empty

**Test:** Run `TestHaeDocTiedotVariables` to see if variables are populated.

If variables are empty, it means:
- DB2 column names don't match the expected names
- Need to update HaeDocTiedot to use correct column names

## Next Actions

1. **Click "Run Check"** and see if Info sheet populates
2. **If it doesn't work**, run these diagnostics:
   - `TestDB2Contents` - shows DB2 data and column names
   - `TestHaeDocTiedotVariables` - shows what variables were set
3. **Share the output** from both diagnostics

Then we can:
- Match the actual DB2 column names to what HaeDocTiedot expects
- Update the code if needed
- Verify Info sheet cell comments are correct
