# Testauksen tarkistuslista / Testing Checklist — v2.0 (64-bit M365)

> **Luokitus / Classification:** `[ACTIVE]` — Kehittäjän testausohje
> **Versio / Version:** 2.0 — 64-bit M365 Compatible
> **Kohderyhmä / Audience:** Kehittäjät / Developers

Status: ✅ Valmis käytettäväksi 64-bit-ympäristössä / Ready for 64-bit environment

## Pre-Testing Setup

- [ ] Open the Excel workbook containing the Kytkentälista macros
- [ ] Ensure you have access to the Access database file
- [ ] Enable macros when prompted
- [ ] Open VBA Editor (Alt+F11) to check for compilation errors

## Compilation Check

- [ ] In VBA Editor: **Debug → Compile VBAProject**
- [ ] Expected: No errors
- [ ] If errors found: Document and report

## Test 1: Get Data (HaeData)

**Purpose:** Verify database connection and data retrieval work with optimized code

### Steps

1. [ ] Click the **"Get Data"** button on Main sheet
2. [ ] Wait for completion message

### Expected Results

- [ ] Message: "Data brought successfully!"
- [ ] DB1 sheet populated with Circuit_Diagrams_IO_Terminals data (should have ~1761 rows)
- [ ] DB2 sheet populated with DOCUMENTS data (should have ~2 rows)
- [ ] No ODBC errors
- [ ] No screen flickering (BeginFastMode working)

### If Errors

- [ ] Note error message
- [ ] Check if database file path is correct in Main sheet cell C6
- [ ] Verify SQL queries in Main sheet cells (rows 8-14)

## Test 2: Run Check (Checkout)

**Purpose:** Verify template validation and Info sheet population work correctly

### Steps

1. [ ] After successful "Get Data", click **"Run Check"** button
2. [ ] Wait for completion

### Expected Results

- [ ] Message: "Check OK!"
- [ ] **Info sheet** populated with document metadata (Customer, Mill, Project Name, Project No, Document ID, Status, Revision, Revision Date)
- [ ] **ERRORS sheet** is empty (no missing headers)
- [ ] **TEMPLATE sheet** has comments on header cells (linking annotations)
- [ ] No freezes or performance issues

### If Warnings/Errors

- [ ] If "WARNING: No document metadata found in DB2": Verify DB2 has data from Test 1
- [ ] If errors in ERRORS sheet: Document which headers are missing
- [ ] If Info sheet is empty: Check DB2 sheet cell comments match expected names

## Test 3: Generate Printout (GenPrintout)

**Purpose:** Verify printout generation works with optimized code

### Steps

1. [ ] After successful "Run Check", click **"Generate Printout"** button
2. [ ] When prompted for filename, enter a test filename
3. [ ] Save the file

### Expected Results

- [ ] New workbook created with sheets: Info, Revisions, [POSheet], Legend
- [ ] [POSheet] contains:
  - Header rows from TEMPLATE
  - Data from DB1 (should match DB1 content)
  - Alternating row colors (every RMAX rows)
  - Footer (if AddFooter option enabled)
- [ ] **Info sheet** has document metadata
- [ ] **Revisions sheet** has revision history
- [ ] **Legend sheet** copied correctly
- [ ] **LINKING sheet** hidden or deleted (based on HideLINKING setting)
- [ ] File saved successfully
- [ ] No performance issues (array-based copy working)

### If Errors

- [ ] Note error message
- [ ] Check if CheckOK was True (from Test 2)
- [ ] Verify TEMPLATE sheet has proper markers (&&PAGE_HEADER_START, etc.)

## Test 4: Error Handling

**Purpose:** Verify graceful error handling

### Test 4A: Missing Database

1. [ ] Temporarily rename the database file (add .bak extension)
2. [ ] Click "Get Data"
3. [ ] Expected: Error message "Database file not found: [path]"
4. [ ] Rename database file back to original name

### Test 4B: Empty DB2

1. [ ] Modify DB2 query in Main sheet to return no results (e.g., `WHERE 1=0`)
2. [ ] Click "Get Data"
3. [ ] Expected: Warning about DB2 returning no data
4. [ ] Restore original DB2 query

### Test 4C: Checkout Without Data

1. [ ] Clear DB1 and DB2 sheets manually
2. [ ] Click "Run Check"
3. [ ] Expected: Warning in ERRORS sheet about missing data
4. [ ] Run "Get Data" again to restore data

## Performance Check

Compare with previous version (if possible):

- [ ] Time "Get Data": _____ seconds (should be fast with BeginFastMode)
- [ ] Time "Run Check": _____ seconds
- [ ] Time "Generate Printout": _____ seconds
- [ ] Screen flickering: None/Minimal/Excessive
- [ ] Overall responsiveness: Good/Acceptable/Poor

## Code Quality Check (VBA Editor)

- [ ] No Debug.Print output in Immediate Window during normal operation
- [ ] Error messages are user-friendly (not technical Debug.Print messages)
- [ ] Comments are concise and explain logic (not verbose)

## Final Verification

- [ ] All tests passed
- [ ] No regressions from previous version
- [ ] Performance is same or better
- [ ] Code is cleaner and more maintainable

## Issues Found

Document any issues below:

### Issue 1

- Test:
- Description:
- Severity: Critical / Major / Minor
- Screenshot/Error message:

### Issue 2

- Test:
- Description:
- Severity: Critical / Major / Minor
- Screenshot/Error message:

## Sign-Off

- Tester: __________________
- Date: __________________
- Result: Pass ✅ / Fail ❌ / Conditional Pass ⚠️
- Notes:

---

## If All Tests Pass

Ready to merge `optimointi` branch to `main`:

```powershell
git checkout main
git merge optimointi
git push origin main
git tag -a v1.0 -m "64-bit compatible, optimized release"
git push origin v1.0
```

## If Issues Found

1. Document issues above
2. Create GitHub issues for tracking
3. Fix issues on `optimointi` branch
4. Re-run affected tests
5. Push fixes and re-test
