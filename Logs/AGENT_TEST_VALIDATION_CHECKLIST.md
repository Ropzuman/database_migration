# Agent Test Branch - Validation Checklist

**Branch:** `agent_test`  
**Purpose:** Validate 64-bit compliance fixes and performance optimizations  
**Modified Files:** 29 files across 6 folders  
**Test Date:** _____________  
**Tester:** _____________

---

## Pre-Testing Setup

### 1. Environment Verification

- [ ] Windows 10/11 64-bit OS
- [ ] Microsoft Office 365 (64-bit) installed
- [ ] Access database: `P:\acaddata\projekti\[project]\tyo\` (network drive accessible)
- [ ] AutoCAD 2019 (64-bit) installed and licensed
- [ ] Git branch checked out: `git checkout agent_test`

### 2. Backup Current State

```powershell
# Create backup before testing
Copy-Item "P:\acaddata\projekti\[project]\tyo\*.accdb" "P:\acaddata\projekti\[project]\tyo\backup_$(Get-Date -Format 'yyyyMMdd')\"
```

### 3. Database Setup

- [ ] Open Access database (.accdb file)
- [ ] Enable VBA macros (Trust Center → Enable all macros for testing)
- [ ] Open VBA Editor (`Alt+F11`)
- [ ] Verify references:
  - [ ] Microsoft DAO 3.6 Object Library ✅
  - [ ] AutoCAD 2019 Type Library ✅
  - [ ] Microsoft Scripting Runtime ✅

---

## PHASE 1: Compilation Verification

### VBA Compilation Test

**Purpose:** Verify all code compiles without errors

**Steps:**

1. Open VBA Editor (`Alt+F11`)
2. Navigate to each folder module/form
3. Menu: Debug → Compile VBA Project
4. Record any errors

#### Compilation Results

| Folder | Status | Errors | Notes |
|--------|--------|--------|-------|
| **PIPE** | ☐ Pass / ☐ Fail | | |
| **Function_descriptions_html** | ☐ Pass / ☐ Fail | | |
| **LoopCircuit** | ☐ Pass / ☐ Fail | | |
| **DOCUMENTS** | ☐ Pass / ☐ Fail | | |
| **instru3** | ☐ Pass / ☐ Fail | | |
| **MAINEQ** | ☐ Pass / ☐ Fail | | |

**Expected Result:** ✅ All folders compile without errors

**If Errors Found:**

- Note exact error message and line number
- Check if DAO prefix is correctly applied
- Verify String$ function syntax

---

## PHASE 2: Critical File Testing

### TEST 2.1: DataToACAD.bas (MAINEQ) ⭐ CRITICAL

**File:** `Access/MAINEQ/DataToACAD.bas`  
**Changes:** 15 string function optimizations (UCase$/Mid$/Left$/Right$)  
**Impact:** AutoCAD LISP attribute generation

**Test Procedure:**

1. Open MAINEQ database
2. Ensure devTbl* tables have test data
3. Run `DataToACAD` function
4. Verify output:
   - [ ] `.txt` files generated in output directory
   - [ ] File format: `[tablename].txt` (8 characters max)
   - [ ] LISP format: `( "ATTRIBUTE.NAME" "value" )`
   - [ ] UCase conversion correct (all attribute names uppercase)
   - [ ] No empty/corrupt files

**Sample LISP Output Verification:**

```lisp
( "ZS-001.DEVTBLINSTRUMENTS.TAG" "FT-101" )
( "ZS-001.DEVTBLINSTRUMENTS.DESCRIPTION" "Flow Transmitter" )
```

**Test Data:**

- [ ] AreaCode: ZS / LoopNo: 001
- [ ] Device table: devTblInstruments (at least 3 records)
- [ ] Verify: Each field generates correct LISP line

**Result:** ☐ Pass / ☐ Fail  
**Notes:**

```


```

---

### TEST 2.2: GlobalVBAs.vba (DOCUMENTS)

**File:** `Access/DOCUMENTS/GlobalVBAs.vba`  
**Changes:** 19 string function optimizations in revision parsing  
**Impact:** Revision notation parsing

**Test Functions:**

#### HaeTekija (Extract Original Author)

```vba
' Test in Immediate Window (Ctrl+G):
? HaeTekija("A 01.02.2025/JohnDoe/JaneSmith/Manager/Initial design" & vbCrLf & "0 15.01.2025/OldAuthor/OldChecker/OldApprover/Preliminary")
' Expected: "OldAuthor"
```

- [ ] Returns correct author from first (oldest) revision
- [ ] Handles Null input (returns "")

#### HaeRevisioija (Extract Latest Author)

```vba
? HaeRevisioija("A 01.02.2025/JohnDoe/JaneSmith/Manager/Updated")
' Expected: "JohnDoe"
```

- [ ] Returns correct author from latest revision
- [ ] Returns "" if only one revision

#### HaeRevisio (Extract Revision Mark)

```vba
? HaeRevisio("B 01.03.2025/Author/Checker/Approver/Description")
' Expected: "B"
```

- [ ] Extracts revision mark correctly
- [ ] Handles Null input

#### HaeViimPaiva (Extract Original Date)

```vba
? HaeViimPaiva("A 01.02.2025/JohnDoe/JaneSmith/Manager/Updated" & vbCrLf & "0 15.01.2025/OrigAuthor/OrigChecker/OrigApprover/Initial")
' Expected: "15.01.2025"
```

- [ ] Returns date from oldest revision
- [ ] Parses multi-line revision strings correctly

**Result:** ☐ Pass / ☐ Fail  
**Notes:**

```


```

---

### TEST 2.3: Form_CopyLoops.cls (instru3)

**File:** `Access/instru3/Form_CopyLoops.cls`  
**Changes:** 8 string function optimizations in database copy loops  
**Impact:** Loop copying between databases

**Test Procedure:**

1. Open instru3 database
2. Open "Form_CopyLoops" form
3. Select source database with test loops
4. Select loops to copy (e.g., ZS-001, ZS-002)
5. Execute copy operation

**Verification:**

- [ ] Source database selection dialog works
- [ ] Loop list populates correctly
- [ ] Table name filtering works (`devTbl*` found, `devTblCommon` excluded)
- [ ] Records copied to devTable
- [ ] TableName field populated correctly (LCase$ comparison works)
- [ ] No error messages during copy
- [ ] Progress status updates displayed

**Test Data:**

- Source DB: At least 2 loops with devices
- Target DB: Empty or with different loops
- Expected: All device records for selected loops copied

**Result:** ☐ Pass / ☐ Fail  
**Notes:**

```


```

---

### TEST 2.4: Form_Linkkien vaihto.cls (instru3)

**File:** `Access/instru3/Form_Linkkien vaihto.cls`  
**Changes:** 5 optimizations in linked table path handling  
**Impact:** Linked table path rewriting

**Test Procedure:**

1. Open instru3 database with linked tables
2. Run "Linkkien vaihto" (Change Links) function
3. Verify:
   - [ ] Current database path extracted correctly (Left$)
   - [ ] Linked table paths parsed (Mid$, Left$)
   - [ ] Path comparison works (LCase$)
   - [ ] Tables with different paths are relinked
   - [ ] Tables in same directory are skipped

**Expected Behavior:**

- Path extraction: `"P:\path\to\database.accdb"` → `"P:\path\to\"`
- Comparison: Case-insensitive path matching
- Relinking: Only tables with wrong path

**Result:** ☐ Pass / ☐ Fail  
**Notes:**

```


```

---

### TEST 2.5: Form_SizingOut.cls (instru3)

**File:** `Access/instru3/Form_SizingOut.cls`  
**Changes:** 3 optimizations in CSV export  
**Impact:** Instrument sizing data export

**Test Procedure:**

1. Open form "SizingOut"
2. Select output directory
3. Export table data to CSV
4. Verify:
   - [ ] Path validation works (Right$)
   - [ ] Decimal comma → period conversion (Left$, Mid$)
   - [ ] CSV format correct: `"field1";"field2";"field3"`
   - [ ] Finnish decimals converted: `"3,14"` → `"3.14"`

**Test Data:**

- Table with decimal fields (e.g., flow rate: 12,5)
- Expected: CSV output has periods: `"12.5"`

**Result:** ☐ Pass / ☐ Fail  
**Notes:**

```


```

---

### TEST 2.6: Form_USysRevText.cls (Multiple Folders)

**Files:**

- `Access/DOCUMENTS/Form_USysRevText.cls` (14 optimizations)
- `Access/MAINEQ/Form_USysRevText.cls` (14 optimizations)

**Changes:** Identical parsing logic in both files  
**Impact:** Revision text form UI

**Test Procedure:**

1. Open form (available in both DOCUMENTS and MAINEQ)
2. Load existing revision data OR add new revision
3. Verify:
   - [ ] Form opens without errors
   - [ ] Revision list displays correctly
   - [ ] Clicking revision populates fields (Mid$, Left$ parsing)
   - [ ] "Add New" increments revision correctly
   - [ ] Date, Author, Checker, Approver fields parse correctly
   - [ ] Description field accepts multi-line text

**Test Revision String:**

```
A 01.02.2025/JohnDoe/JaneSmith/Manager/Updated design parameters
0 15.01.2025/OrigAuthor/OrigChecker/OrigApprover/Initial release
```

**Expected Field Values:**

- Rev: `A`
- Date: `01.02.2025`
- Drawn: `JohnDoe`
- Checked: `JaneSmith`
- Approved: `Manager`
- Description: `Updated design parameters`

**Result (DOCUMENTS):** ☐ Pass / ☐ Fail  
**Result (MAINEQ):** ☐ Pass / ☐ Fail  
**Notes:**

```


```

---

## PHASE 3: DAO Prefix Validation

### TEST 3.1: Database Object Creation

**Purpose:** Verify DAO.Database/DAO.Recordset work correctly

**Test in VBA Immediate Window:**

```vba
Dim DB As DAO.Database
Dim RS As DAO.Recordset
Set DB = CurrentDb
Set RS = DB.OpenRecordset("SELECT * FROM [any_table]", dbOpenDynaset)
? RS.RecordCount
RS.Close
Set RS = Nothing
Set DB = Nothing
```

**Result:** ☐ Pass / ☐ Fail  
**Notes:**

```


```

---

### TEST 3.2: Modified Files DAO Compliance

Test each file that was modified for DAO prefixes:

| File | DAO Objects Used | Test Result | Notes |
|------|------------------|-------------|-------|
| Report_TRANSMITTAL.cls | DAO.Recordset | ☐ Pass / ☐ Fail | |
| Form_USysExcelReport.cls | DAO.Database, 2× DAO.Recordset | ☐ Pass / ☐ Fail | |
| Form_USysAddDocument.cls | DAO.Database, 2× DAO.Recordset | ☐ Pass / ☐ Fail | |
| Form_USysNewDistribution.cls | DAO.Database, 2× DAO.Recordset | ☐ Pass / ☐ Fail | |
| Form_USysAddedDistr.cls | DAO.Database, DAO.Recordset | ☐ Pass / ☐ Fail | |
| instru3/Form_DBUsers.cls | DAO.Database | ☐ Pass / ☐ Fail | |
| MAINEQ/DataToACAD.bas | DAO.TableDef | ☐ Pass / ☐ Fail | |

**Test Method:**

1. Open each form/run each procedure
2. Verify no "Type mismatch" or "Object required" errors
3. Confirm database operations complete successfully

---

## PHASE 4: Form Functionality Testing

### DOCUMENTS Folder Forms

#### Form_DISTRIBUTION.cls

- [ ] Opens without error
- [ ] Excel report generation works
- [ ] Path extraction works (Left$)

#### Form_USysOpenFile.cls

- [ ] File selection dialog works
- [ ] File list parsing correct (Left$, Mid$)
- [ ] Multiple file selection works

#### Form_DBUsers.cls

- [ ] Lock file path created correctly (Left$)
- [ ] Active users displayed

### instru3 Folder Forms

#### Form_DBUsers.cls

- [ ] Lock file detection works (Left$ optimization)
- [ ] User list accurate

### MAINEQ Folder Forms

#### Form_Revisiointi.cls

- [ ] Revision dropdown filters correctly (Left$, LCase$)
- [ ] Revision incrementing works

#### Form_GeneroiMoottorikuvat.cls

- [ ] Motor drawing generation works
- [ ] Case conversion correct (UCase$)

---

## PHASE 5: AutoCAD Integration Testing

### TEST 5.1: AutoCAD Connection

**Prerequisites:**

- AutoCAD 2019 running
- Test drawing open

**Test Code:**

```vba
Sub TestAutoCADConnection()
    Dim acadApp As Object
    Set acadApp = GetObject(, "AutoCAD.Application")
    Debug.Print "AutoCAD Version: " & acadApp.Version
    Debug.Print "Active Document: " & acadApp.ActiveDocument.Name
    Set acadApp = Nothing
End Sub
```

- [ ] Connection established
- [ ] Version displayed correctly

**Result:** ☐ Pass / ☐ Fail

---

### TEST 5.2: Attribute Writing (DataToACAD.bas)

**Full Integration Test:**

1. **Setup:**
   - Open MAINEQ database
   - Ensure devTbl tables have test data
   - AutoCAD drawing with blocks ready

2. **Execute:**
   - Run DataToACAD.bas main procedure
   - Monitor LISP file generation

3. **Verify in AutoCAD:**
   - Load generated LISP files: `(load "path/to/file.lsp")`
   - Check attributes populated correctly
   - Verify UCase$ conversion (all attribute names uppercase)
   - Confirm no empty/null attributes

**Critical Checks:**

- [ ] LISP files created
- [ ] Attribute names uppercase (UCase$)
- [ ] Attribute values correct
- [ ] No AutoCAD errors

**Result:** ☐ Pass / ☐ Fail  
**Notes:**

```


```

---

## PHASE 6: Performance Verification

### String Function Performance Test

**Test Code:**

```vba
Sub TestStringPerformance()
    Dim i As Long
    Dim startTime As Double
    Dim testString As String
    Dim result As String
    
    testString = "ABCDEFGHIJKLMNOPQRSTUVWXYZ"
    
    ' Test Left$ performance
    startTime = Timer
    For i = 1 To 100000
        result = Left$(testString, 5)
    Next i
    Debug.Print "Left$ time: " & (Timer - startTime) & " seconds"
    
    ' Compare with old Left (manual test on main branch)
    ' Expected: Left$ should be faster (no Variant conversion)
End Sub
```

**Expected Results:**

- Left$/Mid$/Right$ should show measurable performance improvement
- Especially noticeable in DataToACAD.bas with large recordsets

**Result:** ☐ Pass / ☐ Fail  
**Performance Gain:** _______ % faster

---

## PHASE 7: Regression Testing

### Comparison Test: agent_test vs main

**Setup:**

1. Checkout main branch: `git checkout main`
2. Run critical procedures, record results
3. Checkout agent_test: `git checkout agent_test`
4. Run same procedures, compare results

**Critical Procedures to Compare:**

| Procedure | main Output | agent_test Output | Match? |
|-----------|-------------|-------------------|--------|
| DataToACAD (MAINEQ) | | | ☐ Yes / ☐ No |
| HaeTekija("test string") | | | ☐ Yes / ☐ No |
| Form_CopyLoops (copy 1 loop) | | | ☐ Yes / ☐ No |
| Form_SizingOut (export CSV) | | | ☐ Yes / ☐ No |

**Expected:** ✅ Identical functional output (only performance difference)

**Result:** ☐ Pass / ☐ Fail

---

## PHASE 8: Error Handling Verification

### Error Recovery Test

**Test Scenarios:**

1. **Null Input Handling:**

```vba
? HaeTekija(Null)  ' Should return ""
? HaeRevisio(Null) ' Should return ""
```

- [ ] No runtime errors
- [ ] Returns empty string

1. **Invalid Database Path:**

- Open Form_CopyLoops
- Select non-existent database
- [ ] Error handler catches it
- [ ] User-friendly error message

1. **Missing AutoCAD:**

- Close AutoCAD
- Run DataToACAD
- [ ] Graceful error handling OR clear error message

**Result:** ☐ Pass / ☐ Fail

---

## PHASE 9: Code Quality Verification

### Code Inspection Checklist

- [ ] All `Dim X As Database` now `Dim X As DAO.Database`
- [ ] All `Dim X As Recordset` now `Dim X As DAO.Recordset`
- [ ] All `Dim X As TableDef` now `Dim X As DAO.TableDef`
- [ ] String functions use $ where appropriate
- [ ] No logic changes (only performance optimizations)
- [ ] Comments preserved
- [ ] Indentation maintained

**Result:** ☐ Pass / ☐ Fail

---

## PHASE 10: Final Validation

### Pre-Merge Checklist

- [ ] All compilation tests pass
- [ ] All critical file tests pass
- [ ] DAO prefix validation complete
- [ ] Form functionality verified
- [ ] AutoCAD integration working
- [ ] Performance improvement confirmed
- [ ] Regression tests show identical output
- [ ] Error handling works correctly
- [ ] Code quality inspection complete

### Sign-Off

**Test Summary:**

- Total Tests: _______
- Passed: _______
- Failed: _______
- Pass Rate: _______ %

**Critical Issues Found:**

```


```

**Recommendation:**

- [ ] ✅ APPROVED - Merge to main
- [ ] ⚠️ CONDITIONAL - Fix issues before merge
- [ ] ❌ REJECTED - Major problems, needs rework

**Tester Signature:** _____________  
**Date:** _____________  
**Approver Signature:** _____________  
**Date:** _____________

---

## Appendix: Quick Test Script

**PowerShell Quick Test:**

```powershell
# Quick compilation check for all modules
cd "c:\database_migration"
git checkout agent_test

# Open Access and run compilation
# (Requires Access 64-bit installed)
# This is a manual step - open Access, press Alt+F11, Debug > Compile

Write-Host "Agent Test Branch - Quick Validation"
Write-Host "======================================"
Write-Host "1. Checkout agent_test branch: DONE"
Write-Host "2. Open Access database"
Write-Host "3. Press Alt+F11 for VBA Editor"
Write-Host "4. Debug > Compile VBA Project"
Write-Host "5. Check for errors"
Write-Host ""
Write-Host "Expected: No compilation errors"
```

---

## Testing Notes

**Environment Details:**

- OS Version: _____________
- Office Version: _____________
- Access Build: _____________
- AutoCAD Version: _____________

**Additional Notes:**

```




```

---

**Document Version:** 1.0  
**Created:** February 26, 2026  
**Purpose:** Validate agent_test branch before merging to main
