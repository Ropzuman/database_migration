# Info Sheet Empty - Diagnostic Report

## Issue
The Info sheet remains completely empty after running Checkout.

## Root Cause Analysis

### Flow Review
```
Checkout() 
  ↓
HaeDocTiedot() - Reads DB2 sheet, populates global variables (DIProject, DIManager, etc.)
  ↓
VaihdaInfo("Info") - Reads comments in Info sheet, fills cells with global variable values
```

### Possible Causes

#### 1. **DB2 Sheet is Empty** ⚠️
**Symptom:** If user hasn't run "Get Data" button first, DB2 sheet has no data.

**Result:** 
- `HaeDocTiedot()` runs but finds no data in DB2
- All DI* global variables remain empty strings ("")
- `VaihdaInfo()` writes empty strings to Info sheet cells

**Test:** Check if DB2 sheet has data in rows 1-2 with headers like "rev", "docno", "project", etc.

#### 2. **Info Sheet Has No Comments** ⚠️
**Symptom:** The Info sheet cells don't have comments with special markers.

**How VaihdaInfo Works:**
```vba
For i = 1 To .Comments.Count
  Select Case LCase(.Comments(i).text)
    Case "project"
      .Comments(i).Parent.Value = DIProject
    Case "manager"
      .Comments(i).Parent.Value = DIManager
    ' etc...
  End Select
Next i
```

**Required:** Each cell in Info sheet that should display data MUST have a comment containing the field name (e.g., "project", "manager", "docno").

**Test:** Open Info sheet and check if cells have comment indicators (red triangles in corners).

#### 3. **Case Sensitivity Issue** ⚠️
**Current Code:** `LCase(.Comments(i).text)` converts to lowercase
**Required:** Comments must match exactly: "project", "manager", "unit", etc. (all lowercase after conversion)

**Test:** Check if comments are spelled correctly and match the Case statements in VaihdaInfo.

#### 4. **Info Sheet Doesn't Exist** ⚠️
**Code Protection:**
```vba
On Error Resume Next
Set ws = Sheets("Info")
On Error GoTo 0

If ws Is Nothing Then Exit Sub
```

If Info sheet is missing or named differently (e.g., "INFO", "info", "Information"), the function silently exits.

**Test:** Verify sheet is named exactly "Info" (case-sensitive in VBA sheet names).

## Diagnostic Steps

### Step 1: Verify DB2 Has Data
1. Run "Get Data" button first
2. Check DB2 sheet - should have headers in row 1 and data in row 2
3. Expected headers: rev, revid, revdate, docno, metsodocno, project, status, docname, etc.

### Step 2: Verify Info Sheet Structure
1. Open Info sheet
2. Look for red comment indicators in cells
3. Right-click cell → "Show/Hide Comments" to see comment text
4. Verify comments contain field names like "project", "manager", "unit", etc.

### Step 3: Add Diagnostic Messages
Add temporary debug code to see what's happening:

#### In HaeDocTiedot (after loop):
```vba
' Safety check: prevent infinite loop (Excel max columns)
If i > 16384 Then Exit Do
Loop

' DEBUG: Show what was loaded
Debug.Print "HaeDocTiedot loaded:"
Debug.Print "  DIProject: " & DIProject
Debug.Print "  DIManager: " & DIManager
Debug.Print "  DIDocNo: " & DIDocNo
Debug.Print "  DIProjNo: " & DIProjNo
End Sub
```

#### In VaihdaInfo (at start):
```vba
If ws Is Nothing Then Exit Sub

' DEBUG: Show sheet info
Debug.Print "VaihdaInfo processing sheet: " & Sheet
Debug.Print "  Comment count: " & ws.Comments.Count
If ws.Comments.Count > 0 Then
  For i = 1 To Application.Min(5, ws.Comments.Count)
    Debug.Print "  Comment " & i & ": " & ws.Comments(i).text
  Next i
End If
```

### Step 4: Manual Test
Manually test if variables are populated:

1. Run Checkout
2. Open VBA Immediate Window (Ctrl+G)
3. Type: `? DIProject` and press Enter
4. Should display the project name (if data exists)
5. Type: `? DIManager` and press Enter
6. Should display the manager name

If these show empty strings, the problem is in `HaeDocTiedot` or DB2 data.
If these show data, the problem is in `VaihdaInfo` or Info sheet structure.

## Quick Fix Options

### Option A: Info Sheet Has No Comments (Most Likely)
**Solution:** Info sheet template needs to be set up with comment markers.

Each cell that should display data needs a comment:
1. Select cell where "Project" should appear
2. Insert → Comment (or Shift+F2)
3. Type exactly: `project` (lowercase)
4. Repeat for all fields

**Required Comments:**
- `unit` - Metso unit name
- `project` - Project name
- `manager` - Project manager
- `contractno` - Contract number
- `projname` - Full project name
- `projno` - Project number
- `date` - Document date
- `status` - Document status
- `mill` - Mill name
- `departname` - Department name
- `customer` - Customer name
- `docname` - Document name
- `metsodocno` - Metso document number
- `rev` - Revision history
- `revid` - Current revision ID
- `revdate` - Current revision date

### Option B: DB2 Has No Data
**Solution:** User must run "Get Data" button before "Run Check".

Add validation to Checkout:
```vba
' Fetch document info from DB2 sheet
HaeDocTiedot

' DEBUG: Check if data was loaded
If DIProject = "" And DIDocNo = "" And DIProjNo = "" Then
  MsgBox "No document data found in DB2 sheet!" & vbCrLf & vbCrLf & _
         "Please click 'Get Data' button first to load data from database.", _
         vbExclamation, "Missing Data"
  Application.ScreenUpdating = True
  Exit Sub
End If

VaihdaInfo   'Populate document info to Info sheet
```

### Option C: Create Fallback Approach
If comment-based approach is too fragile, create a simpler direct-write function:

```vba
Sub PopulateInfoDirect()
  Dim ws As Worksheet
  On Error Resume Next
  Set ws = Sheets("Info")
  On Error GoTo 0
  
  If ws Is Nothing Then Exit Sub
  
  ' Write data directly to fixed cells (adjust cell addresses as needed)
  ws.Range("B2").Value = "Metso Paper - " & DIMunit
  ws.Range("B3").Value = DIProject
  ws.Range("B4").Value = DIManager
  ws.Range("B5").Value = DIContract
  ws.Range("B6").Value = DIProjName
  ws.Range("B7").Value = DIProjNo
  ws.Range("B8").Value = DIDate
  ws.Range("B9").Value = DIStatus
  ws.Range("B10").Value = DIMill
  ws.Range("B11").Value = DIDepartName
  ws.Range("B12").Value = DICustomer
  ws.Range("B13").Value = DIDocName
  ws.Range("B14").Value = DIMetsoDocNo
  ws.Range("B15").Value = DIRev
  ws.Range("B16").Value = "'" & DIRevID
  ws.Range("B17").Value = DIRevDate
End Sub
```

Call this instead of or in addition to `VaihdaInfo`.

## Recommended Fix

Add validation and error messages to help user understand what's wrong:

```vba
Sub Checkout()
  ' ... existing code ...
  
  ' Fetch document info from DB2 sheet
  HaeDocTiedot
  
  ' Validate that data was loaded
  Dim dataLoaded As Boolean
  dataLoaded = (DIProject <> "" Or DIDocNo <> "" Or DIProjNo <> "")
  
  If Not dataLoaded Then
    wsErrors.Range("A1").Value = "WARNING: No document metadata found in DB2 sheet!"
    wsErrors.Range("A2").Value = "Please click 'Get Data' button first to load data from database."
    wsErrors.Range("A1").Font.Bold = True
    wsErrors.Range("A1").Font.ColorIndex = 3 ' Red
  End If
  
  ' Check if Info sheet has comments
  Dim ws As Worksheet
  Set ws = Sheets("Info")
  If ws.Comments.Count = 0 Then
    If wsErrors.Range("A1").Value = "" Then
      wsErrors.Range("A1").Value = "WARNING: Info sheet has no comment markers!"
    Else
      wsErrors.Range("A3").Value = "WARNING: Info sheet has no comment markers!"
    End If
    wsErrors.Range("A4").Value = "Info sheet cells need comments like 'project', 'manager', etc. to populate data."
    wsErrors.Cells(wsErrors.UsedRange.Rows.Count, 1).Font.ColorIndex = 3 ' Red
  End If
  
  VaihdaInfo   'Populate document info to Info sheet
  
  ' ... rest of existing code ...
End Sub
```

## Testing Plan

1. **Test with proper setup:**
   - Click "Get Data" 
   - Verify DB2 has data
   - Click "Run Check"
   - Check Info sheet

2. **Test without Get Data:**
   - Clear DB2 sheet
   - Click "Run Check"
   - Should see warning message

3. **Test with missing comments:**
   - Remove all comments from Info sheet
   - Click "Run Check"  
   - Should see warning message

4. **View Immediate Window:**
   - Press Ctrl+G in VBA editor
   - Run Checkout
   - Check Debug.Print output

## Expected Behavior

**Correct Flow:**
1. User clicks "Get Data" → DB2 populated with document metadata
2. User clicks "Run Check" → HaeDocTiedot reads DB2 → VaihdaInfo writes to Info sheet
3. Info sheet displays all document properties

**Current Issue:**
- Info sheet remains empty
- Most likely cause: Info sheet has no comment markers OR DB2 is empty

## Next Steps

1. Check if DB2 sheet has data (row 1 = headers, row 2 = values)
2. Check if Info sheet cells have comments
3. Add diagnostic Debug.Print statements
4. Run Checkout and check Immediate Window (Ctrl+G)
5. Report findings to determine exact cause
