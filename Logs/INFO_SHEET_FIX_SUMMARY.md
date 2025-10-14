# Info Sheet Empty - Diagnostic Fix Applied

## Issue
Info sheet remains completely empty after running Checkout.

## Root Causes (Most Likely)

### 1. **Info Sheet Has No Comment Markers** ⚠️ HIGH PROBABILITY
The `VaihdaInfo()` function works by reading **comments** in Info sheet cells. Each cell that should display data MUST have a comment containing the field name.

**How it works:**
```vba
For i = 1 To .Comments.Count
  Select Case LCase(.Comments(i).text)
    Case "project"
      .Comments(i).Parent.Value = DIProject  ' Writes to the cell with this comment
```

**If Info sheet has 0 comments → 0 cells are populated → sheet stays empty!**

### 2. **DB2 Sheet is Empty** ⚠️ MEDIUM PROBABILITY
If user hasn't clicked "Get Data" button first, DB2 has no data, so all variables are empty strings.

## Diagnostic Code Added

### Added to HaeDocTiedot (Module2.vba)
```vba
' DEBUG: Report what was loaded
Debug.Print "HaeDocTiedot completed. Loaded " & (i - 1) & " columns from DB2"
Debug.Print "  DIProject: '" & DIProject & "'"
Debug.Print "  DIManager: '" & DIManager & "'"
Debug.Print "  DIDocNo: '" & DIDocNo & "'"
Debug.Print "  DIProjNo: '" & DIProjNo & "'"
```

### Added to VaihdaInfo (Module2.vba)
```vba
' DEBUG: Report sheet info
Debug.Print "VaihdaInfo: Processing sheet '" & Sheet & "' with " & ws.Comments.Count & " comments"
If ws.Comments.Count = 0 Then
  Debug.Print "  WARNING: No comments found in sheet - Info will remain empty!"
End If
```

### Added to Checkout (Module1.vba)
```vba
' Check if data was loaded from DB2
If DIProject = "" And DIDocNo = "" And DIProjNo = "" And DIMetsoDocNo = "" Then
  wsErrors.Range("A1").Value = "WARNING: No document metadata found in DB2 sheet!"
  wsErrors.Range("A2").Value = "Please click 'Get Data' button first to load data from database."
  wsErrors.Range("A1").Font.Bold = True
  wsErrors.Range("A1").Font.ColorIndex = 3 ' Red
  Debug.Print "Checkout: No data found in DB2 - Info sheet will be empty"
End If
```

## How to Use Diagnostics

### Step 1: Open Immediate Window
In VBA Editor: Press **Ctrl+G** to open Immediate Window

### Step 2: Run Checkout
Click "Run Check" button in Excel

### Step 3: Read Debug Output
Check Immediate Window for messages:

**Example Output (Problem #1 - No Comments):**
```
HaeDocTiedot completed. Loaded 15 columns from DB2
  DIProject: 'Mill Upgrade Project'
  DIManager: 'John Smith'
  DIDocNo: 'DOC-12345'
  DIProjNo: 'PROJ-001'
VaihdaInfo: Processing sheet 'Info' with 0 comments
  WARNING: No comments found in sheet - Info will remain empty!
```
→ **Solution:** Info sheet needs comment markers added

**Example Output (Problem #2 - No Data in DB2):**
```
HaeDocTiedot completed. Loaded 0 columns from DB2
  DIProject: ''
  DIManager: ''
  DIDocNo: ''
  DIProjNo: ''
Checkout: No data found in DB2 - Info sheet will be empty
VaihdaInfo: Processing sheet 'Info' with 15 comments
```
→ **Solution:** User must click "Get Data" button first

**Example Output (Both Working):**
```
HaeDocTiedot completed. Loaded 15 columns from DB2
  DIProject: 'Mill Upgrade Project'
  DIManager: 'John Smith'
  DIDocNo: 'DOC-12345'
  DIProjNo: 'PROJ-001'
VaihdaInfo: Processing sheet 'Info' with 15 comments
```
→ Info sheet should now have data!

## Solutions

### Solution A: Add Comment Markers to Info Sheet (If count = 0)

Each cell in Info sheet that should display data needs a comment:

**Required Comment Markers:**
| Comment Text | Data Displayed |
|--------------|----------------|
| `unit` | Metso unit name |
| `project` | Project name |
| `manager` | Project manager |
| `contractno` | Contract number |
| `projname` | Full project name |
| `projno` | Project number |
| `date` | Document date |
| `status` | Document status |
| `mill` | Mill name |
| `departname` | Department name |
| `customer` | Customer name |
| `docname` | Document name (main) |
| `docname1` | Document name line 1 |
| `docname2` | Document name line 2 |
| `docname3` | Document name line 3 |
| `metsodocno` | Metso document number |
| `rev` | Full revision history |
| `revid` | Current revision ID |
| `revdate` | Current revision date |

**How to Add Comments:**
1. Select cell where data should appear (e.g., B2 for project name)
2. Right-click → Insert Comment (or Shift+F2)
3. Type the comment text exactly as shown above (lowercase)
4. Click outside comment to save
5. Repeat for each field

**Quick Setup Script (Optional):**
If you want to automate comment creation, add this temporary sub:

```vba
Sub SetupInfoSheetComments()
  Dim ws As Worksheet
  Set ws = Sheets("Info")
  
  ' Example: Add comments to column B, rows 2-20
  On Error Resume Next
  ws.Range("B2").AddComment "unit"
  ws.Range("B3").AddComment "project"
  ws.Range("B4").AddComment "manager"
  ws.Range("B5").AddComment "contractno"
  ws.Range("B6").AddComment "projname"
  ws.Range("B7").AddComment "projno"
  ws.Range("B8").AddComment "date"
  ws.Range("B9").AddComment "status"
  ws.Range("B10").AddComment "mill"
  ws.Range("B11").AddComment "departname"
  ws.Range("B12").AddComment "customer"
  ws.Range("B13").AddComment "docname"
  ws.Range("B14").AddComment "metsodocno"
  ws.Range("B15").AddComment "rev"
  ws.Range("B16").AddComment "revid"
  ws.Range("B17").AddComment "revdate"
  On Error GoTo 0
  
  MsgBox "Comments added to Info sheet!", vbInformation
End Sub
```

### Solution B: Click "Get Data" First (If DB2 empty)

The correct workflow is:
1. **Get Data** button → Loads data from Access database to DB1 and DB2 sheets
2. **Run Check** button → Validates template and populates Info sheet
3. **Generate Printout** button → Creates formatted workbook

If user skips step 1, DB2 is empty, so Info sheet will be empty even with proper comments.

## Testing Steps

1. **Open VBA Editor** (Alt+F11)
2. **Open Immediate Window** (Ctrl+G)
3. **Click "Run Check" button** in Excel
4. **Read debug output** in Immediate Window
5. **Check ERRORS sheet** for warning messages

### Expected Results

**If Info sheet has no comments:**
- Immediate Window shows: `WARNING: No comments found in sheet`
- Fix: Add comment markers to Info sheet cells

**If DB2 is empty:**
- Immediate Window shows: `Loaded 0 columns from DB2` with empty variable values
- ERRORS sheet shows: `WARNING: No document metadata found`
- Fix: Click "Get Data" button first

**If both are correct:**
- Immediate Window shows data loaded and comment count > 0
- Info sheet should be populated with document metadata

## Removal of Debug Code (Optional)

Once the issue is identified and fixed, you can remove the debug statements:

1. Remove Debug.Print lines from HaeDocTiedot (lines ~105-108)
2. Remove Debug.Print lines from VaihdaInfo (lines ~119-123)
3. Keep the warning in Checkout (helps users remember to click Get Data first)

Or leave them in - they don't affect performance and help with future troubleshooting.

## Summary

✅ **Diagnostic code added** to identify root cause  
✅ **Warning message added** to ERRORS sheet if DB2 is empty  
✅ **Compile status:** No errors  
✅ **Next step:** Run Checkout and check Immediate Window output  

The debug output will tell you exactly what's wrong:
- Comments count = 0 → Need to add comment markers
- Data loaded = 0 → Need to run Get Data first
- Both correct → Different issue (report findings)
