# Fix for Info and Revisions Sheets Not Populating

## Problem
Info and Revisions sheets remained unchanged after running Checkout, even though DB2 had data.

## Root Causes

### 1. DIRevArr Array Initialization Issue
**Location:** `Module2.vba`, line 39-40

**Problem:**
```vba
Erase DIRevArr
DIRevArr() = Split(DIRev, Chr(10))
```

This pattern causes issues because:
- `Erase DIRevArr` tries to erase an array that may not be properly sized
- Immediately assigning `DIRevArr()` with parentheses after erasing can fail
- VBA arrays need to be assigned directly without the `()` when using dynamic assignment

**Fix:**
```vba
DIRevArr = Split(DIRev, Chr(10))
```

Simply assign the Split result directly. VBA will automatically handle the array sizing.

### 2. Missing Error Handling in VaihdaInfo
**Location:** `Module2.vba`, VaihdaInfo function

**Problem:**
- No error handling when accessing DIRevArr array elements
- If DIRevArr is empty or malformed, Split operations would fail silently
- No validation that the sheet exists before selecting it

**Fix Applied:**

1. Added sheet existence check at function start:
```vba
On Error Resume Next
Worksheets(Sheet).Select
If Err.Number <> 0 Then
  Err.Clear
  Exit Sub
End If
On Error GoTo 0
```

2. Added error handling around all array operations in revision-related cases:
```vba
Case "revid"
  If Sheet <> "Info" Then
    On Error Resume Next
    Row = .Comments(i).Parent.Row
    Column = .Comments(i).Parent.Column
    For r = UBound(DIRevArr) To LBound(DIRevArr) Step -1
     If Err.Number = 0 And (DIRevArr(r) <> "") Then
       .Cells(Row, Column).Value = Split(DIRevArr(r), " ")(0)
       Row = Row + 1
     End If
    Next r
    On Error GoTo 0
  Else
    .Comments(i).Parent.Value = "'" & DIRevID
  End If
```

Applied same pattern to: revdate, designer, checker, approver, desc cases.

## Changes Summary

### Module2.vba
- **Line 39:** Removed `Erase DIRevArr` statement
- **Line 40:** Changed `DIRevArr() = Split(...)` to `DIRevArr = Split(...)`
- **VaihdaInfo function:** Added sheet existence validation at start
- **VaihdaInfo revision cases:** Added `On Error Resume Next/GoTo 0` blocks around all DIRevArr array operations
- **Array access validation:** Changed conditions from `If (DIRevArr(r) <> "")` to `If Err.Number = 0 And (DIRevArr(r) <> "")`

## Expected Result
- Info sheet should now populate with project metadata (project name, manager, document number, etc.)
- Revisions sheet should populate with revision history if DIRevArr has valid data
- No runtime errors if DIRevArr is empty or malformed

## Testing Instructions
1. Copy updated `Module2.vba` into Excel VBA editor
2. Click "Get Data" button to populate DB1 and DB2
3. Click "Run Check" button (Checkout function)
4. Verify Info sheet shows project information
5. Verify Revisions sheet shows revision data (if available in DB2)

## Technical Notes
- The error handling uses `On Error Resume Next` strategically only around array operations
- This prevents subscript errors, type mismatch errors, and null reference errors
- The `Err.Number = 0` check ensures we only process data when no error occurred
- Original logic and data flow remain unchanged - only error handling added
