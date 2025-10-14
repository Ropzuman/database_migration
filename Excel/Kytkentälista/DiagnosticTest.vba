'''
' DiagnosticTest.vba - Diagnostic procedures to debug DB2 empty issue
'''

Sub TestDB2Contents()
'''
' Run this after clicking "Get Data" to see what's in DB2
'''
Dim ws As Worksheet
Dim i As Long
Dim msg As String
Dim output As String

  Set ws = Worksheets("DB2")
  
  output = "=========================================" & vbCrLf
  output = output & "DB2 DIAGNOSTIC TEST - " & Now & vbCrLf
  output = output & "=========================================" & vbCrLf & vbCrLf
  
  ' Check row count
  Dim lastRow As Long
  lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
  output = output & "Last row with data: " & lastRow & vbCrLf & vbCrLf
  
  ' Check if row 2 has data
  If ws.Cells(2, 1).Value = "" Then
    output = output & "WARNING: Row 2, Column 1 is EMPTY!" & vbCrLf
  Else
    output = output & "Row 2, Column 1 contains: " & ws.Cells(2, 1).Value & vbCrLf
  End If
  
  ' List all headers in row 1
  output = output & vbCrLf & "Headers in Row 1:" & vbCrLf
  i = 1
  Do While ws.Cells(1, i).Value <> ""
    output = output & "  Col " & i & ": " & ws.Cells(1, i).Value & " = '" & ws.Cells(2, i).Value & "'" & vbCrLf
    i = i + 1
    If i > 50 Then Exit Do ' Safety - show first 50 columns
  Loop
  
  If i = 1 Then
    output = output & "ERROR: DB2 appears completely empty!" & vbCrLf
  Else
    output = output & vbCrLf & "Total columns with headers: " & (i - 1) & vbCrLf
  End If
  
  output = output & "========================================="
  
  ' Also print to debug
  Debug.Print output
  
  ' Show in message box (will be truncated if too long)
  MsgBox output, vbInformation, "DB2 Diagnostic Results"
  
End Sub

Sub TestHaeDocTiedotVariables()
'''
' Run this after HaeDocTiedot to see what variables were populated
'''
Dim output As String

  output = "=========================================" & vbCrLf
  output = output & "HAEDOCTIEDOT VARIABLES TEST" & vbCrLf
  output = output & "=========================================" & vbCrLf
  output = output & "DIProject: '" & DIProject & "'" & vbCrLf
  output = output & "DIManager: '" & DIManager & "'" & vbCrLf
  output = output & "DIDocNo: '" & DIDocNo & "'" & vbCrLf
  output = output & "DIMetsoDocNo: '" & DIMetsoDocNo & "'" & vbCrLf
  output = output & "DIProjNo: '" & DIProjNo & "'" & vbCrLf
  output = output & "DIProjName: '" & DIProjName & "'" & vbCrLf
  output = output & "DIContract: '" & DIContract & "'" & vbCrLf
  output = output & "DIDate: '" & DIDate & "'" & vbCrLf
  output = output & "DIStatus: '" & DIStatus & "'" & vbCrLf
  output = output & "DIMill: '" & DIMill & "'" & vbCrLf
  output = output & "DICustomer: '" & DICustomer & "'" & vbCrLf
  output = output & "DIMunit: '" & DIMunit & "'" & vbCrLf
  output = output & "DIRev: '" & DIRev & "'" & vbCrLf
  output = output & "DIRevID: '" & DIRevID & "'" & vbCrLf
  output = output & "DIRevDate: '" & DIRevDate & "'" & vbCrLf
  output = output & "========================================="
  
  Debug.Print output
  MsgBox output, vbInformation, "Variable Values"
End Sub

Sub TestFullWorkflow()
'''
' Test the complete workflow step by step
'''
  Debug.Print "========================================="
  Debug.Print "FULL WORKFLOW TEST - " & Now
  Debug.Print "========================================="
  
  ' Step 1: Check DB2 before
  Debug.Print ""
  Debug.Print "STEP 1: DB2 contents BEFORE HaeDocTiedot"
  TestDB2Contents
  
  ' Step 2: Run HaeDocTiedot
  Debug.Print ""
  Debug.Print "STEP 2: Running HaeDocTiedot..."
  HaeDocTiedot
  
  ' Step 3: Check variables after
  Debug.Print ""
  Debug.Print "STEP 3: Variables AFTER HaeDocTiedot"
  TestHaeDocTiedotVariables
  
  ' Step 4: Run VaihdaInfo
  Debug.Print ""
  Debug.Print "STEP 4: Running VaihdaInfo for Info sheet..."
  VaihdaInfo "Info"
  
  Debug.Print ""
  Debug.Print "STEP 5: Check Info sheet manually"
  Debug.Print "========================================="
  
  MsgBox "Test complete! Check Immediate Window (Ctrl+G) and Info sheet", vbInformation
End Sub
