'''
' ColumnMappingDiagnostic.vba - Check if DB2 columns match expected names
'''

Sub CheckDB2ColumnNames()
'''
' This shows what columns DB2 has vs what HaeDocTiedot expects
'''
Dim ws As Worksheet
Dim i As Long
Dim output As String
Dim expectedColumns As String

  Set ws = Worksheets("DB2")
  
  output = "=========================================" & vbCrLf
  output = output & "DB2 COLUMN NAME CHECK" & vbCrLf
  output = output & "=========================================" & vbCrLf & vbCrLf
  
  output = output & "EXPECTED COLUMN NAMES (what HaeDocTiedot looks for):" & vbCrLf
  output = output & "  - customer" & vbCrLf
  output = output & "  - mill" & vbCrLf
  output = output & "  - project" & vbCrLf
  output = output & "  - manager" & vbCrLf
  output = output & "  - contractno" & vbCrLf
  output = output & "  - projno" & vbCrLf
  output = output & "  - name (for project name)" & vbCrLf
  output = output & "  - docname, docname1, docname2, docname3" & vbCrLf
  output = output & "  - docno" & vbCrLf
  output = output & "  - metsodocno" & vbCrLf
  output = output & "  - status" & vbCrLf
  output = output & "  - date" & vbCrLf
  output = output & "  - rev" & vbCrLf
  output = output & "  - revid" & vbCrLf
  output = output & "  - revdate" & vbCrLf
  output = output & "  - metsounitname" & vbCrLf
  output = output & "  - departname" & vbCrLf
  output = output & "  - workpath" & vbCrLf
  output = output & "  - file" & vbCrLf
  output = output & vbCrLf
  
  output = output & "ACTUAL COLUMNS IN DB2 (first 30):" & vbCrLf
  i = 1
  Do While ws.Cells(1, i).Value <> "" And i <= 30
    output = output & "  " & i & ". " & LCase(ws.Cells(1, i).Value)
    If ws.Cells(2, i).Value <> "" Then
      output = output & " = '" & Left(ws.Cells(2, i).Value, 50) & "'"
    Else
      output = output & " = (empty)"
    End If
    output = output & vbCrLf
    i = i + 1
  Loop
  
  output = output & vbCrLf & "=========================================" & vbCrLf
  output = output & "MATCHING CHECK:" & vbCrLf
  
  ' Check for key columns
  Dim foundCustomer As Boolean, foundProject As Boolean, foundManager As Boolean
  Dim foundDocNo As Boolean, foundProjNo As Boolean
  
  i = 1
  Do While ws.Cells(1, i).Value <> ""
    Select Case LCase(ws.Cells(1, i).Value)
      Case "customer": foundCustomer = True
      Case "project": foundProject = True
      Case "manager": foundManager = True
      Case "docno": foundDocNo = True
      Case "projno": foundProjNo = True
    End Select
    i = i + 1
    If i > 100 Then Exit Do
  Loop
  
  output = output & "  customer: " & IIf(foundCustomer, "FOUND ✓", "MISSING ✗") & vbCrLf
  output = output & "  project: " & IIf(foundProject, "FOUND ✓", "MISSING ✗") & vbCrLf
  output = output & "  manager: " & IIf(foundManager, "FOUND ✓", "MISSING ✗") & vbCrLf
  output = output & "  docno: " & IIf(foundDocNo, "FOUND ✓", "MISSING ✗") & vbCrLf
  output = output & "  projno: " & IIf(foundProjNo, "FOUND ✓", "MISSING ✗") & vbCrLf
  output = output & "========================================="
  
  Debug.Print output
  MsgBox output, vbInformation, "Column Name Check"
  
End Sub

Sub QuickInfoSheetTest()
'''
' Quick test: Read DB2, populate variables, try to fill Info sheet
'''
Dim output As String

  output = "QUICK INFO SHEET TEST" & vbCrLf & vbCrLf
  
  ' Step 1: Read DB2
  output = output & "Step 1: Reading DB2..." & vbCrLf
  HaeDocTiedot
  
  ' Step 2: Check what we got
  output = output & "Step 2: Variables populated:" & vbCrLf
  output = output & "  DIProject: '" & DIProject & "'" & vbCrLf
  output = output & "  DIManager: '" & DIManager & "'" & vbCrLf
  output = output & "  DICustomer: '" & DICustomer & "'" & vbCrLf
  output = output & "  DIProjNo: '" & DIProjNo & "'" & vbCrLf
  output = output & "  DIDocNo: '" & DIDocNo & "'" & vbCrLf
  output = output & vbCrLf
  
  ' Step 3: Populate Info sheet
  output = output & "Step 3: Populating Info sheet..." & vbCrLf
  VaihdaInfo "Info"
  
  output = output & "Step 4: Done! Check Info sheet now." & vbCrLf
  
  Debug.Print output
  MsgBox output, vbInformation, "Test Complete"
  
  ' Go to Info sheet so user can see result
  Sheets("Info").Select
  
End Sub
