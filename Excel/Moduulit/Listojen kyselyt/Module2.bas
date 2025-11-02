'''
' Module2.vba - Metadata, info, and linking logic for Kytkentälista Excel macro system
' Handles document property extraction, comment-based linking, and error reporting.
' Notes:
' - HaeDocTiedot reads DB2 headers case-insensitively and trims whitespace.
' - WorkPath is recognized via multiple synonyms (workpath, path, work_path, listpath, lists_path, zlistspath, savepath, targetpath, outputpath)
'   and normalized to use backslashes with a guaranteed trailing backslash.
' - File is taken from DB2 (file/filename/file_name) and used as the default Save As name by GenPrintout.
'''

Sub HaeDocTiedot()
'''
' HaeDocTiedot: Extracts document properties from DB2 sheet and stores them in global variables.
' Used for populating headers, footers, and info fields in the printout.
' Optimized: Removed Select/Activate, uses direct worksheet references.
'''
Dim i As Long
Dim Arvo As String
Dim wsDB2 As Worksheet

  ' Initialize all document info variables
  DIRev = ""
  DIRevID = ""
  DIRevDate = ""
  DIDocNo = ""
  DIMetsoDocNo = ""
  DIProject = ""
  DIStatus = ""
  DIDocName = ""
  DIDocName1 = ""
  DIDocName2 = ""
  DIDocName3 = ""
  DIContract = ""
  DIProjNo = ""
  DIProjName = ""
  DIPath = ""
  DIDate = ""
  DIManager = ""
  DIMunit = ""
  DIMill = ""
  DIDepartName = ""
  DICustomer = ""
  DIFile = ""

  On Error Resume Next
  Set wsDB2 = Sheets("DB2")
  On Error GoTo 0
  
  If wsDB2 Is Nothing Then Exit Sub
  
  i = 1
  Do
    ' Normalize header: lowercase and trim to tolerate extra spaces
    Arvo = LCase(Trim(wsDB2.Cells(1, i).Value & ""))
    Select Case Arvo
      Case "rev"
        DIRev = wsDB2.Cells(2, i).Value
        Erase DIRevArr
        DIRevArr() = Split(DIRev, Chr(10))
      Case "revid"
        DIRevID = wsDB2.Cells(2, i).Value
      Case "revdate"
        DIRevDate = wsDB2.Cells(2, i).Value
      Case "date", "dateoriginal"
        DIDate = wsDB2.Cells(2, i).Value
      Case "docno"
        DIDocNo = wsDB2.Cells(2, i).Value
      Case "metsodocno"
        DIMetsoDocNo = wsDB2.Cells(2, i).Value
      Case "project"
        DIProject = wsDB2.Cells(2, i).Value
      Case "status"
        DIStatus = wsDB2.Cells(2, i).Value
      Case "docname"
        DIDocName = wsDB2.Cells(2, i).Value
      Case "docname1"
        DIDocName1 = wsDB2.Cells(2, i).Value
      Case "docname2"
        DIDocName2 = wsDB2.Cells(2, i).Value
      Case "docname3"
        DIDocName3 = wsDB2.Cells(2, i).Value
      Case "contractno"
        DIContract = wsDB2.Cells(2, i).Value
      Case "projno"
        DIProjNo = wsDB2.Cells(2, i).Value
      Case "name"
        DIProjName = wsDB2.Cells(2, i).Value
      Case "workpath", "path", "work_path", "listpath", "lists_path", "zlistspath", "savepath", "targetpath", "outputpath"
        Dim p As String
        p = CStr(wsDB2.Cells(2, i).Value)
        If Len(p) > 0 Then
          ' Normalize separators and ensure trailing slash
          p = Replace(p, "/", "\\")
          DIPath = p & IIf(Right$(p, 1) = "\\", "", "\\")
        End If
      Case "manager"
        DIManager = wsDB2.Cells(2, i).Value
      Case "status"
        DIStatus = wsDB2.Cells(2, i).Value
      Case "mill"
        DIMill = wsDB2.Cells(2, i).Value
      Case "departname"
        DIDepartName = wsDB2.Cells(2, i).Value
      Case "customer"
        DICustomer = wsDB2.Cells(2, i).Value
      Case "metsounitname"
        DIMunit = wsDB2.Cells(2, i).Value
      Case "file", "filename", "file_name"
        DIFile = CStr(wsDB2.Cells(2, i).Value)
      Case ""
        Exit Do
      Case Else
    End Select
    i = i + 1
    ' Safety check: prevent infinite loop (Excel max columns)
    If i > 16384 Then Exit Do
  Loop
  
  ' DEBUG: Report what was loaded
  Debug.Print "HaeDocTiedot completed. Loaded " & (i - 1) & " columns from DB2"
  Debug.Print "  DIProject: '" & DIProject & "'"
  Debug.Print "  DIManager: '" & DIManager & "'"
  Debug.Print "  DIDocNo: '" & DIDocNo & "'"
  Debug.Print "  DIProjNo: '" & DIProjNo & "'"
End Sub
Sub VaihdaInfo(Optional Sheet As String = "Info")
'''
' VaihdaInfo: Updates the specified sheet's comment-annotated cells with document property values.
' Handles Info and Revisions sheets. Uses fast mode for performance.
'''
Dim i As Long
Dim Row As Long
Dim Column As Long
Dim r As Long
Dim ws As Worksheet
Dim processedRevId As Boolean, processedRevDate As Boolean
Dim processedDesigner As Boolean, processedChecker As Boolean
Dim processedApprover As Boolean, processedDesc As Boolean

  On Error Resume Next
  Set ws = Sheets(Sheet)
  On Error GoTo 0
  
  If ws Is Nothing Then
    Debug.Print "VaihdaInfo: Sheet '" & Sheet & "' not found!"
    Exit Sub
  End If
  
  ' DEBUG: Report sheet info
  Debug.Print "VaihdaInfo: Processing sheet '" & Sheet & "' with " & ws.Comments.Count & " comments"
  If ws.Comments.Count = 0 Then
    Debug.Print "  WARNING: No comments found in sheet - Info will remain empty!"
  End If
  
  ' Initialize flags for one-time processing of Revisions sheet arrays
  processedRevId = False
  processedRevDate = False
  processedDesigner = False
  processedChecker = False
  processedApprover = False
  processedDesc = False
  
  With ws
    For i = 1 To .Comments.Count 'Going through all comments
        Select Case LCase(.Comments(i).Text) ' Convert comment text to lowercase
        Case "unit"
          .Comments(i).Parent.Value = "Metso Paper - " & DIMunit
        Case "project"
          .Comments(i).Parent.Value = DIProject
        Case "manager"
          .Comments(i).Parent.Value = DIManager
        Case "contractno"
          .Comments(i).Parent.Value = DIContract
        Case "projname"
          .Comments(i).Parent.Value = DIProjName
        Case "projno"
          .Comments(i).Parent.Value = DIProjNo
        Case "date"
          .Comments(i).Parent.Value = DIDate
        Case "status"
          .Comments(i).Parent.Value = DIStatus
        Case "mill"
          .Comments(i).Parent.Value = DIMill
        Case "departname"
          .Comments(i).Parent.Value = DIDepartName
        Case "customer"
          .Comments(i).Parent.Value = DICustomer
        Case "docname"
          .Comments(i).Parent.Value = DIDocName
        Case "docname1"
          .Comments(i).Parent.Value = DIDocName1
        Case "docname2"
          .Comments(i).Parent.Value = DIDocName2
        Case "docname3"
          .Comments(i).Parent.Value = DIDocName3
        Case "metsodocno"
          .Comments(i).Parent.Value = DIMetsoDocNo
        Case "rev"
          .Comments(i).Parent.Value = DIRev
        Case "revid"
          If Sheet <> "Info" Then
            If Not processedRevId Then
              On Error Resume Next
              ' Only process if DIRevArr has data
              If IsArray(DIRevArr) And UBound(DIRevArr) >= LBound(DIRevArr) Then
                Row = .Comments(i).Parent.Row
                Column = .Comments(i).Parent.Column
                For r = UBound(DIRevArr) To LBound(DIRevArr) Step -1
                 If (DIRevArr(r) <> "") Then
                   .Cells(Row, Column).Value = Split(DIRevArr(r), " ")(0)
                   Row = Row + 1
                 End If
                Next r
              End If
              On Error GoTo 0
              processedRevId = True
            End If
          Else
            .Comments(i).Parent.Value = "'" & DIRevID
          End If
        Case "revdate"
          If Sheet <> "Info" Then
            If Not processedRevDate Then
              On Error Resume Next
              If IsArray(DIRevArr) And UBound(DIRevArr) >= LBound(DIRevArr) Then
                Row = .Comments(i).Parent.Row
                Column = .Comments(i).Parent.Column
                For r = UBound(DIRevArr) To LBound(DIRevArr) Step -1
                  If (DIRevArr(r) <> "") Then
                    .Cells(Row, Column).Value = Mid(DIRevArr(r), InStr(DIRevArr(r), " ") + 1, InStr(DIRevArr(r), "/") - 1 - InStr(DIRevArr(r), " "))
                    Row = Row + 1
                  End If
                Next r
              End If
              On Error GoTo 0
              processedRevDate = True
            End If
          Else
            .Comments(i).Parent.Value = DIRevDate
          End If
        Case "designer"
          If Sheet <> "Info" Then
            If Not processedDesigner Then
              On Error Resume Next
              If IsArray(DIRevArr) And UBound(DIRevArr) >= LBound(DIRevArr) Then
                Row = .Comments(i).Parent.Row
                Column = .Comments(i).Parent.Column
                For r = UBound(DIRevArr) To LBound(DIRevArr) Step -1
                  If (DIRevArr(r) <> "") Then
                   .Cells(Row, Column).Value = Split(DIRevArr(r), "/")(1)
                   Row = Row + 1
                  End If
                Next r
              End If
              On Error GoTo 0
              processedDesigner = True
            End If
          End If
        Case "checker"
          If Sheet <> "Info" Then
            If Not processedChecker Then
              On Error Resume Next
              If IsArray(DIRevArr) And UBound(DIRevArr) >= LBound(DIRevArr) Then
                Row = .Comments(i).Parent.Row
                Column = .Comments(i).Parent.Column
                For r = UBound(DIRevArr) To LBound(DIRevArr) Step -1
                  If (DIRevArr(r) <> "") Then
                   .Cells(Row, Column).Value = Split(DIRevArr(r), "/")(2)
                   Row = Row + 1
                  End If
                Next r
              End If
              On Error GoTo 0
              processedChecker = True
            End If
          End If
        Case "approver"
          If Sheet <> "Info" Then
            If Not processedApprover Then
              On Error Resume Next
              If IsArray(DIRevArr) And UBound(DIRevArr) >= LBound(DIRevArr) Then
                Row = .Comments(i).Parent.Row
                Column = .Comments(i).Parent.Column
                For r = UBound(DIRevArr) To LBound(DIRevArr) Step -1
                  If (DIRevArr(r) <> "") Then
                   .Cells(Row, Column).Value = Split(DIRevArr(r), "/")(3)
                   Row = Row + 1
                  End If
                Next r
              End If
              On Error GoTo 0
              processedApprover = True
            End If
          End If
        Case "desc"
          If Sheet <> "Info" Then
            If Not processedDesc Then
              On Error Resume Next
              If IsArray(DIRevArr) And UBound(DIRevArr) >= LBound(DIRevArr) Then
                Row = .Comments(i).Parent.Row
                Column = .Comments(i).Parent.Column
                For r = UBound(DIRevArr) To LBound(DIRevArr) Step -1
                  If (DIRevArr(r) <> "") Then
                   .Cells(Row, Column).Value = Split(DIRevArr(r), "/")(4)
                   Row = Row + 1
                  End If
                Next r
              End If
              On Error GoTo 0
              processedDesc = True
            End If
          End If
        End Select
      Next i
  End With
End Sub
Function EtsiOts(Otsikko As String, Rivi As Long, Sarake As Long, LRivi As Long) As Boolean
'''
' EtsiOts: Searches for a header (Otsikko) in DB1 and annotates TEMPLATE with a comment if found.
' If not found, logs the missing header in ERRORS sheet. Used for template validation.
' Optimized: Removed all Select/Activate, uses direct worksheet references.
'''
Dim i As Long
Dim j As Long
Dim wsDB1 As Worksheet, wsTemplate As Worksheet, wsErrors As Worksheet

   On Error Resume Next
   Set wsDB1 = Sheets("DB1")
   Set wsTemplate = Sheets("TEMPLATE")
   Set wsErrors = Sheets("ERRORS")
   On Error GoTo 0
   
   If wsDB1 Is Nothing Or wsTemplate Is Nothing Or wsErrors Is Nothing Then
     EtsiOts = False
     Exit Function
   End If
   
   i = 1
   Do
     ' Safety check: prevent infinite loop if sheet is unexpectedly large
     If i > 16384 Then ' Excel max columns
       EtsiOts = False
       Exit Do
     End If
     If LCase(wsDB1.Cells(1, i).Value) = LCase(Otsikko) Then
       ' Found header - add comment to TEMPLATE
       With wsTemplate.Cells(Rivi, Sarake)
         .AddComment
         .Comment.Text Text:=LRivi & ":" & i
         .Comment.Shape.DrawingObject.AutoSize = True
       End With
       EtsiOts = True
       Exit Do
     ElseIf wsDB1.Cells(1, i).Value = "" Then
       ' Not found - log to ERRORS sheet
       If wsErrors.Cells(1, 1).Value = "" Then
         wsErrors.Cells(1, 1).Value = "Following headlines were declared in TEMPLATE, but not found from DB sheet:"
         wsErrors.Cells(2, 1).Value = "HeadLine"
         wsErrors.Cells(2, 2).Value = "Location in TEMPLATE"
         wsErrors.Cells(1, 1).Font.Bold = True
         wsErrors.Cells(2, 1).Font.Bold = True
         wsErrors.Cells(2, 2).Font.Bold = True
         wsErrors.Columns("A:A").ColumnWidth = 30
         wsErrors.Columns("B:B").ColumnWidth = 25
       End If
       j = 3
       Do
         If wsErrors.Cells(j, 1) = "" Then
            wsErrors.Cells(j, 1).Value = Otsikko
            wsErrors.Cells(j, 2).Value = wsTemplate.Cells(Rivi, Sarake).Address
           Exit Do
         End If
         j = j + 1
         ' Safety check: prevent infinite loop
         If j > 10000 Then Exit Do
       Loop
       EtsiOts = False
       Exit Do
     End If
     i = i + 1
   Loop
End Function
Sub VaihdaLinkit(TargetSheet As Worksheet, Alku As Long, Loppu As Long, Kerta As Long)
'''
' VaihdaLinkit: For each comment in the specified range, updates the corresponding cell in LINKING with a formula
' and value, and applies formatting if needed. Used for main linking logic in printout.
' Only processes comments within the Alku:Loppu row range to avoid overwriting previously populated blocks.
'''
Dim TRow As Long, CRow As Long
Dim TCol As Long
Dim i As Long
Dim r As Long, c As Long
Dim Teksti As String
Dim Kaava As String
Dim Osoite As String
Dim cmt As Comment
  With TargetSheet
    ' Process only comments within the specified row range (Alku to Loppu)
    For r = Alku To Loppu
      For c = 1 To Sarakkeita
        Set cmt = .Cells(r, c).Comment
        If Not cmt Is Nothing Then
          ' Found a comment in this cell - process it
          Teksti = cmt.Text
          Osoite = cmt.Parent.Address(rowAbsolute:=False, columnAbsolute:=False)
          TRow = 1 + CInt(Left(Teksti, 1)) + Kerta * RMAX
          TCol = CInt(Mid(Teksti, 3))
          With .Parent.Sheets("LINKING").Cells(TRow, TCol)
            Teksti = .Value
            .Font.ColorIndex = 5
            .Font.Bold = True
            Kaava = "'" & POSheet & "'!" & Osoite
            .Formula = "=IF(" & Kaava & "="""", """"," & Kaava & ")"
          End With
          If cmt.Parent.Value = "££Deleted" Then
            cmt.Parent.Value = Teksti
            If Teksti = "Yes" Then
              CRow = cmt.Parent.Row
              .Rows(CRow).Font.Strikethrough = True
            End If
          Else
            cmt.Parent.Value = Teksti
          End If
          ' Delete this comment after processing
          cmt.Delete
          Set cmt = Nothing
        End If
      Next c
    Next r
  End With
End Sub
Sub PopulateRevisionsSimple()
'''
' PopulateRevisionsSimple: Lightweight function to populate Revisions sheet without comment processing.
' Finds the first cell with revision data markers and writes DIRevArr data directly.
'''
Dim ws As Worksheet
Dim r As Long, startRow As Long
Dim revIdCol As Long, revDateCol As Long, designerCol As Long
Dim checkerCol As Long, approverCol As Long, descCol As Long
Dim i As Long

  On Error Resume Next
  Set ws = Sheets("Revisions")
  If ws Is Nothing Then Exit Sub
  
  ' Check if DIRevArr is valid array with data
  If Not IsArray(DIRevArr) Then Exit Sub
  On Error Resume Next
  Dim arrSize As Long
  arrSize = UBound(DIRevArr) - LBound(DIRevArr) + 1
  If Err.Number <> 0 Or arrSize <= 0 Then
    On Error GoTo 0
    Exit Sub
  End If
  On Error GoTo 0
  
  ' Find columns by searching for comment markers in first 20 rows
  ' This is a simple heuristic - adjust if template structure differs
  startRow = 0
  For r = 1 To 20
    For i = 1 To 10 ' Check first 10 columns
      If ws.Cells(r, i).Comment Is Nothing Then GoTo NextCell
      Select Case LCase(ws.Cells(r, i).Comment.Text)
        Case "revid"
          revIdCol = i: If startRow = 0 Then startRow = r
        Case "revdate"
          revDateCol = i: If startRow = 0 Then startRow = r
        Case "designer"
          designerCol = i: If startRow = 0 Then startRow = r
        Case "checker"
          checkerCol = i: If startRow = 0 Then startRow = r
        Case "approver"
          approverCol = i: If startRow = 0 Then startRow = r
        Case "desc"
          descCol = i: If startRow = 0 Then startRow = r
      End Select
NextCell:
    Next i
  Next r
  
  ' If no columns found, exit
  If startRow = 0 Then Exit Sub
  
  ' Write revision data directly
  On Error Resume Next
  r = startRow
  Dim revParts() As String
  Dim revText As String
  Dim slashParts() As String
  
  For i = UBound(DIRevArr) To LBound(DIRevArr) Step -1
    If DIRevArr(i) <> "" Then
      revText = DIRevArr(i)
      
      ' Parse revid (first part before space)
      If revIdCol > 0 And InStr(revText, " ") > 0 Then
        revParts = Split(revText, " ")
        If UBound(revParts) >= 0 Then ws.Cells(r, revIdCol).Value = revParts(0)
      End If
      
      ' Parse revdate (between space and /)
      If revDateCol > 0 And InStr(revText, " ") > 0 And InStr(revText, "/") > 0 Then
        Dim spacePos As Long, slashPos As Long
        spacePos = InStr(revText, " ")
        slashPos = InStr(revText, "/")
        If slashPos > spacePos Then
          ws.Cells(r, revDateCol).Value = Mid(revText, spacePos + 1, slashPos - spacePos - 1)
        End If
      End If
      
      ' Parse designer, checker, approver, desc (slash-delimited parts)
      If InStr(revText, "/") > 0 Then
        slashParts = Split(revText, "/")
        If designerCol > 0 And UBound(slashParts) >= 1 Then ws.Cells(r, designerCol).Value = slashParts(1)
        If checkerCol > 0 And UBound(slashParts) >= 2 Then ws.Cells(r, checkerCol).Value = slashParts(2)
        If approverCol > 0 And UBound(slashParts) >= 3 Then ws.Cells(r, approverCol).Value = slashParts(3)
        If descCol > 0 And UBound(slashParts) >= 4 Then ws.Cells(r, descCol).Value = slashParts(4)
      End If
      
      r = r + 1
    End If
  Next i
  On Error GoTo 0
  
End Sub
Sub TeeLinkingKommentit()
'''
' TeeLinkingKommentit: Adds comments to all formula cells in the LINKING sheet for traceability.
' Optimized: Removed Select/Activate, uses direct worksheet references.
'''
Dim Solu As Range
Dim wsLinking As Worksheet
Dim formulaCells As Range

  ' Check if LINKING sheet exists
  On Error Resume Next
  Set wsLinking = Sheets("LINKING")
  On Error GoTo 0
  
  If wsLinking Is Nothing Then Exit Sub
  
  ' Find all formula cells
  On Error Resume Next
  Set formulaCells = wsLinking.Cells.SpecialCells(xlCellTypeFormulas)
  On Error GoTo 0
  
  If Not formulaCells Is Nothing Then
    Application.StatusBar = "Setting up comments in LINKING sheet (" & formulaCells.Cells.Count & ")"
    For Each Solu In formulaCells.Cells
      On Error Resume Next
      Solu.AddComment CStr(Solu.Value)
      On Error GoTo 0
    Next
  End If
  
  Application.DisplayCommentIndicator = xlCommentIndicatorOnly
End Sub
