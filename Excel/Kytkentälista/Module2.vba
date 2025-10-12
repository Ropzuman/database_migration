'''
' Module2.vba - Metadata, info, and linking logic for Kytkentälista Excel macro system
' Handles document property extraction, comment-based linking, and error reporting.
'''

Sub HaeDocTiedot()
'''
' HaeDocTiedot: Extracts document properties from DB2 sheet and stores them in global variables.
' Used for populating headers, footers, and info fields in the printout.
'''
Dim i As Long
Dim Arvo As String
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
Dim wsDB2 As Worksheet, wsTemplate As Worksheet

  On Error Resume Next
  Set wsDB2 = Sheets("DB2")
  Set wsTemplate = Sheets("TEMPLATE")
  On Error GoTo 0
  
  If wsDB2 Is Nothing Or wsTemplate Is Nothing Then Exit Sub
  
  wsDB2.Select
  i = 1
  Do
     Arvo = LCase(Cells(1, i).Value) ' Convert cell value to lowercase
    Select Case Arvo
      Case "rev"
        DIRev = Cells(2, i).Value
        Erase DIRevArr
        DIRevArr() = Split(DIRev, Chr(10))
      Case "revid"
        DIRevID = Cells(2, i).Value
      Case "revdate"
        DIRevDate = Cells(2, i).Value
      Case "date"
        DIDate = Cells(2, i).Value
      Case "docno"
        DIDocNo = Cells(2, i).Value
      Case "metsodocno"
        DIMetsoDocNo = Cells(2, i).Value
      Case "project"
        DIProject = Cells(2, i).Value
      Case "status"
        DIStatus = Cells(2, i).Value
      Case "docname"
        DIDocName = Cells(2, i).Value
      Case "docname1"
        DIDocName1 = Cells(2, i).Value
      Case "docname2"
        DIDocName2 = Cells(2, i).Value
       Case "docname3"
        DIDocName3 = Cells(2, i).Value
      Case "contractno"
        DIContract = Cells(2, i).Value
      Case "projno"
        DIProjNo = Cells(2, i).Value
      Case "name"
        DIProjName = Cells(2, i).Value
      Case "workpath"
        DIPath = Cells(2, i).Value & IIf(Right(Cells(2, i).Value, 1) = "\", "", "\")
      Case "manager"
        DIManager = Cells(2, i).Value
      Case "status"
        DIStatus = Cells(2, i).Value
      Case "mill"
        DIMill = Cells(2, i).Value
      Case "departname"
        DIDepartName = Cells(2, i).Value
      Case "customer"
        DICustomer = Cells(2, i).Value
      Case "metsounitname"
        DIMunit = Cells(2, i).Value
      Case "file"
        DIFile = Cells(2, i).Value
      Case ""
        Exit Do
      Case Else
    End Select
    i = i + 1
  Loop
  wsTemplate.Activate
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
  
  If ws Is Nothing Then Exit Sub
  
  ' Initialize flags for one-time processing of Revisions sheet arrays
  processedRevId = False
  processedRevDate = False
  processedDesigner = False
  processedChecker = False
  processedApprover = False
  processedDesc = False
  
  ws.Select
  With ActiveSheet
    For i = 1 To .Comments.Count 'Going through all comments
        Select Case LCase(.Comments(i).text) ' Convert comment text to lowercase
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
  Sheets("TEMPLATE").Select
End Sub
Function EtsiOts(Otsikko As String, Rivi As Long, Sarake As Long, LRivi As Long) As Boolean
'''
' EtsiOts: Searches for a header (Otsikko) in DB1 and annotates TEMPLATE with a comment if found.
' If not found, logs the missing header in ERRORS sheet. Used for template validation.
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
   wsDB1.Select
   Do
     ' Safety check: prevent infinite loop if sheet is unexpectedly large
     If i > 16384 Then ' Excel max columns
       EtsiOts = False
       Exit Do
     End If
     If LCase(Cells(1, i).Value) = LCase(Otsikko) Then
       wsTemplate.Select
       Cells(Rivi, Sarake).Select
       With ActiveCell
         .AddComment
         .Comment.text text:=LRivi & ":" & i
         .Comment.Shape.DrawingObject.AutoSize = True
       End With
       wsTemplate.Select
       EtsiOts = True
       Exit Do
     ElseIf Cells(1, i).Value = "" Then
       wsErrors.Select
       If Cells(1, 1).Value = "" Then
         Cells(1, 1).Value = "Following headlines were declared in TEMPLATE, but not found from DB sheet:"
         Cells(2, 1).Value = "HeadLine"
         Cells(2, 2).Value = "Location in TEMPLATE"
         Cells(1, 1).Font.Bold = True
         Cells(2, 1).Font.Bold = True
         Cells(2, 2).Font.Bold = True
         Columns("A:A").ColumnWidth = 30
         Columns("B:B").ColumnWidth = 25
       End If
       j = 3
       Do
         If Cells(j, 1) = "" Then
            Cells(j, 1).Value = Otsikko
            Cells(j, 2).Value = Cells(Rivi, Sarake).Address
           Exit Do
         End If
         j = j + 1
       Loop
       wsTemplate.Select
       EtsiOts = False
       Exit Do
     End If
     i = i + 1
   Loop
End Function
Sub VaihdaLinkit(Alku As Long, Loppu As Long, Kerta As Long)
'''
' VaihdaLinkit: For each comment in the active sheet, updates the corresponding cell in LINKING with a formula
' and value, and applies formatting if needed. Used for main linking logic in printout.
'''
Dim TRow As Long, CRow As Long
Dim TCol As Long
Dim i As Long
Dim Teksti As String
Dim Kaava As String
Dim Osoite As String
  With ActiveSheet
    For i = 1 To .Comments.Count 'Going through all comments
         Teksti = .Comments(i).text ' Get the comment text
       Osoite = .Comments(i).Parent.Address(rowAbsolute:=False, columnAbsolute:=False)
       TRow = 1 + CInt(Left(Teksti, 1)) + Kerta * RMAX
       TCol = CInt(Mid(Teksti, 3))
       With Sheets("LINKING").Cells(TRow, TCol)
         Teksti = .Value
         .Font.ColorIndex = 5
         .Font.Bold = True
         Kaava = "'" & POSheet & "'!" & Osoite
        .Formula = "=IF(" & Kaava & "="""", """"," & Kaava & ")"
      End With
      If .Comments(i).Parent.Value = "££Deleted" Then
        .Comments(i).Parent.Value = Teksti
        If Teksti = "Yes" Then
            CRow = .Comments(i).Parent.Row
            ActiveSheet.Rows(CRow).Font.Strikethrough = True
        End If
      Else
        .Comments(i).Parent.Value = Teksti
      End If
    Next i
    Cells.ClearComments
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
      Select Case LCase(ws.Cells(r, i).Comment.text)
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
' Uses fast mode for performance.
'''
Dim Solu As Range
Dim wsLinking As Worksheet

  ' Check if LINKING sheet exists
  On Error Resume Next
  Set wsLinking = Sheets("LINKING")
  On Error GoTo 0
  
  If wsLinking Is Nothing Then Exit Sub
  
  Sheets("LINKING").Select
  Cells(1, 1).Activate
  On Error Resume Next
  ActiveCell.SpecialCells(xlCellTypeFormulas).Select
  If Err.Number = 0 Then
    Application.StatusBar = "Setting up comments in LINKING sheet (" & Selection.Cells.Count & ")"
    For Each Solu In Selection.Cells
      Solu.AddComment CStr(Solu.Value)
    Next
  End If
  On Error GoTo 0
  Application.DisplayCommentIndicator = xlCommentIndicatorOnly
  Cells(1, 1).Activate
End Sub