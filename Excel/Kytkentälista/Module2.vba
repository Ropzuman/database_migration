Sub TyhjaaKommentit()
    Cells.ClearComments
End Sub
  '
  ' Fast-mode helpers to reduce flicker and speed up heavy operations
  '''
  ' Module2.vba - Metadata, info, and linking logic for Kytkentälista Excel macro system
  ' Handles document property extraction, comment-based linking, and error reporting.
  '''

  ' Fast-mode helpers to reduce flicker and speed up heavy operations
  Private prevScreenUpdating2 As Boolean
  Private prevCalculation2 As XlCalculation
  Private prevEnableEvents2 As Boolean
  Private prevDisplayAlerts2 As Boolean
  Private prevDisplayStatusBar2 As Boolean

  Private Sub BeginFastMode2()
  '''
  ' BeginFastMode2: Temporarily disables Excel UI updates, events, and sets calculation to manual
  ' to speed up macro execution and prevent screen flicker.
  '''
    prevScreenUpdating2 = Application.ScreenUpdating
    prevCalculation2 = Application.Calculation
    prevEnableEvents2 = Application.EnableEvents
    prevDisplayAlerts2 = Application.DisplayAlerts
    prevDisplayStatusBar2 = Application.DisplayStatusBar
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False
    Application.DisplayAlerts = False
    On Error Resume Next
    Application.DisplayStatusBar = False
    On Error GoTo 0
  End Sub

  Private Sub EndFastMode2()
  '''
  ' EndFastMode2: Restores Excel UI and calculation settings to their previous state.
  '''
    On Error Resume Next
    Application.ScreenUpdating = prevScreenUpdating2
    Application.Calculation = prevCalculation2
    Application.EnableEvents = prevEnableEvents2
    Application.DisplayAlerts = prevDisplayAlerts2
    Application.DisplayStatusBar = prevDisplayStatusBar2
    On Error GoTo 0
  End Sub
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

  Sheets("DB2").Select
  Set wsDB2 = Sheets("DB2")
  Set wsTemplate = Sheets("TEMPLATE")
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
'Dim Row As Range
  Sheets(Sheet).Select
  With ActiveSheet
    For i = 1 To .Comments.Count 'Käydään läpi kaikki kommentit
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
            Row = .Comments(i).Parent.Row
            Column = .Comments(i).Parent.Column
            For r = UBound(DIRevArr) To LBound(DIRevArr) Step -1
             If (DIRevArr(r) <> "") Then
               .Cells(Row, Column).Value = Split(DIRevArr(r), " ")(0)
               Row = Row + 1
             End If
              Dim Row As Long
              Dim Column As Long
              Dim r As Long
              Dim ws As Worksheet
              Set ws = Sheets(Sheet)
              BeginFastMode2
            Next r
          Else
            .Comments(i).Parent.Value = "'" & DIRevID
          End If
        Case "revdate"
          If Sheet <> "Info" Then
            Row = .Comments(i).Parent.Row
            Column = .Comments(i).Parent.Column
            For r = UBound(DIRevArr) To LBound(DIRevArr) Step -1
              If (DIRevArr(r) <> "") Then
                .Cells(Row, Column).Value = Mid(DIRevArr(r), InStr(DIRevArr(r), " ") + 1, InStr(DIRevArr(r), "/") - 1 - InStr(DIRevArr(r), " "))
                Row = Row + 1
              End If
            Next r
          Else
            .Comments(i).Parent.Value = DIRevDate
          End If
        Case "designer"
          If Sheet <> "Info" Then
            Row = .Comments(i).Parent.Row
            Column = .Comments(i).Parent.Column
            For r = UBound(DIRevArr) To LBound(DIRevArr) Step -1
              If (DIRevArr(r) <> "") Then
               .Cells(Row, Column).Value = Split(DIRevArr(r), "/")(1)
               Row = Row + 1
              End If
            Next r
          End If
        Case "checker"
          If Sheet <> "Info" Then
            Row = .Comments(i).Parent.Row
            Column = .Comments(i).Parent.Column
            For r = UBound(DIRevArr) To LBound(DIRevArr) Step -1
              If (DIRevArr(r) <> "") Then
               .Cells(Row, Column).Value = Split(DIRevArr(r), "/")(2)
               Row = Row + 1
              End If
            Next r
          End If
        Case "approver"
          If Sheet <> "Info" Then
            Row = .Comments(i).Parent.Row
            Column = .Comments(i).Parent.Column
            For r = UBound(DIRevArr) To LBound(DIRevArr) Step -1
              If (DIRevArr(r) <> "") Then
               .Cells(Row, Column).Value = Split(DIRevArr(r), "/")(3)
               Row = Row + 1
              End If
            Next r
          End If
        Case "desc"
          If Sheet <> "Info" Then
            Row = .Comments(i).Parent.Row
            Column = .Comments(i).Parent.Column
            For r = UBound(DIRevArr) To LBound(DIRevArr) Step -1
              If (DIRevArr(r) <> "") Then
               .Cells(Row, Column).Value = Split(DIRevArr(r), "/")(4)
               Row = Row + 1
              End If
            Next r
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
i = 1
   Sheets("DB1").Select
   Do
     If LCase(Cells(1, i).Value) = LCase(Otsikko) Then
       Sheets("TEMPLATE").Select
       Cells(Rivi, Sarake).Select
       With ActiveCell
         .AddComment
         .Comment.text text:=LRivi & ":" & i
                EndFastMode2
                Sheets("TEMPLATE").Select
         .Comment.Shape.DrawingObject.AutoSize = True
       End With
       EtsiOts = True
       Exit Do
     ElseIf Cells(1, i).Value = "" Then
       Sheets("ERRORS").Select
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
       Sheets("TEMPLATE").Select
       EtsiOts = False
       Exit Do
      Dim wsDB1 As Worksheet, wsTemplate As Worksheet, wsErrors As Worksheet
      Set wsDB1 = Sheets("DB1")
      Set wsTemplate = Sheets("TEMPLATE")
      Set wsErrors = Sheets("ERRORS")
     End If
     i = i + 1
   Loop
End Function
Sub VaihdaLinkit1(Alku As Long, Loppu As Long, Kerta As Long)
'''
' VaihdaLinkit1: For each cell in the specified range, if it contains a linking marker, copies the value
' from the LINKING sheet and adds a comment/formula for traceability. Used for legacy linking logic.
'''
Dim TRow As Long
Dim TCol As Long
Dim i As Long
Dim j As Long
Dim Teksti As String
Dim Arvo As String
    For i = Alku To Loppu
      For j = 1 To Sarakkeita
        If Left(Cells(i, j).Value, 1) = "£" Then
          Teksti = Cells(i, j).Comment.text
          TRow = 1 + CInt(Left(Teksti, 1)) + Kerta * RMAX
          TCol = CInt(Mid(Teksti, 3))
          With Sheets("LINKING").Cells(TRow, TCol)
            Arvo = .Value
            .Font.ColorIndex = 5
            .Font.Bold = True
            On Error GoTo Virhe_Komment
            .AddComment
Virhe_Komment:
            .Comment.text text:=Arvo
            .FormulaR1C1 = "='" & POSheet & "'!R" & i & "C" & j
            .Comment.Visible = False
            .Comment.Shape.DrawingObject.Shadow = False
            .Comment.Shape.DrawingObject.AutoSize = True
          End With
          Cells(i, j).Value = Arvo
        End If
      Next j
    Next i
End Sub
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
    For i = 1 To .Comments.Count 'Käydään läpi kaikki kommentit
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
Sub VaihdaLinkit_OLD(Alku As Long, Loppu As Long, Kerta As Long)
'''
' VaihdaLinkit_OLD: Legacy version of linking logic, kept for reference. Uses Select/Activate.
'''
Dim TRow As Long
Dim TCol As Long
Dim i As Long
Dim Teksti As String
Dim Arvo As String
Dim Osoite As String
  With ActiveSheet
    For i = 1 To .Comments.Count 'Käydään läpi kaikki kommentit
        Teksti = .Comments(i).text ' Get the comment text
      Osoite = .Comments(i).Parent.Address
      TRow = 1 + CInt(Left(Teksti, 1)) + Kerta * RMAX
      TCol = CInt(Mid(Teksti, 3))
      Sheets("LINKING").Select
      Cells(TRow, TCol).Select
      Arvo = MuutaLinkki(Osoite)
      Sheets(POSheet).Select
      .Comments(i).Parent.Value = Arvo
    Next i
    Cells.ClearComments
  End With
End Sub
Function MuutaLinkki(Kohde As String) As String
'''
' MuutaLinkki: Helper for VaihdaLinkit_OLD. Adds a comment and formula to the active cell for traceability.
'''
Dim Arvo As String
On Error GoTo Virhe_Komment
  With ActiveCell
    Arvo = .Value
    .Font.ColorIndex = 5
    .Font.Bold = True
    .AddComment
Virhe_Komment:
    .Comment.text text:=Arvo
    .Formula = "='" & POSheet & "'!" & Kohde
    .Comment.Visible = False
    .Comment.Shape.DrawingObject.Shadow = False
    .Comment.Shape.DrawingObject.AutoSize = True
    MuutaLinkki = Arvo
End With
End Function
Sub TarkistaVaihto(Vaihto As Long, ViimRivi As Long, Riveja As Long)
'''
' TarkistaVaihto: Ensures page breaks are set at appropriate rows in the printout sheet.
'''
Dim SVRivi As Long
On Error GoTo VirheSivunLuvussa
  
'  SVRivi = CInt(ActiveSheet.HPageBreaks(Vaihto).Location.Row)
  'Automaattinen rivinvaihto tuli huonoon kohtaan, joten tehdään itse uusi edelliseen sopivaan paikkaan
  Cells(ViimRivi, 1).Select
    ActiveSheet.HPageBreaks.Add Before:=ActiveCell ' Add a page break before the active cell

Ulos_TarkistaVaihto:
  Exit Sub
        
VirheSivunLuvussa:
  'Rivinvaihto tuli juuri nappiin kohtaan, vahvistetaan se vielä
  Cells(ViimRivi + Riveja + 1, 1).Select
  ActiveSheet.HPageBreaks.Add Before:=ActiveCell
  Resume Ulos_TarkistaVaihto
  
End Sub
Sub TeeLinkingKommentit()
'''
' TeeLinkingKommentit: Adds comments to all formula cells in the LINKING sheet for traceability.
' Uses fast mode for performance.
'''
Dim Solu As Range
Dim wsLinking As Worksheet
Set wsLinking = Sheets("LINKING")
BeginFastMode2
  Sheets("LINKING").Select
  Cells(1, 1).Activate
  ActiveCell.SpecialCells(xlCellTypeFormulas).Select
  Application.StatusBar = "Setting up comments in LINKING sheet (" & Selection.Cells.Count & ")"
  For Each Solu In Selection.Cells
    Solu.AddComment CStr(Solu.Value)
  Next
  Application.DisplayCommentIndicator = xlCommentIndicatorOnly
  Cells(1, 1).Activate
EndFastMode2
End Sub