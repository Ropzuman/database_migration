Sub TyhjaaKommentit()
    Cells.ClearComments
End Sub
Sub HaeDocTiedot()
Dim i As Integer
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

  Worksheets("DB2").Select
  i = 1
  Do
    Arvo = LCase(Cells(1, i).Value)
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
  Worksheets("TEMPLATE").Select
End Sub
Sub VaihdaInfo(Optional Sheet As String = "Info")
Dim i As Long
'Dim Row As Range
  Worksheets(Sheet).Select
  With ActiveSheet
    For i = 1 To .Comments.Count 'Käydään läpi kaikki kommentit
      Select Case LCase(.Comments(i).Text)
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
  Worksheets("TEMPLATE").Select
End Sub
Function EtsiOts(Otsikko As String, Rivi As Integer, Sarake As Integer, LRivi As Integer) As Boolean
Dim i As Long
Dim j As Long
i = 1
   Worksheets("DB1").Select
   Do
     If LCase(Cells(1, i).Value) = LCase(Otsikko) Then
       Worksheets("TEMPLATE").Select
       Cells(Rivi, Sarake).Select
       With ActiveCell
         .AddComment
         .Comment.Text Text:=LRivi & ":" & i
         .Comment.Visible = False
         .Comment.Shape.DrawingObject.AutoSize = True
       End With
       EtsiOts = True
       Exit Do
     ElseIf Cells(1, i).Value = "" Then
       Worksheets("ERRORS").Select
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
       Worksheets("TEMPLATE").Select
       EtsiOts = False
       Exit Do
     End If
     i = i + 1
   Loop
End Function
Sub VaihdaLinkit1(Alku As Integer, Loppu As Integer, Kerta As Integer)
Dim TRow As Integer
Dim TCol As Integer
Dim i As Integer
Dim j As Integer
Dim Teksti As String
Dim Arvo As String
    For i = Alku To Loppu
      For j = 1 To Sarakkeita
        If Left(Cells(i, j).Value, 1) = "£" Then
          Teksti = Cells(i, j).Comment.Text
          TRow = 1 + CInt(Left(Teksti, 1)) + Kerta * RMAX
          TCol = CInt(Mid(Teksti, 3))
          With Worksheets("LINKING").Cells(TRow, TCol)
            Arvo = .Value
            .Font.ColorIndex = 5
            .Font.Bold = True
            On Error GoTo Virhe_Komment
            .AddComment
Virhe_Komment:
            .Comment.Text Text:=Arvo
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
Sub VaihdaLinkit(Alku As Integer, Loppu As Integer, Kerta As Integer)
Dim TRow, CRow As Integer
Dim TCol As Integer
Dim i As Integer
Dim Teksti As String
Dim Kaava As String
Dim Osoite As String
  With ActiveSheet
    For i = 1 To .Comments.Count 'Käydään läpi kaikki kommentit
       Teksti = .Comments(i).Text
       Osoite = .Comments(i).Parent.Address(rowAbsolute:=False, columnAbsolute:=False)
       TRow = 1 + CInt(Left(Teksti, 1)) + Kerta * RMAX
       TCol = CInt(Mid(Teksti, 3))
       With Worksheets("LINKING").Cells(TRow, TCol)
         Teksti = .Value
         .Font.ColorIndex = 5
         .Font.Bold = True
         Kaava = "'" & POSheet & "'!" & Osoite
        .Formula = "=IF(" & Kaava & "="""", """"," & Kaava & ")"
      End With
      
      If .Comments(i).Parent.Column = 1 Then
        If Teksti <> "" Then
           BGColor = Not BGColor
        End If
      End If
      
      If .Comments(i).Parent.Value = "££Deleted" Or .Comments(i).Parent.Value = "££deldtl" Then
        .Comments(i).Parent.Value = Teksti
        If Teksti = "Yes" Then
            CRow = .Comments(i).Parent.Row
            ActiveSheet.Rows(CRow).Font.Strikethrough = True
        End If
        If BGColor = True Then
            CRow = .Comments(i).Parent.Row
            ActiveSheet.Rows(CRow).Interior.Color = BGValue
        End If
      Else
        .Comments(i).Parent.Value = Teksti
        If BGColor = True Then
            CRow = .Comments(i).Parent.Row
            ActiveSheet.Rows(CRow).Interior.Color = BGValue
        End If
      End If
    Next i
    Cells.ClearComments
  End With
End Sub
Sub VaihdaLinkit_OLD(Alku As Integer, Loppu As Integer, Kerta As Integer)
Dim TRow As Integer
Dim TCol As Integer
Dim i As Integer
Dim Teksti As String
Dim Arvo As String
Dim Osoite As String
  With ActiveSheet
    For i = 1 To .Comments.Count 'Käydään läpi kaikki kommentit
      Teksti = .Comments(i).Text
      Osoite = .Comments(i).Parent.Address
      TRow = 1 + CInt(Left(Teksti, 1)) + Kerta * RMAX
      TCol = CInt(Mid(Teksti, 3))
      Worksheets("LINKING").Select
      Cells(TRow, TCol).Select
      Arvo = MuutaLinkki(Osoite)
      Worksheets(POSheet).Select
      .Comments(i).Parent.Value = Arvo
    Next i
    Cells.ClearComments
  End With
End Sub
Function MuutaLinkki(Kohde As String) As String
Dim Arvo As String
On Error GoTo Virhe_Komment
  With ActiveCell
    Arvo = .Value
    .Font.ColorIndex = 5
    .Font.Bold = True
    .AddComment
Virhe_Komment:
    .Comment.Text Text:=Arvo
    .Formula = "='" & POSheet & "'!" & Kohde
    .Comment.Visible = False
    .Comment.Shape.DrawingObject.Shadow = False
    .Comment.Shape.DrawingObject.AutoSize = True
    MuutaLinkki = Arvo
End With
End Function
Sub TarkistaVaihto(Vaihto As Integer, ViimRivi As Integer, Riveja As Integer)
Dim SVRivi As Integer
On Error GoTo VirheSivunLuvussa
  
'  SVRivi = CInt(ActiveSheet.HPageBreaks(Vaihto).Location.Row)
  'Automaattinen rivinvaihto tuli huonoon kohtaan, joten tehdään itse uusi edelliseen sopivaan paikkaan
  Cells(ViimRivi, 1).Select
  ActiveSheet.HPageBreaks.Add Before:=ActiveCell

Ulos_TarkistaVaihto:
  Exit Sub
        
VirheSivunLuvussa:
  'Rivinvaihto tuli juuri nappiin kohtaan, vahvistetaan se vielä
  Cells(ViimRivi + Riveja + 1, 1).Select
  ActiveSheet.HPageBreaks.Add Before:=ActiveCell
  Resume Ulos_TarkistaVaihto
  
End Sub
Sub TeeLinkingKommentit()
Dim Solu As Range
  Worksheets("LINKING").Select
  Cells(1, 1).Activate
  ActiveCell.SpecialCells(xlCellTypeFormulas).Select
  Application.StatusBar = "Setting up comments in LINKING sheet (" & Selection.Cells.Count & ")"
  For Each Solu In Selection.Cells
    Solu.AddComment CStr(Solu.Value)
  Next
  Application.DisplayCommentIndicator = xlCommentIndicatorOnly
  Cells(1, 1).Activate
End Sub
