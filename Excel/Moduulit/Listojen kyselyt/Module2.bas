'''
' Module2.vba - Metadata-, info- ja linkityslogiikka Kytkentälista Excel-makrojärjestelmälle
' Käsittelee dokumentin ominaisuuksien poiminnan, kommenttipohjaisen linkityksen ja virheraportoinnin.
' Huomiot:
' - HaeDocTiedot lukee DB2-otsikot case-insensitively ja trimm aa whitespacet.
' - WorkPath tunnistetaan useiden synonyymienhyvällä (workpath, path, work_path, listpath, lists_path, zlistspath, savepath, targetpath, outputpath)
'   ja normalisoidaan backslasheilla ja varmistetaan loppuun tulevaksi backslash.
' - File otetaan DB2:sta (file/filename/file_name) ja käytetään oletuksena Save As -nimeen GenPrintoutissa.
' 27.2.2026 - Poistettu duplikaattifunktiot (ne ovat Module1.bas:ssä)
'''

' Suojaus silmukoiden iteraatioille estääkseen ikuiset silmukat
Private Const MAX_EXCEL_COLUMNS As Long = 16384

Sub HaeDocTiedot()
'''
' HaeDocTiedot: Poimii dokumentin ominaisuudet DB2-sheetistä ja tallentaa ne globaaleihin muuttujiin.
' Käytetään otsikoiden,alatunnisteiden ja info-kenttien täyttöön tulosteessa.
' Optimoitu: Poistettu Select/Activate, käytetään suoria worksheetviittauksia.
' 27.2.2026 - Poistettu duplicate case "status"
'''
Dim i As Long
Dim Arvo As String
Dim wsDB2 As Worksheet

  Debug.Print Format(Now, "hh:mm:ss") & " [HaeDocTiedot] Poimitaan dokumentin metadata DB2:sta"

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
    ' Normalisoidaan otsikko: pienet kirjaimet ja trimmaus extra-spacet pois
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
          ' Normalisoidaan erottajat ja varmistetaan loppuun tulevaksi slash
          p = Replace(p, "/", "\\")
          DIPath = p & IIf(Right$(p, 1) = "\\", "", "\\")
        End If
      Case "manager"
        DIManager = wsDB2.Cells(2, i).Value
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
    If i > MAX_EXCEL_COLUMNS Then Exit Do ' Turvatarkistus
  Loop
  
  ' DEBUG: Raportoidaan mitä ladattiin
  Debug.Print "  Ladattu " & (i - 1) & " saraketta DB2:sta"
  Debug.Print "  DIProject: '" & DIProject & "'"
  Debug.Print "  DIManager: '" & DIManager & "'"
  Debug.Print "  DIDocNo: '" & DIDocNo & "'"
  Debug.Print "  DIProjNo: '" & DIProjNo & "'"
  Debug.Print "  DIPath: '" & DIPath & "'"
  Debug.Print "  DIFile: '" & DIFile & "'"
End Sub
Sub VaihdaInfo(Optional SheetName As String = "Info")
'''
' VaihdaInfo: Päivittää määritellyn sheetin kommenttimerkatut solut dokumentin ominaisuuksilla.
' Käsittelee Info- ja Revisions-sheetit. Käyttää fast mode -tilaa suorituskyvyn parantamiseksi.
' 27.2.2026 - Lisätty virheenkäsittely
'''
Dim i As Long
Dim Row As Long
Dim Column As Long
Dim r As Long
Dim ws As Worksheet
Dim processedRevId As Boolean, processedRevDate As Boolean
Dim processedDesigner As Boolean, processedChecker As Boolean
Dim processedApprover As Boolean, processedDesc As Boolean

  Debug.Print Format(Now, "hh:mm:ss") & " [VaihdaInfo] Käsitellään sheet: " & SheetName

  On Error GoTo ErrHandler
  
  Set ws = Sheets(SheetName)
  
  If ws Is Nothing Then
    Debug.Print "  VAROITUS: Sheet '" & SheetName & "' puuttuu!"
    Exit Sub
  End If
  
  ' DEBUG: Raportoidaan sheetin tiedot
  Debug.Print "  Sheetissä '" & SheetName & "' on " & ws.Comments.Count & " kommenttia"
  If ws.Comments.Count = 0 Then
    Debug.Print "  VAROITUS: Ei kommentteja - Info jää tyhjäksi!"
  End If
  
  ' Alustetaan liput yksikertaisesti käsiteltäviksi Revisions-sheet-taulukoiksi
  processedRevId = False
  processedRevDate = False
  processedDesigner = False
  processedChecker = False
  processedApprover = False
  processedDesc = False
  
  With ws
    For i = 1 To .Comments.Count 'Käydään läpi kaikki kommentit
        Select Case LCase(.Comments(i).Text) ' Muutetaan kommenttiteksti pieniksi kirjaimiksi
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
        Case "docname3"          .Comments(i).Parent.Value = DIDocName3
        Case "metsodocno"
          .Comments(i).Parent.Value = DIMetsoDocNo
        Case "rev"
          .Comments(i).Parent.Value = DIRev
        Case "revid"
          If SheetName <> "Info" Then
            If Not processedRevId Then
              On Error Resume Next
              ' Käsitellään vain jos DIRevArr:ssa on dataa
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
              On Error GoTo ErrHandler
              processedRevId = True
            End If
          Else
            .Comments(i).Parent.Value = "'" & DIRevID
          End If
        Case "revdate"
          If SheetName <> "Info" Then
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
              On Error GoTo ErrHandler
              processedRevDate = True
            End If
          Else
            .Comments(i).Parent.Value = DIRevDate
          End If
        Case "designer"
          If SheetName <> "Info" Then
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
              On Error GoTo ErrHandler
              processedDesigner = True
            End If
          End If
        Case "checker"
          If SheetName <> "Info" Then
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
              On Error GoTo ErrHandler
              processedChecker = True
            End If
          End If
        Case "approver"
          If SheetName <> "Info" Then
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
              On Error GoTo ErrHandler
              processedApprover = True
            End If
          End If
        Case "desc"
          If SheetName <> "Info" Then
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
              On Error GoTo ErrHandler
              processedDesc = True
            End If
          End If
        End Select
      Next i
  End With
  
  Debug.Print Format(Now, "hh:mm:ss") & " [VaihdaInfo] Valmis"
  Exit Sub
  
ErrHandler:
  Debug.Print Format(Now, "hh:mm:ss") & " [VaihdaInfo ERROR] " & Err.Number & ": " & Err.Description
  MsgBox "Virhe Info-sheetin päivityksessä:" & vbCrLf & Err.Description & vbCrLf & vbCrLf & _
         "Sheet: " & SheetName, vbCritical, "VaihdaInfo Error"
  Err.Clear
  On Error GoTo 0
End Sub
Function EtsiOts(Otsikko As String, Rivi As Long, Sarake As Long, LRivi As Long) As Boolean
'''
' EtsiOts: Etsii otsikkoa (Otsikko) DB1:stä ja merkitsee TEMPLATEn kommentilla jos löytyi.
' Jos ei löydy, lokitetaan puuttuva otsikko ERRORS-sheettiin. Käytetään templaten validointiin.
' Optimoitu: Poistettu kaikki Select/Activate, käytetään suoria worksheetviittauksia.
'''
Dim i As Long
Dim j As Long
Dim wsDB1 As Worksheet, wsTemplate As Worksheet, wsErrors As Worksheet

   Debug.Print Format(Now, "hh:mm:ss") & " [EtsiOts] Etsitään otsikkoa: '" & Otsikko & "'"

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
     If i > MAX_EXCEL_COLUMNS Then ' Turvatarkistus
       EtsiOts = False
       Debug.Print "  [EtsiOts VAROITUS] Ei löytynyt ennen MAX_EXCEL_COLUMNS"
       Exit Do
     End If
     If LCase(wsDB1.Cells(1, i).Value) = LCase(Otsikko) Then
       ' Löytyi otsikko - lisää kommentti TEMPLATEen
       Debug.Print "  Löytyi sarakkeesta " & i
       With wsTemplate.Cells(Rivi, Sarake)
         .AddComment
         .Comment.Text Text:=LRivi & ":" & i
         .Comment.Shape.DrawingObject.AutoSize = True
       End With
       EtsiOts = True
       Exit Do
     ElseIf wsDB1.Cells(1, i).Value = "" Then
       ' Ei löytynyt - lokitetaan ERRORS-sheettiin
       Debug.Print "  [EtsiOts ERROR] Ei löytynyt - lisätään ERRORS-sheettiin"
       If wsErrors.Cells(1, 1).Value = "" Then
         wsErrors.Cells(1, 1).Value = "Seuraavat otsikot oli määritelty TEMPLATEssa, mutta ei löytynyt DB-sheetistä:"
         wsErrors.Cells(2, 1).Value = "Otsikko"
         wsErrors.Cells(2, 2).Value = "Sijainti TEMPLATEssa"
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
         ' Turvatarkistus: estä ikuinen silmukka
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
' VaihdaLinkit: Jokaiselle kommentille määritellyllä alueella päivitetään vastaava solu LINKINGissä kaavalla
' ja arvolla, ja sovelletaan muotoilua tarvittaessa. Käytetään tulosteen päälinkityslogiikkaan.
' Käsittelee vain kommentit Alku:Loppu -rivialueelta välttäen aiemmin täytettyjen lohkojen ylikirjoittamista.
' OPTIMOITU: Käyttää Comments-kokoelmaa solulta-solulle-iteraation sijaan (30-50% nopeampi).
'''
Dim TRow As Long, CRow As Long
Dim TCol As Long
Dim i As Long
Dim Teksti As String
Dim Kaava As String
Dim Osoite As String
Dim cmt As Comment
Dim commentsToDelete As Collection
Dim parentRow As Long

  Set commentsToDelete = New Collection
  
  With TargetSheet
    ' OPTIMOINTI: Iteroidaan Comments-kokoelmaa kaikkien solujen sijaan
    ' Tämä ohittaa tyhjät solut ja on paljon nopeampi harvoille kommenttijakaumille
    For Each cmt In .Comments
      parentRow = cmt.Parent.Row
      
      ' Käsitellään vain kommentit määritellyllä rivialueella
      If parentRow >= Alku And parentRow <= Loppu Then
        ' Löyttyi kommentti alueella - käsitellään se
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
        ' Merkitään kommentti poistettavaksi (ei voi poistaa iteraation aikana)
        commentsToDelete.Add cmt
      End If
    Next cmt
    
    ' Poistetaan käsitellyt kommentit iteraation valmistuttua
    For i = 1 To commentsToDelete.Count
      commentsToDelete(i).Delete
    Next i
  End With
  
  Set commentsToDelete = Nothing
End Sub
Sub PopulateRevisionsSimple()
'''
' PopulateRevisionsSimple: Kevyt funktio Revisions-sheetin täyttöön ilman kommenttikäsittelyä.
' Etsii ensimmäisen solun revisiodata-merkillä ja kirjoittaa DIRevArr-datan suoraan.
' 27.2.2026 - Rajoitettu On Error Resume Next -käyttö
'''
Dim ws As Worksheet
Dim r As Long, startRow As Long
Dim revIdCol As Long, revDateCol As Long, designerCol As Long
Dim checkerCol As Long, approverCol As Long, descCol As Long
Dim i As Long
Dim arrSize As Long
Dim spacePos As Long, slashPos As Long
Dim revParts() As String
Dim revText As String
Dim slashParts() As String

  Debug.Print Format(Now, "hh:mm:ss") & " [PopulateRevisionsSimple] Täytetään Revisions-sheet"

  On Error Resume Next
  Set ws = Sheets("Revisions")
  On Error GoTo 0
  
  If ws Is Nothing Then
    Debug.Print "  VAROITUS: Revisions-sheet puuttuu - ohitetaan"
    Exit Sub
  End If
  
  ' Tarkistetaan onko DIRevArr validi taulukko datalla
  If Not IsArray(DIRevArr) Then
    Debug.Print "  VAROITUS: DIRevArr ei ole taulukko - ohitetaan"
    Exit Sub
  End If
  
  On Error Resume Next
  arrSize = UBound(DIRevArr) - LBound(DIRevArr) + 1
  On Error GoTo 0
  
  If Err.Number <> 0 Or arrSize <= 0 Then
    Debug.Print "  VAROITUS: DIRevArr tyhjä - ohitetaan"
    Exit Sub
  End If
  
  Debug.Print "  DIRevArr-koko: " & arrSize
  
  ' Etsitään sarakkeet hakemalla kommenttimerkit ensimmäisestä 20 rivistä
  ' Tämä on yksinkertainen heuristiikka -  säädä jos templaten rakenne eroaa
  startRow = 0
  For r = 1 To 20
    For i = 1 To 10 ' Check first 10 columns
      On Error Resume Next
      If ws.Cells(r, i).Comment Is Nothing Then GoTo NextCell
      On Error GoTo 0
      
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
  
  ' Jos sarakkeita ei löytynyt, poistu
  If startRow = 0 Then
    Debug.Print "  VAROITUS: Ei löytynyt revisiomarkkereita - ohitetaan"
    Exit Sub
  End If
  
  Debug.Print "  Revisiotiedot alkavat riviltä " & startRow
  
  ' Kirjoitetaan revisiodata suoraan - rajattu On Error Resume Next vain parsing-osaan
  r = startRow
  
  For i = UBound(DIRevArr) To LBound(DIRevArr) Step -1
    If DIRevArr(i) <> "" Then
      revText = DIRevArr(i)
      
      ' Parse revid (first part before space)
      On Error Resume Next
      If revIdCol > 0 And InStr(revText, " ") > 0 Then
        revParts = Split(revText, " ")
        If UBound(revParts) >= 0 Then ws.Cells(r, revIdCol).Value = revParts(0)
      End If
      On Error GoTo 0
      
      ' Parse revdate (between space and /)
      If revDateCol > 0 And InStr(revText, " ") > 0 And InStr(revText, "/") > 0 Then
        spacePos = InStr(revText, " ")
        slashPos = InStr(revText, "/")
        If slashPos > spacePos Then
          ws.Cells(r, revDateCol).Value = Mid(revText, spacePos + 1, slashPos - spacePos - 1)
        End If
      End If
      
      ' Parse designer, checker, approver, desc (slash-delimited parts)
      On Error Resume Next
      If InStr(revText, "/") > 0 Then
        slashParts = Split(revText, "/")
        If designerCol > 0 And UBound(slashParts) >= 1 Then ws.Cells(r, designerCol).Value = slashParts(1)
        If checkerCol > 0 And UBound(slashParts) >= 2 Then ws.Cells(r, checkerCol).Value = slashParts(2)
        If approverCol > 0 And UBound(slashParts) >= 3 Then ws.Cells(r, approverCol).Value = slashParts(3)
        If descCol > 0 And UBound(slashParts) >= 4 Then ws.Cells(r, descCol).Value = slashParts(4)
      End If
      On Error GoTo 0
      
      r = r + 1
    End If
  Next i
  
  Debug.Print Format(Now, "hh:mm:ss") & " [PopulateRevisionsSimple] Valmis - kirjoitettiin " & (r - startRow) & " revisioriviä"
  
End Sub
Sub TeeLinkingKommentit()
'''
' TeeLinkingKommentit: Lisää kommentit kaikille kaavayksikkösoluille LINKING-sheetissä jäljitettävyyttä varten.
' Optimoitu: Poistettu Select/Activate, käytetään suoria worksheetviittauksia.
'''
Dim Solu As Range
Dim wsLinking As Worksheet
Dim formulaCells As Range

  Debug.Print Format(Now, "hh:mm:ss") & " [TeeLinkingKommentit] Lisätään kommentit LINKING-sheettiin"
  
  ' Tarkistetaan onko LINKING-sheet olemassa
  On Error Resume Next
  Set wsLinking = Sheets("LINKING")
  On Error GoTo 0
  
  If wsLinking Is Nothing Then
    Debug.Print "  VAROITUS: LINKING-sheet puuttuu - ohitetaan kommentit"
    Exit Sub
  End If
  
  ' Etsitään kaikki kaavasolut
  On Error Resume Next
  Set formulaCells = wsLinking.Cells.SpecialCells(xlCellTypeFormulas)
  On Error GoTo 0
  
  If Not formulaCells Is Nothing Then
    Application.StatusBar = "Asetetaan kommentit LINKING-sheettiin (" & formulaCells.Cells.Count & ")"
    Debug.Print "  Käsitellään " & formulaCells.Cells.Count & " kaavasolua"
    For Each Solu in formulaCells.Cells
      On Error Resume Next
      Solu.AddComment CStr(Solu.Value)
      On Error GoTo 0
    Next
  End If
  
  Application.DisplayCommentIndicator = xlCommentIndicatorOnly
  Debug.Print Format(Now, "hh:mm:ss") & " [TeeLinkingKommentit] Valmis"
End Sub
