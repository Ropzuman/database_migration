Option Explicit

'''
' Module2.vba - Metadata-, info- ja linkityslogiikka Kytkentälista Excel-makrojärjestelmälle
' Käsittelee dokumentin ominaisuuksien poiminnan, kommenttipohjaisen linkityksen ja virheraportoinnin.
' Huomiot:
' - HaeDocTiedot lukee DB2-otsikot case-insensitively ja trimmaa whitespacet.
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
Dim p As String
Dim lastCol As Long    ' Viimeinen käytetty sarake DB2:n otsikkoriviltä
Dim hdrArr As Variant  ' Otsikkorivi taulukossa (1 COM-kysely kaikkien sarakkeiden sijaan)
Dim valArr As Variant  ' Datarivi taulukossa (1 COM-kysely kaikkien sarakkeiden sijaan)

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
  
  ' Optimoitu: Luetaan koko otsikko- ja datarivi kerralla taulukkoon (2 COM-kutsua aiemman 2*N sijaan)
  On Error Resume Next
  lastCol = wsDB2.Cells(1, wsDB2.Columns.Count).End(xlToLeft).Column
  On Error GoTo 0
  If lastCol < 1 Or (lastCol = 1 And wsDB2.Cells(1, 1).Value = "") Then
    Debug.Print "  VAROITUS: DB2 otsikkorivi on tyhjä"
    Exit Sub
  End If
  hdrArr = wsDB2.Range(wsDB2.Cells(1, 1), wsDB2.Cells(1, lastCol)).Value
  valArr = wsDB2.Range(wsDB2.Cells(2, 1), wsDB2.Cells(2, lastCol)).Value

  For i = 1 To lastCol
    ' Normalisoidaan otsikko: pienet kirjaimet ja trimmaus extra-spacet pois
    Arvo = LCase(Trim(CStr(hdrArr(1, i))))
    Select Case Arvo
      Case "rev"
        ' Null-turvallinen split: tarkistetaan pituus ennen splittausta (Type Mismatch -esto)
        DIRev = CStr(valArr(1, i) & "")
        Erase DIRevArr
        If Len(DIRev) > 0 Then
          DIRevArr = Split(DIRev, Chr(10))
        Else
          ReDim DIRevArr(0): DIRevArr(0) = ""
        End If
      Case "revid"
        DIRevID = valArr(1, i)
      Case "revdate"
        DIRevDate = valArr(1, i)
      Case "date", "dateoriginal"
        DIDate = valArr(1, i)
      Case "docno"
        DIDocNo = valArr(1, i)
      Case "metsodocno"
        DIMetsoDocNo = valArr(1, i)
      Case "project"
        DIProject = valArr(1, i)
      Case "status"
        DIStatus = valArr(1, i)
      Case "docname"
        DIDocName = valArr(1, i)
      Case "docname1"
        DIDocName1 = valArr(1, i)
      Case "docname2"
        DIDocName2 = valArr(1, i)
      Case "docname3"
        DIDocName3 = valArr(1, i)
      Case "contractno"
        DIContract = valArr(1, i)
      Case "projno"
        DIProjNo = valArr(1, i)
      Case "name"
        DIProjName = valArr(1, i)
      Case "workpath", "path", "work_path", "listpath", "lists_path", "zlistspath", "savepath", "targetpath", "outputpath"
        p = CStr(valArr(1, i) & "")
        If Len(p) > 0 Then
          ' Normalisoidaan erottajat: kauttaviivat kenoviivoiksi, varmistetaan loppukenoviiva
          p = Replace(p, "/", "\")
          DIPath = p & IIf(Right$(p, 1) = "\", "", "\")
        End If
      Case "manager"
        DIManager = valArr(1, i)
      Case "mill"
        DIMill = valArr(1, i)
      Case "departname"
        DIDepartName = valArr(1, i)
      Case "customer"
        DICustomer = valArr(1, i)
      Case "metsounitname"
        DIMunit = valArr(1, i)
      Case "file", "filename", "file_name"
        DIFile = CStr(valArr(1, i) & "")
      Case Else
    End Select
  Next i
  
  ' DEBUG: Raportoidaan mitä ladattiin
  Debug.Print "  Ladattu " & lastCol & " saraketta DB2:sta"
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
Dim cmt As Comment          ' For Each -iteraattori kommenttikokoelmalle (nopeampi kuin indeksiviittaus)
Dim revParts() As String    ' Tilapäistaulukko Split()-rajojen tarkistukseen
Dim processedRevId As Boolean, processedRevDate As Boolean
Dim processedDesigner As Boolean, processedChecker As Boolean
Dim processedApprover As Boolean, processedDesc As Boolean

  Debug.Print Format(Now, "hh:mm:ss") & " [VaihdaInfo] Käsitellään sheet: " & SheetName

  ' Korjattu: Sheets()-kutsu ei palauta Nothing vaan nostaa Error 9 jos sheettiä ei löydy.
  ' On Error Resume Next ennen Set-lausetta mahdollistaa oikean Nothing-tarkistuksen.
  On Error Resume Next
  Set ws = Sheets(SheetName)
  On Error GoTo ErrHandler
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
    ' For Each on nopeampi kuin indeksipohjainen .Comments(i) — COM-kokoelman indeksointi on O(n)
    For Each cmt In .Comments 'Käydään läpi kaikki kommentit
        Select Case LCase(cmt.Text) ' Muutetaan kommenttiteksti pieniksi kirjaimiksi
        Case "unit"
          cmt.Parent.Value = "Metso Paper - " & DIMunit
        Case "project"
          cmt.Parent.Value = DIProject
        Case "manager"
          cmt.Parent.Value = DIManager
        Case "contractno"
          cmt.Parent.Value = DIContract
        Case "projname"
          cmt.Parent.Value = DIProjName
        Case "projno"
          cmt.Parent.Value = DIProjNo
        Case "date"
          cmt.Parent.Value = DIDate
        Case "status"
          cmt.Parent.Value = DIStatus
        Case "mill"
          cmt.Parent.Value = DIMill
        Case "departname"
          cmt.Parent.Value = DIDepartName
        Case "customer"
          cmt.Parent.Value = DICustomer
        Case "docname"
          cmt.Parent.Value = DIDocName
        Case "docname1"
          cmt.Parent.Value = DIDocName1
        Case "docname2"
          cmt.Parent.Value = DIDocName2
        Case "docname3"          
          cmt.Parent.Value = DIDocName3
        Case "metsodocno"
          cmt.Parent.Value = DIMetsoDocNo
        Case "rev"
          cmt.Parent.Value = DIRev
        Case "revid"
          If SheetName <> "Info" Then
            If Not processedRevId Then
              On Error Resume Next
              ' Käsitellään vain jos DIRevArr:ssa on dataa
              If IsArray(DIRevArr) And UBound(DIRevArr) >= LBound(DIRevArr) Then
                Row = cmt.Parent.Row
                Column = cmt.Parent.Column
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
            cmt.Parent.Value = "'" & DIRevID
          End If
        Case "revdate"
          If SheetName <> "Info" Then
            If Not processedRevDate Then
              On Error Resume Next
              If IsArray(DIRevArr) And UBound(DIRevArr) >= LBound(DIRevArr) Then
                Row = cmt.Parent.Row
                Column = cmt.Parent.Column
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
            cmt.Parent.Value = DIRevDate
          End If
        Case "designer"
          If SheetName <> "Info" Then
            If Not processedDesigner Then
              On Error Resume Next
              If IsArray(DIRevArr) And UBound(DIRevArr) >= LBound(DIRevArr) Then
                Row = cmt.Parent.Row
                Column = cmt.Parent.Column
                For r = UBound(DIRevArr) To LBound(DIRevArr) Step -1
                  If (DIRevArr(r) <> "") Then
                    ' Erase estää vanhan arvon läpivuotamisen seuraavaan iteraatioon
                    ' jos merkkijonossa ei ole kauttaviivaa (revParts jäisi edellisen arvoon)
                    Erase revParts
                    If InStr(DIRevArr(r), "/") > 0 Then
                      revParts = Split(DIRevArr(r), "/")
                      If UBound(revParts) >= 1 Then .Cells(Row, Column).Value = revParts(1)
                    End If
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
                Row = cmt.Parent.Row
                Column = cmt.Parent.Column
                For r = UBound(DIRevArr) To LBound(DIRevArr) Step -1
                  If (DIRevArr(r) <> "") Then
                    Erase revParts
                    If InStr(DIRevArr(r), "/") > 0 Then
                      revParts = Split(DIRevArr(r), "/")
                      If UBound(revParts) >= 2 Then .Cells(Row, Column).Value = revParts(2)
                    End If
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
                Row = cmt.Parent.Row
                Column = cmt.Parent.Column
                For r = UBound(DIRevArr) To LBound(DIRevArr) Step -1
                  If (DIRevArr(r) <> "") Then
                    Erase revParts
                    If InStr(DIRevArr(r), "/") > 0 Then
                      revParts = Split(DIRevArr(r), "/")
                      If UBound(revParts) >= 3 Then .Cells(Row, Column).Value = revParts(3)
                    End If
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
                Row = cmt.Parent.Row
                Column = cmt.Parent.Column
                For r = UBound(DIRevArr) To LBound(DIRevArr) Step -1
                  If (DIRevArr(r) <> "") Then
                    Erase revParts
                    If InStr(DIRevArr(r), "/") > 0 Then
                      revParts = Split(DIRevArr(r), "/")
                      If UBound(revParts) >= 4 Then .Cells(Row, Column).Value = revParts(4)
                    End If
                    Row = Row + 1
                  End If
                Next r
              End If
              On Error GoTo ErrHandler
              processedDesc = True
            End If
          End If
        End Select
      Next cmt
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
Sub VaihdaLinkit(TargetSheet As Worksheet, Alku As Long, Loppu As Long, Kerta As Long, _
                 Optional SourceWB As Workbook = Nothing)
'''
' VaihdaLinkit: Kirjoittaa DB1-arvot suoraan kommenttimerkatuille soluille kohdetulosteessa.
' Parametri SourceWB osoittaa lähdetyökirjaan jossa DB1 sijaitsee (yleensä ThisWorkbook/srcWB).
' Jos SourceWB jätetään pois, käytetään ThisWorkbook.Sheets("DB1"):ta oletuksena.
' KORJATTU: Poistettu kaavaketjulogiikka (LINKING-kaavat viittasivat tyhjiin kohdesoluihin).
' OPTIMOITU: Käyttää Comments-kokoelmaa solulta-solulle-iteraation sijaan (30-50% nopeampi).
'''
Dim TRow As Long, CRow As Long
Dim TCol As Long
Dim i As Long
Dim Teksti As String
Dim Osoite As String
Dim cmt As Comment
Dim commentsToDelete As Collection
Dim parentRow As Long
Dim wsDB1 As Worksheet

  Set commentsToDelete = New Collection

  ' Ratkaistaan DB1-viittaus lähdeparametrin tai ThisWorkbook:n kautta
  If SourceWB Is Nothing Then
    Set wsDB1 = ThisWorkbook.Sheets("DB1")
  Else
    Set wsDB1 = SourceWB.Sheets("DB1")
  End If

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
        ' Luetaan arvo suoraan DB1:stä — poistettu kaavaketju joka viittasi tyhjään kohdesoluun
        Teksti = CStr(wsDB1.Cells(TRow, TCol).Value)
        ' Fix4: Kirjoitetaan staattinen arvo LINKINGiin yhdellä COM-kutsulla
        ' (Font-muotoilu poistettu hot path:sta — säästää 2 COM-kutsua per kommenttisolku)
        .Parent.Sheets("LINKING").Cells(TRow, TCol).Value = Teksti
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
  Set wsDB1 = Nothing
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
    ' OERN asetetaan kerran silmukan ulkopuolella — ei per-iteraatio overhead
    ' (AddComment epäonnistuu jos kommentti jo on olemassa; tämä on odotettu tilanne)
    On Error Resume Next
    For Each Solu In formulaCells.Cells
      Solu.AddComment CStr(Solu.Value)
    Next
    On Error GoTo 0
  End If
  
  Application.DisplayCommentIndicator = xlCommentIndicatorOnly
  Debug.Print Format(Now, "hh:mm:ss") & " [TeeLinkingKommentit] Valmis"
End Sub
