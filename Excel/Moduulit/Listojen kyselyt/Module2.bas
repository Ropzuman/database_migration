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
  DocInfo.Rev = ""
  DocInfo.RevID = ""
  DocInfo.RevDate = ""
  DocInfo.DocNo = ""
  DocInfo.MetsoDocNo = ""
  DocInfo.Project = ""
  DocInfo.Status = ""
  DocInfo.DocName = ""
  DocInfo.DocName1 = ""
  DocInfo.DocName2 = ""
  DocInfo.DocName3 = ""
  DocInfo.Contract = ""
  DocInfo.ProjNo = ""
  DocInfo.ProjName = ""
  DocInfo.Path = ""
  DocInfo.Date = ""
  DocInfo.Manager = ""
  DocInfo.Munit = ""
  DocInfo.Mill = ""
  DocInfo.DepartName = ""
  DocInfo.Customer = ""
  DocInfo.File = ""

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
        DocInfo.Rev = CStr(valArr(1, i) & "")
        Erase DIRevArr
        If Len(DocInfo.Rev) > 0 Then
          ' Puhdistetaan piilevät CR-merkit (Chr(13)) ennen jakamista — Windows CR+LF -yhteensopivuus
          DocInfo.Rev = Replace(DocInfo.Rev, vbCr, "")
          DIRevArr = Split(DocInfo.Rev, vbLf)
        Else
          ReDim DIRevArr(0): DIRevArr(0) = ""
        End If
      Case "revid"
        DocInfo.RevID = valArr(1, i)
      Case "revdate"
        DocInfo.RevDate = valArr(1, i)
      Case "date", "dateoriginal"
        DocInfo.Date = valArr(1, i)
      Case "docno"
        DocInfo.DocNo = valArr(1, i)
      Case "metsodocno"
        DocInfo.MetsoDocNo = valArr(1, i)
      Case "project"
        DocInfo.Project = valArr(1, i)
      Case "status"
        DocInfo.Status = valArr(1, i)
      Case "docname"
        DocInfo.DocName = valArr(1, i)
      Case "docname1"
        DocInfo.DocName1 = valArr(1, i)
      Case "docname2"
        DocInfo.DocName2 = valArr(1, i)
      Case "docname3"
        DocInfo.DocName3 = valArr(1, i)
      Case "contractno"
        DocInfo.Contract = valArr(1, i)
      Case "projno"
        DocInfo.ProjNo = valArr(1, i)
      Case "name"
        DocInfo.ProjName = valArr(1, i)
      Case "workpath", "path", "work_path", "listpath", "lists_path", "zlistspath", "savepath", "targetpath", "outputpath"
        p = CStr(valArr(1, i) & "")
        If Len(p) > 0 Then
          ' Normalisoidaan erottajat: kauttaviivat kenoviivoiksi, varmistetaan loppukenoviiva
          p = Replace(p, "/", "\")
          DocInfo.Path = p & IIf(Right$(p, 1) = "\", "", "\")
        End If
      Case "manager"
        DocInfo.Manager = valArr(1, i)
      Case "mill"
        DocInfo.Mill = valArr(1, i)
      Case "departname"
        DocInfo.DepartName = valArr(1, i)
      Case "customer"
        DocInfo.Customer = valArr(1, i)
      Case "metsounitname"
        DocInfo.Munit = valArr(1, i)
      Case "file", "filename", "file_name"
        DocInfo.File = CStr(valArr(1, i) & "")
      Case Else
    End Select
  Next i
  
  ' DEBUG: Raportoidaan mitä ladattiin
  Debug.Print "  Ladattu " & lastCol & " saraketta DB2:sta"
  Debug.Print "  DocInfo.Project: '" & DocInfo.Project & "'"
  Debug.Print "  DocInfo.Manager: '" & DocInfo.Manager & "'"
  Debug.Print "  DocInfo.DocNo: '" & DocInfo.DocNo & "'"
  Debug.Print "  DocInfo.ProjNo: '" & DocInfo.ProjNo & "'"
  Debug.Print "  DocInfo.Path: '" & DocInfo.Path & "'"
  Debug.Print "  DocInfo.File: '" & DocInfo.File & "'"
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
          cmt.Parent.Value = "Metso Paper - " & DocInfo.Munit
        Case "project"
          cmt.Parent.Value = DocInfo.Project
        Case "manager"
          cmt.Parent.Value = DocInfo.Manager
        Case "contractno"
          cmt.Parent.Value = DocInfo.Contract
        Case "projname"
          cmt.Parent.Value = DocInfo.ProjName
        Case "projno"
          cmt.Parent.Value = DocInfo.ProjNo
        Case "date"
          cmt.Parent.Value = DocInfo.Date
        Case "status"
          cmt.Parent.Value = DocInfo.Status
        Case "mill"
          cmt.Parent.Value = DocInfo.Mill
        Case "departname"
          cmt.Parent.Value = DocInfo.DepartName
        Case "customer"
          cmt.Parent.Value = DocInfo.Customer
        Case "docname"
          cmt.Parent.Value = DocInfo.DocName
        Case "docname1"
          cmt.Parent.Value = DocInfo.DocName1
        Case "docname2"
          cmt.Parent.Value = DocInfo.DocName2
        Case "docname3"          
          cmt.Parent.Value = DocInfo.DocName3
        Case "metsodocno"
          cmt.Parent.Value = DocInfo.MetsoDocNo
        Case "rev"
          cmt.Parent.Value = DocInfo.Rev
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
            cmt.Parent.Value = "'" & DocInfo.RevID
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
            cmt.Parent.Value = DocInfo.RevDate
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
        ' CLng estää Integer Overflow -kaatumisen yli 32767 rivin tiedostoissa
        TRow = 1 + CLng(Left(Teksti, 1)) + Kerta * RMAX
        TCol = CLng(Mid(Teksti, 3))
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
Dim r As Long
Dim i As Long
Dim arrSize As Long
Dim spacePos As Long, slashPos As Long
Dim revParts() As String
Dim revText As String
Dim slashParts() As String
Dim remainder As String
Dim token As Variant
Dim descText As String
Dim revIdCols As Collection
Dim revDateCols As Collection
Dim designerCols As Collection
Dim checkerCols As Collection
Dim approverCols As Collection
Dim descCols As Collection
Dim col As Variant
Dim roleIndex As Long
Dim partIndex As Long

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

  Set revIdCols = New Collection
  Set revDateCols = New Collection
  Set designerCols = New Collection
  Set checkerCols = New Collection
  Set approverCols = New Collection
  Set descCols = New Collection
  
  ' Etsitään revisiomarkkerit hakemalla kommentit ensimmäisestä 20 rivistä.
  ' Tämä on kevyt heuristiikka - säädä jos templaten rakenne eroaa.
  Dim revIdRow As Long, revDateRow As Long, designerRow As Long
  Dim checkerRow As Long, approverRow As Long, descRow As Long
  revIdRow = 0: revDateRow = 0: designerRow = 0
  checkerRow = 0: approverRow = 0: descRow = 0

  For r = 1 To 20
    For i = 1 To 10 ' Check first 10 columns
      On Error Resume Next
      If ws.Cells(r, i).Comment Is Nothing Then GoTo NextCell2
      On Error GoTo 0

      ' Normalize comment text (remove CR/LF and trim)
      Dim ctext As String
      ctext = LCase(Replace(Replace(Trim(ws.Cells(r, i).Comment.Text), vbCr, ""), vbLf, ""))

      Select Case ctext
        Case "revid"
          revIdCols.Add i: If revIdRow = 0 Then revIdRow = r
        Case "revdate"
          revDateCols.Add i: If revDateRow = 0 Then revDateRow = r
        Case "designer"
          designerCols.Add i: If designerRow = 0 Then designerRow = r
        Case "checker"
          checkerCols.Add i: If checkerRow = 0 Then checkerRow = r
        Case "approver"
          approverCols.Add i: If approverRow = 0 Then approverRow = r
        Case "desc"
          descCols.Add i: If descRow = 0 Then descRow = r
      End Select
NextCell2:
    Next i
  Next r

  ' Jos sarakkeita ei löytynyt, poistu
  If revIdRow = 0 And revDateRow = 0 And designerRow = 0 And checkerRow = 0 And approverRow = 0 And descRow = 0 Then
    Debug.Print "  VAROITUS: Ei löytynyt revisiomarkkereita - ohitetaan"
    Exit Sub
  End If

  Debug.Print "  Revisiomarkkerit löytyivät riveiltä: revid=" & revIdRow & ", revdate=" & revDateRow & ", designer=" & designerRow

  ' Jos merkintärivit eroavat, käytetään pienintä (ylin) rivinumeroa yhteisenä aloituksena
  Dim commonStart As Long
  commonStart = 0
  If revIdRow > 0 Then commonStart = IIf(commonStart = 0, revIdRow, Application.Min(commonStart, revIdRow))
  If revDateRow > 0 Then commonStart = IIf(commonStart = 0, revDateRow, Application.Min(commonStart, revDateRow))
  If designerRow > 0 Then commonStart = IIf(commonStart = 0, designerRow, Application.Min(commonStart, designerRow))
  If checkerRow > 0 Then commonStart = IIf(commonStart = 0, checkerRow, Application.Min(commonStart, checkerRow))
  If approverRow > 0 Then commonStart = IIf(commonStart = 0, approverRow, Application.Min(commonStart, approverRow))
  If descRow > 0 Then commonStart = IIf(commonStart = 0, descRow, Application.Min(commonStart, descRow))

  If commonStart = 0 Then
    Debug.Print "  VAROITUS: Ei yhteistä aloitusriviä löytynyt - ohitetaan"
    Exit Sub
  End If

  Debug.Print "  Käytetään yhteistä aloitusriviä: " & commonStart

  ' Kirjoitetaan revisiodata käyttäen yhteistä aloitusriviä ja säilytetään aiempi käänteinen järjestys
  Dim rId As Long, rDate As Long, rDesigner As Long, rChecker As Long, rApprover As Long, rDesc As Long
  rId = commonStart: rDate = commonStart: rDesigner = commonStart
  rChecker = commonStart: rApprover = commonStart: rDesc = commonStart

  Dim written As Long: written = 0
  For i = UBound(DIRevArr) To LBound(DIRevArr) Step -1
    If DIRevArr(i) <> "" Then
      revText = DIRevArr(i)
      written = written + 1

      ' Parse revid (first part before space)
      If InStr(revText, " ") > 0 Then
        revParts = Split(revText, " ")
        If UBound(revParts) >= 0 Then
          For Each col In revIdCols
            ws.Cells(rId, CLng(col)).Value = revParts(0)
          Next col
        End If
      End If

      ' Parse revdate (väliosan ja ensimmäisen /-merkin väli) ja kirjoita kaikkiin date-sarakkeisiin.
      spacePos = InStr(revText, " ")
      slashPos = InStr(revText, "/")
      If spacePos > 0 And slashPos > spacePos Then
        If slashPos > spacePos Then
          For Each col In revDateCols
            ws.Cells(rDate, CLng(col)).Value = Trim(Mid(revText, spacePos + 1, slashPos - spacePos - 1))
          Next col
        End If
      End If

      ' Parse designer, checker, approver, desc käyttäen kiinteitä slash-paikkoja.
      ' Tyhjät kentät pitää säilyttää tyhjinä, jotta kuvaus ei siirry väärään sarakkeeseen.
      If slashPos > 0 Then
        remainder = Mid(revText, slashPos + 1)
        slashParts = Split(remainder, "/")

        ' slashParts(0) = Designer, (1) = Checker, (2) = Approver, (3) = Description
        If UBound(slashParts) >= 0 Then
          token = Trim(slashParts(0))
          If token <> "" Then
            For Each col In designerCols
              ws.Cells(rDesigner, CLng(col)).Value = token
            Next col
          End If
        End If

        If UBound(slashParts) >= 1 Then
          token = Trim(slashParts(1))
          If token <> "" Then
            For Each col In checkerCols
              ws.Cells(rChecker, CLng(col)).Value = token
            Next col
          End If
        End If

        If UBound(slashParts) >= 2 Then
          token = Trim(slashParts(2))
          If token <> "" Then
            For Each col In approverCols
              ws.Cells(rApprover, CLng(col)).Value = token
            Next col
          End If
        End If

        descText = ""
        If UBound(slashParts) >= 3 Then
          descText = Trim(slashParts(3))
        End If
        If descText <> "" Then
          For Each col In descCols
            ws.Cells(rDesc, CLng(col)).Value = descText
          Next col
        End If
      End If

      ' Increment all row counters so rows remain aligned
      rId = rId + 1: rDate = rDate + 1: rDesigner = rDesigner + 1
      rChecker = rChecker + 1: rApprover = rApprover + 1: rDesc = rDesc + 1
    End If
  Next i

  Debug.Print Format(Now, "hh:mm:ss") & " [PopulateRevisionsSimple] Valmis - kirjoitettiin " & written & " revisioriviä"
  
End Sub
Sub PrintRevisionDiagnostics()
  '''
  ' PrintRevisionDiagnostics: Tulostaa DIRevArr:n sisällön ja Revisions-sheetin
  ' revisiomarkkerien sijainnit (ei tee muutoksia taulukkoon).
  ' Käytä debug-ikkunaa (Immediate, Ctrl+G) nähdäksesi tulosteet.
  '''
  Dim ws As Worksheet
  Dim i As Long, r As Long
  Dim ctext As String

  Debug.Print "[PrintRevisionDiagnostics] Aloitetaan diagnostikka"

  On Error Resume Next
  Set ws = Sheets("Revisions")
  On Error GoTo 0
  If ws Is Nothing Then
    Debug.Print "  VAROITUS: Revisions-sheet puuttuu"
    Exit Sub
  End If

  ' DIRevArr sisältö
  If Not IsArray(DIRevArr) Then
    Debug.Print "  VAROITUS: DIRevArr ei ole taulukko tai tyhjä"
  Else
    Debug.Print "  DIRevArr sisältö (index:value):"
    For i = LBound(DIRevArr) To UBound(DIRevArr)
      Debug.Print "    [" & i & "] = '" & DIRevArr(i) & "'"
    Next i
  End If

  Debug.Print "  Etsitään revisiomarkkereita Revisions-sheetin ensimmäisiltä riveiltä"
  For r = 1 To 20
    For i = 1 To 10
      On Error Resume Next
      If Not ws.Cells(r, i).Comment Is Nothing Then
        ctext = ws.Cells(r, i).Comment.Text
        ctext = LCase(Replace(Replace(Trim(ctext), vbCr, ""), vbLf, ""))
        Debug.Print "    Löytyi marker '" & ctext & "' osoitteessa " & ws.Cells(r, i).Address(False, False) & " (r=" & r & ", c=" & i & ")"
      End If
      On Error GoTo 0
    Next i
  Next r

  Debug.Print "[PrintRevisionDiagnostics] Valmis"
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
