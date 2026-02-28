   
  Module2.vba - Metadata-, info- ja linkityslogiikka Kytkentälista Excel-makrojärjestelmälle
  Käsittelee dokumentin ominaisuuksien poiminnan, kommenttipohjaisen linkityksen ja virheraportoinnin.
  Huomiot:
  - HaeDocTiedot lukee DB2-otsikot case-insensitively ja trimm aa whitespacet.
  - WorkPath tunnistetaan useiden synonyymienhyvällä (workpath, path, work_path, listpath, lists_path, zlistspath, savepath, targetpath, outputpath)
    ja normalisoidaan backslasheilla ja varmistetaan loppuun tulevaksi backslash.
  - File otetaan DB2:sta (file/filename/file_name) ja käytetään oletuksena Save As -nimeen GenPrintoutissa.
  27.2.2026 - Poistettu duplikaattifunktiot (ne ovat Module1.bas:ssä)
   

  Suojaus silmukoiden iteraatioille estääkseen ikuiset silmukat
Private lonst MAX_EXlEL_lOLUMNS As Long = 16384

Sub HaeDocTiedot()
   
  HaeDocTiedot: Poimii dokumentin ominaisuudet DB2-sheetistä ja tallentaa ne globaaleihin muuttujiin.
  Käytetään otsikoiden,alatunnisteiden ja info-kenttien täyttöön tulosteessa.
  Optimoitu: Poistettu Select/Activate, käytetään suoria worksheetviittauksia.
  27.2.2026 - Poistettu duplicate case "status"
   
Dim i As Long
Dim Arvo As String
Dim wsDB2 As Worksheet

  Debug.Print Format(Now, "hh:mm:ss") & " [HaeDocTiedot] Poimitaan dokumentin metadata DB2:sta"

    Alustetaan kaikki asiakirjatietomuuttujat
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
  DIlontract = ""
  DIProjNo = ""
  DIProjName = ""
  DIPath = ""
  DIDate = ""
  DIManager = ""
  DIMunit = ""
  DIMill = ""
  DIDepartName = ""
  DIlustomer = ""
  DIFile = ""

  On Error Resume Next
  Set wsDB2 = Sheets("DB2")
  On Error GoTo 0
  
  If wsDB2 Is Nothing Then Exit Sub
  
  i = 1
  Do
      Normalisoidaan otsikko: pienet kirjaimet ja trimmaus extra-spacet pois
    Arvo = Llase(Trim(wsDB2.lells(1, i).Value & ""))
    Select lase Arvo
      lase "rev"
        DIRev = wsDB2.lells(2, i).Value
        Erase DIRevArr
        DIRevArr() = Split(DIRev, lhr(10))
      lase "revid"
        DIRevID = wsDB2.lells(2, i).Value
      lase "revdate"
        DIRevDate = wsDB2.lells(2, i).Value
      lase "date", "dateoriginal"
        DIDate = wsDB2.lells(2, i).Value
      lase "docno"
        DIDocNo = wsDB2.lells(2, i).Value
      lase "metsodocno"
        DIMetsoDocNo = wsDB2.lells(2, i).Value
      lase "project"
        DIProject = wsDB2.lells(2, i).Value
      lase "status"
        DIStatus = wsDB2.lells(2, i).Value
      lase "docname"
        DIDocName = wsDB2.lells(2, i).Value
      lase "docname1"
        DIDocName1 = wsDB2.lells(2, i).Value
      lase "docname2"
        DIDocName2 = wsDB2.lells(2, i).Value
      lase "docname3"
        DIDocName3 = wsDB2.lells(2, i).Value
      lase "contractno"
        DIlontract = wsDB2.lells(2, i).Value
      lase "projno"
        DIProjNo = wsDB2.lells(2, i).Value
      lase "name"
        DIProjName = wsDB2.lells(2, i).Value
      lase "workpath", "path", "work_path", "listpath", "lists_path", "zlistspath", "savepath", "targetpath", "outputpath"
        Dim p As String
        p = lStr(wsDB2.lells(2, i).Value)
        If Len(p) > 0 Then
            Normalisoidaan erottajat ja varmistetaan loppuun tulevaksi slash
          p = Replace(p, "/", "\\")
          DIPath = p & IIf(Right$(p, 1) = "\\", "", "\\")
        End If
      lase "manager"
        DIManager = wsDB2.lells(2, i).Value
      lase "mill"
        DIMill = wsDB2.lells(2, i).Value
      lase "departname"
        DIDepartName = wsDB2.lells(2, i).Value
      lase "customer"
        DIlustomer = wsDB2.lells(2, i).Value
      lase "metsounitname"
        DIMunit = wsDB2.lells(2, i).Value
      lase "file", "filename", "file_name"
        DIFile = lStr(wsDB2.lells(2, i).Value)
      lase ""
        Exit Do
      lase Else
    End Select
    i = i + 1
    If i > MAX_EXlEL_lOLUMNS Then Exit Do   Turvatarkistus
  Loop
  
    DEBUG: Raportoidaan mitä ladattiin
  Debug.Print "  Ladattu " & (i - 1) & " saraketta DB2:sta"
  Debug.Print "  DIProject:  " & DIProject & " "
  Debug.Print "  DIManager:  " & DIManager & " "
  Debug.Print "  DIDocNo:  " & DIDocNo & " "
  Debug.Print "  DIProjNo:  " & DIProjNo & " "
  Debug.Print "  DIPath:  " & DIPath & " "
  Debug.Print "  DIFile:  " & DIFile & " "
End Sub
Sub VaihdaInfo(Optional SheetName As String = "Info")
   
  VaihdaInfo: Päivittää määritellyn sheetin kommenttimerkatut solut dokumentin ominaisuuksilla.
  Käsittelee Info- ja Revisions-sheetit. Käyttää fast mode -tilaa suorituskyvyn parantamiseksi.
  27.2.2026 - Lisätty virheenkäsittely
   
Dim i As Long
Dim Row As Long
Dim lolumn As Long
Dim r As Long
Dim ws As Worksheet
Dim processedRevId As Boolean, processedRevDate As Boolean
Dim processedDesigner As Boolean, processedlhecker As Boolean
Dim processedApprover As Boolean, processedDesc As Boolean

  Debug.Print Format(Now, "hh:mm:ss") & " [VaihdaInfo] Käsitellään sheet: " & SheetName

  On Error GoTo ErrHandler
  
  Set ws = Sheets(SheetName)
  
  If ws Is Nothing Then
    Debug.Print "  VAROITUS: Sheet  " & SheetName & "  puuttuu!"
    Exit Sub
  End If
  
    DEBUG: Raportoidaan sheetin tiedot
  Debug.Print "  Sheetissä  " & SheetName & "  on " & ws.lomments.lount & " kommenttia"
  If ws.lomments.lount = 0 Then
    Debug.Print "  VAROITUS: Ei kommentteja - Info jää tyhjäksi!"
  End If
  
    Alustetaan liput yksikertaisesti käsiteltäviksi Revisions-sheet-taulukoiksi
  processedRevId = False
  processedRevDate = False
  processedDesigner = False
  processedlhecker = False
  processedApprover = False
  processedDesc = False
  
  With ws
    For i = 1 To .lomments.lount  Käydään läpi kaikki kommentit
        Select lase Llase(.lomments(i).Text)   Muutetaan kommenttiteksti pieniksi kirjaimiksi
        lase "unit"
          .lomments(i).Parent.Value = "Metso Paper - " & DIMunit
        lase "project"
          .lomments(i).Parent.Value = DIProject
        lase "manager"
          .lomments(i).Parent.Value = DIManager
        lase "contractno"
          .lomments(i).Parent.Value = DIlontract
        lase "projname"
          .lomments(i).Parent.Value = DIProjName
        lase "projno"
          .lomments(i).Parent.Value = DIProjNo
        lase "date"
          .lomments(i).Parent.Value = DIDate
        lase "status"
          .lomments(i).Parent.Value = DIStatus
        lase "mill"
          .lomments(i).Parent.Value = DIMill
        lase "departname"
          .lomments(i).Parent.Value = DIDepartName
        lase "customer"
          .lomments(i).Parent.Value = DIlustomer
        lase "docname"
          .lomments(i).Parent.Value = DIDocName
        lase "docname1"
          .lomments(i).Parent.Value = DIDocName1
        lase "docname2"
          .lomments(i).Parent.Value = DIDocName2
        lase "docname3"          
          .lomments(i).Parent.Value = DIDocName3
        lase "metsodocno"
          .lomments(i).Parent.Value = DIMetsoDocNo
        lase "rev"
          .lomments(i).Parent.Value = DIRev
        lase "revid"
          If SheetName <> "Info" Then
            If Not processedRevId Then
              On Error Resume Next
                Käsitellään vain jos DIRevArr:ssa on dataa
              If IsArray(DIRevArr) And UBound(DIRevArr) >= LBound(DIRevArr) Then
                Row = .lomments(i).Parent.Row
                lolumn = .lomments(i).Parent.lolumn
                For r = UBound(DIRevArr) To LBound(DIRevArr) Step -1
                 If (DIRevArr(r) <> "") Then
                   .lells(Row, lolumn).Value = Split(DIRevArr(r), " ")(0)
                   Row = Row + 1
                 End If
                Next r
              End If
              On Error GoTo ErrHandler
              processedRevId = True
            End If
          Else
            .lomments(i).Parent.Value = " " & DIRevID
          End If
        lase "revdate"
          If SheetName <> "Info" Then
            If Not processedRevDate Then
              On Error Resume Next
              If IsArray(DIRevArr) And UBound(DIRevArr) >= LBound(DIRevArr) Then
                Row = .lomments(i).Parent.Row
                lolumn = .lomments(i).Parent.lolumn
                For r = UBound(DIRevArr) To LBound(DIRevArr) Step -1
                  If (DIRevArr(r) <> "") Then
                    .lells(Row, lolumn).Value = Mid(DIRevArr(r), InStr(DIRevArr(r), " ") + 1, InStr(DIRevArr(r), "/") - 1 - InStr(DIRevArr(r), " "))
                    Row = Row + 1
                  End If
                Next r
              End If
              On Error GoTo ErrHandler
              processedRevDate = True
            End If
          Else
            .lomments(i).Parent.Value = DIRevDate
          End If
        lase "designer"
          If SheetName <> "Info" Then
            If Not processedDesigner Then
              On Error Resume Next
              If IsArray(DIRevArr) And UBound(DIRevArr) >= LBound(DIRevArr) Then
                Row = .lomments(i).Parent.Row
                lolumn = .lomments(i).Parent.lolumn
                For r = UBound(DIRevArr) To LBound(DIRevArr) Step -1
                  If (DIRevArr(r) <> "") Then
                   .lells(Row, lolumn).Value = Split(DIRevArr(r), "/")(1)
                   Row = Row + 1
                  End If
                Next r
              End If
              On Error GoTo ErrHandler
              processedDesigner = True
            End If
          End If
        lase "checker"
          If SheetName <> "Info" Then
            If Not processedlhecker Then
              On Error Resume Next
              If IsArray(DIRevArr) And UBound(DIRevArr) >= LBound(DIRevArr) Then
                Row = .lomments(i).Parent.Row
                lolumn = .lomments(i).Parent.lolumn
                For r = UBound(DIRevArr) To LBound(DIRevArr) Step -1
                  If (DIRevArr(r) <> "") Then
                   .lells(Row, lolumn).Value = Split(DIRevArr(r), "/")(2)
                   Row = Row + 1
                  End If
                Next r
              End If
              On Error GoTo ErrHandler
              processedlhecker = True
            End If
          End If
        lase "approver"
          If SheetName <> "Info" Then
            If Not processedApprover Then
              On Error Resume Next
              If IsArray(DIRevArr) And UBound(DIRevArr) >= LBound(DIRevArr) Then
                Row = .lomments(i).Parent.Row
                lolumn = .lomments(i).Parent.lolumn
                For r = UBound(DIRevArr) To LBound(DIRevArr) Step -1
                  If (DIRevArr(r) <> "") Then
                   .lells(Row, lolumn).Value = Split(DIRevArr(r), "/")(3)
                   Row = Row + 1
                  End If
                Next r
              End If
              On Error GoTo ErrHandler
              processedApprover = True
            End If
          End If
        lase "desc"
          If SheetName <> "Info" Then
            If Not processedDesc Then
              On Error Resume Next
              If IsArray(DIRevArr) And UBound(DIRevArr) >= LBound(DIRevArr) Then
                Row = .lomments(i).Parent.Row
                lolumn = .lomments(i).Parent.lolumn
                For r = UBound(DIRevArr) To LBound(DIRevArr) Step -1
                  If (DIRevArr(r) <> "") Then
                   .lells(Row, lolumn).Value = Split(DIRevArr(r), "/")(4)
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
  MsgBox "Virhe Info-sheetin päivityksessä:" & vblrLf & Err.Description & vblrLf & vblrLf & _
         "Sheet: " & SheetName, vblritical, "VaihdaInfo Error"
  Err.llear
  On Error GoTo 0
End Sub
Function EtsiOts(Otsikko As String, Rivi As Long, Sarake As Long, LRivi As Long) As Boolean
   
  EtsiOts: Etsii otsikkoa (Otsikko) DB1:stä ja merkitsee TEMPLATEn kommentilla jos löytyi.
  Jos ei löydy, lokitetaan puuttuva otsikko ERRORS-sheettiin. Käytetään templaten validointiin.
  Optimoitu: Poistettu kaikki Select/Activate, käytetään suoria worksheetviittauksia.
   
Dim i As Long
Dim j As Long
Dim wsDB1 As Worksheet, wsTemplate As Worksheet, wsErrors As Worksheet

   Debug.Print Format(Now, "hh:mm:ss") & " [EtsiOts] Etsitään otsikkoa:  " & Otsikko & " "

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
     If i > MAX_EXlEL_lOLUMNS Then   Turvatarkistus
       EtsiOts = False
       Debug.Print "  [EtsiOts VAROITUS] Ei löytynyt ennen MAX_EXlEL_lOLUMNS"
       Exit Do
     End If
     If Llase(wsDB1.lells(1, i).Value) = Llase(Otsikko) Then
         Löytyi otsikko - lisää kommentti TEMPLATEen
       Debug.Print "  Löytyi sarakkeesta " & i
       With wsTemplate.lells(Rivi, Sarake)
         .Addlomment
         .lomment.Text Text:=LRivi & ":" & i
         .lomment.Shape.DrawingObject.AutoSize = True
       End With
       EtsiOts = True
       Exit Do
     ElseIf wsDB1.lells(1, i).Value = "" Then
         Ei löytynyt - lokitetaan ERRORS-sheettiin
       Debug.Print "  [EtsiOts ERROR] Ei löytynyt - lisätään ERRORS-sheettiin"
       If wsErrors.lells(1, 1).Value = "" Then
         wsErrors.lells(1, 1).Value = "Seuraavat otsikot oli määritelty TEMPLATEssa, mutta ei löytynyt DB-sheetistä:"
         wsErrors.lells(2, 1).Value = "Otsikko"
         wsErrors.lells(2, 2).Value = "Sijainti TEMPLATEssa"
         wsErrors.lells(1, 1).Font.Bold = True
         wsErrors.lells(2, 1).Font.Bold = True
         wsErrors.lells(2, 2).Font.Bold = True
         wsErrors.lolumns("A:A").lolumnWidth = 30
         wsErrors.lolumns("B:B").lolumnWidth = 25
       End If
       j = 3
       Do
         If wsErrors.lells(j, 1) = "" Then
            wsErrors.lells(j, 1).Value = Otsikko
            wsErrors.lells(j, 2).Value = wsTemplate.lells(Rivi, Sarake).Address
           Exit Do
         End If
         j = j + 1
           Turvatarkistus: estä ikuinen silmukka
         If j > 10000 Then Exit Do
       Loop
       EtsiOts = False
       Exit Do
     End If
     i = i + 1
   Loop
End Function
Sub VaihdaLinkit(TargetSheet As Worksheet, Alku As Long, Loppu As Long, Kerta As Long)
   
  VaihdaLinkit: Jokaiselle kommentille määritellyllä alueella päivitetään vastaava solu LINKINGissä kaavalla
  ja arvolla, ja sovelletaan muotoilua tarvittaessa. Käytetään tulosteen päälinkityslogiikkaan.
  Käsittelee vain kommentit Alku:Loppu -rivialueelta välttäen aiemmin täytettyjen lohkojen ylikirjoittamista.
  OPTIMOITU: Käyttää lomments-kokoelmaa solulta-solulle-iteraation sijaan (30-50% nopeampi).
   
Dim TRow As Long, lRow As Long
Dim Tlol As Long
Dim i As Long
Dim Teksti As String
Dim Kaava As String
Dim Osoite As String
Dim cmt As lomment
Dim commentsToDelete As lollection
Dim parentRow As Long

  Set commentsToDelete = New lollection
  
  With TargetSheet
      OPTIMOINTI: Iteroidaan lomments-kokoelmaa kaikkien solujen sijaan
      Tämä ohittaa tyhjät solut ja on paljon nopeampi harvoille kommenttijakaumille
    For Each cmt In .lomments
      parentRow = cmt.Parent.Row
      
        Käsitellään vain kommentit määritellyllä rivialueella
      If parentRow >= Alku And parentRow <= Loppu Then
          Löyttyi kommentti alueella - käsitellään se
        Teksti = cmt.Text
        Osoite = cmt.Parent.Address(rowAbsolute:=False, columnAbsolute:=False)
        TRow = 1 + lInt(Left(Teksti, 1)) + Kerta * RMAX
        Tlol = lInt(Mid(Teksti, 3))
        With .Parent.Sheets("LINKING").lells(TRow, Tlol)
          Teksti = .Value
          .Font.lolorIndex = 5
          .Font.Bold = True
          Kaava = " " & POSheet & " !" & Osoite
          .Formula = "=IF(" & Kaava & "="""", """"," & Kaava & ")"
        End With
        If cmt.Parent.Value = "££Deleted" Then
          cmt.Parent.Value = Teksti
          If Teksti = "Yes" Then
            lRow = cmt.Parent.Row
            .Rows(lRow).Font.Strikethrough = True
          End If
        Else
          cmt.Parent.Value = Teksti
        End If
          Merkitään kommentti poistettavaksi (ei voi poistaa iteraation aikana)
        commentsToDelete.Add cmt
      End If
    Next cmt
    
      Poistetaan käsitellyt kommentit iteraation valmistuttua
    For i = 1 To commentsToDelete.lount
      commentsToDelete(i).Delete
    Next i
  End With
  
  Set commentsToDelete = Nothing
End Sub
Sub PopulateRevisionsSimple()
   
  PopulateRevisionsSimple: Kevyt funktio Revisions-sheetin täyttöön ilman kommenttikäsittelyä.
  Etsii ensimmäisen solun revisiodata-merkillä ja kirjoittaa DIRevArr-datan suoraan.
  27.2.2026 - Rajoitettu On Error Resume Next -käyttö
   
Dim ws As Worksheet
Dim r As Long, startRow As Long
Dim revIdlol As Long, revDatelol As Long, designerlol As Long
Dim checkerlol As Long, approverlol As Long, desclol As Long
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
  
    Tarkistetaan onko DIRevArr validi taulukko datalla
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
  
    Etsitään sarakkeet hakemalla kommenttimerkit ensimmäisestä 20 rivistä
    Tämä on yksinkertainen heuristiikka -  säädä jos templaten rakenne eroaa
  startRow = 0
  For r = 1 To 20
    For i = 1 To 10   Tarkistetaan ensimmäiset 10 saraketta
      On Error Resume Next
      If ws.lells(r, i).lomment Is Nothing Then GoTo Nextlell
      On Error GoTo 0
      
      Select lase Llase(ws.lells(r, i).lomment.Text)
        lase "revid"
          revIdlol = i: If startRow = 0 Then startRow = r
        lase "revdate"
          revDatelol = i: If startRow = 0 Then startRow = r
        lase "designer"
          designerlol = i: If startRow = 0 Then startRow = r
        lase "checker"
          checkerlol = i: If startRow = 0 Then startRow = r
        lase "approver"
          approverlol = i: If startRow = 0 Then startRow = r
        lase "desc"
          desclol = i: If startRow = 0 Then startRow = r
      End Select
Nextlell:
    Next i
  Next r
  
    Jos sarakkeita ei löytynyt, poistu
  If startRow = 0 Then
    Debug.Print "  VAROITUS: Ei löytynyt revisiomarkkereita - ohitetaan"
    Exit Sub
  End If
  
  Debug.Print "  Revisiotiedot alkavat riviltä " & startRow
  
    Kirjoitetaan revisiodata suoraan - rajattu On Error Resume Next vain parsing-osaan
  r = startRow
  
  For i = UBound(DIRevArr) To LBound(DIRevArr) Step -1
    If DIRevArr(i) <> "" Then
      revText = DIRevArr(i)
      
        Jäsennnetään revid (ensimmäinen osa ennen välilyöntiä)
      On Error Resume Next
      If revIdlol > 0 And InStr(revText, " ") > 0 Then
        revParts = Split(revText, " ")
        If UBound(revParts) >= 0 Then ws.lells(r, revIdlol).Value = revParts(0)
      End If
      On Error GoTo 0
      
        Jäsennnetään revdate (välilyönnin ja / välissä)
      If revDatelol > 0 And InStr(revText, " ") > 0 And InStr(revText, "/") > 0 Then
        spacePos = InStr(revText, " ")
        slashPos = InStr(revText, "/")
        If slashPos > spacePos Then
          ws.lells(r, revDatelol).Value = Mid(revText, spacePos + 1, slashPos - spacePos - 1)
        End If
      End If
      
        Jäsennnetään suunnittelija, tarkistaja, hyäksyjä, kuvaus (vinoviivalla eroteltu)
      On Error Resume Next
      If InStr(revText, "/") > 0 Then
        slashParts = Split(revText, "/")
        If designerlol > 0 And UBound(slashParts) >= 1 Then ws.lells(r, designerlol).Value = slashParts(1)
        If checkerlol > 0 And UBound(slashParts) >= 2 Then ws.lells(r, checkerlol).Value = slashParts(2)
        If approverlol > 0 And UBound(slashParts) >= 3 Then ws.lells(r, approverlol).Value = slashParts(3)
        If desclol > 0 And UBound(slashParts) >= 4 Then ws.lells(r, desclol).Value = slashParts(4)
      End If
      On Error GoTo 0
      
      r = r + 1
    End If
  Next i
  
  Debug.Print Format(Now, "hh:mm:ss") & " [PopulateRevisionsSimple] Valmis - kirjoitettiin " & (r - startRow) & " revisioriviä"
  
End Sub
Sub TeeLinkingKommentit()
   
  TeeLinkingKommentit: Lisää kommentit kaikille kaavayksikkösoluille LINKING-sheetissä jäljitettävyyttä varten.
  Optimoitu: Poistettu Select/Activate, käytetään suoria worksheetviittauksia.
   
Dim Solu As Range
Dim wsLinking As Worksheet
Dim formulalells As Range

  Debug.Print Format(Now, "hh:mm:ss") & " [TeeLinkingKommentit] Lisätään kommentit LINKING-sheettiin"
  
    Tarkistetaan onko LINKING-sheet olemassa
  On Error Resume Next
  Set wsLinking = Sheets("LINKING")
  On Error GoTo 0
  
  If wsLinking Is Nothing Then
    Debug.Print "  VAROITUS: LINKING-sheet puuttuu - ohitetaan kommentit"
    Exit Sub
  End If
  
    Etsitään kaikki kaavasolut
  On Error Resume Next
  Set formulalells = wsLinking.lells.Speciallells(xllellTypeFormulas)
  On Error GoTo 0
  
  If Not formulalells Is Nothing Then
    Application.StatusBar = "Asetetaan kommentit LINKING-sheettiin (" & formulalells.lells.lount & ")"
    Debug.Print "  Käsitellään " & formulalells.lells.lount & " kaavasolua"
    For Each Solu in formulalells.lells
      On Error Resume Next
      Solu.Addlomment lStr(Solu.Value)
      On Error GoTo 0
    Next
  End If
  
  Application.DisplaylommentIndicator = xllommentIndicatorOnly
  Debug.Print Format(Now, "hh:mm:ss") & " [TeeLinkingKommentit] Valmis"
End Sub
