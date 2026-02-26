Public CheckOK As Boolean
Public PHStart As Long
Public PHEnd As Long
Public PFStart As Long
Public PFEnd As Long
Public DocStart As Long
Public DocEnd As Long
Public Sarakkeita As Long
Public RMAX As Long
Public POSheet As String
Public HideLINKING As Boolean
Public AddFooter As Boolean
Public DIContract As String
Public DIMill As String
Public DIDepartName As String
Public DICustomer As String
Public DIProject As String
Public DIProjNo As String
Public DIProjName As String
Public DIMunit As String
Public DIManager As String
Public DIDocNo As String
Public DIMetsoDocNo As String
Public DIDocName As String
Public DIDocName1 As String
Public DIDocName2 As String
Public DIDocName3 As String
Public DIPath As String
Public DIFile As String
Public DIDate As String
Public DIRev As String
Public DIRevArr() As String
Public DIRevID As String
Public DIRevDate As String
Public DIStatus As String

'''
' Module1.vba - Päälogiikka Kytkentälista Excel-makrojärjestelmälle.
' Käsittelee datan haun Accessista, tarkistukset ja tulosteen generoinnin.
'''

' Suorituskyky/UX-tila (minimoi näytön välkkymisen ja nopeuttaa makroja)
Private prevScreenUpdating As Boolean
Private prevCalculation As XlCalculation
Private prevEnableEvents As Boolean
Private prevDisplayAlerts As Boolean
Private prevDisplayStatusBar As Boolean

'''
' BeginFastMode: Poistaa väliaikaisesti Excel UI-päivitykset, tapahtumat ja asettaa laskennan manuaaliseksi
' nopeuttaakseen makron suoritusta ja estääkseen näytön välkkymisen.
'''
Private Sub BeginFastMode()
  Debug.Print Format(Now, "hh:mm:ss") & " [BeginFastMode] Aktivoidaan nopea tila"
  prevScreenUpdating = Application.ScreenUpdating
  prevCalculation = Application.Calculation
  prevEnableEvents = Application.EnableEvents
  prevDisplayAlerts = Application.DisplayAlerts
  prevDisplayStatusBar = Application.DisplayStatusBar
  Application.ScreenUpdating = False
  Application.Calculation = xlCalculationManual
  Application.EnableEvents = False
  Application.DisplayAlerts = False
  On Error Resume Next
  Application.DisplayStatusBar = False
  Application.AskToUpdateLinks = False
  On Error GoTo 0
End Sub

'''
' EndFastMode: Palauttaa Excel UI:n ja laskenta-asetukset aiempaan tilaansa.
'''
Private Sub EndFastMode()
  Debug.Print Format(Now, "hh:mm:ss") & " [EndFastMode] Palautetaan normaalitila"
  On Error Resume Next
  Application.ScreenUpdating = prevScreenUpdating
  Application.Calculation = prevCalculation
  Application.EnableEvents = prevEnableEvents
  Application.DisplayAlerts = prevDisplayAlerts
  Application.DisplayStatusBar = prevDisplayStatusBar
  Application.AskToUpdateLinks = True
  On Error GoTo 0
End Sub

'''
' HaeData: Hakee datan Access-tietokannasta OLE DB:llä käyttäen faceplatessa määriteltyjä SQL-kyselyitä.
' Täyttää DB1 (pääasiallinen body-data) ja DB2 (dokumentin metadata) -sheetit.
' QueryTable elinkaari: Luo → Päivitä → Poista (ei jätä pysyviä yhteyksiä taustalle).
' Diagnostiikka: Rivimäärät näytetään StatusBarissa ja Immediate Windowissa jokaisen kyselyn jälkeen.
'''
Sub HaeData()
Dim Kanta As String
Dim sSQL(1 To 2) As String
Dim Valinta As Long
Dim i As Long
Dim TAULUKKO As QueryTable
Dim Yhteys As String
Dim ws As Worksheet
Dim rc As Long
Dim Provider As String

  Debug.Print Format(Now, "hh:mm:ss") & " [HaeData] Aloitetaan datan haku"
  
  ' Valitaan mikä SQL-kysely käytetään Main-sheetiltä
  On Error Resume Next
  If Sheets("Main").Valinta1.Value = True Then
    Valinta = 0
  ElseIf Sheets("Main").Valinta2.Value = True Then
    Valinta = 1
  Else
    Valinta = 2
  End If
  On Error GoTo 0
  
  Kanta = Sheets("Main").Range("C6").Value
  sSQL(1) = Sheets("Main").Cells(8 + Valinta, 3).Value
  sSQL(2) = Sheets("Main").Cells(12 + Valinta, 3).Value
  
  Debug.Print "  Tietokanta: " & Kanta
  Debug.Print "  SQL-valinta: " & Valinta
  
  BeginFastMode
  
  ' Varmistetaan että tietokantatiedosto on olemassa
  If Dir(Kanta) = "" Then
    MsgBox "Database file not found: " & Kanta, vbCritical, "Database Error"
    EndFastMode
    Exit Sub
  End If
  
  ' OLE DB yhteys ACE provider-fallbackilla (16.0 → 15.0 → 12.0)
  On Error Resume Next
  Provider = "Microsoft.ACE.OLEDB.16.0"
  Yhteys = "OLEDB;Provider=" & Provider & ";Data Source=" & Kanta
  
  ' Testataan yhteyttä - jos epäonnistuu, kokeillaan vanhempaa provideria
  Dim testConn As Object
  Set testConn = CreateObject("ADODB.Connection")
  testConn.Open Yhteys
  If Err.Number <> 0 Then
    Err.Clear
    Provider = "Microsoft.ACE.OLEDB.15.0"
    Yhteys = "OLEDB;Provider=" & Provider & ";Data Source=" & Kanta
    testConn.Open Yhteys
    If Err.Number <> 0 Then
      Err.Clear
      Provider = "Microsoft.ACE.OLEDB.12.0"
      Yhteys = "OLEDB;Provider=" & Provider & ";Data Source=" & Kanta
      testConn.Open Yhteys
    End If
  End If
  testConn.Close
  Set testConn = Nothing
  On Error GoTo ErrorHandler
  
  Debug.Print "  OLE DB Provider: " & Provider
  
  For i = 1 To 2
    Set ws = ThisWorkbook.Sheets("DB" & i)
    ws.Cells.Clear
    
    If sSQL(i) <> "" Then
      Debug.Print Format(Now, "hh:mm:ss") & " [HaeData] Haetaan DB" & i & " dataa..."
      ' Luodaan QueryTable, päivitetään, sitten poistetaan (ei pysyviä yhteyksiä)
      Set TAULUKKO = ws.QueryTables.Add(Connection:=Yhteys, Destination:=ws.Range("A1"))
      With TAULUKKO
        .CommandText = sSQL(i)
        .CommandType = xlCmdSql
        .FieldNames = True
        .RefreshStyle = xlInsertDeleteCells
        .RowNumbers = False
        .FillAdjacentFormulas = False
        .HasAutoFormat = True
        .SaveData = True
        .BackgroundQuery = False
        .Refresh
        .Delete
      End With
      Set TAULUKKO = Nothing
      
      ' Raportoidaan rivimäärä diagnostiikkaa varten
      rc = 0
      On Error Resume Next
      rc = ws.UsedRange.Rows.Count
      On Error GoTo 0
      Application.StatusBar = "DB" & i & " rows: " & rc
      Debug.Print "  DB" & i & " rivejä: " & rc
    End If
  Next i
  EndFastMode
  Debug.Print Format(Now, "hh:mm:ss") & " [HaeData] Valmis!"
  MsgBox "Data brought successfully!", vbOKOnly, "Ready"
  Sheets("Main").Select
  Exit Sub
  
ErrorHandler:
  EndFastMode
  Debug.Print Format(Now, "hh:mm:ss") & " [HaeData ERROR] " & Err.Number & ": " & Err.Description
  MsgBox "OLE DB Error: " & Err.Description & vbCrLf & vbCrLf & _
         "Database: " & Kanta & vbCrLf & _
         "Provider: " & Provider & vbCrLf & _
         "SQL Query " & i & ": " & sSQL(i), vbCritical, "Database Connection Error"
  Err.Clear
  Sheets("Main").Select
End Sub
'''
' GenPrintout: Generoi uuden tuloste-työkirjan käyttäen TEMPLATEa ja DB1-dataa.
' Käyttää template-pohjaista täyttöä: kopioi TEMPLATE-lohkoja per dataryhmä (RMAX riviä),
' sitten kartoittaa arvot LINKING-sheetiltä kommenttipohjaisten merkkien kautta (VaihdaLinkit).
' Tämä säilyttää templaten asettelun, muotoilun ja linkkauslogiikan.
'''
Sub GenPrintout()
  Debug.Print Format(Now, "hh:mm:ss") & " [GenPrintout] Aloitetaan tulosteen generointi"
  
  If CheckOK = False Then
    MsgBox "Check data first!", vbCritical, "Error!"
    Debug.Print Format(Now, "hh:mm:ss") & " [GenPrintout ERROR] CheckOK=False - keskeytetään"
    Exit Sub
  End If
  
  ' Suorituskyvyn mittausmuuttujat
  Dim perfStart As Double, perfTotal As Double
  Dim perfCopy As Double, perfLink As Double, perfShade As Double
  Dim perfIterations As Long
  perfStart = Timer
  perfCopy = 0
  perfLink = 0
  perfShade = 0
  perfIterations = 0
  
  ' Muuttujan määrittelyt
  Dim srcWB As Workbook
  Dim destWB As Workbook
  Dim destSheet As Worksheet
  Dim wsDB1 As Worksheet
  Dim ViimRivi As Long
  Dim Recordeja As Long
  Dim Kerta As Long
  Dim Tiedosto As String
  Dim i As Long
  Dim Oletus As String
  Dim lastCol As Long
  Dim c As Range
  Dim Riveja As Long
  
  On Error GoTo GenPrintoutError
  BeginFastMode
  
  ' Haetaan POSheet-nimi faceplatesta
  POSheet = Sheets("Main").Range("C16").Value
  If Trim(POSheet) = "" Then POSheet = "Printout" ' Oletusnimi jos ei asetettu
  Debug.Print "  POSheet nimi: " & POSheet
  
  ' Varmistetaan että dokumentin tiedot ovat ajantasalla (polku/nimi DB2:sta)
  On Error Resume Next
  If Trim(DIPath) = "" Or Trim(DIFile) = "" Then HaeDocTiedot
  On Error GoTo GenPrintoutError
  
  Application.StatusBar = "Alustetaan tulosteen generointi..."
    
  ' Haetaan käyttäjän asetukset faceplatesta
  AddFooter = Sheets("Main").AddFooter.Value
  On Error Resume Next
  HideLINKING = Sheets("Main").OLEObjects("HLINKING").Object.Value
  On Error GoTo GenPrintoutError
  Debug.Print "  AddFooter: " & AddFooter & ", HideLINKING: " & HideLINKING
    
  ' Asetetaan työkirjaviittaukset
  Set srcWB = ThisWorkbook
  Set wsDB1 = srcWB.Sheets("DB1")
  
  Application.StatusBar = "Luetaan dataa DB1:stä..."
  
  ' Etsitään viimeinen käytetty rivi DB1:stä - null-tarkistuksella
  Dim lastCell As Range
  Set lastCell = wsDB1.Cells.Find(What:="*", _
                                  After:=wsDB1.Range("A1"), _
                                  LookAt:=xlPart, _
                                  LookIn:=xlFormulas, _
                                  SearchOrder:=xlByRows, _
                                  SearchDirection:=xlPrevious, _
                                  MatchCase:=False)
  
  If lastCell Is Nothing Then
    EndFastMode
    Debug.Print Format(Now, "hh:mm:ss") & " [GenPrintout ERROR] DB1 tyhjä"
    MsgBox "DB1 sheet is empty! Please click 'Get Data' first to load data from database.", vbCritical, "No Data Error"
    Exit Sub
  End If
  
  Recordeja = lastCell.Row
  Debug.Print "  DB1 rivejä: " & Recordeja
  
  Application.StatusBar = "Luodaan uusi työkirja..."
  
  ' Luodaan uusi työkirja kopioimalla Info-sheet
  srcWB.Sheets("Info").Copy
  Set destWB = ActiveWorkbook
  destWB.Sheets(1).Cells.ClearComments
  
  ' Kopioidaan TEMPLATE, Legend ja Revisions uuteen työkirjaan
  srcWB.Sheets("TEMPLATE").Copy After:=destWB.Sheets(1)
  destWB.Sheets(2).Name = POSheet
  Set destSheet = destWB.Sheets(POSheet)
  
  srcWB.Sheets("Legend").Copy After:=destWB.Sheets(2)
  srcWB.Sheets("Revisions").Copy After:=destWB.Sheets(1)
  
  Application.StatusBar = "Täytetään revisiot..."
  PopulateRevisionsSimple
  
  ' Tyhjennetään POSheet ja poistetaan muodot
  Application.StatusBar = "Valmistellaan tulostesheet..."
  Debug.Print "  Luotu sheetit: Info, " & POSheet & ", Legend, Revisions"
  destSheet.Cells.Clear
  If destSheet.Shapes.Count > 0 Then
    For i = destSheet.Shapes.Count To 1 Step -1
      destSheet.Shapes(i).Delete
    Next i
  End If
  
  ' Kopioidaan otsikkorivit TEMPLATEsta POSheetiin
  Application.StatusBar = "Kopioidaan otsikot..."
  Debug.Print "  Template alueet: PH=" & PHStart & ":" & PHEnd & ", Doc=" & DocStart & ":" & DocEnd & ", PF=" & PFStart & ":" & PFEnd
  ViimRivi = 1
  srcWB.Sheets("TEMPLATE").Rows(PHStart & ":" & PHEnd).Copy _
      Destination:=destSheet.Rows(ViimRivi & ":" & ViimRivi + PHEnd - PHStart)
  Application.CutCopyMode = False
  
  ' Asetetaan tulostuksen otsikkorivit ja jäädytetään paneelit
  destSheet.PageSetup.PrintTitleRows = "$" & ViimRivi & ":$" & ViimRivi + PHEnd - PHStart
  destSheet.Activate
  destSheet.Cells(ViimRivi + PHEnd - PHStart + 1, 1).Select
  ActiveWindow.FreezePanes = True
  ViimRivi = ViimRivi + 1 + PHEnd - PHStart
  
  ' Asetetaan alatunnisteet kolmelle ensimmäiselle sheetille (Info, POSheet, Legend)
  Application.StatusBar = "Asetetaan alatunnisteet..."
  Debug.Print "  Asennetaan alatunnisteet dokumenttitiedoilla"
  For i = 1 To 3
    With destWB.Sheets(i).PageSetup
      ' Null-turvallinen merkkijonojen yhdistäminen (tyhjät arvot OK - joissakin laitteissa puuttuu tiettyjä attribuutteja)
      .LeftFooter = "&8Document: " & (DIMetsoDocNo & "") & Chr(10) _
                  & "&8Revision: " & (DIRevID & "") & " - " & (DIRevDate & "") & Chr(10) _
                  & "&8Status: " & (DIStatus & "")
      .CenterFooter = "&8 " & (DICustomer & "") & Chr(10) _
                    & "&8 " & (DIMill & "") & Chr(10) _
                    & "&8 " & (DIDepartName & "") & Chr(10) _
                    & "&8 " & (DIDocName2 & "")
      .RightFooter = "&8Project: " & (DIProjNo & "") & Chr(10) _
                   & "&8File: &F" & Chr(10) _
                   & "&8Page &P(&N)"
    End With
  Next i
  
  ' Luodaan LINKING-sheet ja kopioidaan DB1-data
  Application.StatusBar = "Luodaan LINKING-sheet..."
  Debug.Print "  Luodaan LINKING-sheet DB1-datalla"
  With destWB.Sheets.Add(After:=destWB.Sheets(destWB.Sheets.Count))
    .Name = "LINKING"
    wsDB1.Cells.Copy Destination:=.Range("A1")
  End With
  Application.CutCopyMode = False
  
  ' Alkuperäinen linkitys otsikkoalueelle
  Kerta = 0
  VaihdaLinkit destSheet, 1, ViimRivi, Kerta
  
  ' Template-pohjainen täyttö: kopioidaan TEMPLATE-lohkoja ja kartoitetaan arvot VaihdaLinkit-funktiolla
  ' OPTIMOITU: Vähennetään työkirjojen välisiä kopioita käyttämällä väliaikaista range-aluetta + batch clipboard clears
  Application.StatusBar = "Kopioidaan dataa tulosteeseen käyttäen template-lohkoja..."
  Debug.Print "  Aloitetaan template-lohkojen kopiointi (RMAX=" & RMAX & ")"
  Riveja = DocEnd - DocStart
  If RMAX <= 0 Then RMAX = 1
  
  ' Esikopioidaan TEMPLATE-lohko kohteeseen kerran (työkirjojen välinen kopiointi on hidasta)
  Dim templateRange As Range
  Set templateRange = srcWB.Sheets("TEMPLATE").Rows(DocStart & ":" & DocEnd)
  
  ' Iteroidaan DB1-datarivejä RMAX-ryhmissä, kopioidaan TEMPLATE-rivit joka kerralla
  Kerta = 0
  For i = 2 To Recordeja Step RMAX
    perfIterations = perfIterations + 1
    
    ' OPTIMOINTI: Kopioidaan lähdetyökirjasta
    ' (Huom: Ei voi täysin optimoida ilman array-lähestymistapaa koska tarvitaan muotoilua)
    Dim tCopy As Double: tCopy = Timer
    templateRange.Copy Destination:=destSheet.Rows(ViimRivi & ":" & ViimRivi + Riveja)
    perfCopy = perfCopy + (Timer - tCopy)
    
    ' Lisätään vuorottelevatvarjostukset lohkoittain
    Dim tShade As Double: tShade = Timer
    If ((i - 2) \ RMAX) Mod 2 = 1 Then
      With destSheet.Range(destSheet.Cells(ViimRivi, 1), destSheet.Cells(ViimRivi + Riveja, Sarakkeita)).Interior
        .ColorIndex = 19
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
      End With
    End If
    perfShade = perfShade + (Timer - tShade)
    
    ' Kartoitetaan arvot LINKINGistä template-alueelle kommenttimerkkien kautta
    Dim tLink As Double: tLink = Timer
    VaihdaLinkit destSheet, ViimRivi, ViimRivi + Riveja, Kerta
    perfLink = perfLink + (Timer - tLink)
    
    ' Siirrytään seuraavaan lohkoon
    ViimRivi = ViimRivi + Riveja + 1
    Kerta = Kerta + 1
  Next i
  Debug.Print "  Kopioitu " & perfIterations & " template-lohkoa"
  
  ' OPTIMOINTI: Tyhjennetään leikepöytä kerran kaikkien kopioiden jälkeen (ei loopissa)
  Application.CutCopyMode = False
  
  ' Poistetaan ylimääräiset sarakkeet Sarakkeita-määrän jälkeen
  lastCol = destSheet.Cells(1, destSheet.Columns.Count).End(xlToLeft).Column
  If lastCol > Sarakkeita Then
    destSheet.Columns(Sarakkeita + 1 & ":" & lastCol).Delete
  End If
  
  ' Lisätään alatunnistelohko jos pyydetty
  If AddFooter = True Then
    Application.StatusBar = "Lisätään alatunniste..."
    Debug.Print "  Lisätään alatunniste"
    srcWB.Sheets("TEMPLATE").Rows(PFStart & ":" & PFEnd).Copy _
        Destination:=destSheet.Rows(ViimRivi & ":" & ViimRivi + PFEnd - PFStart)
    Application.CutCopyMode = False
    
    ' Asetetaan kaavat alatunnisteen yhteenvetosoluille
    For Each c In destSheet.Range(destSheet.Cells(ViimRivi, 1), destSheet.Cells(ViimRivi + PFEnd - PFStart, Sarakkeita))
      If Len(CStr(c.Value)) >= 2 Then
        If Left(CStr(c.Value), 2) = "&&" Then
          c.Formula = "=SUM(" & destSheet.Cells(PHEnd, c.Column).Address(False, False) & ":" & destSheet.Cells(ViimRivi - 1, c.Column).Address(False, False) & ")"
        End If
      End If
    Next c
  End If
  
  ' Tyhjennetään kommentit ja lisätään LINKING-kommentit
  Application.StatusBar = "Viimeistellään..."
  destSheet.Cells.ClearComments
  TeeLinkingKommentit
  
  ' Käsitellään LINKING-sheetin näkyvyys/poisto
  On Error Resume Next
  If HideLINKING Then
    destWB.Sheets("LINKING").Visible = False
    Debug.Print "  LINKING-sheet piilotettu"
  Else
    Application.DisplayAlerts = False
    destWB.Sheets("LINKING").Delete
    Debug.Print "  LINKING-sheet poistettu"
    Application.DisplayAlerts = True
  End If
  On Error GoTo GenPrintoutError
  
  destSheet.Activate
  Application.StatusBar = False
  EndFastMode
  
  ' Kysytään tiedostonimeä ja tallennetaan
  ' Rakennetaan luotettava oletus polku+tiedosto tallennusdialogille
  Dim defPath As String, defName As String
  
  ' Turvallinen käsittely mahdollisesti tyhjille/null DIPath ja DIFile -arvoille
  On Error Resume Next
    defPath = Trim(DIPath & "")
    defName = Trim(DIFile & "")
  On Error GoTo GenPrintoutError
  
  If defPath = "" Then defPath = ThisWorkbook.Path & Application.PathSeparator
  If Right$(defPath, 1) <> "\\" And Right$(defPath, 1) <> "/" Then defPath = defPath & Application.PathSeparator
  
  ' Vaatimuksen mukaan: tiedostonimi tulee DB2 "File"-sarakkeesta
  If defName = "" Then defName = POSheet ' varatieto vain jos DB2 ei tarjonnut nimeä
  If defName = "" Then defName = "Printout" ' Lopullinen varatieto
  
  ' Varmistetaan .xlsx-pääte jos ei ole tarjottu
  If InStrRev(defName, ".") = 0 Then defName = defName & ".xlsx"
  
  Debug.Print "  Ehdotettu tiedostonimi: " & defPath & defName
  Oletus = defPath & defName
  Tiedosto = InputBox("Give The File Name", "Save File", Oletus)
  If Tiedosto <> "" Then
    destWB.BuiltinDocumentProperties("Title").Value = POSheet
    destWB.SaveAs Tiedosto, xlOpenXMLWorkbook
    Debug.Print Format(Now, "hh:mm:ss") & " [GenPrintout] Tallennettu: " & Tiedosto
  Else
    Debug.Print Format(Now, "hh:mm:ss") & " [GenPrintout] Tallennus peruttu"
  End If
  
  ' Suorituskykydiagnostiikka
  perfTotal = Timer - perfStart
  Debug.Print "=== GenPrintout Suorituskykyraportti ==="
  Debug.Print "Kokonaisaika: " & Format(perfTotal, "0.00") & "s"
  Debug.Print "Iteraatioita: " & perfIterations & " (RMAX=" & RMAX & ", Rivejä=" & Recordeja & ")"
  Debug.Print "  Kopiointiaika: " & Format(perfCopy, "0.00") & "s (" & Format(perfCopy / perfTotal * 100, "0.0") & "%)"
  Debug.Print "  Linkitysaika: " & Format(perfLink, "0.00") & "s (" & Format(perfLink / perfTotal * 100, "0.0") & "%)"
  Debug.Print "  Varjostusaika: " & Format(perfShade, "0.00") & "s (" & Format(perfShade / perfTotal * 100, "0.0") & "%)"
  Debug.Print "  Muu: " & Format(perfTotal - perfCopy - perfLink - perfShade, "0.00") & "s"
  If perfIterations > 0 Then
    Debug.Print "Keskiarvo per iteraatio: " & Format(perfTotal / perfIterations * 1000, "0.0") & "ms"
    Debug.Print "  Kopiointi: " & Format(perfCopy / perfIterations * 1000, "0.0") & "ms"
    Debug.Print "  Linkitys: " & Format(perfLink / perfIterations * 1000, "0.0") & "ms"
    Debug.Print "  Varjostus: " & Format(perfShade / perfIterations * 1000, "0.0") & "ms"
  End If
  Debug.Print "======================================"
  Debug.Print Format(Now, "hh:mm:ss") & " [GenPrintout] VALMIS!"
  
  Exit Sub

GenPrintoutError:
  Application.StatusBar = False
  EndFastMode
  
  ' Parannettu virhekäsittelijä kontekstispesifeillä viesteillä
  Dim errMsg As String
  errMsg = "Error in GenPrintout: " & Err.Description & " (Error " & Err.Number & ")"
  
  ' Lisätään kontekstia virhekoodin perusteella
  Select Case Err.Number
    Case 91
      errMsg = errMsg & vbCrLf & vbCrLf & "Objektimuuttujaa ei ole asetettu." & vbCrLf & _
               "Tämä yleensä tarkoittaa että sheet tai range puuttuu." & vbCrLf & _
               "Tarkista että kaikki vaaditut sheetit ovat olemassa (TEMPLATE, DB1, DB2, Info, Legend, Revisions)."
    Case 1004
      errMsg = errMsg & vbCrLf & vbCrLf & "Sovelluksen tai objektin määrittelemä virhe." & vbCrLf & _
               "Tämä usein tarkoittaa että kopiointi/liittämisoperaatio epäonnistui tai sheetin nimi on virheellinen."
    Case 9
      errMsg = errMsg & vbCrLf & vbCrLf & "Indeksi alueen ulkopuolella." & vbCrLf & _
               "Sheet määritellyllä nimellä ei ole olemassa."
    Case 13
      errMsg = errMsg & vbCrLf & vbCrLf & "Tyyppien yhteensopimattomuus." & vbCrLf & _
               "Yritetään käyttää yhteensopimattomia tietotyyppejä (esim. tekstiä missä odotetaan numeroa)."
  End Select
  
  MsgBox errMsg, vbCritical, "Printout Generation Error"
  
  ' Lokitetaan Immediate Windowiin debuggausta varten
  Debug.Print "GenPrintout ERROR:"
  Debug.Print "  Number: " & Err.Number
  Debug.Print "  Description: " & Err.Description
  Debug.Print "  Source: " & Err.Source
  
  Err.Clear
  On Error GoTo 0
End Sub
'''
' Checkout: Validoi TEMPLATE-rakenteen ja DB1-otsikot.
' - Etsii aluemerkit (&&PAGE_HEADER_START, &&DOC_DATA_START, jne.) TEMPLATEsta
' - Validoi rivimerkit (££ yksirivisille, £1/2/3 monirivisille ryhmille)
' - Luo kommentit TEMPLATE-soluihin linkittääkseen DB1-sarakkeisiin (EtsiOts-funktion kautta)
' - Täyttää Info-sheetin dokumentin metadatalla DB2:sta
' - Asettaa CheckOK-lipun jos validointi läpäisee, muuten raportoi virheet ERRORS-sheettiin
'''
Sub Checkout()
  Debug.Print Format(Now, "hh:mm:ss") & " [Checkout] Aloitetaan validointi"
Dim i As Long
Dim j As Long
Dim Arvo As String
Dim Virhe As Boolean
Dim wsTemplate As Worksheet
Dim wsErrors As Worksheet

  On Error GoTo CheckoutError
  
  CheckOK = False
  RMAX = 0
  Virhe = False
  Application.ScreenUpdating = False
  
  Set wsErrors = Sheets("ERRORS")
  Set wsTemplate = Sheets("TEMPLATE")
  
  ' Tyhjennetään ERRORS-sheet
  wsErrors.Cells.Clear
  
  ' Haetaan vakiot faceplatesta
  POSheet = Sheets("Main").Range("C16").Value
  On Error Resume Next
  HideLINKING = Sheets("Main").OLEObjects("HLINKING").Object.Value
  On Error GoTo CheckoutError
  Debug.Print "  POSheet: " & POSheet & ", HideLINKING: " & HideLINKING
  
  ' Tyhjennetään kommentit TEMPLATEsta
  wsTemplate.Cells.ClearComments
  
  ' Etsitään aluemerkit TEMPLATEsta - null-tarkistuksella
  Debug.Print "  Etsitään template-merkit..."
  Dim foundCell As Range
  With wsTemplate
    Set foundCell = .Cells.Find(What:="&&PAGE_HEADER_START")
    If foundCell Is Nothing Then
      wsErrors.Range("A1").Value = "TEMPLATE ERROR: Marker &&PAGE_HEADER_START not found!"
      wsErrors.Range("A1").Font.Bold = True
      wsErrors.Range("A1").Font.ColorIndex = 3
      wsErrors.Activate
      Application.ScreenUpdating = True
      Debug.Print Format(Now, "hh:mm:ss") & " [Checkout ERROR] &&PAGE_HEADER_START puuttuu"
      MsgBox "TEMPLATE is missing required markers! See ERRORS sheet.", vbCritical, "Template Error"
      Exit Sub
    End If
    PHStart = foundCell.Row + 1
    
    Set foundCell = .Cells.Find(What:="&&PAGE_HEADER_END")
    If foundCell Is Nothing Then Err.Raise vbObjectError + 1, , "&&PAGE_HEADER_END not found"
    PHEnd = foundCell.Row - 1
    
    Set foundCell = .Cells.Find(What:="&&DOC_DATA_START")
    If foundCell Is Nothing Then Err.Raise vbObjectError + 1, , "&&DOC_DATA_START not found"
    DocStart = foundCell.Row + 1
    
    Set foundCell = .Cells.Find(What:="&&DOC_DATA_END")
    If foundCell Is Nothing Then Err.Raise vbObjectError + 1, , "&&DOC_DATA_END not found"
    DocEnd = foundCell.Row - 1
    
    Set foundCell = .Cells.Find(What:="&&END")
    If foundCell Is Nothing Then Err.Raise vbObjectError + 1, , "&&END not found"
    Sarakkeita = foundCell.Column
    
    Set foundCell = .Cells.Find(What:="&&PAGE_FOOTER_START")
    If foundCell Is Nothing Then Err.Raise vbObjectError + 1, , "&&PAGE_FOOTER_START not found"
    PFStart = foundCell.Row + 1
    
    Set foundCell = .Cells.Find(What:="&&PAGE_FOOTER_END")
    If foundCell Is Nothing Then Err.Raise vbObjectError + 1, , "&&PAGE_FOOTER_END not found"
    PFEnd = foundCell.Row - 1
  End With
  Debug.Print "  Template-merkit löydetty: PH=" & PHStart & ":" & PHEnd & ", Doc=" & DocStart & ":" & DocEnd & ", PF=" & PFStart & ":" & PFEnd & ", Cols=" & Sarakkeita
  
  ' Haetaan dokumentin tiedot DB2-sheetiltä
  HaeDocTiedot
  
  ' Tarkistetaan ladattiinko dataa DB2:sta
  If DIProject = "" And DIDocNo = "" And DIProjNo = "" And DIMetsoDocNo = "" Then
    wsErrors.Range("A1").Value = "WARNING: No document metadata found in DB2 sheet!"
    wsErrors.Range("A2").Value = "Please click 'Get Data' button first to load data from database."
    wsErrors.Range("A1").Font.Bold = True
    wsErrors.Range("A1").Font.ColorIndex = 3 ' Punainen
    Debug.Print "  VAROITUS: Ei dataa DB2:ssa - Info-sheet jää tyhjäksi"
  End If
  
  VaihdaInfo   'Täytetään dokumentin tiedot vain Info-sheettiin (ei Revisions checkout-vaiheessa)
  
  ' Etsitään rivimerkit TEMPLATEsta
  Debug.Print "  Etsitään rivimerkit (££ ja £1/2/3)..."
  For i = DocStart To DocEnd
    For j = 1 To Sarakkeita
      Arvo = wsTemplate.Cells(i, j).Value
      If Len(Arvo) > 2 Then 'Solussa on dataa
        If Left(Arvo, 2) = "££" Then
          If RMAX > 1 Then Virhe = True
          RMAX = 1
        ElseIf Left(Arvo, 1) = "£" Then
          If RMAX <> 0 And RMAX <> CInt(Mid(Arvo, 4, 1)) Then Virhe = True
          RMAX = CInt(Mid(Arvo, 4, 1))
        End If
      End If
    Next j
  Next i
  Debug.Print "  RMAX määritetty: " & RMAX
  
  If Virhe Then
    wsErrors.Cells.Clear
    wsErrors.Range("A1").Value = "Virhe määrittelyssä!"
    wsErrors.Range("A1").Font.Bold = True
    wsErrors.Range("A2").Value = "- Et voi käyttää ££ ja £1/2 linkkejä samassa templatessa."
    wsErrors.Range("A3").Value = "- Et myöskään voi käyttää £1/2 ja £1/3 linkkejä samassa templatessa."
    wsErrors.Range("A4").Value = "- Korjaa nämä virheet ja yritä uudelleen!"
    wsErrors.Activate
    Application.ScreenUpdating = True
    Debug.Print Format(Now, "hh:mm:ss") & " [Checkout ERROR] Ristiriitaiset rivimerkit"
    MsgBox "Templatessa oli virheitä, katso ERRORS-sheet!", vbCritical, "Virhe!"
    Exit Sub
  End If
  
  ' Rivimerkit olivat oikein, nyt etsitään otsikot
  Debug.Print "  Tarkistetaan otsikot DB1-sheetistä..."
  For i = DocStart To DocEnd
    For j = 1 To Sarakkeita
      Arvo = wsTemplate.Cells(i, j).Value
      If Len(Arvo) > 2 Then 'Solussa on dataa
        If Left(Arvo, 2) = "££" Then
          If EtsiOts(Mid(Arvo, 3), i, j, 1) = False Then Virhe = True
        ElseIf Left(Arvo, 1) = "£" Then
          If EtsiOts(Mid(Arvo, 5), i, j, CInt(Mid(Arvo, 2, 1))) = False Then Virhe = True
        End If
      End If
    Next j
  Next i
  
  If Virhe Then
    wsErrors.Activate
    Application.ScreenUpdating = True
    Debug.Print Format(Now, "hh:mm:ss") & " [Checkout ERROR] Puuttuvia otsikkoita DB1:ssä"
    MsgBox "Templatessa oli virheitä! Katso ERRORS-sheet.", vbCritical, "Virhe!"
  Else
    Sheets("Main").Activate
    Application.ScreenUpdating = True
    CheckOK = True
    Debug.Print Format(Now, "hh:mm:ss") & " [Checkout] VALMIS - CheckOK=True"
    MsgBox "Tarkistus OK!", vbOKOnly, "OK!"
  End If
  Exit Sub

CheckoutError:
  Application.ScreenUpdating = True
  
  Dim errMsg As String
  errMsg = "Error in Checkout: " & Err.Description & " (Error " & Err.Number & ")"
  
  ' Lisätään hyödyllinen konteksti yleisille virheille
  If Err.Number = vbObjectError + 1 Then
    errMsg = errMsg & vbCrLf & vbCrLf & "TEMPLATE-sheetistä puuttuu vaadittu merkki." & vbCrLf & _
             "Varmista että TEMPLATE sisältää kaikki merkit:" & vbCrLf & _
             "&&PAGE_HEADER_START, &&PAGE_HEADER_END" & vbCrLf & _
             "&&DOC_DATA_START, &&DOC_DATA_END" & vbCrLf & _
             "&&PAGE_FOOTER_START, &&PAGE_FOOTER_END" & vbCrLf & _
             "&&END"
  ElseIf Err.Number = 91 Then
    errMsg = errMsg & vbCrLf & vbCrLf & "Tämä yleensä tarkoittaa että vaadittu sheet tai objekti puuttuu." & vbCrLf & _
             "Varmista että TEMPLATE, DB1, DB2, ERRORS ja Info -sheetit ovat olemassa."
  End If
  
  MsgBox errMsg, vbCritical, "Checkout Error"
  
  ' Lokitetaan Immediate Windowiin
  Debug.Print "Checkout ERROR:"
  Debug.Print "  Number: " & Err.Number
  Debug.Print "  Description: " & Err.Description
  
  Err.Clear
  On Error GoTo 0
End Sub

