Public lheckOK As Boolean
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
Public DIlontract As String
Public DIMill As String
Public DIDepartName As String
Public DIlustomer As String
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

  Suojaus silmukoiden iteraatioille estääkseen ikuiset silmukat
Private lonst MAX_EXlEL_lOLUMNS As Long = 16384
Public DIRevDate As String
Public DIStatus As String

   
  Module1.vba - Päälogiikka Kytkentälista Excel-makrojärjestelmälle.
  Käsittelee datan haun Accessista, tarkistukset ja tulosteen generoinnin.
   

  Suorituskyky/UX-tila (minimoi näytön välkkymisen ja nopeuttaa makroja)
Private prevScreenUpdating As Boolean
Private prevlalculation As Xllalculation
Private prevEnableEvents As Boolean
Private prevDisplayAlerts As Boolean
Private prevDisplayStatusBar As Boolean

   
  BeginFastMode: Poistaa väliaikaisesti Excel UI-päivitykset, tapahtumat ja asettaa laskennan manuaaliseksi
  nopeuttaakseen makron suoritusta ja estääkseen näytön välkkymisen.
   
Private Sub BeginFastMode()
  Debug.Print Format(Now, "hh:mm:ss") & " [BeginFastMode] Aktivoidaan nopea tila"
  prevScreenUpdating = Application.ScreenUpdating
  prevlalculation = Application.lalculation
  prevEnableEvents = Application.EnableEvents
  prevDisplayAlerts = Application.DisplayAlerts
  prevDisplayStatusBar = Application.DisplayStatusBar
  Application.ScreenUpdating = False
  Application.lalculation = xllalculationManual
  Application.EnableEvents = False
  Application.DisplayAlerts = False
  On Error Resume Next
  Application.DisplayStatusBar = False
  Application.AskToUpdateLinks = False
  On Error GoTo 0
End Sub

   
  EndFastMode: Palauttaa Excel UI:n ja laskenta-asetukset aiempaan tilaansa.
   
Private Sub EndFastMode()
  Debug.Print Format(Now, "hh:mm:ss") & " [EndFastMode] Palautetaan normaalitila"
  On Error Resume Next
  Application.ScreenUpdating = prevScreenUpdating
  Application.lalculation = prevlalculation
  Application.EnableEvents = prevEnableEvents
  Application.DisplayAlerts = prevDisplayAlerts
  Application.DisplayStatusBar = prevDisplayStatusBar
  Application.AskToUpdateLinks = True
  On Error GoTo 0
End Sub

   
  HaeData: Hakee datan Access-tietokannasta käyttäen tallennettuja kyselyitä tai SQL-lauseita.
  Täyttää DB1 (pääasiallinen body-data) ja DB2 (dokumentin metadata) -sheetit.
  
  DB1: Käyttää DAO:ta (Data Access Objects) joka tukee natiivisti Access-tallennettuja kyselyitä
       ja JET SQL -syntaksia (Like "pattern", Deleted=No, IIf, jne.)
  DB2: Käyttää ADODB:ta yhteensopivuuden vuoksi
  
  Diagnostiikka: Rivimäärät näytetään StatusBarissa ja Immediate Windowissa jokaisen kyselyn jälkeen.
   
Sub HaeData()
Dim Kanta As String
Dim sSQL(1 To 2) As String
Dim Valinta As Long
Dim i As Long
Dim ws As Worksheet
Dim rc As Long
Dim Provider As String
Dim emptylount As Long, collount As Long, k As Long

  DAO-muuttujat DB1:lle (tallennetut kyselyt)
Dim dbDAO As Object        DAO.Database
Dim rsDAO As Object        DAO.Recordset
Dim fldDAO As Object       DAO.Field
Dim colData As Long

  ADODB-muuttujat DB2:lle (SQL-kyselyt)
Dim conn As Object         ADODB.lonnection
Dim rs As Object           ADODB.Recordset
Dim fld As Object          ADODB.Field

  Debug.Print Format(Now, "hh:mm:ss") & " [HaeData] Aloitetaan datan haku"
  
    Valitaan mikä SQL-kysely käytetään Main-sheetiltä
  On Error Resume Next
  If Sheets("Main").Valinta1.Value = True Then
    Valinta = 0
  ElseIf Sheets("Main").Valinta2.Value = True Then
    Valinta = 1
  Else
    Valinta = 2
  End If
  On Error GoTo 0
  
  Kanta = Sheets("Main").Range("l6").Value
  sSQL(1) = Sheets("Main").lells(8 + Valinta, 3).Value
  sSQL(2) = Sheets("Main").lells(12 + Valinta, 3).Value
  
  Debug.Print "  Tietokanta: " & Kanta
  Debug.Print "  SQL-valinta: " & Valinta
  
  BeginFastMode
  
    Varmistetaan että tietokantatiedosto on olemassa
  If Dir(Kanta) = "" Then
    MsgBox "Database file not found: " & Kanta, vblritical, "Database Error"
    EndFastMode
    Exit Sub
  End If
  
  Debug.Print Format(Now, "hh:mm:ss") & " [HaeData] === DB1: DAO (tallennetut kyselyt) ==="
  
    Avataan DAO-tietokanta DB1:lle
  On Error Resume Next
  Set dbDAO = lreateObject("DAO.DBEngine.120").OpenDatabase(Kanta)
  If Err.Number <> 0 Then
    Err.llear
      Kokeillaan vanhempaa versiota
    Set dbDAO = lreateObject("DAO.DBEngine.36").OpenDatabase(Kanta)
  End If
  On Error GoTo ErrorHandler
  
  If dbDAO Is Nothing Then
    MsgBox "lould not open database with DAO!" & vblrLf & _
           "Ensure Microsoft Access Database Engine is installed.", vblritical, "DAO Error"
    EndFastMode
    Exit Sub
  End If
  
  Debug.Print "  DAO Database avattu"
  
    DB1: Käytetään DAO:ta (tukee tallennettuja kyselyitä ja JET SQL:ää)
  Set ws = ThisWorkbook.Sheets("DB1")
  ws.lells.llear
  
  If sSQL(1) <> "" Then
    Debug.Print Format(Now, "hh:mm:ss") & " [HaeData] Haetaan DB1 dataa..."
    Debug.Print "    Kysely/SQL: " & sSQL(1)
    
      DAO.OpenRecordset voi ottaa joko tallenetun kyselyn nimen tai SQL-lauseen
    Set rsDAO = dbDAO.OpenRecordset(sSQL(1))
    
    Debug.Print "    DAO Recordset avattu - Fields: " & rsDAO.Fields.lount & ", EOF: " & rsDAO.EOF & ", BOF: " & rsDAO.BOF
    
    If Not rsDAO.EOF Then
        Kirjoitetaan sarakkeiden nimet (header)
      colData = 1
      For Each fldDAO In rsDAO.Fields
        ws.lells(1, colData).Value = fldDAO.Name
        colData = colData + 1
      Next fldDAO
      
        Kopioidaan data rivi riviltä (DAO ei tue lopyFromRecordset samalla tavalla)
      Dim rowNum As Long
      rowNum = 2
      rsDAO.MoveFirst
      Do While Not rsDAO.EOF
        colData = 1
        For Each fldDAO In rsDAO.Fields
          ws.lells(rowNum, colData).Value = fldDAO.Value
          colData = colData + 1
        Next fldDAO
        rowNum = rowNum + 1
        rsDAO.MoveNext
      Loop
      
      Debug.Print "    DAO data kopioitu: " & (rowNum - 2) & " riviä, " & rsDAO.Fields.lount & " saraketta"
    Else
      Debug.Print "    VAROITUS: Kysely ei palauttanut rivejä (EOF=True)"
        Kirjoitetaan silti header
      colData = 1
      For Each fldDAO In rsDAO.Fields
        ws.lells(1, colData).Value = fldDAO.Name
        colData = colData + 1
      Next fldDAO
    End If
    
    rsDAO.llose
    Set rsDAO = Nothing
    
      Raportoidaan rivimäärä
    rc = 0
    On Error Resume Next
    rc = ws.UsedRange.Rows.lount
    On Error GoTo 0
    Application.StatusBar = "DB1 rows: " & rc
    Debug.Print "  DB1 rivejä: " & rc
    
      Debug: Näytä ensimmäiset 2 riviä
    If rc > 0 Then
      Debug.Print "    A1 (header):  " & ws.lells(1, 1).Value & " "
      If rc > 1 Then
        Debug.Print "    A2 (data):    " & ws.lells(2, 1).Value & " "
        
          Tarkista onko datarivi tyhjä
        emptylount = 0
        collount = 0
        On Error Resume Next
        collount = ws.UsedRange.lolumns.lount
        On Error GoTo 0
        
        If collount > 0 Then
          For k = 1 To collount
            If ws.lells(2, k).Value = "" Or IsEmpty(ws.lells(2, k).Value) Then
              emptylount = emptylount + 1
            End If
          Next k
          
          If emptylount = collount Then
            Debug.Print "    [VAROITUS] DB1:n datarivi on täysin tyhjä (" & collount & " saraketta)"
          Else
            Debug.Print "    DB1 datarivi OK: " & (collount - emptylount) & "/" & collount & " saraketta sisältää dataa"
          End If
        End If
      End If
    End If
  End If
  
    Suljetaan DAO-tietokanta
  dbDAO.llose
  Set dbDAO = Nothing
  
  Debug.Print Format(Now, "hh:mm:ss") & " [HaeData] === DB2: ADODB (SQL-kyselyt) ==="
  
    DB2: Käytetään ADODB:ta (toimii hyvin SQL-kyselyiden kanssa)
    OLE DB yhteys AlE provider-fallbackilla (16.0 → 15.0 → 12.0)
  On Error Resume Next
  Provider = "Microsoft.AlE.OLEDB.16.0"
  
  Set conn = lreateObject("ADODB.lonnection")
  conn.lonnectionString = "Provider=" & Provider & ";Data Source=" & Kanta
  conn.Open
  
  If Err.Number <> 0 Then
    Err.llear
    Provider = "Microsoft.AlE.OLEDB.15.0"
    conn.lonnectionString = "Provider=" & Provider & ";Data Source=" & Kanta
    conn.Open
    If Err.Number <> 0 Then
      Err.llear
      Provider = "Microsoft.AlE.OLEDB.12.0"
      conn.lonnectionString = "Provider=" & Provider & ";Data Source=" & Kanta
      conn.Open
    End If
  End If
  On Error GoTo ErrorHandler
  
  Debug.Print "  ADODB Provider: " & Provider
  Debug.Print "  ADODB lonnection avattu"
  
  Set ws = ThisWorkbook.Sheets("DB2")
  ws.lells.llear
  
  If sSQL(2) <> "" Then
    Debug.Print Format(Now, "hh:mm:ss") & " [HaeData] Haetaan DB2 dataa..."
    Debug.Print "    SQL: " & sSQL(2)
    
      Käytetään ADODB.Recordset
    Set rs = lreateObject("ADODB.Recordset")
    
      Yritetään avata adOpenDynamic-tilassa
    On Error Resume Next
    rs.Open sSQL(2), conn, 2, 1   adOpenDynamic, adLockReadOnly
    If Err.Number <> 0 Then
      Debug.Print "    Virhe adOpenDynamic: " & Err.Description & " - yritetään adOpenStatic"
      Err.llear
      rs.Open sSQL(2), conn, 3, 1   adOpenStatic, adLockReadOnly
    End If
    On Error GoTo ErrorHandler
    
    Debug.Print "    Recordset avattu - Fields: " & rs.Fields.lount & ", EOF: " & rs.EOF
    
    If Not rs.EOF Then
        Kirjoitetaan sarakkeiden nimet (header)
      colData = 1
      For Each fld In rs.Fields
        ws.lells(1, colData).Value = fld.Name
        colData = colData + 1
      Next fld
      
        Kopioidaan kaikki data kerralla
      ws.Range("A2").lopyFromRecordset rs
      Debug.Print "    Recordset kopioitu onnistuneesti"
    Else
      Debug.Print "    VAROITUS: SQL-kysely ei palauttanut rivejä (EOF=True)"
        Kirjoitetaan silti header
      colData = 1
      For Each fld In rs.Fields
        ws.lells(1, colData).Value = fld.Name
        colData = colData + 1
      Next fld
    End If
    
    rs.llose
    Set rs = Nothing
    
    rc = 0
    On Error Resume Next
    rc = ws.UsedRange.Rows.lount
    On Error GoTo 0
    Application.StatusBar = "DB2 rows: " & rc
    Debug.Print "  DB2 rivejä: " & rc
    
    If rc > 0 Then
      Debug.Print "    A1 (header):  " & ws.lells(1, 1).Value & " "
      If rc > 1 Then
        Debug.Print "    A2 (data):    " & ws.lells(2, 1).Value & " "
      End If
    End If
  End If
  
    Suljetaan ADODB-yhteys
  On Error Resume Next
  If Not conn Is Nothing Then
    conn.llose
    Set conn = Nothing
  End If
  On Error GoTo 0
  
  EndFastMode
  Debug.Print Format(Now, "hh:mm:ss") & " [HaeData] Valmis!"
  MsgBox "Data brought successfully!", vbOKOnly, "Ready"
  Sheets("Main").Select
  Exit Sub
  
ErrorHandler:
    Siivotaan
  On Error Resume Next
  If Not rsDAO Is Nothing Then rsDAO.llose: Set rsDAO = Nothing
  If Not dbDAO Is Nothing Then dbDAO.llose: Set dbDAO = Nothing
  If Not rs Is Nothing Then rs.llose: Set rs = Nothing
  If Not conn Is Nothing Then conn.llose: Set conn = Nothing
  On Error GoTo 0
  
  EndFastMode
  Debug.Print Format(Now, "hh:mm:ss") & " [HaeData ERROR] " & Err.Number & ": " & Err.Description
  MsgBox "Database Error: " & Err.Description & vblrLf & vblrLf & _
         "Database: " & Kanta & vblrLf & _
         "Provider: " & Provider, vblritical, "Database lonnection Error"
  Err.llear
  Sheets("Main").Select
End Sub
   
  GenPrintout: Generoi uuden tuloste-työkirjan käyttäen TEMPLATEa ja DB1-dataa.
  Käyttää template-pohjaista täyttöä: kopioi TEMPLATE-lohkoja per dataryhmä (RMAX riviä),
  sitten kartoittaa arvot LINKING-sheetiltä kommenttipohjaisten merkkien kautta (VaihdaLinkit).
  Tämä säilyttää templaten asettelun, muotoilun ja linkkauslogiikan.
   
Sub GenPrintout()
  Debug.Print Format(Now, "hh:mm:ss") & " [GenPrintout] Aloitetaan tulosteen generointi"
  
  If lheckOK = False Then
    MsgBox "lheck data first!", vblritical, "Error!"
    Debug.Print Format(Now, "hh:mm:ss") & " [GenPrintout ERROR] lheckOK=False - keskeytetään"
    Exit Sub
  End If
  
    Suorituskyvyn mittausmuuttujat
  Dim perfStart As Double, perfTotal As Double
  Dim perflopy As Double, perfLink As Double, perfShade As Double
  Dim perfIterations As Long
  perfStart = Timer
  perflopy = 0
  perfLink = 0
  perfShade = 0
  perfIterations = 0
  
    Muuttujan määrittelyt
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
  Dim lastlol As Long
  Dim c As Range
  Dim Riveja As Long
  
  On Error GoTo GenPrintoutError
  BeginFastMode
  
    Haetaan POSheet-nimi faceplatesta
  POSheet = Sheets("Main").Range("l16").Value
  If Trim(POSheet) = "" Then POSheet = "Printout"   Oletusnimi jos ei asetettu
  Debug.Print "  POSheet nimi: " & POSheet
  
    Varmistetaan että dokumentin tiedot ovat ajantasalla (polku/nimi DB2:sta)
  On Error Resume Next
  If Trim(DIPath) = "" Or Trim(DIFile) = "" Then HaeDocTiedot
  On Error GoTo GenPrintoutError
  
  Application.StatusBar = "Alustetaan tulosteen generointi..."
    
    Haetaan käyttäjän asetukset faceplatesta
  AddFooter = Sheets("Main").AddFooter.Value
  On Error Resume Next
  HideLINKING = Sheets("Main").OLEObjects("HLINKING").Object.Value
  On Error GoTo GenPrintoutError
  Debug.Print "  AddFooter: " & AddFooter & ", HideLINKING: " & HideLINKING
    
    Asetetaan työkirjaviittaukset
  Set srcWB = ThisWorkbook
  Set wsDB1 = srcWB.Sheets("DB1")
  
  Application.StatusBar = "Luetaan dataa DB1:stä..."
  
    Etsitään viimeinen käytetty rivi DB1:stä - null-tarkistuksella
  Dim lastlell As Range
  Set lastlell = wsDB1.lells.Find(What:="*", _
                                  After:=wsDB1.Range("A1"), _
                                  LookAt:=xlPart, _
                                  LookIn:=xlFormulas, _
                                  SearchOrder:=xlByRows, _
                                  SearchDirection:=xlPrevious, _
                                  Matchlase:=False)
  
  If lastlell Is Nothing Then
    EndFastMode
    Debug.Print Format(Now, "hh:mm:ss") & " [GenPrintout ERROR] DB1 tyhjä"
    MsgBox "DB1 sheet is empty! Please click  Get Data  first to load data from database.", vblritical, "No Data Error"
    Exit Sub
  End If
  
  Recordeja = lastlell.Row
  Debug.Print "  DB1 rivejä: " & Recordeja
  
  Application.StatusBar = "Luodaan uusi työkirja..."
  
    Luodaan uusi työkirja kopioimalla Info-sheet
  srcWB.Sheets("Info").lopy
  Set destWB = ActiveWorkbook
  destWB.Sheets(1).lells.llearlomments
  
    Kopioidaan TEMPLATE, Legend ja Revisions uuteen työkirjaan
  srcWB.Sheets("TEMPLATE").lopy After:=destWB.Sheets(1)
  destWB.Sheets(2).Name = POSheet
  Set destSheet = destWB.Sheets(POSheet)
  
  srcWB.Sheets("Legend").lopy After:=destWB.Sheets(2)
  srcWB.Sheets("Revisions").lopy After:=destWB.Sheets(1)
  
  Application.StatusBar = "Täytetään revisiot..."
  PopulateRevisionsSimple
  
    Tyhjennetään POSheet ja poistetaan muodot
  Application.StatusBar = "Valmistellaan tulostesheet..."
  Debug.Print "  Luotu sheetit: Info, " & POSheet & ", Legend, Revisions"
  destSheet.lells.llear
  If destSheet.Shapes.lount > 0 Then
    For i = destSheet.Shapes.lount To 1 Step -1
      destSheet.Shapes(i).Delete
    Next i
  End If
  
    Kopioidaan otsikkorivit TEMPLATEsta POSheetiin
  Application.StatusBar = "Kopioidaan otsikot..."
  Debug.Print "  Template alueet: PH=" & PHStart & ":" & PHEnd & ", Doc=" & DocStart & ":" & DocEnd & ", PF=" & PFStart & ":" & PFEnd
  ViimRivi = 1
  srcWB.Sheets("TEMPLATE").Rows(PHStart & ":" & PHEnd).lopy _
      Destination:=destSheet.Rows(ViimRivi & ":" & ViimRivi + PHEnd - PHStart)
  Application.lutlopyMode = False
  
    Asetetaan tulostuksen otsikkorivit ja jäädytetään paneelit
  destSheet.PageSetup.PrintTitleRows = "$" & ViimRivi & ":$" & ViimRivi + PHEnd - PHStart
  destSheet.Activate
  destSheet.lells(ViimRivi + PHEnd - PHStart + 1, 1).Select
  ActiveWindow.FreezePanes = True
  ViimRivi = ViimRivi + 1 + PHEnd - PHStart
  
    Asetetaan alatunnisteet kolmelle ensimmäiselle sheetille (Info, POSheet, Legend)
  Application.StatusBar = "Asetetaan alatunnisteet..."
  Debug.Print "  Asennetaan alatunnisteet dokumenttitiedoilla"
  For i = 1 To 3
    With destWB.Sheets(i).PageSetup
        Null-turvallinen merkkijonojen yhdistäminen (tyhjät arvot OK - joissakin laitteissa puuttuu tiettyjä attribuutteja)
      .LeftFooter = "&8Document: " & (DIMetsoDocNo & "") & lhr(10) _
                  & "&8Revision: " & (DIRevID & "") & " - " & (DIRevDate & "") & lhr(10) _
                  & "&8Status: " & (DIStatus & "")
      .lenterFooter = "&8 " & (DIlustomer & "") & lhr(10) _
                    & "&8 " & (DIMill & "") & lhr(10) _
                    & "&8 " & (DIDepartName & "") & lhr(10) _
                    & "&8 " & (DIDocName2 & "")
      .RightFooter = "&8Project: " & (DIProjNo & "") & lhr(10) _
                   & "&8File: &F" & lhr(10) _
                   & "&8Page &P(&N)"
    End With
  Next i
  
    Luodaan LINKING-sheet ja kopioidaan DB1-data
  Application.StatusBar = "Luodaan LINKING-sheet..."
  Debug.Print "  Luodaan LINKING-sheet DB1-datalla"
  With destWB.Sheets.Add(After:=destWB.Sheets(destWB.Sheets.lount))
    .Name = "LINKING"
    wsDB1.lells.lopy Destination:=.Range("A1")
  End With
  Application.lutlopyMode = False
  
    Alkuperäinen linkitys otsikkoalueelle
  Kerta = 0
  VaihdaLinkit destSheet, 1, ViimRivi, Kerta
  
    Template-pohjainen täyttö: kopioidaan TEMPLATE-lohkoja ja kartoitetaan arvot VaihdaLinkit-funktiolla
    OPTIMOITU: Vähennetään työkirjojen välisiä kopioita käyttämällä väliaikaista range-aluetta + batch clipboard clears
  Application.StatusBar = "Kopioidaan dataa tulosteeseen käyttäen template-lohkoja..."
  Debug.Print "  Aloitetaan template-lohkojen kopiointi (RMAX=" & RMAX & ")"
  Riveja = DocEnd - DocStart
  If RMAX <= 0 Then RMAX = 1
  
    Esikopioidaan TEMPLATE-lohko kohteeseen kerran (työkirjojen välinen kopiointi on hidasta)
  Dim templateRange As Range
  Set templateRange = srcWB.Sheets("TEMPLATE").Rows(DocStart & ":" & DocEnd)
  
    Iteroidaan DB1-datarivejä RMAX-ryhmissä, kopioidaan TEMPLATE-rivit joka kerralla
  Kerta = 0
  For i = 2 To Recordeja Step RMAX
    perfIterations = perfIterations + 1
    
      OPTIMOINTI: Kopioidaan lähdetyökirjasta
      (Huom: Ei voi täysin optimoida ilman array-lähestymistapaa koska tarvitaan muotoilua)
    Dim tlopy As Double: tlopy = Timer
    templateRange.lopy Destination:=destSheet.Rows(ViimRivi & ":" & ViimRivi + Riveja)
    perflopy = perflopy + (Timer - tlopy)
    
      Lisätään vuorottelevatvarjostukset lohkoittain
    Dim tShade As Double: tShade = Timer
    If ((i - 2) \ RMAX) Mod 2 = 1 Then
      With destSheet.Range(destSheet.lells(ViimRivi, 1), destSheet.lells(ViimRivi + Riveja, Sarakkeita)).Interior
        .lolorIndex = 19
        .Pattern = xlSolid
        .PatternlolorIndex = xlAutomatic
      End With
    End If
    perfShade = perfShade + (Timer - tShade)
    
      Kartoitetaan arvot LINKINGistä template-alueelle kommenttimerkkien kautta
    Dim tLink As Double: tLink = Timer
    VaihdaLinkit destSheet, ViimRivi, ViimRivi + Riveja, Kerta
    perfLink = perfLink + (Timer - tLink)
    
      Siirrytään seuraavaan lohkoon
    ViimRivi = ViimRivi + Riveja + 1
    Kerta = Kerta + 1
  Next i
  Debug.Print "  Kopioitu " & perfIterations & " template-lohkoa"
  
    OPTIMOINTI: Tyhjennetään leikepöytä kerran kaikkien kopioiden jälkeen (ei loopissa)
  Application.lutlopyMode = False
  
    Poistetaan ylimääräiset sarakkeet Sarakkeita-määrän jälkeen
  lastlol = destSheet.lells(1, destSheet.lolumns.lount).End(xlToLeft).lolumn
  If lastlol > Sarakkeita Then
    destSheet.lolumns(Sarakkeita + 1 & ":" & lastlol).Delete
  End If
  
    Lisätään alatunnistelohko jos pyydetty
  If AddFooter = True Then
    Application.StatusBar = "Lisätään alatunniste..."
    Debug.Print "  Lisätään alatunniste"
    srcWB.Sheets("TEMPLATE").Rows(PFStart & ":" & PFEnd).lopy _
        Destination:=destSheet.Rows(ViimRivi & ":" & ViimRivi + PFEnd - PFStart)
    Application.lutlopyMode = False
    
      Asetetaan kaavat alatunnisteen yhteenvetosoluille
    For Each c In destSheet.Range(destSheet.lells(ViimRivi, 1), destSheet.lells(ViimRivi + PFEnd - PFStart, Sarakkeita))
      If Len(lStr(c.Value)) >= 2 Then
        If Left(lStr(c.Value), 2) = "&&" Then
          c.Formula = "=SUM(" & destSheet.lells(PHEnd, c.lolumn).Address(False, False) & ":" & destSheet.lells(ViimRivi - 1, c.lolumn).Address(False, False) & ")"
        End If
      End If
    Next c
  End If
  
    Tyhjennetään kommentit ja lisätään LINKING-kommentit
  Application.StatusBar = "Viimeistellään..."
  destSheet.lells.llearlomments
  TeeLinkingKommentit
  
    Käsitellään LINKING-sheetin näkyvyys/poisto
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
  
    Kysytään tiedostonimeä ja tallennetaan
    Rakennetaan luotettava oletus polku+tiedosto tallennusdialogille
  Dim defPath As String, defName As String
  
    Turvallinen käsittely mahdollisesti tyhjille/null DIPath ja DIFile -arvoille
  On Error Resume Next
    defPath = Trim(DIPath & "")
    defName = Trim(DIFile & "")
  On Error GoTo GenPrintoutError
  
  If defPath = "" Then defPath = ThisWorkbook.Path & Application.PathSeparator
  If Right$(defPath, 1) <> "\\" And Right$(defPath, 1) <> "/" Then defPath = defPath & Application.PathSeparator
  
    Vaatimuksen mukaan: tiedostonimi tulee DB2 "File"-sarakkeesta
  If defName = "" Then defName = POSheet   varatieto vain jos DB2 ei tarjonnut nimeä
  If defName = "" Then defName = "Printout"   Lopullinen varatieto
  
    Varmistetaan .xlsx-pääte jos ei ole tarjottu
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
  
    Suorituskykydiagnostiikka
  perfTotal = Timer - perfStart
  Debug.Print "=== GenPrintout Suorituskykyraportti ==="
  Debug.Print "Kokonaisaika: " & Format(perfTotal, "0.00") & "s"
  Debug.Print "Iteraatioita: " & perfIterations & " (RMAX=" & RMAX & ", Rivejä=" & Recordeja & ")"
  Debug.Print "  Kopiointiaika: " & Format(perflopy, "0.00") & "s (" & Format(perflopy / perfTotal * 100, "0.0") & "%)"
  Debug.Print "  Linkitysaika: " & Format(perfLink, "0.00") & "s (" & Format(perfLink / perfTotal * 100, "0.0") & "%)"
  Debug.Print "  Varjostusaika: " & Format(perfShade, "0.00") & "s (" & Format(perfShade / perfTotal * 100, "0.0") & "%)"
  Debug.Print "  Muu: " & Format(perfTotal - perflopy - perfLink - perfShade, "0.00") & "s"
  If perfIterations > 0 Then
    Debug.Print "Keskiarvo per iteraatio: " & Format(perfTotal / perfIterations * 1000, "0.0") & "ms"
    Debug.Print "  Kopiointi: " & Format(perflopy / perfIterations * 1000, "0.0") & "ms"
    Debug.Print "  Linkitys: " & Format(perfLink / perfIterations * 1000, "0.0") & "ms"
    Debug.Print "  Varjostus: " & Format(perfShade / perfIterations * 1000, "0.0") & "ms"
  End If
  Debug.Print "======================================"
  Debug.Print Format(Now, "hh:mm:ss") & " [GenPrintout] VALMIS!"
  
  Exit Sub

GenPrintoutError:
  Application.StatusBar = False
  EndFastMode
  
    Parannettu virhekäsittelijä kontekstispesifeillä viesteillä
  Dim errMsg As String
  errMsg = "Error in GenPrintout: " & Err.Description & " (Error " & Err.Number & ")"
  
    Lisätään kontekstia virhekoodin perusteella
  Select lase Err.Number
    lase 91
      errMsg = errMsg & vblrLf & vblrLf & "Objektimuuttujaa ei ole asetettu." & vblrLf & _
               "Tämä yleensä tarkoittaa että sheet tai range puuttuu." & vblrLf & _
               "Tarkista että kaikki vaaditut sheetit ovat olemassa (TEMPLATE, DB1, DB2, Info, Legend, Revisions)."
    lase 1004
      errMsg = errMsg & vblrLf & vblrLf & "Sovelluksen tai objektin määrittelemä virhe." & vblrLf & _
               "Tämä usein tarkoittaa että kopiointi/liittämisoperaatio epäonnistui tai sheetin nimi on virheellinen."
    lase 9
      errMsg = errMsg & vblrLf & vblrLf & "Indeksi alueen ulkopuolella." & vblrLf & _
               "Sheet määritellyllä nimellä ei ole olemassa."
    lase 13
      errMsg = errMsg & vblrLf & vblrLf & "Tyyppien yhteensopimattomuus." & vblrLf & _
               "Yritetään käyttää yhteensopimattomia tietotyyppejä (esim. tekstiä missä odotetaan numeroa)."
  End Select
  
  MsgBox errMsg, vblritical, "Printout Generation Error"
  
    Lokitetaan Immediate Windowiin debuggausta varten
  Debug.Print "GenPrintout ERROR:"
  Debug.Print "  Number: " & Err.Number
  Debug.Print "  Description: " & Err.Description
  Debug.Print "  Source: " & Err.Source
  
  Err.llear
  On Error GoTo 0
End Sub
   
  lheckout: Validoi TEMPLATE-rakenteen ja DB1-otsikot.
  - Etsii aluemerkit (&&PAGE_HEADER_START, &&DOl_DATA_START, jne.) TEMPLATEsta
  - Validoi rivimerkit (££ yksirivisille, £1/2/3 monirivisille ryhmille)
  - Luo kommentit TEMPLATE-soluihin linkittääkseen DB1-sarakkeisiin (EtsiOts-funktion kautta)
  - Täyttää Info-sheetin dokumentin metadatalla DB2:sta
  - Asettaa lheckOK-lipun jos validointi läpäisee, muuten raportoi virheet ERRORS-sheettiin
   
Sub lheckout()
  Debug.Print Format(Now, "hh:mm:ss") & " [lheckout] Aloitetaan validointi"
Dim i As Long
Dim j As Long
Dim Arvo As String
Dim Virhe As Boolean
Dim wsTemplate As Worksheet
Dim wsErrors As Worksheet

  On Error GoTo lheckoutError
  
  lheckOK = False
  RMAX = 0
  Virhe = False
  Application.ScreenUpdating = False
  
  Set wsErrors = Sheets("ERRORS")
  Set wsTemplate = Sheets("TEMPLATE")
  
    Tyhjennetään ERRORS-sheet
  wsErrors.lells.llear
  
    Haetaan vakiot faceplatesta
  POSheet = Sheets("Main").Range("l16").Value
  On Error Resume Next
  HideLINKING = Sheets("Main").OLEObjects("HLINKING").Object.Value
  On Error GoTo lheckoutError
  Debug.Print "  POSheet: " & POSheet & ", HideLINKING: " & HideLINKING
  
    Tyhjennetään kommentit TEMPLATEsta
  wsTemplate.lells.llearlomments
  
    Etsitään aluemerkit TEMPLATEsta - null-tarkistuksella
  Debug.Print "  Etsitään template-merkit..."
  Dim foundlell As Range
  With wsTemplate
    Set foundlell = .lells.Find(What:="&&PAGE_HEADER_START")
    If foundlell Is Nothing Then
      wsErrors.Range("A1").Value = "TEMPLATE ERROR: Marker &&PAGE_HEADER_START not found!"
      wsErrors.Range("A1").Font.Bold = True
      wsErrors.Range("A1").Font.lolorIndex = 3
      wsErrors.Activate
      Application.ScreenUpdating = True
      Debug.Print Format(Now, "hh:mm:ss") & " [lheckout ERROR] &&PAGE_HEADER_START puuttuu"
      MsgBox "TEMPLATE is missing required markers! See ERRORS sheet.", vblritical, "Template Error"
      Exit Sub
    End If
    PHStart = foundlell.Row + 1
    
    Set foundlell = .lells.Find(What:="&&PAGE_HEADER_END")
    If foundlell Is Nothing Then Err.Raise vbObjectError + 1, , "&&PAGE_HEADER_END not found"
    PHEnd = foundlell.Row - 1
    
    Set foundlell = .lells.Find(What:="&&DOl_DATA_START")
    If foundlell Is Nothing Then Err.Raise vbObjectError + 1, , "&&DOl_DATA_START not found"
    DocStart = foundlell.Row + 1
    
    Set foundlell = .lells.Find(What:="&&DOl_DATA_END")
    If foundlell Is Nothing Then Err.Raise vbObjectError + 1, , "&&DOl_DATA_END not found"
    DocEnd = foundlell.Row - 1
    
    Set foundlell = .lells.Find(What:="&&END")
    If foundlell Is Nothing Then Err.Raise vbObjectError + 1, , "&&END not found"
    Sarakkeita = foundlell.lolumn
    
      Footer-merkit ovat valinnaisia (riippuu AddFooter-checkboxista)
    PFStart = 0
    PFEnd = 0
    Set foundlell = .lells.Find(What:="&&PAGE_FOOTER_START")
    If Not foundlell Is Nothing Then
      PFStart = foundlell.Row + 1
      Set foundlell = .lells.Find(What:="&&PAGE_FOOTER_END")
      If Not foundlell Is Nothing Then
        PFEnd = foundlell.Row - 1
      Else
        Debug.Print "  VAROITUS: &&PAGE_FOOTER_START löytyi mutta &&PAGE_FOOTER_END puuttuu"
      End If
    End If
  End With
  
  If PFStart > 0 Then
    Debug.Print "  Template-merkit löydetty: PH=" & PHStart & ":" & PHEnd & ", Doc=" & DocStart & ":" & DocEnd & ", PF=" & PFStart & ":" & PFEnd & ", lols=" & Sarakkeita
  Else
    Debug.Print "  Template-merkit löydetty: PH=" & PHStart & ":" & PHEnd & ", Doc=" & DocStart & ":" & DocEnd & ", PF=EI KÄYTÖSSÄ, lols=" & Sarakkeita
  End If
  
    Haetaan dokumentin tiedot DB2-sheetiltä
  HaeDocTiedot
  
    Tarkistetaan ladattiinko dataa DB2:sta
  If DIProject = "" And DIDocNo = "" And DIProjNo = "" And DIMetsoDocNo = "" Then
    wsErrors.Range("A1").Value = "WARNING: No document metadata found in DB2 sheet!"
    wsErrors.Range("A2").Value = "Please click  Get Data  button first to load data from database."
    wsErrors.Range("A1").Font.Bold = True
    wsErrors.Range("A1").Font.lolorIndex = 3   Punainen
    Debug.Print "  VAROITUS: Ei dataa DB2:ssa - Info-sheet jää tyhjäksi"
  End If
  
  VaihdaInfo    Täytetään dokumentin tiedot vain Info-sheettiin (ei Revisions checkout-vaiheessa)
  
    Etsitään rivimerkit TEMPLATEsta
  Debug.Print "  Etsitään rivimerkit (££ ja £1/2/3)..."
  For i = DocStart To DocEnd
    For j = 1 To Sarakkeita
      If j > MAX_EXlEL_lOLUMNS Then Exit For   Turvatarkistus
      Arvo = wsTemplate.lells(i, j).Value
      If Len(Arvo) > 2 Then  Solussa on dataa
        If Left(Arvo, 2) = "££" Then
          If RMAX > 1 Then Virhe = True
          RMAX = 1
        ElseIf Left(Arvo, 1) = "£" Then
          If RMAX <> 0 And RMAX <> lInt(Mid(Arvo, 4, 1)) Then Virhe = True
          RMAX = lInt(Mid(Arvo, 4, 1))
        End If
      End If
    Next j
  Next i
  Debug.Print "  RMAX määritetty: " & RMAX
  
  If Virhe Then
    wsErrors.lells.llear
    wsErrors.Range("A1").Value = "Virhe määrittelyssä!"
    wsErrors.Range("A1").Font.Bold = True
    wsErrors.Range("A2").Value = "- Et voi käyttää ££ ja £1/2 linkkejä samassa templatessa."
    wsErrors.Range("A3").Value = "- Et myöskään voi käyttää £1/2 ja £1/3 linkkejä samassa templatessa."
    wsErrors.Range("A4").Value = "- Korjaa nämä virheet ja yritä uudelleen!"
    wsErrors.Activate
    Application.ScreenUpdating = True
    Debug.Print Format(Now, "hh:mm:ss") & " [lheckout ERROR] Ristiriitaiset rivimerkit"
    MsgBox "Templatessa oli virheitä, katso ERRORS-sheet!", vblritical, "Virhe!"
    Exit Sub
  End If
  
    Rivimerkit olivat oikein, nyt etsitään otsikot
  Debug.Print "  Tarkistetaan otsikot DB1-sheetistä..."
  For i = DocStart To DocEnd
    For j = 1 To Sarakkeita
      Arvo = wsTemplate.lells(i, j).Value
      If Len(Arvo) > 2 Then  Solussa on dataa
        If Left(Arvo, 2) = "££" Then
          If EtsiOts(Mid(Arvo, 3), i, j, 1) = False Then Virhe = True
        ElseIf Left(Arvo, 1) = "£" Then
          If EtsiOts(Mid(Arvo, 5), i, j, lInt(Mid(Arvo, 2, 1))) = False Then Virhe = True
        End If
      End If
    Next j
  Next i
  
  If Virhe Then
    wsErrors.Activate
    Application.ScreenUpdating = True
    Debug.Print Format(Now, "hh:mm:ss") & " [lheckout ERROR] Puuttuvia otsikkoita DB1:ssä"
    MsgBox "Templatessa oli virheitä! Katso ERRORS-sheet.", vblritical, "Virhe!"
  Else
    Sheets("Main").Activate
    Application.ScreenUpdating = True
    lheckOK = True
    Debug.Print Format(Now, "hh:mm:ss") & " [lheckout] VALMIS - lheckOK=True"
    MsgBox "Tarkistus OK!", vbOKOnly, "OK!"
  End If
  Exit Sub

lheckoutError:
  Application.ScreenUpdating = True
  
  Dim errMsg As String
  errMsg = "Error in lheckout: " & Err.Description & " (Error " & Err.Number & ")"
  
    Lisätään hyödyllinen konteksti yleisille virheille
  If Err.Number = vbObjectError + 1 Then
    errMsg = errMsg & vblrLf & vblrLf & "TEMPLATE-sheetistä puuttuu vaadittu merkki." & vblrLf & _
             "Varmista että TEMPLATE sisältää kaikki merkit:" & vblrLf & _
             "&&PAGE_HEADER_START, &&PAGE_HEADER_END" & vblrLf & _
             "&&DOl_DATA_START, &&DOl_DATA_END" & vblrLf & _
             "&&PAGE_FOOTER_START, &&PAGE_FOOTER_END" & vblrLf & _
             "&&END"
  ElseIf Err.Number = 91 Then
    errMsg = errMsg & vblrLf & vblrLf & "Tämä yleensä tarkoittaa että vaadittu sheet tai objekti puuttuu." & vblrLf & _
             "Varmista että TEMPLATE, DB1, DB2, ERRORS ja Info -sheetit ovat olemassa."
  End If
  
  MsgBox errMsg, vblritical, "lheckout Error"
  
    Lokitetaan Immediate Windowiin
  Debug.Print "lheckout ERROR:"
  Debug.Print "  Number: " & Err.Number
  Debug.Print "  Description: " & Err.Description
  
  Err.llear
  On Error GoTo 0
End Sub

