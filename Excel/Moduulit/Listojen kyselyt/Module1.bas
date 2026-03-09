Option Explicit

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
' Dokumentin metatiedot kapseloitu UDT:hen — vältetään globaalin nimiavaruuden saastuminen
Type DocumentMetadata
    Contract   As String
    Mill       As String
    DepartName As String
    Customer   As String
    Project    As String
    ProjNo     As String
    ProjName   As String
    Munit      As String
    Manager    As String
    DocNo      As String
    MetsoDocNo As String
    DocName    As String
    DocName1   As String
    DocName2   As String
    DocName3   As String
    Path       As String
    File       As String
    Date       As String
    Rev        As String
    RevID      As String
    RevDate    As String
    Status     As String
End Type

Public DocInfo As DocumentMetadata
Public DIRevArr() As String  ' Dynaaminen taulukko ei voi olla UDT-jäsen VBA:ssa

' Suojaus silmukoiden iteraatioille estääkseen ikuiset silmukat
Public Const MAX_EXCEL_COLUMNS As Long = 16384

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
' AlustaTila: Nollaa kaikki makron ajokohtaiset globaalitilat.
' Kutsutaan ThisWorkbook.Workbook_Open-tapahtumasta, jotta edellisen ajon
' CheckOK=True-tila ei jaa voimaan uuden istunnon alussa.
'''
Public Sub AlustaTila()
  CheckOK = False
  RMAX = 0
  PHStart = 0: PHEnd = 0: PFStart = 0: PFEnd = 0
  DocStart = 0: DocEnd = 0
  Debug.Print Format(Now, "hh:mm:ss") & " [AlustaTila] Tila nollattu"
End Sub

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
' LuoADODBYhteys: Hakee toimivan ADODB-yhteyden kokeilemalla ACE OLE DB -moottoriversiota prioriteettijärjestyksessä.
' Käyttää Mode=Read -yhteysparametria (adModeRead) estääkseen kirjoitusoperaatiot arkkitehtuuritasolla.
' Palauttaa avatun ADODB.Connection-objektin tai Nothing jos kaikki yritykset epäonnistuvat.
' DRY-apufunktio — korvaa kolminkertaisen yhteysyritysrakenteen kutsupaikoissa.
'''
Private Function LuoADODBYhteys(kantaPolku As String) As Object
    Dim conn As Object
    Dim providerVersions As Variant
    Dim i As Integer
    ' Kokeiltavat versiot prioriteettijärjestyksessä (64-bit / uudemmat ensin)
    providerVersions = Array("16.0", "15.0", "12.0")
    For i = LBound(providerVersions) To UBound(providerVersions)
        Set conn = CreateObject("ADODB.Connection")
        conn.ConnectionString = "Provider=Microsoft.ACE.OLEDB." & providerVersions(i) & ";Data Source=" & kantaPolku & ";Mode=Read;"
        On Error Resume Next
        conn.Open
        If Err.Number = 0 Then
            On Error GoTo 0
            Set LuoADODBYhteys = conn  ' Yhteys onnistui — palautetaan se
            Exit Function
        End If
        Err.Clear
        On Error GoTo 0
        Set conn = Nothing  ' Siivotaan epäonnistunut yritys
    Next i
    Set LuoADODBYhteys = Nothing  ' Kaikki versiot käyty läpi — yhteys ei onnistunut
End Function


'''
' HaeData: Hakee datan Access-tietokannasta käyttäen tallennettuja kyselyitä tai SQL-lauseita.
' Täyttää DB1 (pääasiallinen body-data) ja DB2 (dokumentin metadata) -sheetit.
' 
' DB1: Käyttää DAO:ta (Data Access Objects) joka tukee natiivisti Access-tallennettuja kyselyitä
'      ja JET SQL -syntaksia (Like "pattern", Deleted=No, IIf, jne.)
' DB2: Käyttää ADODB:ta yhteensopivuuden vuoksi (LuoADODBYhteys-apufunktio)
' 
' Diagnostiikka: Rivimäärät näytetään StatusBarissa ja Immediate Windowissa jokaisen kyselyn jälkeen.
'''
Sub HaeData()
Dim Kanta As String
Dim sSQL(1 To 2) As String
Dim Valinta As Long
Dim i As Long
Dim ws As Worksheet
Dim rc As Long
Dim Provider As String
Dim emptyCount As Long, colCount As Long, k As Long

' DAO-muuttujat DB1:lle (tallennetut kyselyt)
Dim dbDAO As Object      ' DAO.Database
Dim rsDAO As Object      ' DAO.Recordset
Dim fldDAO As Object     ' DAO.Field
Dim colData As Long

' ADODB-muuttujat DB2:lle (SQL-kyselyt)
Dim conn As Object       ' ADODB.Connection
Dim rs As Object         ' ADODB.Recordset
Dim fld As Object        ' ADODB.Field

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
  
  On Error GoTo ErrorHandler
  BeginFastMode
  
  ' Varmistetaan että tietokantatiedosto on olemassa
  If Dir(Kanta) = "" Then
    MsgBox "Tietokantatiedostoa ei löydy: " & Kanta, vbCritical, "Tietokantavirhe"
    GoTo SafeExit  ' Try-Finally: SafeExit siivoo aina
  End If
  
  Debug.Print Format(Now, "hh:mm:ss") & " [HaeData] === DB1: DAO (tallennetut kyselyt) ==="
  
  ' Avataan DAO-tietokanta DB1:lle vain luku -tilassa (ReadOnly=True estää kirjoituskyselyt arkkitehtuuritasolla)
  On Error Resume Next
  Set dbDAO = CreateObject("DAO.DBEngine.120").OpenDatabase(Name:=Kanta, Options:=False, ReadOnly:=True)
  If Err.Number <> 0 Then
    Err.Clear  ' DBEngine.120 ei käytettävissä — ACE OLEDB ei ole asennettu
  End If
  On Error GoTo ErrorHandler
  
  If dbDAO Is Nothing Then
    MsgBox "Tietokantaa ei voitu avata DAO:lla!" & vbCrLf & _
           "Varmista että Microsoft Access Database Engine on asennettuna.", vbCritical, "DAO-virhe"
    GoTo SafeExit  ' Try-Finally: SafeExit siivoo aina
  End If
  
  Debug.Print "  DAO Database avattu"
  
  ' DB1: Käytetään DAO:ta (tukee tallennettuja kyselyitä ja JET SQL:ää)
  Set ws = ThisWorkbook.Sheets("DB1")
  ws.Cells.Clear
  
  If sSQL(1) <> "" Then
    Debug.Print Format(Now, "hh:mm:ss") & " [HaeData] Haetaan DB1 dataa..."
    Debug.Print "    Kysely/SQL: " & sSQL(1)


    ' DAO.OpenRecordset voi ottaa joko tallenetun kyselyn nimen tai SQL-lauseen
    Set rsDAO = dbDAO.OpenRecordset(sSQL(1))
    
    Debug.Print "    DAO Recordset avattu - Fields: " & rsDAO.Fields.Count & ", EOF: " & rsDAO.EOF & ", BOF: " & rsDAO.BOF
    
    If Not rsDAO.EOF Then
      ' Kirjoitetaan sarakkeiden nimet (header)
      colData = 1
      For Each fldDAO In rsDAO.Fields
        ws.Cells(1, colData).Value = fldDAO.Name
        colData = colData + 1
      Next fldDAO
      
      ' CopyFromRecordset siirtää koko tulosjoukon natiivillä C++-toteutuksella — korvaa
      ' koko aiemman dataArr-taulukon alustuksen, täyttösilmukan ja levykirjoituksen yhdellä kutsulla
      ws.Range("A2").CopyFromRecordset rsDAO
      
      ' RecordCount raportointia varten (MoveLast vaaditaan dynaset-tyypille)
      Dim totalRows As Long
      rsDAO.MoveLast
      totalRows = rsDAO.RecordCount
      Debug.Print "    DAO data kopioitu (CopyFromRecordset): " & totalRows & " riviä, " & rsDAO.Fields.Count & " saraketta"
    Else
      Debug.Print "    VAROITUS: Kysely ei palauttanut rivejä (EOF=True)"
      ' Kirjoitetaan silti header
      colData = 1
      For Each fldDAO In rsDAO.Fields
        ws.Cells(1, colData).Value = fldDAO.Name
        colData = colData + 1
      Next fldDAO
    End If
    
    rsDAO.Close
    Set rsDAO = Nothing
    
    ' Raportoidaan rivimäärä
    rc = 0
    On Error Resume Next
    rc = ws.UsedRange.Rows.Count
    On Error GoTo 0
    Application.StatusBar = "DB1 rows: " & rc
    Debug.Print "  DB1 rivejä: " & rc
    
    ' Debug: Näytä ensimmäiset 2 riviä
    If rc > 0 Then
      Debug.Print "    A1 (header): '" & ws.Cells(1, 1).Value & "'"
      If rc > 1 Then
        Debug.Print "    A2 (data):   '" & ws.Cells(2, 1).Value & "'"
        
        ' Tarkista onko datarivi tyhjä
        emptyCount = 0
        colCount = 0
        On Error Resume Next
        colCount = ws.UsedRange.Columns.Count
        On Error GoTo 0
        
        If colCount > 0 Then
          For k = 1 To colCount
            If ws.Cells(2, k).Value = "" Or IsEmpty(ws.Cells(2, k).Value) Then
              emptyCount = emptyCount + 1
            End If
          Next k
          
          If emptyCount = colCount Then
            Debug.Print "    [VAROITUS] DB1:n datarivi on täysin tyhjä (" & colCount & " saraketta)"
          Else
            Debug.Print "    DB1 datarivi OK: " & (colCount - emptyCount) & "/" & colCount & " saraketta sisältää dataa"
          End If
        End If
      End If
    End If
  End If
  
  ' Suljetaan DAO-tietokanta
  dbDAO.Close
  Set dbDAO = Nothing
  
  Debug.Print Format(Now, "hh:mm:ss") & " [HaeData] === DB2: ADODB (SQL-kyselyt) ==="
  
  ' DB2: Käytetään ADODB:ta (toimii hyvin SQL-kyselyiden kanssa)
  ' DRY-refaktorointi: provider-fallback (16.0→15.0→12.0) on eristetty LuoADODBYhteys-apufunktioon
  Set conn = LuoADODBYhteys(Kanta)
  On Error GoTo ErrorHandler

  If conn Is Nothing Then
    MsgBox "ADODB-tietokantayhteyttä ei voitu muodostaa!" & vbCrLf & _
           "Tarkista että Microsoft Access Database Engine on asennettuna.", vbCritical, "ADODB-virhe"
    GoTo SafeExit
  End If
  Provider = conn.Provider  ' Tallennetaan diagnostiikkaa varten

  Debug.Print "  ADODB Provider: " & Provider
  Debug.Print "  ADODB Connection avattu"
  
  Set ws = ThisWorkbook.Sheets("DB2")
  ws.Cells.Clear
  
  If sSQL(2) <> "" Then
    Debug.Print Format(Now, "hh:mm:ss") & " [HaeData] Haetaan DB2 dataa..."
    Debug.Print "    SQL: " & sSQL(2)


    ' Käytetään ADODB.Recordset
    Set rs = CreateObject("ADODB.Recordset")
    
    ' Yritetään avata adOpenDynamic-tilassa
    On Error Resume Next
    rs.Open sSQL(2), conn, 2, 1 ' adOpenDynamic, adLockReadOnly
    If Err.Number <> 0 Then
      Debug.Print "    Virhe adOpenDynamic: " & Err.Description & " - yritetään adOpenStatic"
      Err.Clear
      rs.Open sSQL(2), conn, 3, 1 ' adOpenStatic, adLockReadOnly
    End If
    On Error GoTo ErrorHandler
    
    Debug.Print "    Recordset avattu - Fields: " & rs.Fields.Count & ", EOF: " & rs.EOF
    
    If Not rs.EOF Then
      ' Kirjoitetaan sarakkeiden nimet (header)
      colData = 1
      For Each fld In rs.Fields
        ws.Cells(1, colData).Value = fld.Name
        colData = colData + 1
      Next fld
      
      ' Kopioidaan kaikki data kerralla
      ws.Range("A2").CopyFromRecordset rs
      Debug.Print "    Recordset kopioitu onnistuneesti"
    Else
      Debug.Print "    VAROITUS: SQL-kysely ei palauttanut rivejä (EOF=True)"
      ' Kirjoitetaan silti header
      colData = 1
      For Each fld In rs.Fields
        ws.Cells(1, colData).Value = fld.Name
        colData = colData + 1
      Next fld
    End If
    
    rs.Close
    Set rs = Nothing
    
    rc = 0
    On Error Resume Next
    rc = ws.UsedRange.Rows.Count
    On Error GoTo 0
    Application.StatusBar = "DB2 rows: " & rc
    Debug.Print "  DB2 rivejä: " & rc
    
    If rc > 0 Then
      Debug.Print "    A1 (header): '" & ws.Cells(1, 1).Value & "'"
      If rc > 1 Then
        Debug.Print "    A2 (data):   '" & ws.Cells(2, 1).Value & "'"
      End If
    End If
  End If
  
  ' Suljetaan ADODB-yhteys
  On Error Resume Next
  If Not conn Is Nothing Then
    conn.Close
    Set conn = Nothing
  End If
  On Error GoTo 0
  
  Debug.Print Format(Now, "hh:mm:ss") & " [HaeData] Valmis!"
  MsgBox "Data haettu onnistuneesti!", vbOKOnly, "Valmis"
  Sheets("Main").Select

SafeExit:
  ' Try-Finally -malli VBA:ssa: siivoo resurssit ja palauttaa UI-tilan AINA
  ' — oli suoritus onnistunut, varhainen poistuminen tai virhe.
  On Error Resume Next
  If Not rsDAO Is Nothing Then rsDAO.Close: Set rsDAO = Nothing
  If Not dbDAO Is Nothing Then dbDAO.Close: Set dbDAO = Nothing
  If Not rs Is Nothing Then rs.Close: Set rs = Nothing
  If Not conn Is Nothing Then conn.Close: Set conn = Nothing
  On Error GoTo 0
  EndFastMode
  Exit Sub

ErrorHandler:
  Debug.Print Format(Now, "hh:mm:ss") & " [HaeData ERROR] " & Err.Number & ": " & Err.Description
  MsgBox "Database Error: " & Err.Description & vbCrLf & vbCrLf & _
         "Database: " & Kanta & vbCrLf & _
         "Provider: " & Provider, vbCritical, "Database Connection Error"
  Err.Clear
  Resume SafeExit  ' Hyppää aina SafeExit-lohkoon — EndFastMode laukeaa varmasti
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
    MsgBox "Tarkista data ensin!", vbCritical, "Virhe!"
    Debug.Print Format(Now, "hh:mm:ss") & " [GenPrintout ERROR] CheckOK=False - keskeytetään"
    Exit Sub
  End If
  
  ' Suorituskyvyn mittausmuuttujat
  Dim perfStart As Double, perfTotal As Double
  Dim perfCopy As Double, perfLink As Double, perfShade As Double
  Dim perfIterations As Long
  ' Siirretty loopista: Dim-lauseet kuuluvat aliohjelman alkuun (VBA nostaa ne kääntöaikana)
  Dim tCopy As Double, tShade As Double, tLink As Double
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
  ' Siirretty Sub-alkuun yhtenäisyyden vuoksi (review v2: kaikki Dim-lauseet kuuluvat tähän)
  Dim lastCell As Range
  Dim templateRange As Range
  Dim stagingSheet As Worksheet     ' Fix1: cross-WB staging (1 kopio loopille)
  
  On Error GoTo GenPrintoutError
  BeginFastMode
  
  ' Haetaan POSheet-nimi faceplatesta
  POSheet = Sheets("Main").Range("C16").Value
  If Trim(POSheet) = "" Then POSheet = "Printout" ' Oletusnimi jos ei asetettu
  Debug.Print "  POSheet nimi: " & POSheet
  
  ' Varmistetaan että dokumentin tiedot ovat ajantasalla (polku/nimi DB2:sta)
  On Error Resume Next
  If Trim(DocInfo.Path) = "" Or Trim(DocInfo.File) = "" Then HaeDocTiedot
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
    MsgBox "DB1-sheet on tyhjä! Klikkaa ensin 'Hae Data' ladataksesi tiedot tietokannasta.", vbCritical, "Ei dataa"
    Exit Sub
  End If
  
  Recordeja = lastCell.Row
  Debug.Print "  DB1 rivejä: " & Recordeja
  
  Application.StatusBar = "Luodaan uusi työkirja..."
  
  ' Luodaan uusi työkirja kopioimalla Info-sheet
  ' Korjattu: ActiveWorkbook voi pettää jos lisäosa aktivoi toisen työkirjan Copy-operaation jälkeen.
  ' Workbooks(Workbooks.Count) viittaa aina juuri lisättyyn työkirjaan turvallisesti.
  srcWB.Sheets("Info").Copy
  Set destWB = Workbooks(Workbooks.Count)
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
  ' Fix2: PrintCommunication=False estää tulostinohjain-kyselyt per kirjoitus
  Application.PrintCommunication = False
  destSheet.PageSetup.PrintTitleRows = "$" & ViimRivi & ":$" & ViimRivi + PHEnd - PHStart
  Application.PrintCommunication = True
  destSheet.Activate
  destSheet.Cells(ViimRivi + PHEnd - PHStart + 1, 1).Select
  ActiveWindow.FreezePanes = True
  ViimRivi = ViimRivi + 1 + PHEnd - PHStart
  
  ' Asetetaan alatunnisteet kolmelle ensimmäiselle sheetille (Info, POSheet, Legend)
  ' Fix2: PrintCommunication=False niputtaa kaikki 9 PageSetup-kirjoitusta yhteen tulostinohjain-pyyntöön
  Application.PrintCommunication = False
  Application.StatusBar = "Asetetaan alatunnisteet..."
  Debug.Print "  Asennetaan alatunnisteet dokumenttitiedoilla"
  For i = 1 To 3
    With destWB.Sheets(i).PageSetup
      ' Null-turvallinen merkkijonojen yhdistäminen (tyhjät arvot OK - joissakin laitteissa puuttuu tiettyjä attribuutteja)
      .LeftFooter = "&8Document: " & (DocInfo.MetsoDocNo & "") & Chr(10) _
                  & "&8Revision: " & (DocInfo.RevID & "") & " - " & (DocInfo.RevDate & "") & Chr(10) _
                  & "&8Status: " & (DocInfo.Status & "")
      .CenterFooter = "&8 " & (DocInfo.Customer & "") & Chr(10) _
                    & "&8 " & (DocInfo.Mill & "") & Chr(10) _
                    & "&8 " & (DocInfo.DepartName & "") & Chr(10) _
                    & "&8 " & (DocInfo.DocName2 & "")
      .RightFooter = "&8Project: " & (DocInfo.ProjNo & "") & Chr(10) _
                   & "&8File: &F" & Chr(10) _
                   & "&8Page &P(&N)"
    End With
  Next i
  Application.PrintCommunication = True
  
  ' Luodaan LINKING-sheet ja kopioidaan DB1-data
  Application.StatusBar = "Luodaan LINKING-sheet..."
  Debug.Print "  Luodaan LINKING-sheet DB1-datalla"
  With destWB.Sheets.Add(After:=destWB.Sheets(destWB.Sheets.Count))
    .Name = "LINKING"
    wsDB1.UsedRange.Copy Destination:=.Range("A1")  ' Fix5: Cells.Copy → UsedRange.Copy (vain data, ei 1M riviä)
  End With
  Application.CutCopyMode = False
  
  ' Alkuperäinen linkitys otsikkoalueelle
  Kerta = 0
  VaihdaLinkit destSheet, 1, ViimRivi, Kerta, srcWB
  
  ' Template-pohjainen täyttö: kopioidaan TEMPLATE-lohkoja ja kartoitetaan arvot VaihdaLinkit-funktiolla
  ' Fix1: TEMPLATE kopioidaan KERRAN paikalliseen __STAGING__-sheettiin destWB:ssä.
  ' Looppi käyttää saman-WB kopiointia (5-10x nopeampi kuin cross-WB per iteraatio).
  Application.StatusBar = "Kopioidaan dataa tulosteeseen käyttäen template-lohkoja..."
  Debug.Print "  Aloitetaan template-lohkojen kopiointi (RMAX=" & RMAX & ")"
  Riveja = DocEnd - DocStart
  If RMAX <= 0 Then RMAX = 1

  ' Siivotaan mahdollinen aiemman kaatuneen ajon staging-sheet
  Application.DisplayAlerts = False
  On Error Resume Next
  destWB.Sheets("__STAGING__").Delete
  On Error GoTo GenPrintoutError
  Application.DisplayAlerts = True

  ' Yksi cross-WB kopio — kaikki loopin kopiot tapahtuvat tästä eteenpäin saman WB:n sisällä
  Set stagingSheet = destWB.Sheets.Add(After:=destWB.Sheets(destWB.Sheets.Count))
  stagingSheet.Name = "__STAGING__"
  stagingSheet.Visible = xlSheetVeryHidden
  srcWB.Sheets("TEMPLATE").Rows(DocStart & ":" & DocEnd).Copy _
      Destination:=stagingSheet.Rows("1:1")
  Application.CutCopyMode = False
  Set templateRange = stagingSheet.Rows("1:" & (DocEnd - DocStart + 1))
  Debug.Print "  Staging-sheet luotu (__STAGING__) — loop käyttää saman-WB kopiointia"
  
  ' Iteroidaan DB1-datarivejä RMAX-ryhmissä, kopioidaan TEMPLATE-rivit joka kerralla
  Kerta = 0
  For i = 2 To Recordeja Step RMAX
    perfIterations = perfIterations + 1
    
    ' OPTIMOINTI: Kopioidaan lähdetyökirjasta
    ' (Huom: Ei voi täysin optimoida ilman array-lähestymistapaa koska tarvitaan muotoilua)
    tCopy = Timer
    templateRange.Copy Destination:=destSheet.Rows(ViimRivi & ":" & ViimRivi + Riveja)
    perfCopy = perfCopy + (Timer - tCopy)
    
    ' Lisätään vuorottelevatvarjostukset lohkoittain
    tShade = Timer
    If ((i - 2) \ RMAX) Mod 2 = 1 Then
      With destSheet.Range(destSheet.Cells(ViimRivi, 1), destSheet.Cells(ViimRivi + Riveja, Sarakkeita)).Interior
        .ColorIndex = 19
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
      End With
    End If
    perfShade = perfShade + (Timer - tShade)
    
    ' Kartoitetaan arvot DB1:stä template-alueelle kommenttimerkkien kautta (srcWB välitetään suoraa lukua varten)
    tLink = Timer
    VaihdaLinkit destSheet, ViimRivi, ViimRivi + Riveja, Kerta, srcWB
    perfLink = perfLink + (Timer - tLink)
    
    ' Siirrytään seuraavaan lohkoon
    ViimRivi = ViimRivi + Riveja + 1
    Kerta = Kerta + 1
  Next i
  Debug.Print "  Kopioitu " & perfIterations & " template-lohkoa"

  ' Fix1: Poistetaan staging-sheet loopin jälkeen
  Application.DisplayAlerts = False
  On Error Resume Next
  stagingSheet.Delete
  On Error GoTo GenPrintoutError
  Application.DisplayAlerts = True
  Set stagingSheet = Nothing
  Debug.Print "  Staging-sheet poistettu"

  ' Tyhjennetään leikepöytä kerran kaikkien kopioiden jälkeen (ei loopissa)
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
  
  ' Tyhjennetään kommentit — LINKING sisältää nyt staattiset arvot (ei kaavoja), TeeLinkingKommentit ohitetaan
  Application.StatusBar = "Viimeistellään..."
  destSheet.Cells.ClearComments
  
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
  
  ' Turvallinen käsittely mahdollisesti tyhjille/null DocInfo.Path ja DocInfo.File -arvoille
  On Error Resume Next
    defPath = Trim(DocInfo.Path & "")
    defName = Trim(DocInfo.File & "")
  On Error GoTo GenPrintoutError
  
  If defPath = "" Then defPath = ThisWorkbook.Path & Application.PathSeparator
  If Right$(defPath, 1) <> "\" And Right$(defPath, 1) <> "/" Then defPath = defPath & Application.PathSeparator
  
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
    ' Valitaan tallennusformaatti tiedostopäätteen mukaan — .xlsm vaatii makropohjaisen formaatin
    If LCase(Right$(Tiedosto, 5)) = ".xlsm" Then
      destWB.SaveAs Tiedosto, 52  ' xlOpenXMLWorkbookMacroEnabled
    Else
      destWB.SaveAs Tiedosto, xlOpenXMLWorkbook  ' xlOpenXMLWorkbook (51)
    End If
    Debug.Print Format(Now, "hh:mm:ss") & " [GenPrintout] Tallennettu: " & Tiedosto
  Else
    Debug.Print Format(Now, "hh:mm:ss") & " [GenPrintout] Tallennus peruttu"
  End If
  
  ' Suorituskykydiagnostiikka
  ' Keskiyön ylittyminen: Timer nollaantuu päivänvaihteessa — käytetään turvallista laskentaa
  Dim endTimer As Double
  endTimer = Timer
  If endTimer < perfStart Then
    perfTotal = (86400 - perfStart) + endTimer  ' Päivänvaihde ylittyi
  Else
    perfTotal = endTimer - perfStart
  End If
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
  ' Siivotaan staging-sheet jos se jäi kesken virheen sattuessa
  On Error Resume Next
  If Not stagingSheet Is Nothing Then
    Application.DisplayAlerts = False
    stagingSheet.Delete
    Application.DisplayAlerts = True
    Set stagingSheet = Nothing
  End If
  On Error GoTo 0
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
  ' Käytetään BeginFastMode/EndFastMode yhtenäisesti kuten muissakin makroissa
  BeginFastMode
  
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
      EndFastMode
      Debug.Print Format(Now, "hh:mm:ss") & " [Checkout ERROR] &&PAGE_HEADER_START puuttuu"
      MsgBox "TEMPLATE-sheetistä puuttuu vaaditut merkit! Katso ERRORS-sheet.", vbCritical, "Template-virhe"
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
    
    ' Footer-merkit ovat valinnaisia (riippuu AddFooter-checkboxista)
    PFStart = 0
    PFEnd = 0
    Set foundCell = .Cells.Find(What:="&&PAGE_FOOTER_START")
    If Not foundCell Is Nothing Then
      PFStart = foundCell.Row + 1
      Set foundCell = .Cells.Find(What:="&&PAGE_FOOTER_END")
      If Not foundCell Is Nothing Then
        PFEnd = foundCell.Row - 1
      Else
        Debug.Print "  VAROITUS: &&PAGE_FOOTER_START löytyi mutta &&PAGE_FOOTER_END puuttuu"
      End If
    End If
  End With
  
  If PFStart > 0 Then
    Debug.Print "  Template-merkit löydetty: PH=" & PHStart & ":" & PHEnd & ", Doc=" & DocStart & ":" & DocEnd & ", PF=" & PFStart & ":" & PFEnd & ", Cols=" & Sarakkeita
  Else
    Debug.Print "  Template-merkit löydetty: PH=" & PHStart & ":" & PHEnd & ", Doc=" & DocStart & ":" & DocEnd & ", PF=EI KÄYTÖSSÄ, Cols=" & Sarakkeita
  End If
  
  ' Haetaan dokumentin tiedot DB2-sheetiltä
  HaeDocTiedot
  
  ' Tarkistetaan ladattiinko dataa DB2:sta
  If DocInfo.Project = "" And DocInfo.DocNo = "" And DocInfo.ProjNo = "" And DocInfo.MetsoDocNo = "" Then
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
      If j > MAX_EXCEL_COLUMNS Then Exit For ' Turvatarkistus
      Arvo = wsTemplate.Cells(i, j).Value
      If Len(Arvo) > 2 Then 'Solussa on dataa
        If Left(Arvo, 2) = "££" Then
          If RMAX > 1 Then Virhe = True
          RMAX = 1
        ElseIf Left(Arvo, 1) = "£" Then
          ' Korjattu: käytetään Mid(Arvo, 2, 1) joka lukee rivinumeron £-merkin jälkeen
          ' (Mid(Arvo, 4, 1) luki välilyönnin "£1: "-muodossa ja antoi CInt(" ")=0)
          If RMAX <> 0 And RMAX <> CLng(Mid(Arvo, 2, 1)) Then Virhe = True
          RMAX = CLng(Mid(Arvo, 2, 1))
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
    EndFastMode
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
          ' Merkkimuoto: "£D SARAKE" — £=1, numero=2, kaksoispiste=3, välilyönti=4, sarakenimi alkaen 5
          ' Mid(Arvo, 2, 1) lukee rivinumeron D (korjattu RMAX-bugi), Mid(Arvo, 5) lukee sarakenimen
          ' Molemmat offsetit ovat johdettu samasta dokumentoidusta formaatista.
          If EtsiOts(Mid(Arvo, 5), i, j, CLng(Mid(Arvo, 2, 1))) = False Then Virhe = True
        End If
      End If
    Next j
  Next i
  
  If Virhe Then
    wsErrors.Activate
    EndFastMode
    Debug.Print Format(Now, "hh:mm:ss") & " [Checkout ERROR] Puuttuvia otsikkoita DB1:ssä"
    MsgBox "Templatessa oli virheitä! Katso ERRORS-sheet.", vbCritical, "Virhe!"
  Else
    Sheets("Main").Activate
    EndFastMode
    ' Varmistetaan että RMAX on asetettu — tyhjä datasetti jättäisi sen 0:ksi
    If RMAX = 0 Then
      CheckOK = False
      wsErrors.Range("A1").Value = "TEMPLATE ERROR: Yhtään rivimerkkiä (££ tai £1/2/3) ei löytynyt!"
      wsErrors.Range("A1").Font.Bold = True
      wsErrors.Range("A1").Font.ColorIndex = 3
      wsErrors.Activate
      EndFastMode
      Debug.Print Format(Now, "hh:mm:ss") & " [Checkout ERROR] RMAX=0 — template ei sisällä rivimerkkejä"
      MsgBox "Templatessa ei löytynyt rivimerkkejä (££ tai £1)! Katso ERRORS-sheet.", vbCritical, "RMAX-virhe"
      Exit Sub
    End If
    CheckOK = True
    Debug.Print Format(Now, "hh:mm:ss") & " [Checkout] VALMIS - CheckOK=True"
    MsgBox "Tarkistus OK!", vbOKOnly, "OK!"
  End If
  Exit Sub

CheckoutError:
  EndFastMode
  
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

