Option Compare Database
Option Explicit

'================================================================================
' Lomake: CopyLoops
' Tarkoitus: Tuo looppirivejä ja laitetaulukoita ulkoisesta Access-tietokannasta
' Päivitetty: 2025-11-11 - VBA7/64-bit-tuki lisätty
'             2026-03-03 - Me.-etuliitteet ja kommentit suomeksi
'
' Kuvaus:
'   Looppien tuontityökalu, joka:
'   - Linkittää ulkoiseen instru3-tietokantaan
'   - Näyttää saatavilla olevat loopit lähdetietokannasta
'   - Tuo valitut loopit LOOPS-taulukkoon
'   - Tuo laiterivit devTbl*-taulukoista
'   - Estää duplikaattilooppien tuonnin
'   - Antaa yksityiskohtaisen tilapalautteen
'
' Riippuvuudet:
'   - LOOPS-taulu (looppien kohdetaulu)
'   - LOOPLINK-väliaikainen linkitetty taulu
'   - devTable-taulu (laitteiden kohdetaulu)
'   - devTbl*-taulut (lähteen laitetaulut ulkoisessa tietokannassa)
'   - comdlg32.dll (Windows Common Dialog tiedostovalintaan)
'
' Kontrollit:
'   - Kanta: TextBox tietokannan polkua varten
'   - Loopit: Alilomake, joka näyttää saatavilla olevat loopit
'   - TTiedot: TextBox tilaviesteille
'   - ValitseKanta: Painike lähdetietokannan valintaan
'   - HaeLoopit: Painike looppien tuontiin
'   - Ulos: Painike lomakkeen sulkemiseen
'================================================================================

' Moduulitason muuttuja lähdetietokantayhteydelle – pysyy auki koko session ajan
Dim LahdeDB As DAO.Database

'--------------------------------------------------------------------------------
' Tapahtuma: Form_Unload
' Tarkoitus: Siivotaan väliaikaiset taulut ja yhteydet lomakkeen sulkeutuessa
'
' Prosessi:
'   1. Poistetaan väliaikainen LOOPLINK-taulu, jos se on olemassa
'   2. Suljetaan lähdetietokantayhteys
'   3. Tyhjennetään kaikki objektiviittaukset
'
' Huomiot:
'   - Resume Next estää virheen, jos taulua ei löydy
'   - Suoritetaan aina lomakkeen sulkeutuessa
'   - Varmistaa, ettei jää avoimia väliaikaistauluja tai yhteyksiä
'--------------------------------------------------------------------------------
Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    ' Vapautetaan alilomakkeen lukitukset ennen taulun poistoa — estää Access-kaatumisen
    Me.Loopit.Form.RecordSource = ""
    DoEvents  ' Annetaan Accessille hetki vapauttaa tiedostolukot
    CurrentDb.TableDefs.Delete "LOOPLINK"  ' Poistetaan väliaikainen LOOPLINK-taulu
    Err.Clear
    If Not LahdeDB Is Nothing Then
        LahdeDB.Close
        Set LahdeDB = Nothing
    End If
    On Error GoTo 0
End Sub
'--------------------------------------------------------------------------------
' Tapahtuma: HaeLoopit_Click
' Tarkoitus: Tuo valitut loopit ja laitedata lähdetietokannasta
'
' Prosessi:
'   1. Tarkistetaan, että lähdetietokanta on valittu
'   2. Kopioidaan loopit LOOPLINK-taulusta LOOPS-taulukkoon
'   3. Iteroidaan kaikki devTbl*-taulut lähdetietokannassa
'   4. Tuodaan laiterivit valituille loopeille
'   5. Asetetaan TABLEName-kenttä tuoduille laiteriveille
'   6. Seurataan duplikaatit ja annetaan tilapalaute
'
' Parametrit:
'   - Käyttää Loopit-alilomaketta lähteenä (linkitetty LOOPLINK-tauluun)
'   - Tuo LOOPS- ja devTable-tauluihin
'
' Huomiot:
'   - Estää duplikaattilooppien tuonnin (tarkistaa LOOPS-taulusta ensin)
'   - Tuo devTbl*- ja devTable-tauluista (kirjainkoosta riippumaton)
'   - Sulkee devTblCommon-taulun pois tuonnista
'   - Antaa yksityiskohtaisia tilaviestejä TTiedot-tekstikenttään
'   - Käyttää SysCmd-funktiota edistymisen näyttämiseen
'--------------------------------------------------------------------------------
Private Sub HaeLoopit_Click()
On Error GoTo ErrorHandler
    Dim OliJo As String  ' Kerää jo olemassa olevat (duplikaatti)loopit
    Dim LahdeTaulu As DAO.Recordset  ' Lähteen laitetaulun recordset
    Dim KohdeTaulu As DAO.Recordset  ' Kohdetaulun recordset (nykyinen tietokanta)
    Dim Kopioitavat() As String  ' Kopioitavien looppien lista (muoto: "AreaCode-LoopNo")
    Dim i As Integer, j As Integer, a As Integer  ' Silmukkalaskurit
    Dim Taulukko As String  ' Käsiteltävän taulun nimi
    
    Debug.Print "========================================"
    Debug.Print "HaeLoopit: Starting loop copy operation"
    Debug.Print "========================================"
    
    ReDim Kopioitavat(0)  ' Initialize array
    
    ' Validate source database is selected
    If LahdeDB Is Nothing Then
        Debug.Print "ERROR: No source database selected"
        MsgBox "Valitse ensin tietokanta!", vbCritical, "Hae loopit"
        Exit Sub
    End If
    
    Debug.Print "Source DB: " & LahdeDB.Name
  
  'Kopioidaan ensin loopit LOOPS-taulukkoon
  Me.TTiedot.Value = ""  ' Tyhjennetään tilatieto
  
  Debug.Print "Copying selected loops to LOOPS table..."
  
  ' Iterate through selected loops in subform
  With Me.Loopit.Form.RecordsetClone
    .MoveFirst
    Do While Not .EOF
      ' Check if loop already exists in target database
      Set KohdeTaulu = CurrentDb.OpenRecordset("SELECT * FROM LOOPS WHERE Areacode='" & .Fields("AreaCode").Value & "' AND LoopNo='" & .Fields("LoopNo").Value & "'")
      If KohdeTaulu.RecordCount = 0 Then
        ' New loop - add to import list and copy record
        Debug.Print "  Adding loop: " & .Fields("AreaCode").Value & "-" & .Fields("LoopNo").Value
        ReDim Preserve Kopioitavat(UBound(Kopioitavat) + 1)
        Kopioitavat(UBound(Kopioitavat)) = .Fields("AreaCode").Value & "-" & .Fields("LoopNo").Value
        KohdeTaulu.AddNew
        On Error Resume Next
        ' Copy all fields from source to target
        For i = 0 To .Fields.Count - 1
          KohdeTaulu.Fields(.Fields(i).Name).Value = .Fields(i).Value
        Next i
        KohdeTaulu.Update
        Err.Clear
        On Error GoTo 0
      Else
        ' Loop already exists - track for duplicate report
        OliJo = OliJo & .Fields("AreaCode") & "-" & .Fields("LoopNo") & vbCrLf
      End If
      .MoveNext
    Loop
  End With
  
  ' Import device records from devTbl* tables
  Debug.Print "Importing device records from devTbl* tables..."
  For i = 0 To LahdeDB.TableDefs.Count - 1
    Taulukko = LCase$(LahdeDB.TableDefs(i).Name)
    Set KohdeTaulu = CurrentDb.OpenRecordset("devTable")
    
    ' Check if table is a device table (devTbl* or devTable, excluding devTblCommon)
    If (Left$(Taulukko, 6) = "devtbl" And Left$(Taulukko, 12) <> "devtblcommon") Or (LCase$(Taulukko) = "devtable") Then
      Debug.Print "  Processing table: " & LahdeDB.TableDefs(i).Name
      SysCmd acSysCmdSetStatus, "Updating table: " & LahdeDB.TableDefs(i).Name  ' Show progress
      Set LahdeTaulu = LahdeDB.OpenRecordset(Taulukko)
      
      ' Iterate through all records in source device table
      Do While Not LahdeTaulu.EOF  'Käydään läpi kaikki laitetaulukon rivit
        ' Check if this record belongs to one of the imported loops
        For j = 1 To UBound(Kopioitavat)  'Tarkistetaan, kuuluuko rivi kopioitaviin looppeihin
          If Kopioitavat(j) = LahdeTaulu.Fields("AreaCode").Value & "-" & LahdeTaulu.Fields("LoopNo").Value Then
            ' Record belongs to imported loop - copy it
            KohdeTaulu.AddNew
            ' Copy all fields
            For a = 0 To LahdeTaulu.Fields.Count - 1
              KohdeTaulu.Fields(LahdeTaulu.Fields(a).Name).Value = LahdeTaulu.Fields(a).Value
              If Err <> 0 Then
                ' Log field copy errors
                Me.TTiedot.Value = Me.TTiedot.Value & "Ongelma kentässä: " & LahdeDB.TableDefs(i).Name & "!" & LahdeTaulu.Fields(a).Name
                Err.Clear
              End If
            Next a
            ' Set TableName field for device records (except devTable which already has it)
            If LCase$(Taulukko) <> "devtable" Then
              KohdeTaulu.Fields("TABLEName").Value = LahdeDB.TableDefs(i).Name
            End If
            KohdeTaulu.Update
            Exit For  ' Found match, move to next record
          End If
        Next j
        LahdeTaulu.MoveNext
      Loop
    End If
  Next i
  
  ' Display duplicate loop report
  If OliJo <> "" Then
    Debug.Print "Duplicate loops found (already in database):"
    Debug.Print OliJo
    Me.TTiedot.Value = Me.TTiedot.Value & "Seuraavat loopit olivat jo tietokannassa:" & vbCrLf & OliJo
  End If
  
  Debug.Print "HaeLoopit: COMPLETED successfully"
  Debug.Print "  Total loops copied: " & (UBound(Kopioitavat))
  Debug.Print "========================================"
  
  Me.TTiedot.Value = Me.TTiedot.Value & "Valmis!"
  SysCmd acSysCmdClearStatus
  MsgBox "Valmis!", vbInformation, "Hae loopit"
  
  ' Cleanup recordsets
  If Not KohdeTaulu Is Nothing Then
    KohdeTaulu.Close
    Set KohdeTaulu = Nothing
  End If
  If Not LahdeTaulu Is Nothing Then
    LahdeTaulu.Close
    Set LahdeTaulu = Nothing
  End If
  Exit Sub

ErrorHandler:
  Debug.Print "*** ERROR in HaeLoopit: " & Err.Number & " - " & Err.Description
  Debug.Print "    Source: " & Err.Source & ", Line: " & Erl
  Debug.Print "========================================"
  SysCmd acSysCmdClearStatus
  MsgBox "Virhe looppien tuonnissa: " & Err.Description, vbCritical, "Hae loopit"
  On Error Resume Next
  If Not KohdeTaulu Is Nothing Then KohdeTaulu.Close
  If Not LahdeTaulu Is Nothing Then LahdeTaulu.Close
  Set KohdeTaulu = Nothing
  Set LahdeTaulu = Nothing
  On Error GoTo 0
End Sub
'--------------------------------------------------------------------------------
' Tapahtuma: Ulos_Click
' Tarkoitus: Sulkee lomakkeen
'
' Huomiot:
'   - Yksinkertainen sulkemispainike
'   - Form_Unload-tapahtuma hoitaa siivousoperaatiot
'--------------------------------------------------------------------------------
Private Sub Ulos_Click()
  DoCmd.Close
End Sub
'--------------------------------------------------------------------------------
' Tapahtuma: ValitseKanta_Click
' Tarkoitus: Avaa tiedostovalintaikkuna lähde-Access-tietokannan valintaan
'
' Prosessi:
'   1. Määritellään alkuhakemisto (nykyinen tai aiemmin valittu kansio)
'   2. Näytetään Windows Common Dialog -tiedostovalinta
'   3. Tallennetaan valittu polku Kanta-tekstikenttään
'   4. Avataan tietokantayhteys (tallennetaan LahdeDB-moduulimuuttujaan)
'   5. Poistetaan mahdollinen olemassa oleva LOOPLINK-väliaikainen taulu
'   6. Luodaan uusi LOOPLINK, joka viittaa lähteen Loops-tauluun
'   7. Asetetaan alilomakkeen tietolähde näyttämään saatavilla olevat loopit
'
' Huomiot:
'   - Käyttää GetOpenFileName-APIa comdlg32.dll:stä
'   - Suodattaa vain .accdb-tiedostot
'   - Luo väliaikaisen linkitetyn taulun looppien selailua varten
'   - Lähdetietokanta pysyy auki, kunnes lomake sulkeutuu
'   - LOOPLINK poistetaan Form_Unload-tapahtumassa
'--------------------------------------------------------------------------------
Private Sub ValitseKanta_Click()
On Error GoTo ErrorHandler
    ' Avataan Office-natiivi tiedostovalintaikkuna tietokannan valitsemiseksi
    Dim fd As Object                   ' FileDialog-objekti
    Dim ValittuTiedosto As String      ' Käyttäjän valitsema tiedostopolku
    Dim AHakem As String               ' Alkuhakemisto dialogille
    Dim Linkki As DAO.TableDef         ' Väliaikaisen linkitetyn taulun määritys

    Debug.Print "ValitseKanta: Opening database selection dialog"

    ' Määritetään alkuhakemisto aiemmin valitusta tai nykykannasta
    If InStr(Nz(Me.Kanta.Value), "\") Then
        AHakem = Me.Kanta.Value
    Else
        AHakem = Application.CurrentDb.Name
    End If
    AHakem = Left$(AHakem, InStrRev(AHakem, "\"))

    ' Application.FileDialog korvaa haurauden OPENFILENAME-API-rakenteen (kaataa 64-bit Accessin)
    Set fd = Application.FileDialog(3)  ' msoFileDialogFilePicker = 3
    With fd
        .Title = "Choose Instru Database"
        .AllowMultiSelect = False
        .Filters.Clear
        .Filters.Add "Microsoft Access Databases", "*.accdb"
        .InitialFileName = AHakem
        If .Show = -1 Then
            ValittuTiedosto = .SelectedItems(1)
        End If
    End With
    Set fd = Nothing

    If ValittuTiedosto <> "" Then
        Me.Kanta.Value = ValittuTiedosto
        Debug.Print "  Valittu tietokanta: " & ValittuTiedosto

        ' Avataan lähdetietokantayhteys
        Set LahdeDB = OpenDatabase(ValittuTiedosto)
        Debug.Print "  Tietokanta avattu onnistuneesti"

        ' Poistetaan mahdollinen olemassa oleva LOOPLINK-taulu
        On Error Resume Next
        CurrentDb.TableDefs.Delete "LOOPLINK"
        Err.Clear
        On Error GoTo ErrorHandler

        ' Luodaan linkitetty taulu lähteen Loops-tauluun
        Set Linkki = CurrentDb.CreateTableDef("LOOPLINK")
        Linkki.Connect = ";DATABASE=" & ValittuTiedosto
        Linkki.SourceTableName = "Loops"
        CurrentDb.TableDefs.Append Linkki
        Set Linkki = Nothing

        ' Näytetään saatavilla olevat loopit alilomakkeessa
        Me.Loopit.Form.RecordSource = "SELECT * FROM LOOPLINK"
    End If
    Me.Refresh
    Exit Sub

ErrorHandler:
    Debug.Print "*** ERROR in ValitseKanta: " & Err.Number & " - " & Err.Description
    MsgBox "Virhe tietokannan valinnassa: " & Err.Description, vbCritical, "Valitse tietokanta"
    On Error Resume Next
    Set fd = Nothing
    Set Linkki = Nothing
    On Error GoTo 0
End Sub
'--------------------------------------------------------------------------------
' Tapahtuma: Form_Resize
' Tarkoitus: Säätää alilomakkeen koko lomakkeen ikkunakoon muuttuessa
'
' Prosessi:
'   - Muutetaan Loopit-alilomakkeen leveys vastaamaan lomakeikkunaa
'   - Säädetään korkeus, jos pystysuunnassa on riittävästi tilaa
'
' Huomiot:
'   - Ylläpitää 400 twipsin marginaalin oikealla puolella
'   - Korkeus säädetään vain, jos ikkuna on yli 3300 twipsiä korkea
'   - Tarjoaa mukautuvan käyttöliittymän looppien katseluun
'--------------------------------------------------------------------------------
Private Sub Form_Resize()
  Me.Loopit.Width = Me.WindowWidth - 400  ' Säilytetään oikea marginaali
  If Me.WindowHeight > 3300 Then
    Me.Loopit.Height = Me.WindowHeight - 3300  ' Suurennetaan korkeus, kun tilaa on
  End If
End Sub

