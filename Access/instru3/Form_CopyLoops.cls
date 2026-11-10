Option Compare Database
Option Explicit

'================================================================================
' Lomake: CopyLoops
' Tarkoitus: Tuo looppirivej‰ ja laitetaulukoita ulkoisesta Access-tietokannasta
' P‰ivitetty: 2025-11-11 - VBA7/64-bit-tuki lis‰tty
'             2026-03-03 - Me.-etuliitteet ja kommentit suomeksi
'             2026-11-10 - Transaktio, virheenk‰sittely ja SQL-heittomerkkien suojaus,
'                          debug-tulosteet poistettu, kommenttien viimeistely
'
' Kuvaus:
'   Looppien tuontityˆkalu, joka:
'   - Linkitt‰‰ ulkoiseen instru3-tietokantaan
'   - N‰ytt‰‰ saatavilla olevat loopit l‰hdetietokannasta
'   - Tuo valitut loopit LOOPS-taulukkoon
'   - Tuo laiterivit devTbl*-taulukoista
'   - Est‰‰ duplikaattilooppien tuonnin
'   - Antaa yksityiskohtaisen tilapalautteen
'
' Riippuvuudet:
'   - LOOPS-taulu (looppien kohdetaulu)
'   - LOOPLINK-v‰liaikainen linkitetty taulu
'   - devTable-taulu (laitteiden kohdetaulu)
'   - devTbl*-taulut (l‰hteen laitetaulut ulkoisessa tietokannassa)
'   - Application.FileDialog (Office-tiedostovalinta)
'
' Kontrollit:
'   - Kanta: TextBox tietokannan polkua varten
'   - Loopit: Alilomake, joka n‰ytt‰‰ saatavilla olevat loopit
'   - TTiedot: TextBox tilaviesteille
'   - ValitseKanta: Painike l‰hdetietokannan valintaan
'   - HaeLoopit: Painike looppien tuontiin
'   - Ulos: Painike lomakkeen sulkemiseen
'================================================================================

' Moduulitason muuttuja l‰hdetietokantayhteydelle ñ pysyy auki koko session ajan
Dim LahdeDB As DAO.Database

'--------------------------------------------------------------------------------
' Tapahtuma: Form_Unload
' Tarkoitus: Siivotaan v‰liaikaiset taulut ja yhteydet lomakkeen sulkeutuessa
'
' Prosessi:
'   1. Poistetaan v‰liaikainen LOOPLINK-taulu, jos se on olemassa
'   2. Suljetaan l‰hdetietokantayhteys
'   3. Tyhjennet‰‰n kaikki objektiviittaukset
'
' Huomiot:
'   - Resume Next est‰‰ virheen, jos taulua ei lˆydy
'   - Suoritetaan aina lomakkeen sulkeutuessa
'   - Varmistaa, ettei j‰‰ avoimia v‰liaikaistauluja tai yhteyksi‰
'--------------------------------------------------------------------------------
Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    ' Vapautetaan alilomakkeen lukitukset ennen taulun poistoa ó est‰‰ Access-kaatumisen
    Me.Loopit.Form.RecordSource = ""
    DoEvents  ' Annetaan Accessille hetki vapauttaa tiedostolukot
    CurrentDb.TableDefs.Delete "LOOPLINK"  ' Poistetaan v‰liaikainen LOOPLINK-taulu
    Err.Clear
    If Not LahdeDB Is Nothing Then
        LahdeDB.Close
        Set LahdeDB = Nothing
    End If
    On Error GoTo 0
End Sub
'--------------------------------------------------------------------------------
' Tapahtuma: HaeLoopit_Click
' Tarkoitus: Tuo valitut loopit ja laitedata l‰hdetietokannasta
'
' Prosessi:
'   1. Tarkistetaan, ett‰ l‰hdetietokanta on valittu
'   2. Kopioidaan loopit LOOPLINK-taulusta LOOPS-taulukkoon
'   3. Iteroidaan kaikki devTbl*-taulut l‰hdetietokannassa
'   4. Tuodaan laiterivit valituille loopeille
'   5. Asetetaan TABLEName-kentt‰ tuoduille laiteriveille
'   6. Seurataan duplikaatit ja annetaan tilapalaute
'
' Parametrit:
'   - K‰ytt‰‰ Loopit-alilomaketta l‰hteen‰ (linkitetty LOOPLINK-tauluun)
'   - Tuo LOOPS- ja devTable-tauluihin
'
' Huomiot:
'   - Est‰‰ duplikaattilooppien tuonnin (tarkistaa LOOPS-taulusta ensin)
'   - Tuo devTbl*- ja devTable-tauluista (kirjainkoosta riippumaton)
'   - Sulkee devTblCommon-taulun pois tuonnista
'   - Antaa yksityiskohtaisia tilaviestej‰ TTiedot-tekstikentt‰‰n
'   - K‰ytt‰‰ SysCmd-funktiota edistymisen n‰ytt‰miseen
'   - Koko tuonti on yhdess‰ transaktiossa: virhe perutaan kokonaan
'--------------------------------------------------------------------------------
Private Sub HaeLoopit_Click()
On Error GoTo ErrorHandler
    Dim OliJo As String  ' Ker‰‰ jo olemassa olevat (duplikaatti)loopit
    Dim LahdeTaulu As DAO.Recordset  ' L‰hteen laitetaulun recordset
    Dim KohdeTaulu As DAO.Recordset  ' Kohdetaulun recordset (nykyinen tietokanta)
    Dim Kopioitavat() As String  ' Kopioitavien looppien lista (muoto: "AreaCode-LoopNo")
    Dim i As Integer, j As Integer, a As Integer  ' Silmukkalaskurit
    Dim Taulukko As String  ' K‰sitelt‰v‰n taulun nimi
    Dim Avain As String  ' Laiterivin looppiavain (AreaCode-LoopNo)
    Dim Tyotila As DAO.Workspace  ' Tyˆtila transaktiota varten
    Dim TransAktiivinen As Boolean  ' Onko transaktio k‰ynniss‰
    Dim virhenumero As Long  ' Virheen numero k‰sittelij‰ss‰
    Dim virhekuvaus As String  ' Virheen kuvaus k‰sittelij‰ss‰

    ReDim Kopioitavat(0)

    If LahdeDB Is Nothing Then
        MsgBox "Valitse ensin tietokanta!", vbCritical, "Hae loopit"
        Exit Sub
    End If

    Me.TTiedot.Value = ""  ' Tyhjennet‰‰n tilatieto

    ' Koko tuonti tehd‰‰n yhdess‰ transaktiossa, jottei virhe j‰t‰ osittaista tulosta
    Set Tyotila = DBEngine.Workspaces(0)
    Tyotila.BeginTrans
    TransAktiivinen = True

    ' Kopioidaan valitut loopit LOOPS-taulukkoon
    With Me.Loopit.Form.RecordsetClone
        If Not .EOF Then .MoveFirst
        Do While Not .EOF
            ' Heittomerkit tuplataan, jottei arvo riko SQL-lausetta
            Set KohdeTaulu = CurrentDb.OpenRecordset("SELECT * FROM LOOPS WHERE Areacode='" & Replace(CStr(Nz(.Fields("AreaCode").Value, "")), "'", "''") & "' AND LoopNo='" & Replace(CStr(Nz(.Fields("LoopNo").Value, "")), "'", "''") & "'")
            If KohdeTaulu.RecordCount = 0 Then
                KohdeTaulu.AddNew
                ' Kohteesta puuttuvat kent‰t ohitetaan; Update-virhe sen sijaan keskeytt‰‰ tuonnin
                On Error Resume Next
                For i = 0 To .Fields.Count - 1
                    KohdeTaulu.Fields(.Fields(i).Name).Value = .Fields(i).Value
                Next i
                Err.Clear
                On Error GoTo ErrorHandler
                KohdeTaulu.Update
                ' Looppi lis‰t‰‰n laiterivien tuontilistalle vasta, kun tallennus onnistui
                ReDim Preserve Kopioitavat(UBound(Kopioitavat) + 1)
                Kopioitavat(UBound(Kopioitavat)) = .Fields("AreaCode").Value & "-" & .Fields("LoopNo").Value
            Else
                OliJo = OliJo & .Fields("AreaCode") & "-" & .Fields("LoopNo") & vbCrLf
            End If
            KohdeTaulu.Close
            Set KohdeTaulu = Nothing
            .MoveNext
        Loop
    End With

    ' Tuodaan laiterivit devTbl*-tauluista
    Set KohdeTaulu = CurrentDb.OpenRecordset("devTable")
    For i = 0 To LahdeDB.TableDefs.Count - 1
        Taulukko = LCase$(LahdeDB.TableDefs(i).Name)

        ' Laitetaulu on devTbl* tai devTable (devTblCommon pois lukien)
        If (Left$(Taulukko, 6) = "devtbl" And Left$(Taulukko, 12) <> "devtblcommon") Or Taulukko = "devtable" Then
            SysCmd acSysCmdSetStatus, "P‰ivitet‰‰n taulua: " & LahdeDB.TableDefs(i).Name
            Set LahdeTaulu = LahdeDB.OpenRecordset(LahdeDB.TableDefs(i).Name)

            Do While Not LahdeTaulu.EOF
                Avain = Nz(LahdeTaulu.Fields("AreaCode").Value, "") & "-" & Nz(LahdeTaulu.Fields("LoopNo").Value, "")
                ' Kopioidaan vain tuotuihin looppeihin kuuluvat rivit
                For j = 1 To UBound(Kopioitavat)
                    If Kopioitavat(j) = Avain Then
                        KohdeTaulu.AddNew
                        For a = 0 To LahdeTaulu.Fields.Count - 1
                            On Error Resume Next
                            KohdeTaulu.Fields(LahdeTaulu.Fields(a).Name).Value = LahdeTaulu.Fields(a).Value
                            If Err.Number <> 0 Then
                                Me.TTiedot.Value = Me.TTiedot.Value & "Ongelma kent‰ss‰: " & LahdeDB.TableDefs(i).Name & "!" & LahdeTaulu.Fields(a).Name & vbCrLf
                                Err.Clear
                            End If
                            On Error GoTo ErrorHandler
                        Next a
                        ' devTable-l‰hteess‰ TABLEName on jo mukana
                        If Taulukko <> "devtable" Then
                            KohdeTaulu.Fields("TABLEName").Value = LahdeDB.TableDefs(i).Name
                        End If
                        KohdeTaulu.Update
                        Exit For
                    End If
                Next j
                LahdeTaulu.MoveNext
            Loop
            LahdeTaulu.Close
            Set LahdeTaulu = Nothing
        End If
    Next i
    KohdeTaulu.Close
    Set KohdeTaulu = Nothing

    Tyotila.CommitTrans
    TransAktiivinen = False
    Set Tyotila = Nothing

    If OliJo <> "" Then
        Me.TTiedot.Value = Me.TTiedot.Value & "Seuraavat loopit olivat jo tietokannassa:" & vbCrLf & OliJo
    End If
    Me.TTiedot.Value = Me.TTiedot.Value & "Valmis!"
    SysCmd acSysCmdClearStatus
    MsgBox "Valmis!", vbInformation, "Hae loopit"
    Exit Sub

ErrorHandler:
    virhenumero = Err.Number
    virhekuvaus = Err.Description
    On Error Resume Next
    ' Perutaan kaikki tuonnin muutokset
    If TransAktiivinen Then Tyotila.Rollback
    If Not KohdeTaulu Is Nothing Then KohdeTaulu.Close
    If Not LahdeTaulu Is Nothing Then LahdeTaulu.Close
    Set KohdeTaulu = Nothing
    Set LahdeTaulu = Nothing
    Set Tyotila = Nothing
    SysCmd acSysCmdClearStatus
    On Error GoTo 0
    MsgBox "Virhe looppien tuonnissa, muutokset peruttiin (" & virhenumero & "): " & virhekuvaus, vbCritical, "Hae loopit"
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
' Tarkoitus: Avaa tiedostovalintaikkuna l‰hde-Access-tietokannan valintaan
'
' Prosessi:
'   1. M‰‰ritell‰‰n alkuhakemisto (nykyinen tai aiemmin valittu kansio)
'   2. N‰ytet‰‰n Office FileDialog -tiedostovalinta
'   3. Tallennetaan valittu polku Kanta-tekstikentt‰‰n
'   4. Avataan tietokantayhteys (tallennetaan LahdeDB-moduulimuuttujaan)
'   5. Poistetaan mahdollinen olemassa oleva LOOPLINK-v‰liaikainen taulu
'   6. Luodaan uusi LOOPLINK, joka viittaa l‰hteen Loops-tauluun
'   7. Asetetaan alilomakkeen tietol‰hde n‰ytt‰m‰‰n saatavilla olevat loopit
'
' Huomiot:
'   - K‰ytt‰‰ Application.FileDialog-tiedostovalintaa (ei Windows API -kutsuja)
'   - Suodattaa vain .accdb-tiedostot
'   - Luo v‰liaikaisen linkitetyn taulun looppien selailua varten
'   - L‰hdetietokanta pysyy auki, kunnes lomake sulkeutuu
'   - LOOPLINK poistetaan Form_Unload-tapahtumassa
'--------------------------------------------------------------------------------
Private Sub ValitseKanta_Click()
On Error GoTo ErrorHandler
    ' Avataan Office-natiivi tiedostovalintaikkuna tietokannan valitsemiseksi
    Dim fd As Object                   ' FileDialog-objekti
    Dim ValittuTiedosto As String      ' K‰ytt‰j‰n valitsema tiedostopolku
    Dim AHakem As String               ' Alkuhakemisto dialogille
    Dim Linkki As DAO.TableDef         ' V‰liaikaisen linkitetyn taulun m‰‰ritys

    ' M‰‰ritet‰‰n alkuhakemisto aiemmin valitusta tai nykykannasta
    If InStr(Nz(Me.Kanta.Value), "\") Then
        AHakem = Me.Kanta.Value
    Else
        AHakem = Application.CurrentDb.Name
    End If
    AHakem = Left$(AHakem, InStrRev(AHakem, "\"))

    ' Application.FileDialog korvaa ep‰vakaan OPENFILENAME-API-rakenteen (kaatoi 64-bittisen Accessin)
    Set fd = Application.FileDialog(3)  ' msoFileDialogFilePicker = 3
    With fd
        .Title = "Valitse instru-tietokanta"
        .AllowMultiSelect = False
        .Filters.Clear
        .Filters.Add "Access-tietokannat", "*.accdb"
        .InitialFileName = AHakem
        If .Show = -1 Then
            ValittuTiedosto = .SelectedItems(1)
        End If
    End With
    Set fd = Nothing

    If ValittuTiedosto <> "" Then
        Me.Kanta.Value = ValittuTiedosto

        ' Suljetaan aiempi l‰hdeyhteys, jottei se j‰‰ vuotamaan uutta valintaa teht‰ess‰
        If Not LahdeDB Is Nothing Then
            LahdeDB.Close
            Set LahdeDB = Nothing
        End If

        ' Avataan l‰hdetietokantayhteys
        Set LahdeDB = OpenDatabase(ValittuTiedosto)

        ' Poistetaan mahdollinen olemassa oleva LOOPLINK-taulu
        On Error Resume Next
        CurrentDb.TableDefs.Delete "LOOPLINK"
        Err.Clear
        On Error GoTo ErrorHandler

        ' Luodaan linkitetty taulu l‰hteen Loops-tauluun
        Set Linkki = CurrentDb.CreateTableDef("LOOPLINK")
        Linkki.Connect = ";DATABASE=" & ValittuTiedosto
        Linkki.SourceTableName = "Loops"
        CurrentDb.TableDefs.Append Linkki
        Set Linkki = Nothing

        ' N‰ytet‰‰n saatavilla olevat loopit alilomakkeessa
        Me.Loopit.Form.RecordSource = "SELECT * FROM LOOPLINK"
    End If
    Me.Refresh
    Exit Sub

ErrorHandler:
    MsgBox "Virhe tietokannan valinnassa: " & Err.Description, vbCritical, "Valitse tietokanta"
    On Error Resume Next
    Set fd = Nothing
    Set Linkki = Nothing
    On Error GoTo 0
End Sub
'--------------------------------------------------------------------------------
' Tapahtuma: Form_Resize
' Tarkoitus: S‰‰t‰‰ alilomakkeen koko lomakkeen ikkunakoon muuttuessa
'
' Prosessi:
'   - Muutetaan Loopit-alilomakkeen leveys vastaamaan lomakeikkunaa
'   - S‰‰det‰‰n korkeus, jos pystysuunnassa on riitt‰v‰sti tilaa
'
' Huomiot:
'   - Yll‰pit‰‰ 400 twipsin marginaalin oikealla puolella
'   - Korkeus s‰‰det‰‰n vain, jos ikkuna on yli 3300 twipsi‰ korkea
'   - Tarjoaa mukautuvan k‰yttˆliittym‰n looppien katseluun
'--------------------------------------------------------------------------------
Private Sub Form_Resize()
  Me.Loopit.Width = Me.WindowWidth - 400  ' S‰ilytet‰‰n oikea marginaali
  If Me.WindowHeight > 3300 Then
    Me.Loopit.Height = Me.WindowHeight - 3300  ' Suurennetaan korkeus, kun tilaa on
  End If
End Sub

