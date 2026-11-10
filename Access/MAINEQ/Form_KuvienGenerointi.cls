Option Compare Database
Option Explicit

'==============================================================================
' Lomake: KuvienGenerointi
' Tarkoitus: Generoi AutoCAD-kuvia tietokannan tiedoista:
'   - GenRev:       revisiosivut (indeksikuvat) piirustuksista
'   - Teetaulukko:  poimii kuviin päivitettävät tekstit TiedotKuviin-tauluun
'   - GenKuvat:     päivittää tekstit pohjakuviin ja tallentaa valmiit kuvat
' Päivitetty: 2025-11-11 - Myöhäinen sidonta AutoCADiin, DAO-tyypit
'             2026-11-10 - Virheenkäsittely ja siivous, SQL-heittomerkkien suojaus,
'                          transaktio taulun tekoon, muiden avointen piirustusten
'                          sulkeminen poistettu
'
' Riippuvuudet:
'   - Käynnissä oleva AutoCAD (GetObject)
'   - Taulut/kyselyt: zTITLE_Tbl, zTitle_, zRevisio, Tekstit, TiedotKuviin, qryKuviin
'==============================================================================

' Myöhäinen sidonta AutoCADiin - tyyppikirjastoviittausta ei tarvita
Private oAcad As Object  ' AcadApplication
Private oDoc As Object   ' AcadDocument

' Yhdistää käynnissä olevaan AutoCADiin; palauttaa False, jos sitä ei löydy
Private Function YhdistaAutoCAD() As Boolean
    On Error Resume Next
    Set oAcad = GetObject(, "AutoCAD.Application")
    If Err.Number <> 0 Or oAcad Is Nothing Then
        Err.Clear
        Set oAcad = Nothing
        MsgBox "Käynnissä olevaa AutoCADiä ei löytynyt!", vbCritical, "Generoi kuvat"
    Else
        YhdistaAutoCAD = True
    End If
End Function

' Tarkistaa, onko piirustus jo auki AutoCADissa (ettei käyttäjän avointa kuvaa suljeta)
Private Function OnkoAvoinna(ByVal Polku As String) As Boolean
    Dim Dok As Object  ' Avoin piirustus

    For Each Dok In oAcad.Documents
        If StrComp(Dok.FullName, Polku, vbTextCompare) = 0 Then
            OnkoAvoinna = True
            Exit Function
        End If
    Next Dok
End Function

' Tarkistaa, että tallennettavan tiedoston hakemisto on olemassa
Private Function HakemistoOlemassa(ByVal Polku As String) As Boolean
    Dim Pos As Long  ' Viimeisen kenoviivan sijainti

    Pos = InStrRev(Polku, "\")
    If Pos = 0 Then
        HakemistoOlemassa = True
    Else
        HakemistoOlemassa = (Len(Dir$(Left$(Polku, Pos), vbDirectory)) > 0)
    End If
End Function

' Tuplaa heittomerkit, jotta arvo voidaan liittää SQL-merkkijonoon
Private Function SqlTeksti(ByVal Arvo As Variant) As String
    SqlTeksti = Replace(CStr(Nz(Arvo, "")), "'", "''")
End Function

' Hakee kyselyn ensimmäisen rivin ensimmäisen kentän; palauttaa False, jos rivejä ei ole
Private Function HaeEnsimmainen(ByVal DB As DAO.Database, ByVal Kysely As String, ByRef Arvo As Variant) As Boolean
    Dim rs As DAO.Recordset  ' Kyselyn tulos

    Set rs = DB.OpenRecordset(Kysely, dbOpenSnapshot)
    If Not rs.EOF Then
        Arvo = rs.Fields(0).Value
        HaeEnsimmainen = True
    End If
    rs.Close
    Set rs = Nothing
End Function

'------------------------------------------------------------------------------
' Tapahtuma: GenRev_Click
' Tarkoitus: Generoi revisiosivut (indeksikuvat) jokaiselle piirustukselle
'------------------------------------------------------------------------------
Private Sub GenRev_Click()
On Error GoTo ErrorHandler
    Dim Kuvat As DAO.Recordset  ' Generoitavat piirustukset
    Dim Revisiot As DAO.Recordset  ' Piirustuksen revisiot
    Dim DB As DAO.Database  ' Nykyinen tietokanta
    Dim Pohja As String  ' Indeksipohjan tiedosto
    Dim Talletus As String  ' Valmiin kuvan tiedosto
    Dim i As Integer  ' Revisiosivun rivi
    Dim Sarake As Integer  ' Revisiokirjaimen sarake
    Dim Spositio As String  ' Edellisen rivin Spositio
    Dim virhenumero As Long  ' Virheen numero käsittelijässä
    Dim virhekuvaus As String  ' Virheen kuvaus käsittelijässä

    If Not YhdistaAutoCAD Then Exit Sub

    Set DB = CurrentDb
    Set Kuvat = DB.OpenRecordset("SELECT TyyppipHak, Hakemisto, PiirNro FROM zTITLE_Tbl GROUP BY TyyppipHak, Hakemisto, PiirNro;", dbOpenSnapshot)
    If Kuvat.EOF Then
        MsgBox "Generoitavia kuvia ei löytynyt.", vbInformation, "Generoi kuvat"
        GoTo Siivous
    End If

    Do While Not Kuvat.EOF
        SysCmd acSysCmdSetStatus, "Tekee kuvaa: " & Kuvat.Fields("PiirNro").Value
        Pohja = Kuvat.Fields("TyyppipHak").Value & "IndexForRevision.DWG"
        Talletus = Kuvat.Fields("Hakemisto").Value & Kuvat.Fields("PiirNro").Value & "_INDICE.DWG"
        If OnkoAvoinna(Pohja) Then Err.Raise vbObjectError + 1001, , "Pohjakuva on jo auki AutoCADissa, sulje se ensin: " & Pohja
        If Not HakemistoOlemassa(Talletus) Then Err.Raise vbObjectError + 1002, , "Tallennushakemistoa ei löydy: " & Talletus

        Set oDoc = oAcad.Documents.Open(Pohja)

        ' Otsikkorivi ja ensimmäisen sarakkeen X-merkki
        i = 1
        LisaaTeksti i, 0, "01/INDICE"
        LisaaTeksti i, 1, "X"

        ' Haetaan kuvaan lisättävät revisiot (Spositio nollataan jokaiselle kuvalle)
        Spositio = ""
        Set Revisiot = DB.OpenRecordset("SELECT zRevisio.Spositio, FOLHA, RevKirjain, Ext2 FROM zRevisio INNER JOIN zTITLE_Tbl ON zRevisio.Spositio = zTITLE_Tbl.Spositio WHERE (((zTITLE_Tbl.PiirNro) = '" & SqlTeksti(Kuvat.Fields("PiirNro").Value) & "')) GROUP BY zRevisio.Spositio, zRevisio.FOLHA, zRevisio.RevKirjain, zTITLE_Tbl.Ext2 ORDER BY zTITLE_Tbl.Ext2;", dbOpenSnapshot)
        Do While Not Revisiot.EOF
            ' Uusi Spositio aloittaa uuden rivin
            If Spositio <> CStr(Nz(Revisiot.Fields("Spositio").Value, "")) Then
                i = i + 1
                Spositio = CStr(Nz(Revisiot.Fields("Spositio").Value, ""))
                LisaaTeksti i, 0, CStr(Nz(Revisiot.Fields("FOLHA").Value, ""))
            End If
            Sarake = AnnaSarake(CStr(Nz(Revisiot.Fields("RevKirjain").Value, "")))
            ' Tuntematon revisiokirjain (0) ohitetaan, ettei se ylikirjoita sivun nimeä
            If Sarake > 0 Then LisaaTeksti i, Sarake, "X"
            Revisiot.MoveNext
        Loop
        Revisiot.Close
        Set Revisiot = Nothing

        oDoc.SaveAs Talletus
        oDoc.Close
        Set oDoc = Nothing
        Kuvat.MoveNext
    Loop
    MsgBox "Valmis", vbInformation, "Generoi kuvat"

Siivous:
    On Error Resume Next
    SysCmd acSysCmdClearStatus
    If Not Revisiot Is Nothing Then Revisiot.Close
    If Not Kuvat Is Nothing Then Kuvat.Close
    Set Revisiot = Nothing
    Set Kuvat = Nothing
    Set DB = Nothing
    Set oDoc = Nothing
    Set oAcad = Nothing
    If virhenumero <> 0 Then
        MsgBox "Virhe revisiosivujen generoinnissa (" & virhenumero & "): " & virhekuvaus, vbCritical, "Generoi kuvat"
    End If
    Exit Sub

ErrorHandler:
    virhenumero = Err.Number
    virhekuvaus = Err.Description
    On Error Resume Next
    ' Suljetaan keskeneräinen piirustus tallentamatta
    If Not oDoc Is Nothing Then oDoc.Close False
    Resume Siivous
End Sub

'------------------------------------------------------------------------------
' Tapahtuma: Teetaulukko_Click
' Tarkoitus: Poimii jokaiselle kuvalle päivitettävät tekstit (TiedotKuviin)
'
' Huomiot:
'   - Poistetaan vain sulkemattomat rivit (Frozen=0)
'   - Koko poiminta on yhdessä transaktiossa: virhe palauttaa taulun entiselleen
'------------------------------------------------------------------------------
Private Sub Teetaulukko_Click()
On Error GoTo ErrorHandler
    Dim Spositio As String  ' Kuvan Spositio
    Dim Tyyppipiir As String  ' Pohjakuvan nimi
    Dim Maara As Long  ' Kuvien kokonaismäärä
    Dim i As Long  ' Käsitelty kuva
    Dim DB As DAO.Database  ' Nykyinen tietokanta
    Dim Tyotila As DAO.Workspace  ' Työtila transaktiota varten
    Dim TransAktiivinen As Boolean  ' Onko transaktio käynnissä
    Dim zTitleTbl As DAO.Recordset  ' Kuvien otsikkotiedot
    Dim KuviinTbl As DAO.Recordset  ' Kohdetaulu
    Dim TEKSTITTbl As DAO.Recordset  ' Kuvan päivitettävät tekstit
    Dim EntType As String  ' Tekstin tyyppi
    Dim Kentta1 As String  ' Tekstit.Field1
    Dim Kentta2 As String  ' Tekstit.Field2
    Dim Kentta3 As String  ' Tekstit.Field3
    Dim Tieto As Variant  ' Kuvaan kirjoitettava tieto
    Dim Lisaa As Boolean  ' Lisätäänkö rivi TiedotKuviin-tauluun
    Dim virhenumero As Long  ' Virheen numero käsittelijässä
    Dim virhekuvaus As String  ' Virheen kuvaus käsittelijässä

    Set DB = CurrentDb
    Set Tyotila = DBEngine.Workspaces(0)
    Tyotila.BeginTrans
    TransAktiivinen = True

    DB.Execute "DELETE * FROM TiedotKuviin WHERE Frozen=0", dbFailOnError

    Set KuviinTbl = DB.OpenRecordset("TiedotKuviin", dbOpenDynaset)
    Set zTitleTbl = DB.OpenRecordset("zTitle_", dbOpenDynaset)

    If Not zTitleTbl.EOF Then
        zTitleTbl.MoveLast
        zTitleTbl.MoveFirst
        Maara = zTitleTbl.RecordCount
    End If

    Do While Not zTitleTbl.EOF
        i = i + 1
        Spositio = CStr(Nz(zTitleTbl("Spositio").Value, ""))
        Tyyppipiir = CStr(Nz(zTitleTbl("Tyyppipiir").Value, ""))
        Set TEKSTITTbl = DB.OpenRecordset("SELECT * FROM Tekstit WHERE DWG='" & SqlTeksti(Tyyppipiir) & "' AND EntType Is Not Null", dbOpenSnapshot)
        SysCmd acSysCmdSetStatus, "Poimii tietoja kuvaan (" & i & "/" & Maara & "): " & zTitleTbl("FileName").Value & "..."

        Do While Not TEKSTITTbl.EOF
            Tieto = ""
            EntType = CStr(TEKSTITTbl("EntType").Value)
            Kentta1 = CStr(Nz(TEKSTITTbl("Field1").Value, ""))
            Kentta2 = CStr(Nz(TEKSTITTbl("Field2").Value, ""))
            Kentta3 = CStr(Nz(TEKSTITTbl("Field3").Value, ""))

            Select Case UCase$(EntType)
                Case "TITLE"
                    Tieto = zTitleTbl(Kentta1).Value
                Case "REVISIO", "EMISIO", "I_O"
                    ' Haetaan tieto revisio-, emisio- tai I_O-taulukosta
                    HaeEnsimmainen DB, "SELECT [" & Kentta1 & "] FROM z" & EntType & " WHERE Spositio='" & SqlTeksti(Spositio) & "' AND " & EntType & "_ID='" & SqlTeksti(Kentta3) & "'", Tieto
                Case "DETAIL", "NOTE"
                    ' Haetaan tieto detail- tai note-taulukosta
                    HaeEnsimmainen DB, "SELECT [" & Kentta1 & "] FROM z" & EntType & " WHERE Spositio='" & SqlTeksti(Spositio) & "'", Tieto
                Case "DEVICE", "CABLE"
                    ' Haetaan tieto zDevice- tai zCable-taulukosta
                    If LCase$(Kentta1) = "tagid" Then
                        If HaeEnsimmainen(DB, "SELECT TAG FROM z" & EntType & " WHERE Spositio='" & SqlTeksti(Spositio) & "' AND TagID='" & SqlTeksti(Kentta3) & "' AND Suffix='" & SqlTeksti(Kentta2) & "'", Tieto) Then
                            If IsNull(Tieto) Then Tieto = TEKSTITTbl("VALUE").Value
                        Else
                            Tieto = TEKSTITTbl("VALUE").Value
                        End If
                    Else
                        HaeEnsimmainen DB, "SELECT [" & Kentta1 & "] FROM z" & EntType & " WHERE Spositio='" & SqlTeksti(Spositio) & "' AND TagID='" & SqlTeksti(Kentta3) & "' AND Suffix='" & SqlTeksti(Kentta2) & "'", Tieto
                    End If
            End Select

            ' Tyhjää merkkijonoa ("tietoa ei löytynyt") ei kirjoiteta, mutta Null kirjoitetaan
            If IsNull(Tieto) Then
                Lisaa = True
            Else
                Lisaa = (CStr(Tieto) <> "")
            End If
            If Lisaa Then
                KuviinTbl.AddNew
                KuviinTbl("PHakem").Value = zTitleTbl("TyyppipHak").Value
                KuviinTbl("PohjaKuva").Value = zTitleTbl("Tyyppipiir").Value
                KuviinTbl("HNDL").Value = TEKSTITTbl("HANDLES").Value
                KuviinTbl("UHakem").Value = zTitleTbl("Hakemisto").Value
                KuviinTbl("Talletuskuva").Value = zTitleTbl("FileName").Value
                KuviinTbl("EntType").Value = TEKSTITTbl("EntType").Value
                KuviinTbl("Tieto").Value = Tieto
                KuviinTbl.Update
            End If
            TEKSTITTbl.MoveNext
        Loop
        TEKSTITTbl.Close
        Set TEKSTITTbl = Nothing
        zTitleTbl.MoveNext
    Loop

    Tyotila.CommitTrans
    TransAktiivinen = False

Siivous:
    On Error Resume Next
    SysCmd acSysCmdClearStatus
    If Not TEKSTITTbl Is Nothing Then TEKSTITTbl.Close
    If Not zTitleTbl Is Nothing Then zTitleTbl.Close
    If Not KuviinTbl Is Nothing Then KuviinTbl.Close
    Set KuviinTbl = Nothing
    Set zTitleTbl = Nothing
    Set TEKSTITTbl = Nothing
    Set Tyotila = Nothing
    Set DB = Nothing
    If virhenumero <> 0 Then
        MsgBox "Virhe tietojen poiminnassa, muutokset peruttiin (" & virhenumero & "): " & virhekuvaus, vbCritical, "Poimi tiedot"
    End If
    Exit Sub

ErrorHandler:
    virhenumero = Err.Number
    virhekuvaus = Err.Description
    On Error Resume Next
    If TransAktiivinen Then Tyotila.Rollback
    Resume Siivous
End Sub

'------------------------------------------------------------------------------
' Tapahtuma: GenKuvat_Click
' Tarkoitus: Avaa pohjakuvat vain luku -tilassa, päivittää niiden tekstit
'            TiedotKuviin-taulun mukaan ja tallentaa valmiit kuvat
'
' Huomiot:
'   - Käsitellään vain oma piirustus kerrallaan; käyttäjän muita avoimia
'     piirustuksia ei suljeta
'   - Puuttuvat tekstikahvat ohitetaan ja niiden määrä ilmoitetaan lopuksi
'------------------------------------------------------------------------------
Private Sub GenKuvat_Click()
On Error GoTo ErrorHandler
    Dim Kuvat As DAO.Recordset  ' Generoitavat kuvat
    Dim Tiedot As DAO.Recordset  ' Kuvan päivitettävät tekstit
    Dim oEntity As Object  ' Päivitettävä tekstiobjekti
    Dim DocTila As Boolean  ' AutoCADin SingleDocumentMode ennen ajoa
    Dim TilaTallennettu As Boolean  ' Onko SingleDocumentMode tallennettu
    Dim DB As DAO.Database  ' Nykyinen tietokanta
    Dim i As Long, Maara As Long  ' Edistymislaskurit
    Dim Puuttuvat As Long  ' Päivittämättä jääneet tekstit
    Dim Avaa As String  ' Avattava pohjakuva
    Dim Talleta As String  ' Tallennettava kuva
    Dim virhenumero As Long  ' Virheen numero käsittelijässä
    Dim virhekuvaus As String  ' Virheen kuvaus käsittelijässä

    If Not YhdistaAutoCAD Then Exit Sub

    Set DB = CurrentDb
    Set Kuvat = DB.OpenRecordset("SELECT TalletusKuva, [phakem] & [pohjakuva] AS Avaa, [UHakem] & [talletuskuva] AS Talleta FROM qryKuviin GROUP BY TalletusKuva, [phakem] & [pohjakuva], [UHakem] & [talletuskuva]", dbOpenSnapshot)
    If Kuvat.EOF Then
        MsgBox "Generoitavia kuvia ei löytynyt.", vbInformation, "Generoi kuvat"
        GoTo Siivous
    End If
    Kuvat.MoveLast
    Kuvat.MoveFirst
    Maara = Kuvat.RecordCount

    DocTila = oAcad.Preferences.System.SingleDocumentMode
    TilaTallennettu = True
    oAcad.Preferences.System.SingleDocumentMode = False

    Do While Not Kuvat.EOF
        i = i + 1
        Avaa = CStr(Kuvat.Fields("Avaa").Value)
        Talleta = CStr(Kuvat.Fields("Talleta").Value)
        SysCmd acSysCmdSetStatus, "Käsittelee kuvaa (" & i & "/" & Maara & "): " & Kuvat("Talletuskuva").Value & "..."
        If OnkoAvoinna(Avaa) Then Err.Raise vbObjectError + 1001, , "Pohjakuva on jo auki AutoCADissa, sulje se ensin: " & Avaa
        If Not HakemistoOlemassa(Talleta) Then Err.Raise vbObjectError + 1002, , "Tallennushakemistoa ei löydy: " & Talleta

        Set Tiedot = DB.OpenRecordset("SELECT HNDL, Tieto FROM qryKuviin WHERE TalletusKuva='" & SqlTeksti(Kuvat("Talletuskuva").Value) & "'", dbOpenSnapshot)
        Set oDoc = oAcad.Documents.Open(Avaa, True)

        Do While Not Tiedot.EOF
            ' Yksittäisen tekstin virhe (esim. puuttuva kahva) ei keskeytä koko ajoa
            On Error Resume Next
            Set oEntity = Nothing
            Set oEntity = oDoc.HandleToObject(CStr(Tiedot("HNDL").Value))
            If Err.Number = 0 And Not oEntity Is Nothing Then
                oEntity.TextString = CStr(Nz(Tiedot("Tieto").Value, ""))
                If Err.Number <> 0 Then Puuttuvat = Puuttuvat + 1
            Else
                Puuttuvat = Puuttuvat + 1
            End If
            Err.Clear
            On Error GoTo ErrorHandler
            Tiedot.MoveNext
        Loop
        Tiedot.Close
        Set Tiedot = Nothing

        oDoc.SaveAs Talleta
        oDoc.Close False
        Set oDoc = Nothing
        Kuvat.MoveNext
    Loop

    If Puuttuvat > 0 Then
        MsgBox "Valmis. " & Puuttuvat & " tekstiä jäi päivittämättä (kahvaa ei löytynyt kuvasta).", vbExclamation, "Generoi kuvat"
    Else
        MsgBox "Valmis", vbOKOnly, "Generoi kuvat"
    End If

Siivous:
    On Error Resume Next
    SysCmd acSysCmdClearStatus
    If TilaTallennettu Then oAcad.Preferences.System.SingleDocumentMode = DocTila
    If Not Tiedot Is Nothing Then Tiedot.Close
    If Not Kuvat Is Nothing Then Kuvat.Close
    Set Tiedot = Nothing
    Set Kuvat = Nothing
    Set oEntity = Nothing
    Set oDoc = Nothing
    Set oAcad = Nothing
    Set DB = Nothing
    If virhenumero <> 0 Then
        MsgBox "Virhe kuvien generoinnissa (" & virhenumero & "): " & virhekuvaus, vbCritical, "Generoi kuvat"
    End If
    Exit Sub

ErrorHandler:
    virhenumero = Err.Number
    virhekuvaus = Err.Description
    On Error Resume Next
    ' Suljetaan keskeneräinen piirustus tallentamatta
    If Not oDoc Is Nothing Then oDoc.Close False
    Resume Siivous
End Sub

'------------------------------------------------------------------------------
' Alirutiini: LisaaTeksti
' Tarkoitus: Lisää tekstin revisiosivun ruudukkoon (rivi ja sarake laskettuna
'            Paikka-numerosta)
' Parametrit:
'   Paikka - Rivin järjestysnumero (1-alkuinen)
'   Sarake - 0 = sivun nimi, 1... = revisiokirjaimen sarake
'   teksti - Kirjoitettava teksti
'------------------------------------------------------------------------------
Private Sub LisaaTeksti(ByVal Paikka As Integer, Sarake As Integer, teksti As String)
  Dim Rivi As Integer  ' Rivi sarakkeen sisällä
  Dim Sar As Integer  ' Ruudukon sarakeryhmä
  Dim iPoint(2) As Double  ' Lisäyspiste
  Dim XSiir As Double  ' Sarakkeen siirtymä X-suunnassa
  Const AloitusX = 26.5
  Const AloitusY = 277
  Const RiviKork = 5.5
  Const SarLev = 5
  Const EkaLev = 19.5
  Const ColLev = 64.6
  Const Riveja = 45
  Paikka = Paikka - 1
  Rivi = Paikka Mod Riveja
  Sar = Paikka \ Riveja
  If Sarake = 0 Then
   XSiir = 0
  Else
   XSiir = EkaLev + (Sarake - 1) * SarLev
  End If
  iPoint(0) = AloitusX + ColLev * Sar + XSiir
  iPoint(1) = AloitusY - (Rivi) * RiviKork
  oDoc.ModelSpace.AddText teksti, iPoint, 2
End Sub

' Muuntaa revisiokirjaimen sarakenumeroksi (0 = tuntematon)
Private Function AnnaSarake(Kirjain As String) As Integer
Select Case UCase$(Kirjain)
  Case "0"
    AnnaSarake = 1
  Case "A"
    AnnaSarake = 2
  Case "B"
    AnnaSarake = 3
  Case "C"
    AnnaSarake = 4
  Case "D"
    AnnaSarake = 5
  Case "E"
    AnnaSarake = 6
  Case "F"
    AnnaSarake = 7
  Case "G"
    AnnaSarake = 8
  Case "H"
    AnnaSarake = 9
End Select
End Function
