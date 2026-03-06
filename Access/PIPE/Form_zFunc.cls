Option Compare Database
Option Explicit

'================================================================================
' Lomake: zFunc
' Tarkoitus: Utility functions for zDetails table management and batch operations
' Päivitetty: 2025-11-12 - VBA7/64-bit-tuki lisätty
'
' Kuvaus:
'   Kokoelma apufunktioita zDetails-hallintatoiminnoille:
'   - zDetails-taulun ylläpito (lisäys, päivitys, tyhjennys)
'   - Eräajokyselyt instrumenttipiireille
'   - Tyhjien merkkijonojen muunto Nulliksi kaikissa tauluissa
'   - Blokkisijoitusten numerointi
'   - Seuraavan tietueen etsintä
'
' Riippuvuudet:
'   - zDetails-taulu (instrumenttitiedot)
'   - Function_zdetails*-kyselyt
'   - InstrumentIndex-taulu
'   - Useat Tag*-kyselyt
'   - KaikkiTaulukot-yhdistelmäruutu (täytetty taulunimillä)
'
' Kontrollit:
'   - KaikkiTaulukot: yhdistelmäruutu taulun valintaan
'   - Useita komentopainikkeita eri toiminnoille
'================================================================================

Private Sub Command1_Click()
On Error GoTo Err_Command1_Click


    Screen.PreviousControl.SetFocus
    DoCmd.FindNext

Exit_Command1_Click:
    Exit Sub

Err_Command1_Click:
    MsgBox Err.Description
    Resume Exit_Command1_Click
    
End Sub


Private Sub Command10_Click()
DoCmd.OpenQuery "Function_zdetails15_loopsprintselect", acNormal, acEdit
End Sub

'--------------------------------------------------------------------------------
' Tapahtuma: Command11_Click
' Tarkoitus: Convert empty strings to Null across selected table
'
' Toiminta:
'   1. Avataan valittu taulu KaikkiTaulukot-yhdistelmäruudusta
'   2. Käydään läpi kaikki tietueet
'   3. Jos kentän arvo on tyhjä merkkijono, asetetaan se Nulliksi
'   4. Päivitetään tietue tapahtumatuen kanssa
'
' Huom:
'   - Parantaa tietokannan normalisointia
'   - Estää tyhjien merkkijonojen tallentumisen
'   - Voi olla aikaa vievä suurilla tauluilla
'   - Optimoitu: Edit-kutsu vain kerran tietuetta kohden, tapahtumatuki
'--------------------------------------------------------------------------------
Private Sub Command11_Click()
On Error GoTo ErrorHandler
    Dim DB As DAO.Database      ' Tietokantaviittaus
    Dim WS As DAO.Workspace     ' Työtila tapahtumatukea varten
    Dim Taulu As DAO.Recordset  ' Valittu taulu
    Dim i As Integer            ' Kenttälaskuri
    Dim RecordsProcessed As Long ' Käsiteltyjen tietueiden laskuri
    Dim RecordsUpdated As Long  ' Päivitettyjen tietueiden laskuri
    Dim NeedsUpdate As Boolean  ' Lippu: tarvitaanko tietueen päivitys
    
    Set DB = CurrentDb
    Set WS = DBEngine.Workspaces(0)  ' Haetaan oletustyötila
    Set Taulu = DB.OpenRecordset(KaikkiTaulukot.Value, dbOpenDynaset)
    
    If Taulu.EOF Then
        MsgBox "Taulu on tyhjä", vbInformation, "Tyhjennetään"
        Taulu.Close
        Set Taulu = Nothing
        DB.Close
        Set DB = Nothing
        Exit Sub
    End If
    
    ' Aloitetaan tapahtuma suorituskyvyn parantamiseksi
    WS.BeginTrans
    
    Do While Not Taulu.EOF
      NeedsUpdate = False
      RecordsProcessed = RecordsProcessed + 1
      
      ' Tarkistetaan kaikki kentät ennen muokkausta
      For i = 0 To Taulu.Fields.Count - 1
        If Taulu.Fields(i) = "" Then
          If Not NeedsUpdate Then
            Taulu.Edit
            NeedsUpdate = True
          End If
          Taulu.Fields(i) = Null
        End If
      Next i
      
      ' Päivitetään vain jos muutoksia tehtiin
      If NeedsUpdate Then
        Taulu.Update
        RecordsUpdated = RecordsUpdated + 1
      End If
      
      Taulu.MoveNext
    Loop
    
    ' Vahvistetaan tapahtuma
    WS.CommitTrans
    
    MsgBox "Valmis: " & RecordsProcessed & " riviä käyty, " & RecordsUpdated & " päivitetty", vbInformation, "Tyhjennetään"
    
    ' Siivotaan objektit
    Taulu.Close
    Set Taulu = Nothing
    Set WS = Nothing
    DB.Close
    Set DB = Nothing
    Exit Sub

ErrorHandler:
    ' Peruutetaan tapahtuma virhetilanteessa
    On Error Resume Next
    If Not WS Is Nothing Then WS.Rollback
    If Not Taulu Is Nothing Then Taulu.Close
    Set Taulu = Nothing
    Set WS = Nothing
    If Not DB Is Nothing Then DB.Close
    Set DB = Nothing
    On Error GoTo 0
    MsgBox "Virhe tyhjiä merkkijonoja muunnettaessa: " & Err.Description, vbCritical, "Virhe"
End Sub

'--------------------------------------------------------------------------------
' Tapahtuma: Command3_Click
' Tarkoitus: Assign block placement numbers to zDetails records
'
' Toiminta:
'   1. Avataan lajiteltu tietuejoukko (Luupin, sitten AllDevTagin mukaan)
'   2. Jokaiselle luuppiryhmälle annetaan peräkkäiset sijoitusnumerot (14, 15, 16...)
'   3. Nollataan laskuri jokaisen uuden luupin alussa
'
' Huom:
'   - Blokkisijoitus alkaa 14:stä
'   - Ryhmittely luuppi-tunnisteen mukaan ([AreaCode] & [Loopno] & [pageno])
'   - Käytetään AutoCAD-blokkien sijoitteluun
'--------------------------------------------------------------------------------
Private Sub Command3_Click()
On Error GoTo Err_Command3_Click
    Dim tbl As DAO.Recordset  ' zDetails-taulun tietuejoukko
    Dim q As String            ' Nykyinen luuppitunniste
    Dim q1 As String           ' Edellinen luuppi (käyttämätön)
    Dim qq As String           ' Seuraava luuppitunniste
    Dim qq1 As String          ' Seuraavan luupin vertailu (käyttämätön)
    Dim i As Integer           ' Blokkisijoituslaskuri
    
    ' Avataan lajiteltu tietuejoukko
    Set tbl = CurrentDb.OpenRecordset("SELECT [AreaCode] & [Loopno] & [pageno] AS Loop, zDetails.Blockplace, zDetails.AllDevTag FROM zDetails ORDER BY [AreaCode] & [Loopno] & [pageno], zDetails.AllDevTag;")
  
    tbl.MoveFirst
    Do
      i = 14  ' Blokkisijoitus alkaa 14:stä
      Do
        q = tbl!Loop
        tbl.Edit
        tbl!blockplace = i
        tbl.Update
        i = i + 1
        tbl.MoveNext
        If tbl.EOF Then Exit Do
        qq = tbl!Loop
        If q <> qq Then Exit Do  ' Uusi luuppiryhmä, nollataan laskuri
      Loop
    Loop Until tbl.EOF
  
    MsgBox "Homma hoidettu", vbInformation, "Sijoitukset"
    
    ' Siivotaan objektit
    tbl.Close
    Set tbl = Nothing
    Exit Sub

Err_Command3_Click:
    MsgBox "Virhe blokkisijoitusten määrittämisessä: " & Err.Description, vbCritical, "Virhe"
    On Error Resume Next
    If Not tbl Is Nothing Then tbl.Close
    Set tbl = Nothing
    On Error GoTo 0
End Sub

Private Sub Command4_Click()
DoCmd.OpenQuery "Function_zDetails03_lisättävät", acNormal, acEdit
DoCmd.OpenQuery "Function_zDetails04_FileID", acNormal, acEdit
DoCmd.OpenQuery "Function_zDetails09_update", acNormal, acEdit
DoCmd.OpenQuery "Function_zDetails14_Fileid2", acNormal, acEdit
End Sub


'--------------------------------------------------------------------------------
' Tapahtuma: Command5_Click
' Tarkoitus: Poistaa zDetails-tietueet, joita ei enää löydy InstrumentIndex-taulukosta
'
' Toiminta:
'   1. Etsitään zDetails-tietueet, joilla ei ole vastaavaa Function_zdetails01-tietuetta
'   2. Poistetaan jokainen orpo tietue zDetails-taulusta
'   3. Käytetään tiimalasikursoria edistymisen osoittamiseen
'   4. Tapahtumatuki tietojen eheyden varmistamiseksi
'
' Huom:
'   - Siivoaa vanhentuneet instrumenttitiedot
'   - Käyttää LEFT JOIN -kyselyä löytämään vastaattomat tietueet
'   - Tapahtumatuki tietojen eheyden varmistamiseksi
'--------------------------------------------------------------------------------
Private Sub Command5_Click()
On Error GoTo ErrorHandler
    ' Siivotaan orpot zDetails-tietueet
    Dim DB As DAO.Database      ' Tietokantaviittaus
    Dim WS As DAO.Workspace     ' Työtila tapahtumatukea varten
    Dim tbl As DAO.Recordset    ' Kyselyn tulokset (orpot tietueet)
    Dim qdf As DAO.QueryDef     ' Parametrisoitu DELETE-kysely (SQL-injektiosuojaus)
    Dim sSQL As String          ' Pääkyselyn SQL
    Dim acode As String         ' Aluekoodi
    Dim LoopNo As String        ' Luuppinumero
    Dim LoopSymb As String      ' Luuppimerkki
    Dim Suffix As String        ' Tägisuffiksi
    Dim DeleteCount As Long     ' Poistettujen tietueiden lukumäärä
    DeleteCount = 0

    Set DB = CurrentDb
    Set WS = DBEngine.Workspaces(0)  ' Haetaan oletustyötila

    ' Rakennetaan kysely orpojen zDetails-tietueiden löytämiseksi
    sSQL = ""
    sSQL = sSQL & "SELECT zDetails.AreaCode, zDetails.LoopNo, zDetails.LoopSymb, zDetails.Suffix"
    sSQL = sSQL & " FROM zDetails LEFT JOIN Function_zdetails01 ON (zDetails.Suffix = Function_zdetails01.Suffix) AND (zDetails.LoopSymb = Function_zdetails01.LoopSymb) AND (zDetails.LoopNo = Function_zdetails01.LoopNo) AND (zDetails.AreaCode = Function_zdetails01.AreaCode)"
    sSQL = sSQL & " WHERE (((Function_zdetails01.AreaCode) Is Null) AND ((Function_zdetails01.LoopNo) Is Null) AND ((Function_zdetails01.LoopSymb) Is Null) AND ((Function_zdetails01.Suffix) Is Null));"

    Set tbl = DB.OpenRecordset(sSQL)

    If Not tbl.EOF Then
      DoCmd.Hourglass True
      WS.BeginTrans  ' Aloitetaan tapahtuma
      
      tbl.MoveFirst
      Do Until tbl.EOF
          acode = tbl!AreaCode
          LoopNo = tbl!LoopNo
          Suffix = tbl!Suffix
          LoopSymb = tbl!LoopSymb

          ' Poistetaan tietue turvallisella parametrisoidulla kyselyllä (SQL-injektiosuojaus)
          Set qdf = DB.CreateQueryDef("", "DELETE FROM zDetails WHERE AreaCode=[pArea] AND LoopNo=[pLoop] AND LoopSymb=[pSymb] AND Suffix=[pSuffix]")
          qdf.Parameters("pArea").Value = acode
          qdf.Parameters("pLoop").Value = LoopNo
          qdf.Parameters("pSymb").Value = LoopSymb
          qdf.Parameters("pSuffix").Value = Suffix
          qdf.Execute dbFailOnError
          Set qdf = Nothing
          DeleteCount = DeleteCount + 1
          tbl.MoveNext
      Loop
      
      WS.CommitTrans  ' Vahvistetaan tapahtuma
      DoCmd.Hourglass False
      
      tbl.Close
      Set tbl = Nothing
      MsgBox "Poistettu " & DeleteCount & " vanhentunutta tietuetta", vbInformation, "Siivous valmis"
    Else
      MsgBox "Ei vanhentuneita tietueita", vbInformation, "Siivous valmis"
    End If
    
    ' Siivotaan objektit
    Set WS = Nothing
    Set DB = Nothing
    Exit Sub

ErrorHandler:
    DoCmd.Hourglass False
    On Error Resume Next
    If Not WS Is Nothing Then WS.Rollback
    If Not tbl Is Nothing Then tbl.Close
    Set tbl = Nothing
    Set qdf = Nothing
    Set DB = Nothing
    On Error GoTo 0
    MsgBox "Virhe vanhentuneita tietueita poistettaessa: " & Err.Description, vbCritical, "Virhe"
End Sub

Private Sub Command6_Click()
DoCmd.OpenQuery "Function_zDetails06_clearprint", acNormal, acEdit
End Sub

Private Sub Command7_Click()
DoCmd.OpenQuery "Function_zDetails05_printselect", acNormal, acEdit
End Sub

Private Sub Command8_Click()
DoCmd.OpenQuery "Function_Zdetails08_Clearzdetailskuviin", acNormal, acEdit
End Sub

Private Sub Command9_Click()
DoCmd.OpenQuery "Tag_1-7_ZDetailsKuviin", acNormal, acEdit
DoCmd.OpenQuery "Tag1_CR0-1_FC0-1_F0-2_PWC0_CCM0-1_JB0120", acNormal, acEdit
DoCmd.OpenQuery "Tag1_DCS", acNormal, acEdit
DoCmd.OpenQuery "Tag1_F10-11", acNormal, acEdit
DoCmd.OpenQuery "Tag2_CCM0-1_PWC0", acNormal, acEdit
DoCmd.OpenQuery "Tag1_JB10-12", acNormal, acEdit
DoCmd.OpenQuery "TAG1_FC10-11", acNormal, acEdit
'DoCmd.OpenQuery "Tag_1-7_ZDetailsKuviin", acNormal, acEdit
End Sub

'--------------------------------------------------------------------------------
' Tapahtuma: Form_Load
' Tarkoitus: Täyttää KaikkiTaulukot-yhdistelmäruudun käyttäjätaulukoiden nimillä
'
' Toiminta:
'   1. Käydään läpi kaikki TableDef-objektit tietokannasta
'   2. Jätetään pois järjestelmätaulukot (USys*, MSys*, Dev*) ja vanhat taulukot
'   3. Rakennetaan puolipisteeroteltu luettelo yhdistelmäruutua varten
'   4. Asetetaan oletusvalinta ensimmäiseen taulukkoon
'
' Huom:
'   - Suodattaa pois: USys*, MSys*, Dev*, "old"-tekstin sisältävät taulut
'   - Täyttää KaikkiTaulukot-yhdistelmäruudun RowSource-ominaisuuden
'--------------------------------------------------------------------------------
Private Sub Form_Load()
On Error GoTo ErrorHandler
    Dim DB As DAO.Database  ' Tietokantaviittaus
    Dim i As Long           ' Silmukkalasku
    Dim Taulukko As String  ' Nykyinen taulun nimi
    Dim TJoukko As String   ' Puolipisteeroteltu taulukkoluettelo
    
    Set DB = CurrentDb
    For i = 0 To DB.TableDefs.Count - 1
      Taulukko = DB.TableDefs(i).Name
      ' Suodatetaan järjestelmätaulukot ja vanhat taulukot
      If Left$(Taulukko, 4) <> "USys" And Left$(Taulukko, 4) <> "MSys" And Left$(Taulukko, 3) <> "Dev" And InStr(Taulukko, "old") = 0 Then
        TJoukko = TJoukko & Taulukko & ";"
      End If
    Next i
    
    ' Asetetaan yhdistelmäruudun tietolähde ja oletusarvo
    KaikkiTaulukot.RowSource = TJoukko
    KaikkiTaulukot.Value = Left$(TJoukko, InStr(TJoukko, ";") - 1)
    
    Set DB = Nothing
    Exit Sub

ErrorHandler:
    MsgBox "Virhe lomakkeen latauksessa: " & Err.Description, vbCritical, "Virhe"
    On Error Resume Next
    Set DB = Nothing
    On Error GoTo 0
End Sub
