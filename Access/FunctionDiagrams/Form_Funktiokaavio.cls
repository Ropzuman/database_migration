Option Compare Database
Option Explicit
'==========================================================================
' LOMAKE    : Form_Funktiokaavio
' SOVELLUS : FunctionDiagrams - Funktiokaavion päälomake
' KUVAUS    : Päälomake funktio- ja moottoripiirustusjärjestelmälle.
'            Hallinnoi revisioita (ADDNEWREV), ohjausrivilistauksia
'            (Command50, Komento47), reseptejä (Komento67, Komento80,
'            Command83) sekä AutoCAD-indeksisivujen luontia (Command98,
'            Lukituskaavioiden sivutiedot päivitetään
'            Intpage-tauluun (Komento46). Näkymänvaihtot MOTORS/LOOPS
'            (Komento57, Komento58) säädeltävät RecordSource-kytkinnällä.
' HUOM     : Option Explicit; Late Binding kaikissa olioissa ja MsgBox-
'            vakioissa (ei riippuvuuksia ulkoisista kirjastoista).
' PÄIVITETTY: 2026-09-23 (MsgBox-vakiot korvattu numeerisilla arvoilla)
'==========================================================================

' AutoCAD-yhteys ja aktiivinen piirustus (late binding, 64-bit yhteens.)
Private oAcad   As Object
Private oDoc    As Object

' Polkumuuttujat indeksisivujen tallennukseen ja pohjakuviin
Private BasePath As String
Private SavePath As String

Private Sub ADDNEWREV_Click()
On Error GoTo ErrorHandler
    Debug.Print "ADDNEWREV_Click: Toiminto alkaa"
    Dim DB As Object
    Dim Taulu As Object
    
    Dim urev As String
    Dim ALUE As String
    Dim LooppiNo As String
    Dim LooppiSymb As String
    Dim paiva As String
    Dim kk As String
    Dim vuosi As String
    Dim revpaiva As String
    Dim revteki As String

    Set DB = CurrentDb
    Debug.Print "  Tietokantayhteys avattu"

    revteki = Nz(Me!RevBy.Value, "")
    urev = Nz(Me!NEWREV.Value, "")
    ALUE = Nz(Me!AreaCode.Value, "")
    LooppiNo = Nz(Me!LoopNo.Value, "")
    LooppiSymb = Nz(Me!LoopSymb.Value, "")
    
    paiva = Day(Date)
    kk = Month(Date)
    vuosi = Year(Date)
    revpaiva = paiva & "." & kk & "." & vuosi

    Debug.Print "  Uusi revisio: " & urev & ", Alue: " & ALUE & ", Silmukkanumero: " & LooppiNo

    If Me!FReV = urev Then
        Debug.Print "  VAROITUS: Revisio on jo olemassa: " & urev
        ' 0 = vbOKOnly
        MsgBox "Uusi revisio on jo olemassa!", 0
        GoTo CleanExit
    End If

    ' 2 = dbOpenDynaset (dbOpenTable ei toimi linkitetylle taululle)
    Set Taulu = DB.OpenRecordset("rev", 2)
    Debug.Print "  Lisätään uusi revisio rev-tauluun"
    
    Taulu.AddNew
    Taulu.Fields("Area") = ALUE
    Taulu.Fields("no") = LooppiNo
    Taulu.Fields("symb") = LooppiSymb
    Taulu.Fields("rev") = urev
    Taulu.Fields("revdate") = revpaiva
    Taulu.Fields("revby") = revteki
    Taulu.Fields("doctype") = "F"
    Taulu.Fields("revdesc") = Nz(Me!RevDesc.Value, "")
    Taulu.Update
    Debug.Print "  Revisiotietue päivitetty onnistuneesti"

    Me!FReV = urev
    Debug.Print "ADDNEWREV_Click: VALMIS onnistuneesti"

CleanExit:
    If Not Taulu Is Nothing Then
        Taulu.Close
        Set Taulu = Nothing
    End If
    Set DB = Nothing
    Exit Sub

ErrorHandler:
    Debug.Print "*** VIRHE aliohjelmassa ADDNEWREV_Click: " & Err.Number & " - " & Err.Description
    Debug.Print "    Alue: " & ALUE & ", Silmukkanumero: " & LooppiNo & ", Uusi revisio: " & urev
    Debug.Print "    Lähde: " & Err.Source & ", Rivi: " & Erl
    ' 16 = vbCritical
    MsgBox "Virhe revision lisäyksessä: " & Err.Description, 16
    Resume CleanExit
End Sub
Private Sub Command50_Click()
On Error GoTo ErrorHandler
    Debug.Print "Command50_Click: Ohjaustietojen päivitys alkaa"

    Dim sAreaCode As String, sLoopNo As String, sTagID As String, sSivu As String, sUser As String
    Dim strWhere As String, strSQL As String
    Dim DB As Object

    ' 1. Luetaan arvot lomakkeen kentistä
    sAreaCode = Replace(CStr(Nz(Me!Muokkaus0.Value, "")), "'", "''")
    sLoopNo = Replace(CStr(Nz(Me!Muokkaus4.Value, "")), "'", "''")
    sTagID = Replace(CStr(Nz(Me!Muokkaus3.Value, "")), "'", "''")
    
    ' Luetaan KAYTTAJA()-funktion koko paluuarvo (täsmää Funtitle_Func HAVING-ehdon kanssa!)
    sUser = Replace(CStr(KAYTTAJA()), "'", "''")

    ' Luetaan sivunumero
    On Error Resume Next
    sSivu = Replace(CStr(Nz(Me!Sivu.Value, "")), "'", "''")
    If sSivu = "" Then sSivu = Replace(CStr(Nz(Me!Page.Value, "")), "'", "''")
    If sSivu = "" Then sSivu = Replace(CStr(Nz(Me!Muokkaus5.Value, "")), "'", "''")
    On Error GoTo ErrorHandler

    Debug.Print "  Aluetunnus: " & sAreaCode & ", Silmukkanumero: " & sLoopNo & ", Tagitunnus: " & sTagID & ", Sivu: " & sSivu & ", Käyttäjä: " & sUser

    ' 2. Suoritetaan vanhojen valintojen poisto kyselyllä (Control_Poistatulostus)
    DoCmd.SetWarnings False
    DoCmd.OpenQuery "Control_Poistatulostus"
    DoCmd.SetWarnings True

    ' 3. Muodostetaan Hakuehto.
    ' Jos sivunumero on lomakkeella, rajataan päivitys vain auki olevaan sivuun!
    strWhere = "AreaCode = '" & sAreaCode & "' AND LoopNo = '" & sLoopNo & "' AND TagID = '" & sTagID & "'"
    If sSivu <> "" Then
        strWhere = strWhere & " AND Sivu = '" & sSivu & "'"
    End If

    strSQL = "UPDATE Control SET Sel = '" & sUser & "' WHERE " & strWhere

    ' 4. Suoritetaan UPDATE (Late Binding, 128 = dbFailOnError)
    Set DB = CurrentDb
    DB.Execute strSQL, 128

    Debug.Print "  Ohjaustauluun päivitettiin käyttäjä '" & sUser & "', päivitetyt tietueet: " & DB.RecordsAffected

    ' Virkistetään välimuisti ja lomake
    DBEngine.Idle 8
    Set DB = Nothing
    Me.Refresh

    Debug.Print "Command50_Click: VALMIS onnistuneesti"
    Exit Sub

ErrorHandler:
    DoCmd.SetWarnings True
    Debug.Print "*** VIRHE aliohjelmassa Command50_Click: " & Err.Number & " - " & Err.Description
    Debug.Print "    Lähde: " & Err.Source & ", Rivi: " & Erl
    Set DB = Nothing
    MsgBox "Virhe ohjausrivin päivityksessä: " & Err.Description, 16
End Sub

Private Sub Command83_Click()
On Error GoTo ErrorHandler
    Debug.Print "Command83_Click: Reseptin päivitys alkaa"
    Dim Vastaus As Integer
    Dim strPrompt As String
    Dim Transaktio As Boolean

    Debug.Print "  Reseptitunnus: " & RECIPES.Form.ID
    
    strPrompt = "Tämä päivittää reseptin (ID=" & RECIPES.Form.ID & ")niille piireille, joissa on alla mainittu resepti." & vbCrLf & "Poistetaanko vanhat blokit piireistä, joille on määritelty resepti?"
    ' 3 = vbYesNoCancel
    Vastaus = MsgBox(strPrompt, 3, "PÄIVITÄ VALITTU RESEPTI KAIKKIIN PIIREIHIN")
    
    ' Poisto ja lisäys tehdään samassa transaktiossa, ettei virhe jätä piirejä ilman blokkeja
    If Vastaus <> 2 Then
        DBEngine.BeginTrans
        Transaktio = True
    End If
    ' 6 = vbYes
    If Vastaus = 6 Then
        Debug.Print "  Poistetaan vanhat blokit Control-taulusta"
        CurrentDb.Execute "DELETE * FROM Control where Recipeid = " & RECIPES.Form.ID, 128
    End If
    
    ' 2 = vbCancel
    If Vastaus <> 2 Then
        Debug.Print "  Lisätään reseptin tiedot Control-tauluun"
        CurrentDb.Execute "INSERT INTO control ( RecipeID, AreaCode, LoopNo, FBlock, LoopSymb, Suffix, BlockPlace, Description2, [Note], Sivu, Sivuja ) SELECT * FROM RECIPES_ADD_ID where RecipeID=" & RECIPES.Form.ID, 128
        DBEngine.CommitTrans
        Transaktio = False
        Debug.Print "  Reseptin päivitys valmis"
        ' 0 = vbOKOnly
        MsgBox "Piirit päivitetty!", 0, "PÄIVITÄ VALITTU RESEPTI KAIKKIIN PIIREIHIN"
        Me.Refresh
    End If
    Debug.Print "Command83_Click: VALMIS onnistuneesti"
    Exit Sub

ErrorHandler:
    Debug.Print "*** VIRHE aliohjelmassa Command83_Click: " & Err.Number & " - " & Err.Description
    Debug.Print "    Lähde: " & Err.Source & ", Rivi: " & Erl
    ' 16 = vbCritical
    If Transaktio Then DBEngine.Rollback: Transaktio = False
    MsgBox "Virhe reseptin päivityksessä: " & Err.Description, 16
End Sub

Private Sub Command97_Click()
On Error GoTo ErrorHandler
    Debug.Print "Command97_Click: Linkkikyselyiden ajo alkaa"
    DoCmd.SetWarnings False
    DoCmd.OpenQuery "Delete tolink from FunckblockAttr"
    DoCmd.OpenQuery "Delete todesc from FunckblockAttr"
    DoCmd.OpenQuery "Delete fromdesc from FunckblockAttr"
    DoCmd.OpenQuery "LinkTo1ADD"
    DoCmd.OpenQuery "LinkTo2ADD"
    DoCmd.OpenQuery "LinkTo3ADD"
    DoCmd.OpenQuery "LinkTo4ADD"
    DoCmd.OpenQuery "LinkTo5ADD"
    DoCmd.OpenQuery "LinkTo6ADD"
    DoCmd.OpenQuery "TODESCADD"
    DoCmd.OpenQuery "FROMDESCADD"
    DoCmd.SetWarnings True
    Debug.Print "Command97_Click: VALMIS onnistuneesti - All link queries executed"
    Exit Sub

ErrorHandler:
    Debug.Print "*** VIRHE aliohjelmassa Command97_Click: " & Err.Number & " - " & Err.Description
    Debug.Print "    Lähde: " & Err.Source & ", Rivi: " & Erl
    DoCmd.SetWarnings True
    ' 16 = vbCritical
    MsgBox "Virhe linkkikyselyjen ajossa: " & Err.Description, 16
End Sub

Private Sub Command98_Click()
On Error GoTo ErrorHandler
    Debug.Print "Command98_Click: Indeksisivujen luonti alkaa (AutoCAD)"
    Dim Taul As Object
    Dim edPOS As String
    Dim Sivut As String
    Dim IndexBase As String
    Dim edPiir As String
    Dim viimPiir As String
    Dim Piste(2) As Double
    Dim Piste2(2) As Double
    Dim ind As Integer
    Dim aX As Double
    Dim aY As Double
    Dim PosSivu As Double
    Dim Seuraava As Double
    Dim MinY As Double
    Dim MaxX As Double
    Dim RiviVali As Double
    Dim TextSize As Double

    IndexBase = Nz(DLookup("[VALUE]", "[SETTINGS]", "SETTING = 'INDEXBASE'"), "")
    If IndexBase = "" Then
        ' 16 = vbCritical
        MsgBox "INDEXBASE puuttuu SETTINGS-taulusta! Indeksisivuja ei voi luoda.", 16
        Exit Sub
    End If
    Debug.Print "  Indeksipohja: " & IndexBase
    
    ' Pohjakuvien hakemisto luetaan SETTINGS-taulusta, oletuksena aiempi kiinteä polku
    BasePath = Nz(DLookup("[VALUE]", "[SETTINGS]", "SETTING = 'INDEXPATH'"), "Y:\base\int\")
    If Right$(BasePath, 1) <> "\" Then BasePath = BasePath & "\"
    SavePath = Nz(DLookup("[VALUE]", "[SETTINGS]", "SETTING = 'SAVEPATH'"), "")
    Debug.Print "  Pohjapolku: " & BasePath & "  tallennuspolku: " & SavePath
    
    If SavePath = "" Then
        ' 16 = vbCritical
        MsgBox "SAVEPATH puuttuu SETTINGS-taulusta! Indeksisivuja ei voi tallentaa.", 16
        Exit Sub
    End If
    
    Set oAcad = GetObject(, "AutoCAD.Application")
    Debug.Print "  AutoCAD-yhteys avattu"
    
    Set Taul = CurrentDb.OpenRecordset("Funtitle_Func")
    Set oDoc = oAcad.Documents.Open(BasePath & IndexBase & ".dwg")
    Debug.Print "  Pohjapiirustus avattu: " & BasePath & IndexBase & ".dwg"

    aX = 30
    aY = 285
    PosSivu = 40
    Seuraava = 90
    MinY = 35
    MaxX = 333
    RiviVali = 5
    TextSize = 2.5

    Piste(0) = aX
    Piste(1) = aY
    Piste2(0) = aX + PosSivu
    Piste2(1) = aY

    Debug.Print "  Käsitellään tietueita..."
    Do While Not Taul.EOF
        viimPiir = Nz(Taul.Fields("Piir"), "")
        If edPOS <> Taul.Fields("position") Then
            If Sivut <> "" Then
                If Piste(1) = aY Then
                    oDoc.ModelSpace.AddText "POSITION", Piste, TextSize + 0.5
                    oDoc.ModelSpace.AddText "DESCRIPTION", Piste2, TextSize + 0.5
                    Piste(1) = Piste(1) - RiviVali - 1
                    Piste2(1) = Piste(1)
                End If
                
                oDoc.ModelSpace.AddText edPOS, Piste, TextSize
                oDoc.ModelSpace.AddText Sivut, Piste2, TextSize
                Piste(1) = Piste(1) - RiviVali
                Piste2(1) = Piste(1)
                
                If Piste(1) < MinY Then
                    Piste(1) = aY
                    Piste2(1) = aY
                    Piste(0) = Piste(0) + Seuraava
                    Piste2(0) = Piste2(0) + Seuraava
                End If
                
                If Piste(0) > MaxX Then
                    If Taul.Fields("Piir") <> edPiir Then
                        ind = 0
                    End If
                    ind = ind + 1
                    oDoc.SaveAs SavePath & Taul.Fields("Piir") & "_IDX_" & ind
                    edPiir = Taul.Fields("Piir")
                    oDoc.Close
                    Set oDoc = oAcad.Documents.Open(BasePath & IndexBase & ".dwg")
                    Piste(0) = aX
                    Piste2(0) = aX + PosSivu
                End If
            End If
            
            Sivut = Taul.Fields("Description1")
            edPOS = Taul.Fields("Position")
        Else
            Sivut = Sivut & "/" & Taul.Fields("PAGE")
        End If
        Taul.MoveNext
    Loop
    
    oDoc.ModelSpace.AddText edPOS, Piste, TextSize
    oDoc.ModelSpace.AddText Sivut, Piste2, TextSize
    ' Viimeisen sivun nimeen käytetään viimeisen rivin piiriä
    If viimPiir <> edPiir Then ind = 0
    ind = ind + 1
    Debug.Print "  Tallennetaan indeksisivu: " & SavePath & viimPiir & "_IDX_" & ind
    oDoc.SaveAs SavePath & viimPiir & "_IDX_" & ind
    oDoc.Close
    
    Debug.Print "Command98_Click: VALMIS onnistuneesti - indeksisivut luotu"
    MsgBox "Valmis"

CleanExit:
    If Not Taul Is Nothing Then
        Taul.Close
        Set Taul = Nothing
    End If
    Set oDoc = Nothing
    Set oAcad = Nothing
    Exit Sub

ErrorHandler:
    Debug.Print "*** VIRHE aliohjelmassa Command98_Click: " & Err.Number & " - " & Err.Description
    ' 16 = vbCritical
    MsgBox "Virhe indeksisivujen luonnissa: " & Err.Description, 16
    Resume CleanExit
End Sub

Private Sub Komento46_Click()
On Error GoTo ErrorHandler
    Debug.Print "Komento46_Click: Lukituskaavion sivutietojen päivitys alkaa"
    Dim DB As DAO.Database
    Dim Taul As DAO.Recordset
    Dim Intpage As DAO.Recordset
    Dim EdLoop As String, PIIR As String, PIIRI As String
    Dim NRO As String, Tyyppi As String
    Dim underscorePos As Integer
    Dim lSrc As Long, lOut As Long
    Dim Transaktio As Boolean

    Set DB = CurrentDb

    ' 1) Read the source FIRST, as a static snapshot, before touching Intpage
    Set Taul = DB.OpenRecordset( _
        "SELECT DWG, TEXT_4, PAGE FROM POSANDPAGES " & _
        "GROUP BY DWG, TEXT_4, PAGE HAVING (TEXT_4<>'') ORDER BY TEXT_4, PAGE;", _
        dbOpenSnapshot)
    If Not Taul.EOF Then Taul.MoveLast: lSrc = Taul.RecordCount: Taul.MoveFirst
    Debug.Print "  Lähderivejä: " & lSrc

    If lSrc = 0 Then
        MsgBox "POSANDPAGES palautti 0 riviä - Intpagea ei muutettu.", vbExclamation
        GoTo CleanExit
    End If

    ' 2) Only now clear the target (tyhjennys ja täyttö samassa transaktiossa)
    DBEngine.BeginTrans
    Transaktio = True
    DB.Execute "DELETE * FROM Intpage", dbFailOnError
    Set Intpage = DB.OpenRecordset("Intpage", dbOpenDynaset)

    Do While Not Taul.EOF
        PIIRI = Nz(Taul.Fields("TEXT_4").Value, "")

        If EdLoop <> PIIRI Then
            If EdLoop <> "" Then Intpage.Update: lOut = lOut + 1
            Intpage.AddNew
            EdLoop = PIIRI
        End If

        PIIR = Nz(Taul.Fields("DWG").Value, "")
        underscorePos = InStr(10, PIIR, "_")
        If underscorePos > 0 Then PIIR = Left$(PIIR, underscorePos - 1)

        If Len(PIIRI) >= 5 Then Tyyppi = Mid$(PIIRI, 5, 1) Else Tyyppi = ""
        If Tyyppi = "-" Then
            NRO = Right$(PIIRI, 3)
        Else
            NRO = PIIRI
        End If

        If Nz(Intpage.Fields("Ensimmäinenkohteestapage").Value, "") <> "" Then
            Intpage.Fields("Ensimmäinenkohteestapage").Value = _
                Intpage.Fields("Ensimmäinenkohteestapage").Value & "/" & Nz(Taul.Fields("PAGE").Value, "")
        Else
            Intpage.Fields("Ensimmäinenkohteestapage").Value = Nz(Taul.Fields("PAGE").Value, "")
        End If

        Intpage.Fields("PIIRNRO").Value = PIIR
        Intpage.Fields("looppisymb").Value = PIIRI
        Intpage.Fields("looppinro").Value = NRO
        Taul.MoveNext
    Loop

    If EdLoop <> "" Then Intpage.Update: lOut = lOut + 1
    Intpage.Close
    DBEngine.CommitTrans
    Transaktio = False

    Debug.Print "  Intpage-tauluun kirjoitettuja rivejä: " & lOut & _
                ", rows now in table: " & DCount("*", "Intpage")
    DBEngine.Idle dbRefreshCache
    MsgBox "Lukituskaavioiden sivutiedot päivitetty Intpage taulukkoon (" & lOut & " riviä).", _
           vbInformation, "Lukituskaavion sivut"

CleanExit:
    On Error Resume Next
    If Not Intpage Is Nothing Then Intpage.Close
    If Not Taul Is Nothing Then Taul.Close
    Set Intpage = Nothing: Set Taul = Nothing: Set DB = Nothing
    Exit Sub

ErrorHandler:
    Debug.Print "*** VIRHE aliohjelmassa Komento46_Click: " & Err.Number & " - " & Err.Description
    If Transaktio Then DBEngine.Rollback: Transaktio = False
    MsgBox "Virhe lukituskaaviosivujen päivityksessä: " & Err.Description, vbCritical
    Resume CleanExit
End Sub
Private Sub Komento47_Click()
On Error GoTo ErrorHandler
    Debug.Print "Komento47_Click: Ohjauslistauksen päivitys alkaa"
    DoCmd.SetWarnings False
    DoCmd.OpenQuery "Control_Poistatulostus"
    DoCmd.OpenQuery "Control_Lisaa_tulostus"
    DoCmd.SetWarnings True
    Debug.Print "Komento47_Click: VALMIS onnistuneesti"
    Exit Sub

ErrorHandler:
    Debug.Print "*** VIRHE aliohjelmassa Komento47_Click: " & Err.Number & " - " & Err.Description
    DoCmd.SetWarnings True
    ' 16 = vbCritical
    MsgBox "Virhe tulostuslistauksen päivityksessä: " & Err.Description, 16
End Sub

Private Sub hae_Click()
On Error GoTo Err_hae_Click
    Screen.PreviousControl.SetFocus
    DoCmd.FindNext
Exit_hae_Click:
    Exit Sub
Err_hae_Click:
    MsgBox Err.Description
    Resume Exit_hae_Click
End Sub

Private Sub Komento57_Click()
On Error GoTo ErrorHandler
    Debug.Print "Komento57_Click: Vaihdetaan MOTORS-näkymään"
    LTITLE.Caption = "MOTORS"
    Me.RecordSource = "FuncDiag_Motors"
    Me.Requery
    Debug.Print "Komento57_Click: VALMIS onnistuneesti"
    Exit Sub

ErrorHandler:
    Debug.Print "*** VIRHE aliohjelmassa Komento57_Click: " & Err.Number & " - " & Err.Description
    ' 16 = vbCritical
    MsgBox "Virhe moottorinäkymän avauksessa: " & Err.Description, 16
End Sub

Private Sub Komento58_Click()
On Error GoTo ErrorHandler
    Debug.Print "Komento58_Click: Vaihdetaan LOOPS-näkymään"
    LTITLE.Caption = "LOOPS"
    Me.RecordSource = "FuncDiag_Loops"
    Me.Requery
    Debug.Print "Komento58_Click: VALMIS onnistuneesti"
    Exit Sub

ErrorHandler:
    Debug.Print "*** VIRHE aliohjelmassa Komento58_Click: " & Err.Number & " - " & Err.Description
    ' 16 = vbCritical
    MsgBox "Virhe piirinäkymän avauksessa: " & Err.Description, 16
End Sub

Private Sub Komento67_Click()
On Error GoTo ErrorHandler
    Debug.Print "Komento67_Click: Reseptin luonti blokeista alkaa"
    Dim DB As Object
    Dim Taul As Object
    Dim RID As Long
    Dim sSQL As String
    Dim sProcess As String
    Dim sCode As String
    Dim sAreaCode2 As String
    Dim sLoopNo2 As String
    Dim Transaktio As Boolean

    If IsNull(RProcess.Value) Or IsNull(RCode.Value) Then
        Debug.Print "  VIRHE: Process- tai Code-arvo puuttuu"
        ' 16 = vbCritical
        MsgBox "Syötä ensin Process ja Code reseptille", 16, "Tee resepti"
        Exit Sub
    End If

    Set DB = CurrentDb
    sProcess = Replace(CStr(RProcess.Value), "'", "''")
    sCode = Replace(CStr(RCode.Value), "'", "''")

    Set Taul = DB.OpenRecordset("SELECT * FROM RECIPES WHERE Process='" & sProcess & "' and Code = '" & sCode & "'")
    If Not Taul.EOF Then
        Debug.Print "  VIRHE: Resepti on jo olemassa"
        ' 16 = vbCritical
        MsgBox "Anna reseptille Process ja Code, jota ei ole vielä käytössä.", 16, "Tee resepti"
        GoTo CleanExit
    End If

    Debug.Print "  Luodaan uusi resepti RECIPES-tauluun"
    ' Resepti ja siihen liittyvät rivit tallennetaan samassa transaktiossa
    DBEngine.BeginTrans
    Transaktio = True
    Taul.AddNew
    Taul.Fields("PROCESS") = RProcess.Value
    Taul.Fields("CODE") = RCode.Value
    Taul.Fields("INFO") = RInfo.Value
    Taul.Update
    ' AutoNumber luetaan vasta päivityksen jälkeen
    Taul.Bookmark = Taul.LastModified
    RID = Taul.Fields("ID")

    Debug.Print "  New Reseptitunnus: " & RID

    sAreaCode2 = Replace(CStr(Me.AreaCode), "'", "''")
    sLoopNo2 = Replace(CStr(Me.LoopNo), "'", "''")

    sSQL = "INSERT INTO RECIPE_TMP ( RecipeID, FBlock, LoopSymb, Suffix, Description2, BlockPlace, [Note], Sivu, Sivuja ) SELECT " & RID & " AS ID, Control.FBlock, Control.LoopSymb, Control.Suffix, Control.Description2, Control.BlockPlace, Control.Note, Control.Sivu, Control.Sivuja FROM Control WHERE (((Control.AreaCode)='" & sAreaCode2 & "') AND ((Control.LoopNo)='" & sLoopNo2 & "'));"
    DB.Execute sSQL, 128

    sSQL = "UPDATE Control SET Control.RecipeID = " & RID & " WHERE (((Control.AreaCode)='" & sAreaCode2 & "') AND ((Control.LoopNo)='" & sLoopNo2 & "'));"
    DB.Execute sSQL, 128

    DBEngine.CommitTrans
    Transaktio = False
    Me.RecipeID = RID
    Debug.Print "  Resepti liitetty ohjauspiireihin"
    Me.Refresh
    Debug.Print "Komento67_Click: VALMIS onnistuneesti - resepti luotu"

CleanExit:
    If Not Taul Is Nothing Then
        Taul.Close
        Set Taul = Nothing
    End If
    Set DB = Nothing
    Exit Sub

ErrorHandler:
    Debug.Print "*** VIRHE aliohjelmassa Komento67_Click: " & Err.Number & " - " & Err.Description
    If Transaktio Then DBEngine.Rollback: Transaktio = False
    ' 16 = vbCritical
    MsgBox "Virhe reseptin luonnissa: " & Err.Description, 16
    Resume CleanExit
End Sub

Private Sub Komento80_Click()
On Error GoTo ErrorHandler
    Debug.Print "Komento80_Click: Reseptin päivitys kaikkiin piireihin alkaa"
    Dim Vastaus As Integer
    Dim Transaktio As Boolean

    ' 3 = vbYesNoCancel
    Vastaus = MsgBox("Poistetaanko vanhat blokit piireistä, joille on määritelty resepti?", 3, "PÄIVITÄ RESEPTIT KAIKKIIN PIIREIHIN")

    ' Poisto ja lisäys tehdään samassa transaktiossa, ettei virhe jätä piirejä ilman blokkeja
    If Vastaus <> 2 Then
        DBEngine.BeginTrans
        Transaktio = True
    End If
    ' 6 = vbYes
    If Vastaus = 6 Then
        Debug.Print "  Poistetaan vanhat blokit niistä Control-taulun tietueista, joilla on RecipeID"
        CurrentDb.Execute "DELETE * FROM Control where Recipeid is not null", 128
    End If
    ' 2 = vbCancel
    If Vastaus <> 2 Then
        Debug.Print "  Suoritetaan RECIPES_ADD-kysely"
        CurrentDb.Execute "RECIPES_ADD", 128
        DBEngine.CommitTrans
        Transaktio = False
        Debug.Print "  Kaikki piirit päivitetty resepteillä"
        ' 0 = vbOKOnly
        MsgBox "Piirit päivitetty!", 0, "PÄIVITÄ RESEPTIT KAIKKIIN PIIREIHIN"
        Me.Refresh
    End If
    Debug.Print "Komento80_Click: VALMIS onnistuneesti"
    Exit Sub

ErrorHandler:
    Debug.Print "*** VIRHE aliohjelmassa Komento80_Click: " & Err.Number & " - " & Err.Description
    ' 16 = vbCritical
    If Transaktio Then DBEngine.Rollback: Transaktio = False
    MsgBox "Virhe reseptien päivityksessä: " & Err.Description, 16
End Sub

Private Sub CommandJANI_Click()
On Error GoTo ErrorHandler
    Debug.Print "CommandJANI_Click: Indeksisivujen luonti Intpage-taulusta alkaa"
    Dim Taul As Object
    Dim edPOS As String
    Dim Sivut As String
    Dim IndexBase As String
    Dim edPiir As String
    Dim Piste(2) As Double
    Dim Piste2(2) As Double
    Dim ind As Integer
    Dim aX As Double
    Dim aY As Double
    Dim PosSivu As Double
    Dim Seuraava As Double
    Dim MinY As Double
    Dim MaxX As Double
    Dim RiviVali As Double
    Dim TextSize As Double

    IndexBase = Nz(DLookup("[VALUE]", "[SETTINGS]", "SETTING = 'INDEXBASE'"), "")
    If IndexBase = "" Then
        ' 16 = vbCritical
        MsgBox "INDEXBASE puuttuu SETTINGS-taulusta! Indeksisivuja ei voi luoda.", 16
        Exit Sub
    End If
    Debug.Print "  Indeksipohja: " & IndexBase

    If oAcad Is Nothing Then
        Set oAcad = GetObject(, "AutoCAD.Application")
        Debug.Print "  AutoCAD-yhteys avattu (CommandJANI)"
    End If

    ' Pohjakuvien hakemisto luetaan SETTINGS-taulusta, oletuksena aiempi kiinteä polku
    BasePath = Nz(DLookup("[VALUE]", "[SETTINGS]", "SETTING = 'INDEXPATH'"), "Y:\base\int\")
    If Right$(BasePath, 1) <> "\" Then BasePath = BasePath & "\"
    SavePath = Nz(DLookup("[VALUE]", "[SETTINGS]", "SETTING = 'SAVEPATH'"), "")
    Debug.Print "  Pohjapolku: " & BasePath & "  tallennuspolku: " & SavePath
    
    If SavePath = "" Then
        ' 16 = vbCritical
        MsgBox "SAVEPATH puuttuu SETTINGS-taulusta! Indeksisivuja ei voi tallentaa.", 16
        Exit Sub
    End If

    Set Taul = CurrentDb.OpenRecordset("Intpage")
    Debug.Print "  Intpage-tietuejoukko avattu"
    Set oDoc = oAcad.Documents.Open(BasePath & IndexBase & ".dwg")
    Debug.Print "  Pohjapiirustus avattu: " & BasePath & IndexBase & ".dwg"

    aX = 30
    aY = 275
    PosSivu = 40
    Seuraava = 90
    MinY = 40
    MaxX = 333
    RiviVali = 5
    TextSize = 2.5

    Piste(0) = aX
    Piste(1) = aY
    Piste2(0) = aX + PosSivu
    Piste2(1) = aY

    Debug.Print "  Käsitellään Intpage-tietueita..."
    Do While Not Taul.EOF
        If edPOS <> Taul.Fields("TEXT_4") Then
            If Sivut <> "" Then
                If Piste(1) = aY Then
                    oDoc.ModelSpace.AddText "POSITION", Piste, TextSize + 0.5
                    oDoc.ModelSpace.AddText "PAGE", Piste2, TextSize + 0.5
                    Piste(1) = Piste(1) - RiviVali - 1
                    Piste2(1) = Piste(1)
                End If
                
                oDoc.ModelSpace.AddText edPOS, Piste, TextSize
                oDoc.ModelSpace.AddText Sivut, Piste2, TextSize
                Piste(1) = Piste(1) - RiviVali
                Piste2(1) = Piste(1)
                
                If Piste(1) < MinY Then
                    Piste(1) = aY
                    Piste2(1) = aY
                    Piste(0) = Piste(0) + Seuraava
                    Piste2(0) = Piste2(0) + Seuraava
                End If
                
                If Piste(0) > MaxX Then
                    If Taul.Fields("Piir") <> edPiir Then
                        ind = 0
                    End If
                    ind = ind + 1
                    oDoc.SaveAs SavePath & Taul.Fields("Piir") & "_IDX_" & ind
                    edPiir = Taul.Fields("Piir")
                    oDoc.Close
                    Set oDoc = oAcad.Documents.Open(BasePath & IndexBase & ".dwg")
                    Piste(0) = aX
                    Piste2(0) = aX + PosSivu
                End If
            End If
            
            Sivut = Taul.Fields("PAGE")
            edPOS = Taul.Fields("TEXT_4")
        Else
            Sivut = Sivut & "/" & Taul.Fields("PAGE")
        End If
        Taul.MoveNext
    Loop
    
    oDoc.ModelSpace.AddText edPOS, Piste, TextSize
    oDoc.ModelSpace.AddText Sivut, Piste2, TextSize
    ind = ind + 1
    Debug.Print "  Tallennetaan indeksisivu: " & SavePath & edPiir & "_IDX_" & ind
    oDoc.SaveAs SavePath & edPiir & "_IDX_" & ind
    oDoc.Close
    
    Debug.Print "CommandJANI_Click: VALMIS onnistuneesti - indeksisivut luotu"
    MsgBox "Valmis"

CleanExit:
    If Not Taul Is Nothing Then
        Taul.Close
        Set Taul = Nothing
    End If
    Set oDoc = Nothing
    Set oAcad = Nothing
    Exit Sub

ErrorHandler:
    Debug.Print "*** VIRHE aliohjelmassa CommandJANI_Click: " & Err.Number & " - " & Err.Description
    ' 16 = vbCritical
    MsgBox "Virhe indeksisivujen luonnissa: " & Err.Description, 16
    Resume CleanExit
End Sub

