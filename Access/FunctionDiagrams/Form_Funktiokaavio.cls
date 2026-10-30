Option Compare Database
Option Explicit
'==========================================================================
' LOMAKE    : Form_Funktiokaavio
' SOVELLUS : FunctionDiagrams – Funktiokaavion päälomake
' KUVAUS    : Päälomake funktio- ja moottoripiirustusjärjestelmälle.
'            Hallinnoi revisioita (ADDNEWREV), ohjausrivilistauksia
'            (Command50, Komento47), reseptejä (Komento67, Komento80,
'            Command83) sekä AutoCAD-indeksisivujen luontia (Command98,
'            CommandJANI). Lukituskaavioiden sivutiedot päivitetään
'            Intpage-tauluun (Komento46). Näkymänvaihtot MOTORS/LOOPS
'            (Komento57, Komento58) säädeltävät RecordSource-kytkinnällä.
' DIPLE     : Option Explicit; Late Binding kaikissa olioissa ja MsgBox-
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
    Debug.Print "ADDNEWREV_Click: Starting operation"
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
    Debug.Print "  Database connection established"

    revteki = Nz(Me!RevBy.Value, "")
    urev = Nz(Me!NEWREV.Value, "")
    ALUE = Nz(Me!AreaCode.Value, "")
    LooppiNo = Nz(Me!LoopNo.Value, "")
    LooppiSymb = Nz(Me!LoopSymb.Value, "")
    
    paiva = Day(Date)
    kk = Month(Date)
    vuosi = Year(Date)
    revpaiva = paiva & "." & kk & "." & vuosi

    Debug.Print "  NewRev: " & urev & ", Area: " & ALUE & ", LoopNo: " & LooppiNo

    If Me!FReV = urev Then
        Debug.Print "  WARNING: Revision already exists: " & urev
        ' 0 = vbOKOnly
        MsgBox "Uusi revisio on jo olemassa!", 0
        GoTo CleanExit
    End If

    ' 1 = dbOpenTable
    Set Taulu = DB.OpenRecordset("rev", 1)
    Debug.Print "  Adding new revision to 'rev' table"
    
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
    Debug.Print "  Revision record updated successfully"

    Me!FReV = urev
    Debug.Print "ADDNEWREV_Click: COMPLETED successfully"

CleanExit:
    If Not Taulu Is Nothing Then
        Taulu.Close
        Set Taulu = Nothing
    End If
    Set DB = Nothing
    Exit Sub

ErrorHandler:
    Debug.Print "*** ERROR in ADDNEWREV_Click: " & Err.Number & " - " & Err.Description
    Debug.Print "    Area: " & ALUE & ", LoopNo: " & LooppiNo & ", NewRev: " & urev
    Debug.Print "    Source: " & Err.Source & ", Line: " & Erl
    ' 16 = vbCritical
    MsgBox "Virhe revision lisäyksessä: " & Err.Description, 16
    Resume CleanExit
End Sub
Private Sub Command50_Click()
On Error GoTo ErrorHandler
    Debug.Print "Command50_Click: Starting control update operation"

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

    Debug.Print "  AreaCode: " & sAreaCode & ", LoopNo: " & sLoopNo & ", TagID: " & sTagID & ", Sivu: " & sSivu & ", User: " & sUser

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

    Debug.Print "  Control table updated with '" & sUser & "', Affected records: " & DB.RecordsAffected

    ' Virkistetään välimuisti ja lomake
    DBEngine.Idle 8
    Set DB = Nothing
    Me.Refresh

    Debug.Print "Command50_Click: COMPLETED successfully"
    Exit Sub

ErrorHandler:
    DoCmd.SetWarnings True
    Debug.Print "*** ERROR in Command50_Click: " & Err.Number & " - " & Err.Description
    Debug.Print "    Source: " & Err.Source & ", Line: " & Erl
    Set DB = Nothing
    MsgBox "Virhe ohjausrivin päivityksessä: " & Err.Description, 16
End Sub

Private Sub Command83_Click()
On Error GoTo ErrorHandler
    Debug.Print "Command83_Click: Starting recipe update operation"
    Dim Vastaus As Integer
    Dim strPrompt As String

    Debug.Print "  RecipeID: " & RECIPES.Form.ID
    
    strPrompt = "Tämä päivittää reseptin (ID=" & RECIPES.Form.ID & ")niille piireille, joissa on alla mainittu resepti." & vbCrLf & "Poistetaanko vanhat blokit piireistä, joille on määritelty resepti?"
    ' 3 = vbYesNoCancel
    Vastaus = MsgBox(strPrompt, 3, "PÄIVITÄ VALITTU RESEPTI KAIKKIIN PIIREIHIN")
    
    ' 6 = vbYes
    If Vastaus = 6 Then
        Debug.Print "  Deleting old blocks from Control table"
        CurrentDb.Execute "DELETE * FROM Control where Recipeid = " & RECIPES.Form.ID, 128
    End If
    
    ' 2 = vbCancel
    If Vastaus <> 2 Then
        Debug.Print "  Inserting recipe data into control table"
        CurrentDb.Execute "INSERT INTO control ( RecipeID, AreaCode, LoopNo, FBlock, LoopSymb, Suffix, BlockPlace, Description2, [Note], Sivu, Sivuja ) SELECT * FROM RECIPES_ADD_ID where RecipeID=" & RECIPES.Form.ID, 128
        Debug.Print "  Recipe update completed"
        ' 0 = vbOKOnly
        MsgBox "Piirit päivitetty!", 0, "PÄIVITÄ VALITTU RESEPTI KAIKKIIN PIIREIHIN"
        Me.Refresh
    End If
    Debug.Print "Command83_Click: COMPLETED successfully"
    Exit Sub

ErrorHandler:
    Debug.Print "*** ERROR in Command83_Click: " & Err.Number & " - " & Err.Description
    Debug.Print "    Source: " & Err.Source & ", Line: " & Erl
    ' 16 = vbCritical
    MsgBox "Virhe reseptin päivityksessä: " & Err.Description, 16
End Sub

Private Sub Command97_Click()
On Error GoTo ErrorHandler
    Debug.Print "Command97_Click: Starting link cleanup operation"
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
    Debug.Print "Command97_Click: COMPLETED successfully - All link queries executed"
    Exit Sub

ErrorHandler:
    Debug.Print "*** ERROR in Command97_Click: " & Err.Number & " - " & Err.Description
    Debug.Print "    Source: " & Err.Source & ", Line: " & Erl
    DoCmd.SetWarnings True
    ' 16 = vbCritical
    MsgBox "Virhe linkkikyselyjen ajossa: " & Err.Description, 16
End Sub

Private Sub Command98_Click()
On Error GoTo ErrorHandler
    Debug.Print "Command98_Click: Starting index page creation (AutoCAD)"
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

    IndexBase = DLookup("[VALUE]", "[SETTINGS]", "SETTING = 'INDEXBASE'")
    Debug.Print "  IndexBase: " & IndexBase
    
    BasePath = "Y:\base\int\"
    SavePath = Nz(DLookup("[VALUE]", "[SETTINGS]", "SETTING = 'SAVEPATH'"), "")
    Debug.Print "  BasePath: " & BasePath & "  SavePath: " & SavePath
    
    If SavePath = "" Then
        ' 16 = vbCritical
        MsgBox "SAVEPATH puuttuu SETTINGS-taulusta! Indeksisivuja ei voi tallentaa.", 16
        Exit Sub
    End If
    
    Set oAcad = GetObject(, "AutoCAD.Application")
    Debug.Print "  AutoCAD connection established"
    
    Set Taul = CurrentDb.OpenRecordset("Funtitle_Func")
    Set oDoc = oAcad.Documents.Open(BasePath & IndexBase & ".dwg")
    Debug.Print "  Base document opened: " & BasePath & IndexBase & ".dwg"

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

    Debug.Print "  Processing records..."
    Do While Not Taul.EOF
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
    ind = ind + 1
    Debug.Print "  Saving index page: " & SavePath & edPiir & "_IDX_" & ind
    oDoc.SaveAs SavePath & edPiir & "_IDX_" & ind
    
    Debug.Print "Command98_Click: COMPLETED successfully - Index pages created"
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
    Debug.Print "*** ERROR in Command98_Click: " & Err.Number & " - " & Err.Description
    ' 16 = vbCritical
    MsgBox "Virhe indeksisivujen luonnissa: " & Err.Description, 16
    Resume CleanExit
End Sub

Private Sub Komento46_Click()
On Error GoTo ErrorHandler
    Debug.Print "Komento46_Click: Starting interlocking page update"
    Dim DB As DAO.Database
    Dim Taul As DAO.Recordset
    Dim Intpage As DAO.Recordset
    Dim EdLoop As String, PIIR As String, PIIRI As String
    Dim NRO As String, Tyyppi As String
    Dim underscorePos As Integer
    Dim lSrc As Long, lOut As Long

    Set DB = CurrentDb

    ' 1) Read the source FIRST, as a static snapshot, before touching Intpage
    Set Taul = DB.OpenRecordset( _
        "SELECT DWG, TEXT_4, PAGE FROM POSANDPAGES " & _
        "GROUP BY DWG, TEXT_4, PAGE HAVING (TEXT_4<>'') ORDER BY TEXT_4, PAGE;", _
        dbOpenSnapshot)
    If Not Taul.EOF Then Taul.MoveLast: lSrc = Taul.RecordCount: Taul.MoveFirst
    Debug.Print "  Source rows: " & lSrc

    If lSrc = 0 Then
        MsgBox "POSANDPAGES palautti 0 riviä - Intpagea ei muutettu.", vbExclamation
        GoTo CleanExit
    End If

    ' 2) Only now clear the target
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

    Debug.Print "  Intpage rows written: " & lOut & _
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
    Debug.Print "*** ERROR in Komento46_Click: " & Err.Number & " - " & Err.Description
    MsgBox "Virhe lukituskaaviosivujen päivityksessä: " & Err.Description, vbCritical
    Resume CleanExit
End Sub
Private Sub Komento47_Click()
On Error GoTo ErrorHandler
    Debug.Print "Komento47_Click: Starting control print update"
    DoCmd.SetWarnings False
    DoCmd.OpenQuery "Control_Poistatulostus"
    DoCmd.OpenQuery "Control_Lisaa_tulostus"
    DoCmd.SetWarnings True
    Debug.Print "Komento47_Click: COMPLETED successfully"
    Exit Sub

ErrorHandler:
    Debug.Print "*** ERROR in Komento47_Click: " & Err.Number & " - " & Err.Description
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
    Debug.Print "Komento57_Click: Switching to MOTORS view"
    LTITLE.Caption = "MOTORS"
    Me.RecordSource = "FuncDiag_Motors"
    Me.Requery
    Debug.Print "Komento57_Click: COMPLETED successfully"
    Exit Sub

ErrorHandler:
    Debug.Print "*** ERROR in Komento57_Click: " & Err.Number & " - " & Err.Description
    ' 16 = vbCritical
    MsgBox "Virhe moottorinäkymän avauksessa: " & Err.Description, 16
End Sub

Private Sub Komento58_Click()
On Error GoTo ErrorHandler
    Debug.Print "Komento58_Click: Switching to LOOPS view"
    LTITLE.Caption = "LOOPS"
    Me.RecordSource = "FuncDiag_Loops"
    Me.Requery
    Debug.Print "Komento58_Click: COMPLETED successfully"
    Exit Sub

ErrorHandler:
    Debug.Print "*** ERROR in Komento58_Click: " & Err.Number & " - " & Err.Description
    ' 16 = vbCritical
    MsgBox "Virhe piirinäkymän avauksessa: " & Err.Description, 16
End Sub

Private Sub Komento67_Click()
On Error GoTo ErrorHandler
    Debug.Print "Komento67_Click: Starting recipe creation from blocks"
    Dim DB As Object
    Dim Taul As Object
    Dim RID As Long
    Dim sSQL As String
    Dim sProcess As String
    Dim sCode As String
    Dim sAreaCode2 As String
    Dim sLoopNo2 As String

    If IsNull(RProcess.Value) Or IsNull(RCode.Value) Then
        Debug.Print "  ERROR: Process or Code missing"
        ' 16 = vbCritical
        MsgBox "Syötä ensin Process ja Code reseptille", 16, "Tee resepti"
        Exit Sub
    End If

    Set DB = CurrentDb
    sProcess = Replace(CStr(RProcess.Value), "'", "''")
    sCode = Replace(CStr(RCode.Value), "'", "''")

    Set Taul = DB.OpenRecordset("SELECT * FROM RECIPES WHERE Process='" & sProcess & "' and Code = '" & sCode & "'")
    If Not Taul.EOF Then
        Debug.Print "  ERROR: Recipe already exists"
        ' 16 = vbCritical
        MsgBox "Anna reseptille Process ja Code, jota ei ole vielä käytössä.", 16, "Tee resepti"
        GoTo CleanExit
    End If

    Debug.Print "  Creating new recipe in RECIPES table"
    Taul.AddNew
    Taul.Fields("PROCESS") = RProcess.Value
    Taul.Fields("CODE") = RCode.Value
    Taul.Fields("INFO") = RInfo.Value
    RID = Taul.Fields("ID")
    Taul.Update

    Debug.Print "  New RecipeID: " & RID
    Me.RecipeID = RID

    sAreaCode2 = Replace(CStr(Me.AreaCode), "'", "''")
    sLoopNo2 = Replace(CStr(Me.LoopNo), "'", "''")

    sSQL = "INSERT INTO RECIPE_TMP ( RecipeID, FBlock, LoopSymb, Suffix, Description2, BlockPlace, [Note], Sivu, Sivuja ) SELECT " & RID & " AS ID, Control.FBlock, Control.LoopSymb, Control.Suffix, Control.Description2, Control.BlockPlace, Control.Note, Control.Sivu, Control.Sivuja FROM Control WHERE (((Control.AreaCode)='" & sAreaCode2 & "') AND ((Control.LoopNo)='" & sLoopNo2 & "'));"
    DB.Execute sSQL, 128

    sSQL = "UPDATE Control SET Control.RecipeID = " & RID & " WHERE (((Control.AreaCode)='" & sAreaCode2 & "') AND ((Control.LoopNo)='" & sLoopNo2 & "'));"
    DB.Execute sSQL, 128

    Debug.Print "  Recipe assigned to control loops"
    Me.Refresh
    Debug.Print "Komento67_Click: COMPLETED successfully - Recipe created"

CleanExit:
    If Not Taul Is Nothing Then
        Taul.Close
        Set Taul = Nothing
    End If
    Set DB = Nothing
    Exit Sub

ErrorHandler:
    Debug.Print "*** ERROR in Komento67_Click: " & Err.Number & " - " & Err.Description
    ' 16 = vbCritical
    MsgBox "Virhe reseptin luonnissa: " & Err.Description, 16
    Resume CleanExit
End Sub

Private Sub Komento80_Click()
On Error GoTo ErrorHandler
    Debug.Print "Komento80_Click: Starting recipe update to all circuits"
    Dim Vastaus As Integer

    ' 3 = vbYesNoCancel
    Vastaus = MsgBox("Poistetaanko vanhat blokit piireistä, joille on määritelty resepti?", 3, "PÄIVITÄ RESEPTIT KAIKKIIN PIIREIHIN")

    ' 6 = vbYes
    If Vastaus = 6 Then
        Debug.Print "  Deleting old blocks from Control where RecipeID exists"
        CurrentDb.Execute "DELETE * FROM Control where Recipeid is not null", 128
    End If
    ' 2 = vbCancel
    If Vastaus <> 2 Then
        Debug.Print "  Executing RECIPES_ADD query"
        CurrentDb.Execute "RECIPES_ADD", 128
        Debug.Print "  All circuits updated with recipes"
        ' 0 = vbOKOnly
        MsgBox "Piirit päivitetty!", 0, "PÄIVITÄ RESEPTIT KAIKKIIN PIIREIHIN"
        Me.Refresh
    End If
    Debug.Print "Komento80_Click: COMPLETED successfully"
    Exit Sub

ErrorHandler:
    Debug.Print "*** ERROR in Komento80_Click: " & Err.Number & " - " & Err.Description
    ' 16 = vbCritical
    MsgBox "Virhe reseptien päivityksessä: " & Err.Description, 16
End Sub

Private Sub CommandJANI_Click()
On Error GoTo ErrorHandler
    Debug.Print "CommandJANI_Click: Starting index page generation (Intpage)"
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

    IndexBase = DLookup("[VALUE]", "[SETTINGS]", "SETTING = 'INDEXBASE'")
    Debug.Print "  IndexBase: " & IndexBase

    If oAcad Is Nothing Then
        Set oAcad = GetObject(, "AutoCAD.Application")
        Debug.Print "  AutoCAD connection established (CommandJANI)"
    End If

    BasePath = "Y:\base\int\"
    SavePath = Nz(DLookup("[VALUE]", "[SETTINGS]", "SETTING = 'SAVEPATH'"), "")
    Debug.Print "  BasePath: " & BasePath & "  SavePath: " & SavePath
    
    If SavePath = "" Then
        ' 16 = vbCritical
        MsgBox "SAVEPATH puuttuu SETTINGS-taulusta! Indeksisivuja ei voi tallentaa.", 16
        Exit Sub
    End If

    Set Taul = CurrentDb.OpenRecordset("Intpage")
    Debug.Print "  Intpage recordset opened"
    Set oDoc = oAcad.Documents.Open(BasePath & IndexBase & ".dwg")
    Debug.Print "  Base document opened: " & BasePath & IndexBase & ".dwg"

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

    Debug.Print "  Processing Intpage records..."
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
    Debug.Print "  Saving index page: " & SavePath & edPiir & "_IDX_" & ind
    oDoc.SaveAs SavePath & edPiir & "_IDX_" & ind
    oDoc.Close
    
    Debug.Print "CommandJANI_Click: COMPLETED successfully - Index pages created"
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
    Debug.Print "*** ERROR in CommandJANI_Click: " & Err.Number & " - " & Err.Description
    ' 16 = vbCritical
    MsgBox "Virhe indeksisivujen luonnissa: " & Err.Description, 16
    Resume CleanExit
End Sub

