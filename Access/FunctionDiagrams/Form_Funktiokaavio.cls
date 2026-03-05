Option Compare Database
Option Explicit
'==========================================================================
' LOMAKE   : Form_Funktiokaavio
' SOVELLUS : FunctionDiagrams — Funktiokaavion päälomake
' KUVAUS   : Päälomake funktio- ja moottoripiirustusjärjestelmälle.
'            Hallinnoi revisioita (ADDNEWREV), ohjausrivilistauksia
'            (Command50, Komento47), reseptejä (Komento67, Komento80,
'            Command83) sekä AutoCAD-indeksisivujen luontia (Command98,
'            CommandJANI). Lukituskaavioiden sivutiedot päivitetään
'            Intpage-tauluun (Komento46). Näkymänvaihdot MOTORS/LOOPS
'            (Komento57, Komento58) sääteltävät RecordSource-kytkinnällä.
' DIPLE    : Option Explicit; moduulitason oAcad/oDoc/BasePath/SavePath;
'            kuollut api_GetUserName-julistus poistettu; kaikki
'            On Error GoTo -lohkot yhdenmukaistettu; Set = Nothing kaikkialle.
' PÄIVITETTY: 2026-03-03
'==========================================================================

' AutoCAD-yhteys ja aktiivinen piirustus (late binding, 64-bit yhteens.)
Private oAcad   As Object
Private oDoc    As Object

' Polkumuuttujat indeksisivujen tallennukseen ja pohjakuviin
' Nämä tulee asettaa ennen Command98_Click / CommandJANI_Click-kutsujen tekemistä
' (esim. DLookup SETTINGS-taulusta tai lomakken kenttästä)
Private BasePath As String
Private SavePath As String

Private Sub ADDNEWREV_Click()
On Error GoTo ErrorHandler
Debug.Print "ADDNEWREV_Click: Starting operation"
Dim DB As DAO.Database
Dim urev As String
Dim Taulu As DAO.Recordset
Dim ALUE As String
Dim LooppiSymb As String
Dim LooppiNo As String
Dim paiva As String
Dim kk As String
Dim vuosi As String
Dim revpaiva As String
Dim revteki As String
Set DB = CurrentDb
Debug.Print "  Database connection established"

'BuffSize = 256
'NBuffer = Space$(BuffSize)
'
'    If api_GetUserName(NBuffer, BuffSize) Then
'     revteki = Left$(NBuffer, InStr(NBuffer, Chr(0)) - 1)
'    Else
'      revteki = "Unknown"
'    End If

revteki = Me!RevBy
Debug.Print "  RevBy: " & revteki

urev = Me!NEWREV
ALUE = Me!AreaCode
LooppiNo = Me!LoopNo
Debug.Print "  NewRev: " & urev & ", Area: " & ALUE & ", LoopNo: " & LooppiNo
paiva = Day(Date)
kk = Month(Date)
'vuosi = Right(Year(Date), 2)
vuosi = Year(Date)
revpaiva = paiva & "." & kk & "." & vuosi
LooppiSymb = Me!LoopSymb

If Me!FReV = urev Then
    Debug.Print "  WARNING: Revision already exists: " & urev
    ' Paluuarvoa ei käytetä — kutsutaan suoraan MsgBox-proseduurina
    MsgBox "Uusi revisio on jo olemassa!", vbOKOnly
    Exit Sub
End If

Set Taulu = DB.OpenRecordset("rev", dbOpenTable)
Debug.Print "  Adding new revision to 'rev' table"
Taulu.AddNew
Taulu.Fields("Area") = ALUE
Taulu.Fields("no") = LooppiNo
Taulu.Fields("symb") = LooppiSymb
Taulu.Fields("rev") = urev
Taulu.Fields("revdate") = revpaiva
Taulu.Fields("revby") = revteki
Taulu.Fields("doctype") = "F"
Taulu.Fields("revdesc") = Me!RevDesc
Taulu.Update
Debug.Print "  Revision record updated successfully"

'If IsNull(Forms!funktiokaavio.FReV.Value) Then
Me!FReV = urev

Debug.Print "ADDNEWREV_Click: COMPLETED successfully"

' Vapautetaan DAO-resurssit muistivuodon ehkäisemiseksi
Set Taulu = Nothing
Set DB = Nothing
Exit Sub

ErrorHandler:
    Debug.Print "*** ERROR in ADDNEWREV_Click: " & Err.Number & " - " & Err.Description
    Debug.Print "    Area: " & ALUE & ", LoopNo: " & LooppiNo & ", NewRev: " & urev
    Debug.Print "    Source: " & Err.Source & ", Line: " & Erl
    Set Taulu = Nothing
    Set DB = Nothing
    MsgBox "Virhe revision lisäyksessä: " & Err.Description, vbCritical
End Sub

Private Sub Command50_Click()
On Error GoTo ErrorHandler
Debug.Print "Command50_Click: Starting control update operation"

Debug.Print "  AreaCode: " & Me!Muokkaus0.Value & ", LoopNo: " & Me!Muokkaus4.Value & ", TagID: " & Me!Muokkaus3.Value
DoCmd.SetWarnings False
DoCmd.OpenQuery ("Control_Poistatulostus")
'sql = "UPDATE Control SET Control.Sel = ""x"" WHERE Control.AreaCode = " & Me!Muokkaus0.Value & " And Control.LoopNo = " & Me!Muokkaus4.Value & ""

CurrentDb.Execute "UPDATE Control SET Sel = '" & KAYTTAJA() & "' WHERE AreaCode = """ & Me!Muokkaus0.Value & """ And LoopNo = """ & Me!Muokkaus4.Value & """ And TagID = """ & Me!Muokkaus3.Value & """"
Debug.Print "  Control table updated"

'CurrentDb.Execute "UPDATE Control SET Control.Sel = ""x"" WHERE Control.AreaCode = " & Me!Muokkaus0.Value & " And Control.LoopNo = " & Me!Muokkaus4.Value & ""


'Set qry = db.OpenRecordset(sql)

'DoCmd.OpenQuery ("Control_Lisaa_tulostus")
DoCmd.SetWarnings True
Debug.Print "Command50_Click: COMPLETED successfully"
Exit Sub

ErrorHandler:
    Debug.Print "*** ERROR in Command50_Click: " & Err.Number & " - " & Err.Description
    Debug.Print "    AreaCode: " & Me!Muokkaus0.Value & ", LoopNo: " & Me!Muokkaus4.Value
    Debug.Print "    Source: " & Err.Source & ", Line: " & Erl
    DoCmd.SetWarnings True
    MsgBox "Virhe ohjausrivin päivityksessä: " & Err.Description, vbCritical
End Sub


Private Sub Command83_Click()
'Tämä päivittää reseptit kaikkiin piireihin, joissa on käytetty tiettyä reseptiä.
On Error GoTo ErrorHandler
Debug.Print "Command83_Click: Starting recipe update operation"
Dim Vastaus As VbMsgBoxResult

Debug.Print "  RecipeID: " & RECIPES.Form.ID
Vastaus = MsgBox("Tämä päivittää reseptin (ID=" & RECIPES.Form.ID & ")niille piireille, joissa on alla mainittu resepti." & vbCrLf & "Poistetaanko vanhat blokit piireistä, joille on määritelty resepti?", vbYesNoCancel, "PÄIVITÄ VALITTU RESEPTI KAIKKIIN PIIREIHIN")
If Vastaus = vbYes Then 'Poistetaan ensin kaikki vanhat blokit niistä piireistä, joissa on resepti
  Debug.Print "  Deleting old blocks from Control table"
  CurrentDb.Execute "DELETE * FROM Control where Recipeid = " & RECIPES.Form.ID
End If
If Vastaus <> vbCancel Then
  Debug.Print "  Inserting recipe data into control table"
  CurrentDb.Execute "INSERT INTO control ( RecipeID, AreaCode, LoopNo, FBlock, LoopSymb, Suffix, BlockPlace, Description2, [Note], Sivu, Sivuja ) SELECT * FROM RECIPES_ADD_ID where RecipeID=" & RECIPES.Form.ID
  Debug.Print "  Recipe update completed"
  MsgBox "Piirit päivitetty!", vbOKOnly, "PÄIVITÄ VALITTU RESEPTI KAIKKIIN PIIREIHIN"
  Me.Refresh
End If
Debug.Print "Command83_Click: COMPLETED successfully"
Exit Sub

ErrorHandler:
    Debug.Print "*** ERROR in Command83_Click: " & Err.Number & " - " & Err.Description
    Debug.Print "    RecipeID: " & RECIPES.Form.ID
    Debug.Print "    Source: " & Err.Source & ", Line: " & Erl
    MsgBox "Virhe reseptin päivityksessä: " & Err.Description, vbCritical
End Sub

Private Sub Command97_Click()
On Error GoTo ErrorHandler
Debug.Print "Command97_Click: Starting link cleanup operation"
DoCmd.SetWarnings False
DoCmd.OpenQuery ("Delete tolink from FunckblockAttr")
DoCmd.OpenQuery ("Delete todesc from FunckblockAttr")
DoCmd.OpenQuery ("Delete fromdesc from FunckblockAttr")
DoCmd.OpenQuery ("LinkTo1ADD")
DoCmd.OpenQuery ("LinkTo2ADD")
DoCmd.OpenQuery ("LinkTo3ADD")
DoCmd.OpenQuery ("LinkTo4ADD")
DoCmd.OpenQuery ("LinkTo5ADD")
DoCmd.OpenQuery ("LinkTo6ADD")
DoCmd.OpenQuery ("TODESCADD")
DoCmd.OpenQuery ("FROMDESCADD")
DoCmd.SetWarnings True
Debug.Print "Command97_Click: COMPLETED successfully - All link queries executed"
Exit Sub

ErrorHandler:
    Debug.Print "*** ERROR in Command97_Click: " & Err.Number & " - " & Err.Description
    Debug.Print "    Source: " & Err.Source & ", Line: " & Erl
    DoCmd.SetWarnings True
    MsgBox "Virhe linkkikyselyjen ajossa: " & Err.Description, vbCritical
End Sub

Private Sub Command98_Click()
'Luodaan indeksisivut
On Error GoTo ErrorHandler
Debug.Print "Command98_Click: Starting index page creation (AutoCAD)"
Dim Taul As DAO.Recordset
Dim edPOS As String
Dim Sivut As String
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
Dim IndexBase As String
Dim edPiir As String

IndexBase = DLookup("[VALUE]", "[SETTINGS]", "SETTING = 'INDEXBASE'")
Debug.Print "  IndexBase: " & IndexBase
' Pohjakuvien hakemisto (Y-asema, INT-kansio) ja tallennuspolku SETTINGS-taulusta
BasePath = "Y:\base\int\"
SavePath = Nz(DLookup("[VALUE]", "[SETTINGS]", "SETTING = 'SAVEPATH'"), "")
Debug.Print "  BasePath: " & BasePath & "  SavePath: " & SavePath
If SavePath = "" Then
    MsgBox "SAVEPATH puuttuu SETTINGS-taulusta! Indeksisivuja ei voi tallentaa.", vbCritical
    GoTo ErrorHandler
End If
Set oAcad = GetObject(, "AutoCAD.Application")
Debug.Print "  AutoCAD connection established"
Set Taul = CurrentDb.OpenRecordset("Funtitle_Func")
Set oDoc = oAcad.Documents.Open(BasePath & IndexBase & ".dwg")
Debug.Print "  Base document opened: " & BasePath & IndexBase & ".dwg"

aX = 30        ' Aloituspiste X
aY = 285       ' Aloituspiste Y
PosSivu = 40   ' Position ja sivun välinen etäisyys x-akselilla
Seuraava = 90  ' Seuraavan positiosarakkeen etäisyys edellisestä
MinY = 35      ' Minimi Y-koordinaatti, johon asti tekstejä tehdään
MaxX = 333     ' Maksimi X-koordinaatti, johon sarakkeet tehdään ennen sivunvaihtoa
RiviVali = 5   ' Tekstien riviinjako
TextSize = 2.5 ' Tekstien koko


Piste(0) = aX
Piste(1) = aY
Piste2(0) = aX + PosSivu
Piste2(1) = aY

Debug.Print "  Processing records..."
Do While Not Taul.EOF
  If edPOS <> Taul.Fields("position") Then
    If Sivut <> "" Then
      If Piste(1) = aY Then
        'Lisätään ensin otsikot
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
  ' Viimeinen indeksisivu jää auki AutoCAD:issa (alkuperäinen käyttäytyminen)
  ' oDoc.Close
  Debug.Print "Command98_Click: COMPLETED successfully - Index pages created"
  MsgBox "Valmis"

  ' Vapautetaan resurssit
  Set Taul = Nothing
  Set oDoc = Nothing
  Set oAcad = Nothing
  Exit Sub

ErrorHandler:
    Debug.Print "*** ERROR in Command98_Click: " & Err.Number & " - " & Err.Description
    Debug.Print "    IndexBase: " & IndexBase
    Debug.Print "    Source: " & Err.Source & ", Line: " & Erl
    Set Taul = Nothing
    Set oDoc = Nothing
    Set oAcad = Nothing
    MsgBox "Virhe indeksisivujen luonnissa: " & Err.Description, vbCritical
End Sub


Private Sub Komento46_Click()
On Error GoTo ErrorHandler
Debug.Print "Komento46_Click: Starting interlocking page update"
Dim DB As DAO.Database
Dim Taul As DAO.Recordset
Dim Intpage As DAO.Recordset
Dim EdLoop As String
Dim PIIR As String
Dim PIIRI As String
Dim ALUE As String
Dim SYMB As String
Dim NRO As String
Dim OS As Long
Dim OS2 As Long
Dim Tyyppi As String

Set DB = CurrentDb
Debug.Print "  Clearing Intpage table"
DB.Execute "DELETE * FROM Intpage"
Set Taul = DB.OpenRecordset("SELECT DWG, TEXT_4, PAGE FROM POSANDPAGES GROUP BY DWG, TEXT_4, PAGE HAVING (TEXT_4<>'') ORDER BY TEXT_4, PAGE;")
Debug.Print "  Recordset opened with " & Taul.RecordCount & " records"
Set Intpage = DB.OpenRecordset("Intpage")


Do While Not Taul.EOF
    If EdLoop <> Taul.Fields("TEXT_4") Then
      If EdLoop <> "" Then
        Intpage.Update
      End If
      Intpage.AddNew
      EdLoop = Taul.Fields("TEXT_4")
    End If
    PIIR = Taul.Fields("DWG")
    PIIR = Left$(PIIR, InStr(10, PIIR, "_") - 1)
    PIIRI = Taul.Fields("TEXT_4")
    Tyyppi = Mid$(PIIRI, 5, 1)
    OS = InStr(PIIRI, Tyyppi)
    ALUE = Mid$(PIIRI, 1, 4)
    If Tyyppi = "-" Then
      NRO = Right$(PIIRI, 3)
    Else
      OS2 = InStr(PIIRI, "-")
      NRO = PIIRI
      SYMB = Mid$(PIIRI, OS2 + 1)
    End If
    
    If Intpage.Fields("Ensimmäinenkohteestapage") <> "" Then
      Intpage.Fields("Ensimmäinenkohteestapage") = Intpage.Fields("Ensimmäinenkohteestapage") & "/" & Taul.Fields("PAGE")
    Else
      Intpage.Fields("Ensimmäinenkohteestapage") = Taul.Fields("PAGE")
    End If
    Intpage.Fields("PIIRNRO") = PIIR
    Intpage.Fields("looppisymb") = PIIRI
    Intpage.Fields("PIIRNRO") = PIIR
   ' Intpage.Fields("Areacode") = ALUE
'    Intpage.Fields("Loopsymb") = SYMB
    Intpage.Fields("looppinro") = NRO
    Taul.MoveNext
Loop
Intpage.Update
Debug.Print "  Intpage table updated successfully"
Set Taul = Nothing
Set Intpage = Nothing
Set DB = Nothing
Debug.Print "Komento46_Click: COMPLETED successfully"
MsgBox "Lukituskaavioiden sivutiedot päivitetty Intpage taulukkoon.", vbInformation, "Lukituskaavion sivut"
Exit Sub

ErrorHandler:
    Debug.Print "*** ERROR in Komento46_Click: " & Err.Number & " - " & Err.Description
    Debug.Print "    Source: " & Err.Source & ", Line: " & Erl
    ' Vapautetaan resurssit myös virhetilanteessa
    Set Taul = Nothing
    Set Intpage = Nothing
    Set DB = Nothing
    MsgBox "Virhe lukituskaaviosivujen päivityksessä: " & Err.Description, vbCritical
End Sub

Private Sub Komento47_Click()
On Error GoTo ErrorHandler
Debug.Print "Komento47_Click: Starting control print update"
DoCmd.SetWarnings False
DoCmd.OpenQuery ("Control_Poistatulostus")
DoCmd.OpenQuery ("Control_Lisaa_tulostus")
DoCmd.SetWarnings True
Debug.Print "Komento47_Click: COMPLETED successfully"
Exit Sub

ErrorHandler:
    Debug.Print "*** ERROR in Komento47_Click: " & Err.Number & " - " & Err.Description
    Debug.Print "    Source: " & Err.Source & ", Line: " & Erl
    DoCmd.SetWarnings True
    MsgBox "Virhe tulostuslistauksen päivityksessä: " & Err.Description, vbCritical
End Sub

Private Sub hae_Click()
' Siirretään fokus edelliseen kentään (hakukenttä) ja suoritetaan löydä seuraava
On Error GoTo Err_hae_Click

    Screen.PreviousControl.SetFocus  ' SetFocus ei palauta arvoa - .LoopNo-ketjutus oli virheellinen
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
    Debug.Print "    Source: " & Err.Source & ", Line: " & Erl
    MsgBox "Virhe moottorinäkymän avauksessa: " & Err.Description, vbCritical
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
    Debug.Print "    Source: " & Err.Source & ", Line: " & Erl
    MsgBox "Virhe piirinäkymän avauksessa: " & Err.Description, vbCritical
End Sub

Private Sub Komento67_Click()
'Luodaan valituista blokeista resepti
On Error GoTo ErrorHandler
Debug.Print "Komento67_Click: Starting recipe creation from blocks"
Dim DB As DAO.Database
Dim RID As Long
Dim Taul As DAO.Recordset
Dim sSQL As String
If IsNull(RProcess.Value) Or IsNull(RCode.Value) Then
  Debug.Print "  ERROR: Process or Code missing"
  MsgBox "Syötä ensin Process ja Code reseptille", vbCritical, "Tee resepti"
  Exit Sub
End If
Debug.Print "  Process: " & RProcess.Value & ", Code: " & RCode.Value
Set DB = CurrentDb
Set Taul = DB.OpenRecordset("SELECT * FROM RECIPES WHERE Process='" & RProcess.Value & "' and Code = '" & RCode.Value & "'")
If Taul.RecordCount > 0 Then
  Debug.Print "  ERROR: Recipe already exists"
  MsgBox "Anna reseptille Process ja Code, jota ei ole vielä käytössä.", vbCritical, "Tee resepti"
  Exit Sub
End If
Debug.Print "  Creating new recipe in RECIPES table"
'Luodaan uusi resepti RECIPES taulukkoon.
Taul.AddNew
Taul.Fields("PROCESS") = RProcess.Value
Taul.Fields("CODE") = RCode.Value
Taul.Fields("INFO") = RInfo.Value
'Otetaan ylös luodun reseptin ID
RID = Taul.Fields("ID")
Taul.Update
Debug.Print "  New RecipeID: " & RID
'Kirjoitetana reseptin ID piirille (Loops tai drives)
Me.RecipeID = RID
'Luodaan tiedoista rivit RECIPE_TMP taulukkoon
sSQL = "INSERT INTO RECIPE_TMP ( RecipeID, FBlock, LoopSymb, Suffix, Description2, BlockPlace, [Note], Sivu, Sivuja )"
sSQL = sSQL & "SELECT " & RID & " AS ID, Control.FBlock, Control.LoopSymb, Control.Suffix, Control.Description2, Control.BlockPlace, Control.Note, Control.Sivu, Control.Sivuja "
sSQL = sSQL & "FROM Control WHERE (((Control.AreaCode)='" & Me.AreaCode & "') AND ((Control.LoopNo)='" & Me.LoopNo & "'));"
DB.Execute sSQL
'Kirjoitetaan vielä Control taulukkoon, mikä resepti piirillä on käytössä
sSQL = "UPDATE Control SET Control.RecipeID = " & RID & " "
sSQL = sSQL & "WHERE (((Control.AreaCode)='" & Me.AreaCode & "') AND ((Control.LoopNo)='" & Me.LoopNo & "'));"
DB.Execute sSQL
Debug.Print "  Recipe assigned to control loops"
Me.Refresh
Debug.Print "Komento67_Click: COMPLETED successfully - Recipe created"

' Vapautetaan DAO-resurssit
Set Taul = Nothing
Set DB = Nothing
Exit Sub

ErrorHandler:
    Debug.Print "*** ERROR in Komento67_Click: " & Err.Number & " - " & Err.Description
    Debug.Print "    Process: " & RProcess.Value & ", Code: " & RCode.Value
    Debug.Print "    Source: " & Err.Source & ", Line: " & Erl
    Set Taul = Nothing
    Set DB = Nothing
    MsgBox "Virhe reseptin luonnissa: " & Err.Description, vbCritical
End Sub


Private Sub Komento80_Click()
'Tämä päivittää reseptit kaikkiin piireihin, joissa on käytetty tiettyä reseptiä.
On Error GoTo ErrorHandler
Debug.Print "Komento80_Click: Starting recipe update to all circuits"
Dim DB As DAO.Database
Dim Vastaus As VbMsgBoxResult

Vastaus = MsgBox("Poistetaanko vanhat blokit piireistä, joille on määritelty resepti?", vbYesNoCancel, "PÄIVITÄ RESEPTIT KAIKKIIN PIIREIHIN")

If Vastaus = vbYes Then ' Poistetaan ensin kaikki vanhat blokit niistä piireistä, joissa on resepti
  Debug.Print "  Deleting old blocks from Control where RecipeID exists"
  CurrentDb.Execute "DELETE * FROM Control where Recipeid is not null"
End If
If Vastaus <> vbCancel Then
  Debug.Print "  Executing RECIPES_ADD query"
  CurrentDb.Execute "RECIPES_ADD"
  Debug.Print "  All circuits updated with recipes"
  MsgBox "Piirit päivitetty!", vbOKOnly, "PÄIVITÄ RESEPTIT KAIKKIIN PIIREIHIN"
  Me.Refresh
End If
Debug.Print "Komento80_Click: COMPLETED successfully"
Exit Sub

ErrorHandler:
    Debug.Print "*** ERROR in Komento80_Click: " & Err.Number & " - " & Err.Description
    Debug.Print "    Source: " & Err.Source & ", Line: " & Erl
    MsgBox "Virhe reseptien päivityksessä: " & Err.Description, vbCritical
End Sub
Private Sub CommandJANI_Click()
' Luodaan indeksisivut Intpage-taulun tiedoista AutoCAD-piirustuksiin
On Error GoTo ErrorHandler
Debug.Print "CommandJANI_Click: Starting index page generation (Intpage)"
Dim Taul As DAO.Recordset
Dim edPOS As String
Dim Sivut As String
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
Dim IndexBase As String
Dim edPiir As String

IndexBase = DLookup("[VALUE]", "[SETTINGS]", "SETTING = 'INDEXBASE'")
Debug.Print "  IndexBase: " & IndexBase
' Varmistetaan AutoCAD-yhteys — CommandJANI voidaan ajaa itsenäisesti
If oAcad Is Nothing Then
    Set oAcad = GetObject(, "AutoCAD.Application")
    Debug.Print "  AutoCAD connection established (CommandJANI)"
End If
' Pohjakuvien hakemisto ja tallennuspolku SETTINGS-taulusta
BasePath = "Y:\base\int\"
SavePath = Nz(DLookup("[VALUE]", "[SETTINGS]", "SETTING = 'SAVEPATH'"), "")
Debug.Print "  BasePath: " & BasePath & "  SavePath: " & SavePath
If SavePath = "" Then
    MsgBox "SAVEPATH puuttuu SETTINGS-taulusta! Indeksisivuja ei voi tallentaa.", vbCritical
    GoTo ErrorHandler
End If
Set Taul = CurrentDb.OpenRecordset("Intpage")
Debug.Print "  Intpage recordset opened"
Set oDoc = oAcad.Documents.Open(BasePath & IndexBase & ".dwg")
Debug.Print "  Base document opened: " & BasePath & IndexBase & ".dwg"

aX = 30        ' Aloituspiste X
aY = 275       ' Aloituspiste Y
PosSivu = 40   ' Position ja sivun välinen etäisyys x-akselilla
Seuraava = 90  ' Seuraavan positiosarakkeen etäisyys edellisestä
MinY = 40      ' Minimi Y-koordinaatti, johon asti tekstejä tehdään
MaxX = 333     ' Maksimi X-koordinaatti, johon sarakkeet tehdään ennen sivunvaihtoa
RiviVali = 5   ' Tekstien riviinjako
TextSize = 2.5 ' Tekstien koko


Piste(0) = aX
Piste(1) = aY
Piste2(0) = aX + PosSivu
Piste2(1) = aY

Debug.Print "  Processing Intpage records..."
Do While Not Taul.EOF
  If edPOS <> Taul.Fields("TEXT_4") Then
    If Sivut <> "" Then
      If Piste(1) = aY Then
        'Lisätään ensin otsikot
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

  ' Vapautetaan DAO-recordset ja AutoCAD-viitteet
  Set Taul = Nothing
  Set oDoc = Nothing
  Set oAcad = Nothing
  Exit Sub

ErrorHandler:
    Debug.Print "*** ERROR in CommandJANI_Click: " & Err.Number & " - " & Err.Description
    Debug.Print "    IndexBase: " & IndexBase
    Debug.Print "    Source: " & Err.Source & ", Line: " & Erl
    Set Taul = Nothing
    Set oDoc = Nothing
    Set oAcad = Nothing
    MsgBox "Virhe indeksisivujen luonnissa: " & Err.Description, vbCritical
End Sub
