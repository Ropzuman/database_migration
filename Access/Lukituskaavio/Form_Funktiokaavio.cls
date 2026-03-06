Option Compare Database
Option Explicit

' HUOM: api_GetUserName-kutsu on poistettu käytöstä — käytetään Me!RevBy-kentän arvoa suoraan
Private Sub ADDNEWREV_Click()
Dim DB As DAO.Database
Dim vrev As String
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
Dim BuffSize As Long
Dim NBuffer As String
On Error GoTo ErrorHandler

Set DB = CurrentDb

'BuffSize = 256
'NBuffer = Space$(BuffSize)
'
'    If api_GetUserName(NBuffer, BuffSize) Then
'     revteki = Left$(NBuffer, InStr(NBuffer, Chr(0)) - 1)
'    Else
'      revteki = "Unknown"
'    End If

revteki = Me!RevBy

urev = Me!NEWREV
ALUE = Me!AreaCode
LooppiNo = Me!LoopNo
paiva = Day(Date)
kk = Month(Date)
'vuosi = Right(Year(Date), 2)
vuosi = Year(Date)
revpaiva = paiva & "." & kk & "." & vuosi
LooppiSymb = Me!LoopSymb

If Me!FReV = urev Then
    MsgBox "Uusi revisio on jo olemassa!", vbOKOnly
    Exit Sub
End If

Set Taulu = DB.OpenRecordset("rev", dbOpenTable)
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

'If IsNull(Forms!funktiokaavio.FReV.Value) Then
Me!FReV = urev
Exit Sub

ErrorHandler:
  MsgBox "Error: " & Err.Description, vbCritical
End Sub

Private Sub Command50_Click()
Dim qry As DAO.Recordset
Dim sql As String
Dim DB As Object ' Database
On Error GoTo ErrorHandler


Set DB = CurrentDb

DoCmd.SetWarnings False
DoCmd.OpenQuery ("Control_Poistatulostus")
'sql = "UPDATE Control SET Control.Sel = ""x"" WHERE Control.AreaCode = " & Me!Muokkaus0.Value & " And Control.LoopNo = " & Me!Muokkaus4.Value & ""

CurrentDb.Execute "UPDATE Control SET Sel = '" & KAYTTAJA() & "' WHERE AreaCode = """ & Me!Muokkaus0.VALUE & """ And LoopNo = """ & Me!Muokkaus4.VALUE & """ And TagID = """ & Me!Muokkaus3.VALUE & """"


'CurrentDb.Execute "UPDATE Control SET Control.Sel = ""x"" WHERE Control.AreaCode = " & Me!Muokkaus0.Value & " And Control.LoopNo = " & Me!Muokkaus4.Value & ""


'Set qry = db.OpenRecordset(sql)

'DoCmd.OpenQuery ("Control_Lisaa_tulostus")
DoCmd.SetWarnings True
Exit Sub

ErrorHandler:
  MsgBox "Error: " & Err.Description, vbCritical
  DoCmd.SetWarnings True
End Sub


Private Sub Command83_Click()
'Tämä päivittää reseptit kaikkiin piireihin, joissa on käytetty tiettyä reseptiä.
Dim DB As DAO.Database
Dim Vastaus As VbMsgBoxResult
On Error GoTo ErrorHandler

Vastaus = MsgBox("Tämä päivittää reseptin (ID=" & Me.RECIPES.Form.ID & ")niille piireille, joissa on alla mainittu resepti." & vbCrLf & "Poistetaanko vanhat blokit piireistä, joille on määritelty resepti?", vbYesNoCancel, "PÄIVITÄ VALITTU RESEPTI KAIKKIIN PIIREIHIN")
If Vastaus = vbYes Then 'Poistetaan ensin kaikki vanhat blokit niistä piireistä, joissa on resepti
  CurrentDb.Execute "DELETE * FROM Control where Recipeid = " & Me.RECIPES.Form.ID
End If
If Vastaus <> vbCancel Then
  CurrentDb.Execute "INSERT INTO control ( RecipeID, AreaCode, LoopNo, FBlock, LoopSymb, Suffix, BlockPlace, Description2, [Note], Sivu, Sivuja ) SELECT * FROM RECIPES_ADD_ID where RecipeID=" & Me.RECIPES.Form.ID
  MsgBox "Piirit päivitetty!", vbOKOnly, "PÄIVITÄ VALITTU RESEPTI KAIKKIIN PIIREIHIN"
  Me.Refresh
End If
Exit Sub

ErrorHandler:
  MsgBox "Error: " & Err.Description, vbCritical
End Sub

Private Sub Command97_Click()
On Error GoTo ErrorHandler

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
Exit Sub

ErrorHandler:
  MsgBox "Error: " & Err.Description, vbCritical
  DoCmd.SetWarnings True
End Sub

Private Sub Command98_Click()
'Luodaan indeksisivut
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
Set oACAD = GetObject(, "AutoCAD.Application")
Set Taul = CurrentDb.OpenRecordset("Funtitle_Func")
Set oDOC = oACAD.Documents.Open("Y:\base\int\" & IndexBase & ".dwg")

aX = 30        'Aloituspiste Xl

aY = 275       'Aloituspiste Y
PosSivu = 40   'Position ja sivun välinen etäisyys x akselilla
Seuraava = 90  'Seuraavan positiosarakkeen etäisyys edellisestä
MinY = 40      'Mini Y koordinaatti johon asti tekstejä tehdään
MaxX = 333     'Maksimi X koordinaatti, johon sarakkeita tehdään ennen sivunvaihtoa
RiviVali = 5   'Tekstien riviväli
TextSize = 2.5 'Tekstien koko


Piste(0) = aX
Piste(1) = aY
Piste2(0) = aX + PosSivu
Piste2(1) = aY

Do While Not Taul.EOF
If Taul.Fields("Position") = "=231P891A-A" Then
End If
  If edPOS <> Taul.Fields("position") Then
    If Sivut <> "" Then
      If Piste(1) = aY Then
        'Lisätään ensin otsikot
        oDOC.ModelSpace.AddText "POSITION", Piste, TextSize + 0.5
        oDOC.ModelSpace.AddText "DESCRIPTION", Piste2, TextSize + 0.5
        Piste(1) = Piste(1) - RiviVali - 1
        Piste2(1) = Piste(1)
      End If
      
      oDOC.ModelSpace.AddText edPOS, Piste, TextSize
      oDOC.ModelSpace.AddText Sivut, Piste2, TextSize
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
        oDOC.SaveAs SavePath & Taul.Fields("Piir") & "_IDX_" & ind
        edPiir = Taul.Fields("Piir")
        oDOC.Close
        Set oDOC = oACAD.Documents.Open(BasePath & IndexBase & ".dwg")
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
    oDOC.ModelSpace.AddText edPOS, Piste, TextSize
    oDOC.ModelSpace.AddText Sivut, Piste2, TextSize
  ind = ind + 1
  oDOC.SaveAs SavePath & edPiir & "_IDX_" & ind
  'oDoc.Close
  MsgBox "Valmis"
End Sub


Private Sub Komento46_Click()
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
Dim RecordCount As Long
On Error GoTo ErrorHandler


Set DB = CurrentDb
DB.Execute "DELETE * FROM Intpage"
Set Taul = DB.OpenRecordset("SELECT DWG, TEXT_4, PAGE FROM POSANDPAGES GROUP BY DWG, TEXT_4, PAGE HAVING (TEXT_4<>'') ORDER BY TEXT_4, PAGE;")
Set Intpage = DB.OpenRecordset("Intpage")

RecordCount = 0
Do While Not Taul.EOF
    If EdLoop <> Taul.Fields("TEXT_4") Then
      If EdLoop <> "" Then
        Intpage.Update
        RecordCount = RecordCount + 1
      End If
      Intpage.AddNew
      EdLoop = Taul.Fields("TEXT_4")
    End If
    PIIR = Taul.Fields("DWG")
    PIIR = Left$(PIIR, InStr(PIIR, "_") - 1)
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
    Intpage.Fields("Areacode") = ALUE
'    Intpage.Fields("Loopsymb") = SYMB
    Intpage.Fields("looppinro") = NRO
    Taul.MoveNext
Loop
Intpage.Update
RecordCount = RecordCount + 1
Set Taul = Nothing
Set Intpage = Nothing
Set DB = Nothing
MsgBox "Lukituskaavioiden sivutiedot päivitetty Intpage taulukkoon.", vbInformation, "Lukituskaavion sivut"
Exit Sub

ErrorHandler:
  MsgBox "Error: " & Err.Description, vbCritical
End Sub



'DoCmd.Hourglass True

'Dim DB As Database
'Dim MyQuery As QueryDef
'Dim qry As Recordset
'Dim tbl As Recordset
'Dim sSQL As String
'Dim ssSQL As String
'Dim qry2 As Recordset
'
'Set DB = CurrentDb
'
''L = Chr(34)
'
'DoCmd.OpenQuery ("Interlockingpage2")
'Set qry = DB.OpenRecordset("Interlockinpage")
'
'
'Set tbl = DB.OpenRecordset("Intpage")
'
'edel = ""
'qry.MoveFirst
'Do Until qry.EOF
'
'Do Until tbl.EOF
'            If tbl!LooppiSymb = qry!text_4 Then
'                tbl.Edit
'             '   OnJo = InStr(tbl!Page, qry!Page)
'              '  If OnJo = 0 Then
'                    tbl!Ensimmäinenkohteestapage = tbl!Ensimmäinenkohteestapage & "/" & qry!Page
'               '     tbl!Filename = tbl!Filename & "/" & qry!Filename
'               tbl.Update
'               'Else
'               'tbl.MoveNext
'            End If
'   tbl.MoveNext
'Loop
'tbl.MoveFirst
'    qry.MoveNext
'
'Loop
'
'
'DoCmd.Hourglass False
'
'tbl.Close
'qry.Close

End Sub

Private Sub Komento47_Click()
On Error GoTo ErrorHandler

DoCmd.SetWarnings False
DoCmd.OpenQuery ("Control_Poistatulostus")
DoCmd.OpenQuery ("Control_Lisaa_tulostus")
DoCmd.SetWarnings True
Exit Sub

ErrorHandler:
  MsgBox "Error: " & Err.Description, vbCritical
  DoCmd.SetWarnings True
End Sub

Private Sub hae_Click()
On Error GoTo Err_hae_Click


    Screen.PreviousControl.SetFocus.LoopNo
    DoCmd.FindNext

Exit_hae_Click:
    Exit Sub

Err_hae_Click:
    MsgBox Err.Description
    Resume Exit_hae_Click
    
End Sub

Private Sub Komento57_Click()
  Me.LTITLE.Caption = "MOTORS"
  Me.RecordSource = "FuncDiag_Motors"
  Me.Requery
End Sub

Private Sub Komento58_Click()
  Me.LTITLE.Caption = "LOOPS"
  Me.RecordSource = "FuncDiag_Loops"
  Me.Requery
End Sub

Private Sub Komento67_Click()
'Luodaan valituista blokeista resepti
Dim DB As DAO.Database
Dim RID As Long
Dim Taul As DAO.Recordset
Dim sSQL As String
On Error GoTo ErrorHandler


If IsNull(Me.RProcess.Value) Or IsNull(Me.RCode.Value) Then
  MsgBox "Syötä ensin Process ja Code reseptille", vbCritical, "Tee resepti"
  Exit Sub
End If

Set DB = CurrentDb
Set Taul = DB.OpenRecordset("SELECT * FROM RECIPES WHERE Process='" & Me.RProcess.Value & "' and Code = '" & Me.RCode.Value & "'")
If Taul.RecordCount > 0 Then
  MsgBox "Anna reseptille Process ja Code, jota ei ole vielä käytössä.", vbCritical, "Tee resepti"
  Exit Sub
End If
'Luodaan uusi resepti RECIPES taulukkoon.
Taul.AddNew
Taul.Fields("PROCESS") = Me.RProcess.Value
Taul.Fields("CODE") = Me.RCode.Value
Taul.Fields("INFO") = Me.RInfo.Value
'Otetaan ylös luodun reseptin ID
RID = Taul.Fields("ID")
Taul.Update
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
Me.Refresh
Exit Sub

ErrorHandler:
  MsgBox "Error: " & Err.Description, vbCritical
End Sub


Private Sub Komento80_Click()
'Tämä päivittää reseptit kaikkiin piireihin, joissa on käytetty tiettyä reseptiä.
Dim DB As DAO.Database
Dim Vastaus As VbMsgBoxResult
On Error GoTo ErrorHandler

Vastaus = MsgBox("Poistetaanko vanhat blokit piireistä, joille on määritelty resepti?", vbYesNoCancel, "PÄIVITÄ RESEPTIT KAIKKIIn PIIREIHIN")

If Vastaus = vbYes Then 'Poistetaan ensin kaikki vanhat blokit niistä piireistä, joissa on resepti
  CurrentDb.Execute "DELETE * FROM Control where Recipeid is not null"
End If
If Vastaus <> vbCancel Then
  CurrentDb.Execute "RECIPES_ADD"
  MsgBox "Piirit päivitetty!", vbOKOnly, "PÄIVITÄ RESEPTIT KAIKKIIn PIIREIHIN"
  Me.Refresh
End If
Exit Sub

ErrorHandler:
  MsgBox "Error: " & Err.Description, vbCritical
End Sub

Private Sub Tiedot_Click()
'Private Sub Command152_Click()
'Luodaan indeksisivut
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
Set Taul = CurrentDb.OpenRecordset("Interlocking_Pages")
Set oDOC = oACAD.Documents.Open(BasePath & IndexBase & ".dwg")

aX = 30        'Aloituspiste Xl

aY = 275       'Aloituspiste Y
PosSivu = 40   'Position ja sivun välinen etäisyys x akselilla
Seuraava = 90  'Seuraavan positiosarakkeen etäisyys edellisestä
MinY = 40      'Mini Y koordinaatti johon asti tekstejä tehdään
MaxX = 333     'Maksimi X koordinaatti, johon sarakkeita tehdään ennen sivunvaihtoa
RiviVali = 5   'Tekstien riviväli
TextSize = 2.5 'Tekstien koko


Piste(0) = aX
Piste(1) = aY
Piste2(0) = aX + PosSivu
Piste2(1) = aY

Do While Not Taul.EOF
If Taul.Fields("TEXT_4") = "=231P891A-A" Then
End If
  If edPOS <> Taul.Fields("TEXT_4") Then
    If Sivut <> "" Then
      If Piste(1) = aY Then
        'Lisätään ensin otsikot
        oDOC.ModelSpace.AddText "POSITION", Piste, TextSize + 0.5
        oDOC.ModelSpace.AddText "PAGE", Piste2, TextSize + 0.5
        Piste(1) = Piste(1) - RiviVali - 1
        Piste2(1) = Piste(1)
      End If
      
      oDOC.ModelSpace.AddText edPOS, Piste, TextSize
      oDOC.ModelSpace.AddText Sivut, Piste2, TextSize
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
        oDOC.SaveAs SavePath & Taul.Fields("Piir") & "_IDX_" & ind
        edPiir = Taul.Fields("Piir")
        oDOC.Close
        Set oDOC = oACAD.Documents.Open(BasePath & IndexBase & ".dwg")
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
    oDOC.ModelSpace.AddText edPOS, Piste, TextSize
    oDOC.ModelSpace.AddText Sivut, Piste2, TextSize
  ind = ind + 1
  oDOC.SaveAs SavePath & edPiir & "_IDX_" & ind
  oDOC.Close
  MsgBox "Valmis"


End Sub
