Option Compare Database
Private Declare Function api_GetUserName _
                Lib "advapi32.dll" _
                Alias "GetUserNameA" _
                (ByVal lpBuffer As String, nSize As Long) As Long
Private Sub ADDNEWREV_Click()
Dim DB As Database
Dim vrev As String
Dim urev As String
Dim Taulu As Recordset
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
    Viesti = MsgBox("Uusi revisio on jo olemassa!", vbOKOnly)
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
'Else
'vrev = Me!FReV
'Me!FReV = vrev & "/" & urev
'End If


End Sub

Private Sub Command50_Click()
Dim qry As Recordset
Dim sql As String
Dim DB As Object ' Database
Set DB = CurrentDb

DoCmd.SetWarnings False
DoCmd.OpenQuery ("Control_Poistatulostus")
'sql = "UPDATE Control SET Control.Sel = ""x"" WHERE Control.AreaCode = " & Me!Muokkaus0.Value & " And Control.LoopNo = " & Me!Muokkaus4.Value & ""

CurrentDb.Execute "UPDATE Control SET Sel = '" & KAYTTAJA() & "' WHERE AreaCode = """ & Me!Muokkaus0.Value & """ And LoopNo = """ & Me!Muokkaus4.Value & """ And TagID = """ & Me!Muokkaus3.Value & """"


'CurrentDb.Execute "UPDATE Control SET Control.Sel = ""x"" WHERE Control.AreaCode = " & Me!Muokkaus0.Value & " And Control.LoopNo = " & Me!Muokkaus4.Value & ""


'Set qry = db.OpenRecordset(sql)

'DoCmd.OpenQuery ("Control_Lisaa_tulostus")
DoCmd.SetWarnings True
End Sub


Private Sub Command83_Click()
'T�m� p�ivitt�� reseptit kaikkiin piireihin, joissa on k�ytetty tietty� resepti�.
Dim DB As Database
Dim Vastaus As VbMsgBoxResult

Vastaus = MsgBox("T�m� p�ivitt�� reseptin (ID=" & RECIPES.Form.ID & ")niille piireille, joissa on alla mainittu resepti." & vbCrLf & "Poistetaanko vanhat blokit piireist�, joille on m��ritely resepti?", vbYesNoCancel, "P�IVIT� VALITTU RESEPTI KAIKKIIN PIIREIHIN")
If Vastaus = vbYes Then 'Poistetaan ensin kaikki vanhat blokit niist� piireist�, joissa on resepti
  CurrentDb.Execute "DELETE * FROM Control where Recipeid = " & RECIPES.Form.ID
End If
If Vastaus <> vbCancel Then
  CurrentDb.Execute "INSERT INTO control ( RecipeID, AreaCode, LoopNo, FBlock, LoopSymb, Suffix, BlockPlace, Description2, [Note], Sivu, Sivuja ) SELECT * FROM RECIPES_ADD_ID where RecipeID=" & RECIPES.Form.ID
  MsgBox "Piirit p�ivitetty!", vbOKOnly, "P�IVIT� VALITTU RESEPTI KAIKKIIN PIIREIHIN"
  Me.Refresh
End If
End Sub

Private Sub Command97_Click()
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
End Sub

Private Sub Command98_Click()
'Luodaan indeksisivut
Dim Taul As Recordset
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
Set oAcad = GetObject(, "AutoCAD.Application")
Set Taul = CurrentDb.OpenRecordset("Funtitle_Func")
Set oDoc = oAcad.Documents.Open("Y:\base\int\" & IndexBase & ".dwg")

aX = 30        'Aloituspiste Xl

aY = 275       'Aloituspiste Y
PosSivu = 40   'Position ja sivun v�linen et�isyys x akselilla
Seuraava = 90  'Seuraavan positiosarakkeen et�isyys edellisest�
MinY = 40      'Mini Y koordinaatti johon asti tektej� tehd��n
MaxX = 333     'Maksimi X koordinaatti, johon sarakkeita tehd��n ennen sivunvaihtoa
RiviVali = 5   'Tekstien riviv�li
TextSize = 2.5 'Tekstien koko


Piste(0) = aX
Piste(1) = aY
Piste2(0) = aX + PosSivu
Piste2(1) = aY

Do While Not Taul.EOF
If Taul.Fields("Position") = "=231P891A-A" Then
  Debug.Print ""
End If
  If edPOS <> Taul.Fields("position") Then
    If Sivut <> "" Then
      If Piste(1) = aY Then
        'Lis�t��n ensin otsikot
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
  oDoc.SaveAs SavePath & edPiir & "_IDX_" & ind
  'oDoc.Close
  MsgBox "Valmis"
End Sub


Private Sub Komento46_Click()
Dim DB As Database
Dim Taul As Recordset
Dim Intpage As Recordset
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
DB.Execute "DELETE * FROM Intpage"
Set Taul = DB.OpenRecordset("SELECT DWG, TEXT_4, PAGE FROM POSANDPAGES GROUP BY DWG, TEXT_4, PAGE HAVING (TEXT_4<>'') ORDER BY TEXT_4, PAGE;")
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
    PIIR = Left(PIIR, InStr(PIIR, "_") - 1)
    PIIRI = Taul.Fields("TEXT_4")
    Tyyppi = Mid(PIIRI, 5, 1)
    OS = InStr(PIIRI, Tyyppi)
    ALUE = Mid(PIIRI, 1, 4)
    If Tyyppi = "-" Then
      NRO = Right(PIIRI, 3)
    Else
      OS2 = InStr(PIIRI, "-")
      NRO = PIIRI
      SYMB = Mid(PIIRI, OS2 + 1)
    End If
    
    If Intpage.Fields("Ensimm�inenkohteestapage") <> "" Then
      Intpage.Fields("Ensimm�inenkohteestapage") = Intpage.Fields("Ensimm�inenkohteestapage") & "/" & Taul.Fields("PAGE")
    Else
      Intpage.Fields("Ensimm�inenkohteestapage") = Taul.Fields("PAGE")
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
Set Taul = Nothing
Set Intpage = Nothing
Set DB = Nothing
MsgBox "Lukituskaavioiden sivutiedot p�ivitetty Intpage taulukkoon.", vbInformation, "Lukituskaavion sivut"



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
'                    tbl!Ensimm�inenkohteestapage = tbl!Ensimm�inenkohteestapage & "/" & qry!Page
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
DoCmd.SetWarnings False
DoCmd.OpenQuery ("Control_Poistatulostus")
DoCmd.OpenQuery ("Control_Lisaa_tulostus")
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
  LTITLE.Caption = "MOTORS"
  Me.RecordSource = "FuncDiag_Motors"
  Me.Requery
End Sub

Private Sub Komento58_Click()
  LTITLE.Caption = "LOOPS"
  Me.RecordSource = "FuncDiag_Loops"
  Me.Requery
End Sub

Private Sub Komento67_Click()
'Luodaan valituista blokeista resepti
Dim DB As Database
Dim RID As Long
Dim Taul As Recordset
Dim sSQL As String
If IsNull(RProcess.Value) Or IsNull(RCode.Value) Then
  MsgBox "Sy�t� ensin Process ja Code reseptille", vbCritical, "Tee resepti"
  Exit Sub
End If
Set DB = CurrentDb
Set Taul = DB.OpenRecordset("SELECT * FROM RECIPES WHERE Process='" & RProcess.Value & "' and Code = '" & RCode.Value & "'")
If Taul.RecordCount > 0 Then
  MsgBox "Anna reseptille Process ja Code, jota ei ole viel� k�yt�ss�.", vbCritical, "Tee resepti"
  Exit Sub
End If
'Luodaan uusi resepti RECIPES taulukkoon.
Taul.AddNew
Taul.Fields("PROCESS") = RProcess.Value
Taul.Fields("CODE") = RCode.Value
Taul.Fields("INFO") = RInfo.Value
'Otetaan yl�s luodun reseptin ID
RID = Taul.Fields("ID")
Taul.Update
'Kirjoitetana reseptin ID piirille (Loops tai drives)
Me.RecipeID = RID
'Luodaan tiedoista rivit RECIPE_TMP taulukkoon
sSQL = "INSERT INTO RECIPE_TMP ( RecipeID, FBlock, LoopSymb, Suffix, Description2, BlockPlace, [Note], Sivu, Sivuja )"
sSQL = sSQL & "SELECT " & RID & " AS ID, Control.FBlock, Control.LoopSymb, Control.Suffix, Control.Description2, Control.BlockPlace, Control.Note, Control.Sivu, Control.Sivuja "
sSQL = sSQL & "FROM Control WHERE (((Control.AreaCode)='" & Me.AreaCode & "') AND ((Control.LoopNo)='" & Me.LoopNo & "'));"
DB.Execute sSQL
'Kirjoitetaan viel� Control taulukkoon, mik� resepti piirill� on k�yt�ss�
sSQL = "UPDATE Control SET Control.RecipeID = " & RID & " "
sSQL = sSQL & "WHERE (((Control.AreaCode)='" & Me.AreaCode & "') AND ((Control.LoopNo)='" & Me.LoopNo & "'));"
DB.Execute sSQL
Me.Refresh


End Sub


Private Sub Komento80_Click()
'T�m� p�ivitt�� reseptit kaikkiin piireihin, joissa on k�ytetty tietty� resepti�.
Dim DB As Database
Dim Vastaus As VbMsgBoxResult

Vastaus = MsgBox("Poistetaanko vanhat blokit piireist�, joille on m��ritely resepti?", vbYesNoCancel, "P�IVIT� RESEPTIT KAIKKIIn PIIREIHIN")

If Vastaus = vbYes Then 'Poistetaan ensin kaikki vanhat blokit niist� piireist�, joissa on resepti
  CurrentDb.Execute "DELETE * FROM Control where Recipeid is not null"
End If
If Vastaus <> vbCancel Then
  CurrentDb.Execute "RECIPES_ADD"
  MsgBox "Piirit p�ivitetty!", vbOKOnly, "P�IVIT� RESEPTIT KAIKKIIn PIIREIHIN"
  Me.Refresh
End If
End Sub

Private Sub Tiedot_Click()
'Private Sub Command152_Click()
'Luodaan indeksisivut
Dim Taul As Recordset
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
Set oDoc = oAcad.Documents.Open(BasePath & IndexBase & ".dwg")

aX = 30        'Aloituspiste Xl

aY = 275       'Aloituspiste Y
PosSivu = 40   'Position ja sivun v�linen et�isyys x akselilla
Seuraava = 90  'Seuraavan positiosarakkeen et�isyys edellisest�
MinY = 40      'Mini Y koordinaatti johon asti tektej� tehd��n
MaxX = 333     'Maksimi X koordinaatti, johon sarakkeita tehd��n ennen sivunvaihtoa
RiviVali = 5   'Tekstien riviv�li
TextSize = 2.5 'Tekstien koko


Piste(0) = aX
Piste(1) = aY
Piste2(0) = aX + PosSivu
Piste2(1) = aY

Do While Not Taul.EOF
If Taul.Fields("TEXT_4") = "=231P891A-A" Then
  Debug.Print ""
End If
  If edPOS <> Taul.Fields("TEXT_4") Then
    If Sivut <> "" Then
      If Piste(1) = aY Then
        'Lis�t��n ensin otsikot
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
  oDoc.SaveAs SavePath & edPiir & "_IDX_" & ind
  oDoc.Close
  MsgBox "Valmis"


End Sub
