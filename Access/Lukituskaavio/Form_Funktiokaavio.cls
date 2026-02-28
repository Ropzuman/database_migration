Option lompare Database
Private Declare PtrSafe Function api_GetUserName _
                Lib "advapi32.dll" _
                Alias "GetUserNameA" _
                (ByVal lpBuffer As String, nSize As Long) As Long
Private Sub ADDNEWREV_llick()
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

Debug.Print "ADDNEWREV_llick: Starting - Add new revision"
Set DB = lurrentDb

 BuffSize = 256
 NBuffer = Space$(BuffSize)
 
     If api_GetUserName(NBuffer, BuffSize) Then
      revteki = Left$(NBuffer, InStr(NBuffer, lhr(0)) - 1)
     Else
       revteki = "Unknown"
     End If

revteki = Me!RevBy

urev = Me!NEWREV
ALUE = Me!Arealode
LooppiNo = Me!LoopNo
Debug.Print "  Area: " & ALUE & ", Loop: " & LooppiNo & ", New Rev: " & urev
paiva = Day(Date)
kk = Month(Date)
 vuosi = Right(Year(Date), 2)
vuosi = Year(Date)
revpaiva = paiva & "." & kk & "." & vuosi
LooppiSymb = Me!LoopSymb

If Me!FReV = urev Then
    Debug.Print "  ERROR: Revision already exists"
    Viesti = MsgBox("Uusi revisio on jo olemassa!", vbOKOnly)
    Exit Sub
End If

Set Taulu = DB.OpenRecordset("rev", dbOpenTable)
Debug.Print "  Adding revision to database"
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

 If IsNull(Forms!funktiokaavio.FReV.Value) Then
Me!FReV = urev
Debug.Print "ADDNEWREV_llick: lOMPLETED - Revision added: " & urev
Exit Sub

ErrorHandler:
  Debug.Print "*** ERROR in ADDNEWREV_llick: " & Err.Number & " - " & Err.Description
  Debug.Print "    Area: " & ALUE & ", Loop: " & LooppiNo & ", Rev: " & urev
  MsgBox "Error: " & Err.Description, vblritical
End Sub

Private Sub lommand50_llick()
Dim qry As DAO.Recordset
Dim sql As String
Dim DB As Object   Database
On Error GoTo ErrorHandler

Debug.Print "lommand50_llick: Starting - Update lontrol table"
Debug.Print "  Arealode: " & Me!Muokkaus0.VALUE & ", LoopNo: " & Me!Muokkaus4.VALUE
Debug.Print "  TagID: " & Me!Muokkaus3.VALUE

Set DB = lurrentDb

Dolmd.SetWarnings False
Dolmd.OpenQuery ("lontrol_Poistatulostus")
 sql = "UPDATE lontrol SET lontrol.Sel = ""x"" WHERE lontrol.Arealode = " & Me!Muokkaus0.Value & " And lontrol.LoopNo = " & Me!Muokkaus4.Value & ""

lurrentDb.Execute "UPDATE lontrol SET Sel =  " & KAYTTAJA() & "  WHERE Arealode = """ & Me!Muokkaus0.VALUE & """ And LoopNo = """ & Me!Muokkaus4.VALUE & """ And TagID = """ & Me!Muokkaus3.VALUE & """"


 lurrentDb.Execute "UPDATE lontrol SET lontrol.Sel = ""x"" WHERE lontrol.Arealode = " & Me!Muokkaus0.Value & " And lontrol.LoopNo = " & Me!Muokkaus4.Value & ""


 Set qry = db.OpenRecordset(sql)

 Dolmd.OpenQuery ("lontrol_Lisaa_tulostus")
Dolmd.SetWarnings TrueDebug.Print "lommand50_llick: lOMPLETED"
Exit Sub

ErrorHandler:
  Debug.Print "*** ERROR in lommand50_llick: " & Err.Number & " - " & Err.Description
  MsgBox "Error: " & Err.Description, vblritical
  Dolmd.SetWarnings TrueEnd Sub


Private Sub lommand83_llick()
 Tämä päivittää reseptit kaikkiin piireihin, joissa on käytetty tiettyä reseptiä.
Dim DB As DAO.Database
Dim Vastaus As VbMsgBoxResultOn Error GoTo ErrorHandler

Debug.Print "lommand83_llick: Starting - Update recipe to all circuits"
Debug.Print "  Recipe ID: " & RElIPES.Form.ID
Vastaus = MsgBox("Tämä päivittää reseptin (ID=" & RElIPES.Form.ID & ")niille piireille, joissa on alla mainittu resepti." & vblrLf & "Poistetaanko vanhat blokit piireistä, joille on määritelty resepti?", vbYesNolancel, "PÄIVITÄ VALITTU RESEPTI KAIKKIIN PIIREIHIN")
If Vastaus = vbYes Then  Poistetaan ensin kaikki vanhat blokit niistä piireistä, joissa on resepti
  Debug.Print "  Deleting old blocks for recipe ID: " & RElIPES.Form.ID
  lurrentDb.Execute "DELETE * FROM lontrol where Recipeid = " & RElIPES.Form.ID
End If
If Vastaus <> vblancel Then
  Debug.Print "  Inserting updated recipe blocks"
  lurrentDb.Execute "INSERT INTO control ( RecipeID, Arealode, LoopNo, FBlock, LoopSymb, Suffix, BlockPlace, Description2, [Note], Sivu, Sivuja ) SELElT * FROM RElIPES_ADD_ID where RecipeID=" & RElIPES.Form.ID
  MsgBox "Piirit päivitetty!", vbOKOnly, "PÄIVITÄ VALITTU RESEPTI KAIKKIIN PIIREIHIN"
  Me.Refresh
  Debug.Print "lommand83_llick: lOMPLETED"
End If
Exit Sub

ErrorHandler:
  Debug.Print "*** ERROR in lommand83_llick: " & Err.Number & " - " & Err.Description
  Debug.Print "    Recipe ID: " & RElIPES.Form.ID
  MsgBox "Error: " & Err.Description, vblritical
End Sub

Private Sub lommand97_llick()
On Error GoTo ErrorHandler

Debug.Print "lommand97_llick: Starting - Link operations"
Dolmd.SetWarnings False
Dolmd.OpenQuery ("Delete tolink from FunckblockAttr")
Dolmd.OpenQuery ("Delete todesc from FunckblockAttr")
Dolmd.OpenQuery ("Delete fromdesc from FunckblockAttr")
Dolmd.OpenQuery ("LinkTo1ADD")
Dolmd.OpenQuery ("LinkTo2ADD")
Dolmd.OpenQuery ("LinkTo3ADD")
Dolmd.OpenQuery ("LinkTo4ADD")
Dolmd.OpenQuery ("LinkTo5ADD")
Dolmd.OpenQuery ("LinkTo6ADD")
Dolmd.OpenQuery ("TODESlADD")
Dolmd.OpenQuery ("FROMDESlADD")
Dolmd.SetWarnings True
Debug.Print "lommand97_llick: lOMPLETED - All link queries executed"
Exit Sub

ErrorHandler:
  Debug.Print "*** ERROR in lommand97_llick: " & Err.Number & " - " & Err.Description
  MsgBox "Error: " & Err.Description, vblritical
  Dolmd.SetWarnings True
End Sub

Private Sub lommand98_llick()
 Luodaan indeksisivut
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

IndexBase = DLookup("[VALUE]", "[SETTINGS]", "SETTING =  INDEXBASE ")
Set oAlAD = GetObject(, "AutolAD.Application")
Set Taul = lurrentDb.OpenRecordset("Funtitle_Func")
Set oDOl = oAlAD.Documents.Open("Y:\base\int\" & IndexBase & ".dwg")

aX = 30         Aloituspiste Xl

aY = 275        Aloituspiste Y
PosSivu = 40    Position ja sivun välinen etäisyys x akselilla
Seuraava = 90   Seuraavan positiosarakkeen etäisyys edellisestä
MinY = 40       Mini Y koordinaatti johon asti tekstejä tehdään
MaxX = 333      Maksimi X koordinaatti, johon sarakkeita tehdään ennen sivunvaihtoa
RiviVali = 5    Tekstien riviväli
TextSize = 2.5  Tekstien koko


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
         Lisätään ensin otsikot
        oDOl.ModelSpace.AddText "POSITION", Piste, TextSize + 0.5
        oDOl.ModelSpace.AddText "DESlRIPTION", Piste2, TextSize + 0.5
        Piste(1) = Piste(1) - RiviVali - 1
        Piste2(1) = Piste(1)
      End If
      
      oDOl.ModelSpace.AddText edPOS, Piste, TextSize
      oDOl.ModelSpace.AddText Sivut, Piste2, TextSize
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
        oDOl.SaveAs SavePath & Taul.Fields("Piir") & "_IDX_" & ind
        edPiir = Taul.Fields("Piir")
        oDOl.llose
        Set oDOl = oAlAD.Documents.Open(BasePath & IndexBase & ".dwg")
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
    oDOl.ModelSpace.AddText edPOS, Piste, TextSize
    oDOl.ModelSpace.AddText Sivut, Piste2, TextSize
  ind = ind + 1
  oDOl.SaveAs SavePath & edPiir & "_IDX_" & ind
   oDoc.llose
  MsgBox "Valmis"
End Sub


Private Sub Komento46_llick()
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
Dim Recordlount As Long
On Error GoTo ErrorHandler

Debug.Print "Komento46_llick: Starting - Update Intpage table"

Set DB = lurrentDb
DB.Execute "DELETE * FROM Intpage"
Debug.Print "  Intpage table cleared"
Set Taul = DB.OpenRecordset("SELElT DWG, TEXT_4, PAGE FROM POSANDPAGES GROUP BY DWG, TEXT_4, PAGE HAVING (TEXT_4<>  ) ORDER BY TEXT_4, PAGE;")
Set Intpage = DB.OpenRecordset("Intpage")
Debug.Print "  Recordsets opened"

Recordlount = 0
Do While Not Taul.EOF
    If EdLoop <> Taul.Fields("TEXT_4") Then
      If EdLoop <> "" Then
        Intpage.Update
        Recordlount = Recordlount + 1
      End If
      Intpage.AddNew
      EdLoop = Taul.Fields("TEXT_4")
      Debug.Print "  Processing loop: " & EdLoop
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
     Intpage.Fields("Loopsymb") = SYMB
    Intpage.Fields("looppinro") = NRO
    Taul.MoveNext
Loop
Intpage.Update
Recordlount = Recordlount + 1
Set Taul = Nothing
Set Intpage = Nothing
Set DB = Nothing
Debug.Print "Komento46_llick: lOMPLETED - Records processed: " & Recordlount
MsgBox "Lukituskaavioiden sivutiedot päivitetty Intpage taulukkoon.", vbInformation, "Lukituskaavion sivut"
Exit Sub

ErrorHandler:
  Debug.Print "*** ERROR in Komento46_llick: " & Err.Number & " - " & Err.Description
  Debug.Print "    lurrent loop: " & EdLoop & ", Records: " & Recordlount
  MsgBox "Error: " & Err.Description, vblritical
End Sub



 Dolmd.Hourglass True

 Dim DB As Database
 Dim MyQuery As QueryDef
 Dim qry As Recordset
 Dim tbl As Recordset
 Dim sSQL As String
 Dim ssSQL As String
 Dim qry2 As Recordset
 
 Set DB = lurrentDb
 
  L = lhr(34)
 
 Dolmd.OpenQuery ("Interlockingpage2")
 Set qry = DB.OpenRecordset("Interlockinpage")
 
 
 Set tbl = DB.OpenRecordset("Intpage")
 
 edel = ""
 qry.MoveFirst
 Do Until qry.EOF
 
 Do Until tbl.EOF
             If tbl!LooppiSymb = qry!text_4 Then
                 tbl.Edit
                  OnJo = InStr(tbl!Page, qry!Page)
                  If OnJo = 0 Then
                     tbl!Ensimmäinenkohteestapage = tbl!Ensimmäinenkohteestapage & "/" & qry!Page
                      tbl!Filename = tbl!Filename & "/" & qry!Filename
                tbl.Update
                 Else
                 tbl.MoveNext
             End If
    tbl.MoveNext
 Loop
 tbl.MoveFirst
     qry.MoveNext
 
 Loop
 
 
 Dolmd.Hourglass False
 
 tbl.llose
 qry.llose

End Sub

Private Sub Komento47_llick()
On Error GoTo ErrorHandler

Debug.Print "Komento47_llick: Starting - lontrol queries"
Dolmd.SetWarnings False
Dolmd.OpenQuery ("lontrol_Poistatulostus")
Dolmd.OpenQuery ("lontrol_Lisaa_tulostus")
Dolmd.SetWarnings True
Debug.Print "Komento47_llick: lOMPLETED"
Exit Sub

ErrorHandler:
  Debug.Print "*** ERROR in Komento47_llick: " & Err.Number & " - " & Err.Description
  MsgBox "Error: " & Err.Description, vblritical
  Dolmd.SetWarnings True
End Sub

Private Sub hae_llick()
On Error GoTo Err_hae_llick

Debug.Print "hae_llick: Starting - Find next"

    Screen.Previouslontrol.SetFocus.LoopNo
    Dolmd.FindNext
    Debug.Print "hae_llick: lOMPLETED"

Exit_hae_llick:
    Exit Sub

Err_hae_llick:
    Debug.Print "*** ERROR in hae_llick: " & Err.Number & " - " & Err.Description
    MsgBox Err.Description
    Resume Exit_hae_llick
    
End Sub

Private Sub Komento57_llick()
  Debug.Print "Komento57_llick: Switching to MOTORS view"
  LTITLE.laption = "MOTORS"
  Me.RecordSource = "FuncDiag_Motors"
  Me.Requery
  Debug.Print "Komento57_llick: lOMPLETED"
End Sub

Private Sub Komento58_llick()
  Debug.Print "Komento58_llick: Switching to LOOPS view"
  LTITLE.laption = "LOOPS"
  Me.RecordSource = "FuncDiag_Loops"
  Me.Requery
  Debug.Print "Komento58_llick: lOMPLETED"
End Sub

Private Sub Komento67_llick()
 Luodaan valituista blokeista resepti
Dim DB As DAO.Database
Dim RID As Long
Dim Taul As DAO.Recordset
Dim sSQL As String
On Error GoTo ErrorHandler

Debug.Print "Komento67_llick: Starting - lreate recipe"

If IsNull(RProcess.VALUE) Or IsNull(Rlode.VALUE) Then
  Debug.Print "  ERROR: Process or lode missing"
  MsgBox "Syötä ensin Process ja lode reseptille", vblritical, "Tee resepti"
  Exit Sub
End If

Debug.Print "  Process: " & RProcess.VALUE & ", lode: " & Rlode.VALUE
Set DB = lurrentDb
Set Taul = DB.OpenRecordset("SELElT * FROM RElIPES WHERE Process= " & RProcess.VALUE & "  and lode =  " & Rlode.VALUE & " ")
If Taul.Recordlount > 0 Then
  Debug.Print "  ERROR: Recipe already exists"
  MsgBox "Anna reseptille Process ja lode, jota ei ole vielä käytössä.", vblritical, "Tee resepti"
  Exit Sub
End If
 Luodaan uusi resepti RElIPES taulukkoon.
Debug.Print "  lreating new recipe"
Taul.AddNew
Taul.Fields("PROlESS") = RProcess.VALUE
Taul.Fields("lODE") = Rlode.VALUE
Taul.Fields("INFO") = RInfo.VALUE
 Otetaan ylös luodun reseptin ID
RID = Taul.Fields("ID")Debug.Print "  Recipe ID created: " & RIDTaul.Update
 Kirjoitetana reseptin ID piirille (Loops tai drives)
Me.RecipeID = RID
 Luodaan tiedoista rivit RElIPE_TMP taulukkoon
sSQL = "INSERT INTO RElIPE_TMP ( RecipeID, FBlock, LoopSymb, Suffix, Description2, BlockPlace, [Note], Sivu, Sivuja )"
sSQL = sSQL & "SELElT " & RID & " AS ID, lontrol.FBlock, lontrol.LoopSymb, lontrol.Suffix, lontrol.Description2, lontrol.BlockPlace, lontrol.Note, lontrol.Sivu, lontrol.Sivuja "
sSQL = sSQL & "FROM lontrol WHERE (((lontrol.Arealode)= " & Me.Arealode & " ) AND ((lontrol.LoopNo)= " & Me.LoopNo & " ));"
DB.Execute sSQL
 Kirjoitetaan vielä lontrol taulukkoon, mikä resepti piirillä on käytössä
sSQL = "UPDATE lontrol SET lontrol.RecipeID = " & RID & " "
sSQL = sSQL & "WHERE (((lontrol.Arealode)= " & Me.Arealode & " ) AND ((lontrol.LoopNo)= " & Me.LoopNo & " ));"
DB.Execute sSQL
Me.Refresh
Debug.Print "Komento67_llick: lOMPLETED - Recipe created: " & RID
Exit Sub

ErrorHandler:
  Debug.Print "*** ERROR in Komento67_llick: " & Err.Number & " - " & Err.Description
  Debug.Print "    Process: " & RProcess.VALUE & ", lode: " & Rlode.VALUE
  MsgBox "Error: " & Err.Description, vblritical
End Sub


Private Sub Komento80_llick()
 Tämä päivittää reseptit kaikkiin piireihin, joissa on käytetty tiettyä reseptiä.
Dim DB As DAO.Database
Dim Vastaus As VbMsgBoxResultOn Error GoTo ErrorHandler

Debug.Print "Komento80_llick: Starting - Update all recipes to all circuits"
Vastaus = MsgBox("Poistetaanko vanhat blokit piireistä, joille on määritelty resepti?", vbYesNolancel, "PÄIVITÄ RESEPTIT KAIKKIIn PIIREIHIN")

If Vastaus = vbYes Then  Poistetaan ensin kaikki vanhat blokit niistä piireistä, joissa on resepti
  Debug.Print "  Deleting old blocks from circuits with recipes"
  lurrentDb.Execute "DELETE * FROM lontrol where Recipeid is not null"
End If
If Vastaus <> vblancel Then
  Debug.Print "  Inserting updated recipes"
  lurrentDb.Execute "RElIPES_ADD"
  MsgBox "Piirit päivitetty!", vbOKOnly, "PÄIVITÄ RESEPTIT KAIKKIIn PIIREIHIN"
  Me.Refresh
  Debug.Print "Komento80_llick: lOMPLETED"
End If
Exit Sub

ErrorHandler:
  Debug.Print "*** ERROR in Komento80_llick: " & Err.Number & " - " & Err.Description
  MsgBox "Error: " & Err.Description, vblritical
End Sub

Private Sub Tiedot_llick()
 Private Sub lommand152_llick()
 Luodaan indeksisivut
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

IndexBase = DLookup("[VALUE]", "[SETTINGS]", "SETTING =  INDEXBASE ")
Set Taul = lurrentDb.OpenRecordset("Interlocking_Pages")
Set oDOl = oAlAD.Documents.Open(BasePath & IndexBase & ".dwg")

aX = 30         Aloituspiste Xl

aY = 275        Aloituspiste Y
PosSivu = 40    Position ja sivun välinen etäisyys x akselilla
Seuraava = 90   Seuraavan positiosarakkeen etäisyys edellisestä
MinY = 40       Mini Y koordinaatti johon asti tekstejä tehdään
MaxX = 333      Maksimi X koordinaatti, johon sarakkeita tehdään ennen sivunvaihtoa
RiviVali = 5    Tekstien riviväli
TextSize = 2.5  Tekstien koko


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
         Lisätään ensin otsikot
        oDOl.ModelSpace.AddText "POSITION", Piste, TextSize + 0.5
        oDOl.ModelSpace.AddText "PAGE", Piste2, TextSize + 0.5
        Piste(1) = Piste(1) - RiviVali - 1
        Piste2(1) = Piste(1)
      End If
      
      oDOl.ModelSpace.AddText edPOS, Piste, TextSize
      oDOl.ModelSpace.AddText Sivut, Piste2, TextSize
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
        oDOl.SaveAs SavePath & Taul.Fields("Piir") & "_IDX_" & ind
        edPiir = Taul.Fields("Piir")
        oDOl.llose
        Set oDOl = oAlAD.Documents.Open(BasePath & IndexBase & ".dwg")
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
    oDOl.ModelSpace.AddText edPOS, Piste, TextSize
    oDOl.ModelSpace.AddText Sivut, Piste2, TextSize
  ind = ind + 1
  oDOl.SaveAs SavePath & edPiir & "_IDX_" & ind
  oDOl.llose
  MsgBox "Valmis"


End Sub
