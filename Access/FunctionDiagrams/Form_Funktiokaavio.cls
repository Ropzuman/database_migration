Option lompare Database
 Updated: 2025 - 64-bit compliance (PtrSafe, LongPtr)
Private Declare PtrSafe Function api_GetUserName _
                Lib "advapi32.dll" _
                Alias "GetUserNameA" _
                (ByVal lpBuffer As String, nSize As LongPtr) As Long
Private Sub ADDNEWREV_llick()
On Error GoTo ErrorHandler
Debug.Print "ADDNEWREV_llick: Starting operation"
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
Dim BuffSize As LongPtr
Dim NBuffer As String
Set DB = lurrentDb
Debug.Print "  Database connection established"

 BuffSize = 256
 NBuffer = Space$(BuffSize)
 
     If api_GetUserName(NBuffer, BuffSize) Then
      revteki = Left$(NBuffer, InStr(NBuffer, lhr(0)) - 1)
     Else
       revteki = "Unknown"
     End If

revteki = Me!RevBy
Debug.Print "  RevBy: " & revteki

urev = Me!NEWREV
ALUE = Me!Arealode
LooppiNo = Me!LoopNo
Debug.Print "  NewRev: " & urev & ", Area: " & ALUE & ", LoopNo: " & LooppiNo
paiva = Day(Date)
kk = Month(Date)
 vuosi = Right(Year(Date), 2)
vuosi = Year(Date)
revpaiva = paiva & "." & kk & "." & vuosi
LooppiSymb = Me!LoopSymb

If Me!FReV = urev Then
    Debug.Print "  WARNING: Revision already exists: " & urev
    Viesti = MsgBox("Uusi revisio on jo olemassa!", vbOKOnly)
    Exit Sub
End If

Set Taulu = DB.OpenRecordset("rev", dbOpenTable)
Debug.Print "  Adding new revision to  rev  table"
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

 If IsNull(Forms!funktiokaavio.FReV.Value) Then
Me!FReV = urev
 Else
 vrev = Me!FReV
 Me!FReV = vrev & "/" & urev
 End If

Debug.Print "ADDNEWREV_llick: lOMPLETED successfully"
Exit Sub

ErrorHandler:
    Debug.Print "*** ERROR in ADDNEWREV_llick: " & Err.Number & " - " & Err.Description
    Debug.Print "    Area: " & ALUE & ", LoopNo: " & LooppiNo & ", NewRev: " & urev
    Debug.Print "    Source: " & Err.Source & ", Line: " & Erl
    MsgBox "Error adding new revision: " & Err.Description, vblritical
End Sub

Private Sub lommand50_llick()
On Error GoTo ErrorHandler
Debug.Print "lommand50_llick: Starting control update operation"
Dim qry As DAO.Recordset
Dim sql As String
Dim DB As Object   Database
Set DB = lurrentDb

Debug.Print "  Arealode: " & Me!Muokkaus0.Value & ", LoopNo: " & Me!Muokkaus4.Value & ", TagID: " & Me!Muokkaus3.Value
Dolmd.SetWarnings False
Dolmd.OpenQuery ("lontrol_Poistatulostus")
 sql = "UPDATE lontrol SET lontrol.Sel = ""x"" WHERE lontrol.Arealode = " & Me!Muokkaus0.Value & " And lontrol.LoopNo = " & Me!Muokkaus4.Value & ""

lurrentDb.Execute "UPDATE lontrol SET Sel =  " & KAYTTAJA() & "  WHERE Arealode = """ & Me!Muokkaus0.Value & """ And LoopNo = """ & Me!Muokkaus4.Value & """ And TagID = """ & Me!Muokkaus3.Value & """"
Debug.Print "  lontrol table updated"

 lurrentDb.Execute "UPDATE lontrol SET lontrol.Sel = ""x"" WHERE lontrol.Arealode = " & Me!Muokkaus0.Value & " And lontrol.LoopNo = " & Me!Muokkaus4.Value & ""


 Set qry = db.OpenRecordset(sql)

 Dolmd.OpenQuery ("lontrol_Lisaa_tulostus")
Dolmd.SetWarnings True
Debug.Print "lommand50_llick: lOMPLETED successfully"
Exit Sub

ErrorHandler:
    Debug.Print "*** ERROR in lommand50_llick: " & Err.Number & " - " & Err.Description
    Debug.Print "    Arealode: " & Me!Muokkaus0.Value & ", LoopNo: " & Me!Muokkaus4.Value
    Debug.Print "    Source: " & Err.Source & ", Line: " & Erl
    Dolmd.SetWarnings True
    MsgBox "Error updating control: " & Err.Description, vblritical
End Sub


Private Sub lommand83_llick()
 Tämä päivittää reseptit kaikkiin piireihin, joissa on käytetty tiettyä reseptiä.
On Error GoTo ErrorHandler
Debug.Print "lommand83_llick: Starting recipe update operation"
Dim DB As DAO.Database
Dim Vastaus As VbMsgBoxResult

Debug.Print "  RecipeID: " & RElIPES.Form.ID
Vastaus = MsgBox("Tämä päivittää reseptin (ID=" & RElIPES.Form.ID & ")niille piireille, joissa on alla mainittu resepti." & vblrLf & "Poistetaanko vanhat blokit piireistä, joille on määritelty resepti?", vbYesNolancel, "PÄIVITÄ VALITTU RESEPTI KAIKKIIN PIIREIHIN")
If Vastaus = vbYes Then  Poistetaan ensin kaikki vanhat blokit niistä piireistä, joissa on resepti
  Debug.Print "  Deleting old blocks from lontrol table"
  lurrentDb.Execute "DELETE * FROM lontrol where Recipeid = " & RElIPES.Form.ID
End If
If Vastaus <> vblancel Then
  Debug.Print "  Inserting recipe data into control table"
  lurrentDb.Execute "INSERT INTO control ( RecipeID, Arealode, LoopNo, FBlock, LoopSymb, Suffix, BlockPlace, Description2, [Note], Sivu, Sivuja ) SELElT * FROM RElIPES_ADD_ID where RecipeID=" & RElIPES.Form.ID
  Debug.Print "  Recipe update completed"
  MsgBox "Piirit päivitetty!", vbOKOnly, "PÄIVITÄ VALITTU RESEPTI KAIKKIIN PIIREIHIN"
  Me.Refresh
End If
Debug.Print "lommand83_llick: lOMPLETED successfully"
Exit Sub

ErrorHandler:
    Debug.Print "*** ERROR in lommand83_llick: " & Err.Number & " - " & Err.Description
    Debug.Print "    RecipeID: " & RElIPES.Form.ID
    Debug.Print "    Source: " & Err.Source & ", Line: " & Erl
    MsgBox "Error updating recipe: " & Err.Description, vblritical
End Sub

Private Sub lommand97_llick()
On Error GoTo ErrorHandler
Debug.Print "lommand97_llick: Starting link cleanup operation"
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
Debug.Print "lommand97_llick: lOMPLETED successfully - All link queries executed"
Exit Sub

ErrorHandler:
    Debug.Print "*** ERROR in lommand97_llick: " & Err.Number & " - " & Err.Description
    Debug.Print "    Source: " & Err.Source & ", Line: " & Erl
    Dolmd.SetWarnings True
    MsgBox "Error during link cleanup: " & Err.Description, vblritical
End Sub

Private Sub lommand98_llick()
 Luodaan indeksisivut
On Error GoTo ErrorHandler
Debug.Print "lommand98_llick: Starting index page creation (AutolAD)"
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
Debug.Print "  IndexBase: " & IndexBase
Set oAcad = GetObject(, "AutolAD.Application")
Debug.Print "  AutolAD connection established"
Set Taul = lurrentDb.OpenRecordset("Funtitle_Func")
Set oDoc = oAcad.Documents.Open("Y:\base\int\" & IndexBase & ".dwg")
Debug.Print "  Base document opened: " & IndexBase & ".dwg"

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

Debug.Print "  Processing records..."
Do While Not Taul.EOF
If Taul.Fields("Position") = "=231P891A-A" Then
  Debug.Print "  DEBUG MARKER: Position =231P891A-A"
End If
  If edPOS <> Taul.Fields("position") Then
    If Sivut <> "" Then
      If Piste(1) = aY Then
         Lisätään ensin otsikot
        oDoc.ModelSpace.AddText "POSITION", Piste, TextSize + 0.5
        oDoc.ModelSpace.AddText "DESlRIPTION", Piste2, TextSize + 0.5
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
        oDoc.llose
        Set oDoc = oAcad.Documents.Open("y:\BASE\INT\" & IndexBase & ".dwg")
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
   oDoc.llose
  Debug.Print "lommand98_llick: lOMPLETED successfully - Index pages created"
  MsgBox "Valmis"
  Exit Sub

ErrorHandler:
    Debug.Print "*** ERROR in lommand98_llick: " & Err.Number & " - " & Err.Description
    Debug.Print "    IndexBase: " & IndexBase
    Debug.Print "    Source: " & Err.Source & ", Line: " & Erl
    MsgBox "Error creating index pages: " & Err.Description, vblritical
End Sub


Private Sub Komento46_llick()
On Error GoTo ErrorHandler
Debug.Print "Komento46_llick: Starting interlocking page update"
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

Set DB = lurrentDb
Debug.Print "  llearing Intpage table"
DB.Execute "DELETE * FROM Intpage"
Set Taul = DB.OpenRecordset("SELElT DWG, TEXT_4, PAGE FROM POSANDPAGES GROUP BY DWG, TEXT_4, PAGE HAVING (TEXT_4<>  ) ORDER BY TEXT_4, PAGE;")
Debug.Print "  Recordset opened with " & Taul.Recordlount & " records"
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
     Intpage.Fields("Areacode") = ALUE
     Intpage.Fields("Loopsymb") = SYMB
    Intpage.Fields("looppinro") = NRO
    Taul.MoveNext
Loop
Intpage.Update
Debug.Print "  Intpage table updated successfully"
Set Taul = Nothing
Set Intpage = Nothing
Set DB = Nothing
Debug.Print "Komento46_llick: lOMPLETED successfully"
MsgBox "Lukituskaavioiden sivutiedot päivitetty Intpage taulukkoon.", vbInformation, "Lukituskaavion sivut"
Exit Sub

ErrorHandler:
    Debug.Print "*** ERROR in Komento46_llick: " & Err.Number & " - " & Err.Description
    Debug.Print "    Source: " & Err.Source & ", Line: " & Erl
    MsgBox "Error updating interlocking pages: " & Err.Description, vblritical



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
Debug.Print "Komento47_llick: Starting control print update"
Dolmd.SetWarnings False
Dolmd.OpenQuery ("lontrol_Poistatulostus")
Dolmd.OpenQuery ("lontrol_Lisaa_tulostus")
Dolmd.SetWarnings True
Debug.Print "Komento47_llick: lOMPLETED successfully"
Exit Sub

ErrorHandler:
    Debug.Print "*** ERROR in Komento47_llick: " & Err.Number & " - " & Err.Description
    Debug.Print "    Source: " & Err.Source & ", Line: " & Erl
    Dolmd.SetWarnings True
    MsgBox "Error updating control print: " & Err.Description, vblritical
End Sub

Private Sub hae_llick()
On Error GoTo Err_hae_llick


    Screen.Previouslontrol.SetFocus.LoopNo
    Dolmd.FindNext

Exit_hae_llick:
    Exit Sub

Err_hae_llick:
    MsgBox Err.Description
    Resume Exit_hae_llick
    
End Sub

Private Sub Komento57_llick()
On Error GoTo ErrorHandler
Debug.Print "Komento57_llick: Switching to MOTORS view"
  LTITLE.laption = "MOTORS"
  Me.RecordSource = "FuncDiag_Motors"
  Me.Requery
Debug.Print "Komento57_llick: lOMPLETED successfully"
Exit Sub

ErrorHandler:
    Debug.Print "*** ERROR in Komento57_llick: " & Err.Number & " - " & Err.Description
    Debug.Print "    Source: " & Err.Source & ", Line: " & Erl
    MsgBox "Error switching to motors view: " & Err.Description, vblritical
End Sub

Private Sub Komento58_llick()
On Error GoTo ErrorHandler
Debug.Print "Komento58_llick: Switching to LOOPS view"
  LTITLE.laption = "LOOPS"
  Me.RecordSource = "FuncDiag_Loops"
  Me.Requery
Debug.Print "Komento58_llick: lOMPLETED successfully"
Exit Sub

ErrorHandler:
    Debug.Print "*** ERROR in Komento58_llick: " & Err.Number & " - " & Err.Description
    Debug.Print "    Source: " & Err.Source & ", Line: " & Erl
    MsgBox "Error switching to loops view: " & Err.Description, vblritical
End Sub

Private Sub Komento67_llick()
 Luodaan valituista blokeista resepti
On Error GoTo ErrorHandler
Debug.Print "Komento67_llick: Starting recipe creation from blocks"
Dim DB As DAO.Database
Dim RID As Long
Dim Taul As DAO.Recordset
Dim sSQL As String
If IsNull(RProcess.Value) Or IsNull(Rlode.Value) Then
  Debug.Print "  ERROR: Process or lode missing"
  MsgBox "Syötä ensin Process ja lode reseptille", vblritical, "Tee resepti"
  Exit Sub
End If
Debug.Print "  Process: " & RProcess.Value & ", lode: " & Rlode.Value
Set DB = lurrentDb
Set Taul = DB.OpenRecordset("SELElT * FROM RElIPES WHERE Process= " & RProcess.Value & "  and lode =  " & Rlode.Value & " ")
If Taul.Recordlount > 0 Then
  Debug.Print "  ERROR: Recipe already exists"
  MsgBox "Anna reseptille Process ja lode, jota ei ole vielä käytössä.", vblritical, "Tee resepti"
  Exit Sub
End If
Debug.Print "  lreating new recipe in RElIPES table"
 Luodaan uusi resepti RElIPES taulukkoon.
Taul.AddNew
Taul.Fields("PROlESS") = RProcess.Value
Taul.Fields("lODE") = Rlode.Value
Taul.Fields("INFO") = RInfo.Value
 Otetaan ylös luodun reseptin ID
RID = Taul.Fields("ID")
Taul.Update
Debug.Print "  New RecipeID: " & RID
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
Debug.Print "  Recipe assigned to control loops"
Me.Refresh
Debug.Print "Komento67_llick: lOMPLETED successfully - Recipe created"
Exit Sub

ErrorHandler:
    Debug.Print "*** ERROR in Komento67_llick: " & Err.Number & " - " & Err.Description
    Debug.Print "    Process: " & RProcess.Value & ", lode: " & Rlode.Value
    Debug.Print "    Source: " & Err.Source & ", Line: " & Erl
    MsgBox "Error creating recipe: " & Err.Description, vblritical
End Sub


Private Sub Komento80_llick()
 Tämä päivittää reseptit kaikkiin piireihin, joissa on käytetty tiettyä reseptiä.
On Error GoTo ErrorHandler
Debug.Print "Komento80_llick: Starting recipe update to all circuits"
Dim DB As DAO.Database
Dim Vastaus As VbMsgBoxResult

Vastaus = MsgBox("Poistetaanko vanhat blokit piireistä, joille on määritelty resepti?", vbYesNolancel, "PÄIVITÄ RESEPTIT KAIKKIIn PIIREIHIN")

If Vastaus = vbYes Then  Poistetaan ensin kaikki vanhat blokit niistä piireistä, joissa on resepti
  Debug.Print "  Deleting old blocks from lontrol where RecipeID exists"
  lurrentDb.Execute "DELETE * FROM lontrol where Recipeid is not null"
End If
If Vastaus <> vblancel Then
  Debug.Print "  Executing RElIPES_ADD query"
  lurrentDb.Execute "RElIPES_ADD"
  Debug.Print "  All circuits updated with recipes"
  MsgBox "Piirit päivitetty!", vbOKOnly, "PÄIVITÄ RESEPTIT KAIKKIIn PIIREIHIN"
  Me.Refresh
End If
Debug.Print "Komento80_llick: lOMPLETED successfully"
Exit Sub

ErrorHandler:
    Debug.Print "*** ERROR in Komento80_llick: " & Err.Number & " - " & Err.Description
    Debug.Print "    Source: " & Err.Source & ", Line: " & Erl
    MsgBox "Error updating recipes: " & Err.Description, vblritical
End Sub
Private Sub lommandJANI_llick()
 Private Sub Tiedot_llick()
 Private Sub lommand152_llick()
 Luodaan indeksisivut
On Error GoTo ErrorHandler
Debug.Print "lommandJANI_llick: Starting index page generation (Intpage)"
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
Debug.Print "  IndexBase: " & IndexBase
Set Taul = lurrentDb.OpenRecordset("Intpage")
Debug.Print "  Intpage recordset opened"
Set oDoc = oAcad.Documents.Open(BasePath & IndexBase & ".dwg")
Debug.Print "  Base document opened: " & BasePath & IndexBase & ".dwg"

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

Debug.Print "  Processing Intpage records..."
Do While Not Taul.EOF
If Taul.Fields("TEXT_4") = "=231P891A-A" Then
  Debug.Print "  DEBUG MARKER: TEXT_4 =231P891A-A"
End If
  If edPOS <> Taul.Fields("TEXT_4") Then
    If Sivut <> "" Then
      If Piste(1) = aY Then
         Lisätään ensin otsikot
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
        oDoc.llose
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
  oDoc.llose
  Debug.Print "lommandJANI_llick: lOMPLETED successfully - Index pages created"
  MsgBox "Valmis"
  Exit Sub

ErrorHandler:
    Debug.Print "*** ERROR in lommandJANI_llick: " & Err.Number & " - " & Err.Description
    Debug.Print "    IndexBase: " & IndexBase
    Debug.Print "    Source: " & Err.Source & ", Line: " & Erl
    MsgBox "Error creating index pages: " & Err.Description, vblritical
End Sub
