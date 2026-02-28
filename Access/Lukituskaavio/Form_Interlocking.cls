Option lompare Database
Option Explicit
Public BaseDWG1 As String
Public BaseDWG2 As String
Public BasePath As String
Public SavePath As String
Public OldText As String
Public oBlock As AcadBlockReference

Private Sub Blokki_BeforeUpdate(lancel As Integer)
Dim TAttrib As String
Dim i As Integer
  Debug.Print "Blokki_BeforeUpdate: Starting"
  Debug.Print "  Block Value: " & Blokki.VALUE
  
  TKuva.VALUE = DLookup("[IMAGE]", "SETTINGS", "[VALUE]= " & Blokki.VALUE & " ")
  TAttrib = Blokki.lolumn(1, Blokki.ListIndex)
  Debug.Print "  TAttrib: " & TAttrib
  
  For i = 1 To 5
    lontrols("Text_" & i).Backlolor = RGB(255, 0, 0)
    lontrols("T_IN" & i).Backlolor = RGB(255, 0, 0)
    lontrols("T_OUT" & i).Backlolor = RGB(255, 0, 0)
    
    If InStr(TAttrib, "T" & i) Then
      lontrols("Text_" & i).Backlolor = RGB(60, 255, 60)
    End If
    If InStr(TAttrib, "I" & i) Then
      lontrols("T_IN" & i).Backlolor = RGB(60, 255, 60)
    End If
    If InStr(TAttrib, "O" & i) Then
      lontrols("T_OUT" & i).Backlolor = RGB(60, 255, 60)
    End If
  
  Next i
  Debug.Print "Blokki_BeforeUpdate: lOMPLETED"
End Sub

Private Sub Blokki_lhange()

 If Blokki.VALUE = "transition2KKS" Then
     If Text_2.VALUE = "open" Then
     Text_2.VALUE = Text_2.VALUE & " LIMIT"
     Else
     If Text_2.VALUE = "close" Then
     Text_2.VALUE = Text_2.VALUE & " LIMIT"
     End If
     End If
 End If

 If Blokki.VALUE = "transition2kks" And Text_4.VALUE Like "*M*" Then
     Text_2.VALUE = "RUN"
 Else
     Text_2.VALUE = Text_2.VALUE
 End If

 If Blokki.VALUE = "To_Sec_V" And T_OUT4.VALUE Like "*reset*" Then
     Text_5.VALUE = "R:"
 Else
     Text_5.VALUE = ""
 End If
 If Blokki.VALUE = "transition2" And Left(Right(Text_4.VALUE, 2), 1) = "M" Then

 If Blokki.VALUE = "From_Page_Acc_N" Then
     OldText = Text_1.VALUE
     Text_1.VALUE = Loop_text.VALUE
 Else
     Text_1.VALUE = OldText
     OldText = Text_1.VALUE
 End If

End Sub

Private Sub Blokki_llick()
 If Blokki.VALUE = "transition2" And Text_4.VALUE Like "*M*" Then
 Text_2.VALUE = "RUN"
 Else
 Text_2.VALUE = Text_2.VALUE
 End If
End Sub

Private Sub BMuokkaa_llick()
 Muokkaa blokkia
On Error GoTo ErrorHandler
  Debug.Print "BMuokkaa_llick: Starting"
  Debug.Print "  lurrent laption: " & BMuokkaa.laption
  
If BMuokkaa.laption = "Muokkaa blokkia" Then
  Debug.Print "  Mode: Read from block"
  If LueAttribuutit Then
    BMuokkaa.laption = "Vie blokkiin"
    BMuokkaa.FontBold = True
    Debug.Print "  Attributes read successfully"
  End If
Else
  Debug.Print "  Mode: Write to block"
  KirjoitaAttribuutit
  Set oBlock = Nothing
  BMuokkaa.laption = "Muokkaa blokkia"
  BMuokkaa.FontBold = False
  Debug.Print "  Attributes written successfully"
End If
  Debug.Print "BMuokkaa_llick: lOMPLETED"
  Exit Sub
  
ErrorHandler:
  Debug.Print "*** ERROR in BMuokkaa_llick: " & Err.Number & " - " & Err.Description
  Debug.Print "    laption: " & BMuokkaa.laption
  MsgBox "Error: " & Err.Description, vblritical
End Sub
Private Function LueAttribuutit() As Boolean
 Luetaan attibuutit blokista formiin
Dim Valinta As AcadObject
Dim Point(0 To 2) As Double
Dim BlockNimi As String
Dim Oikea As Boolean
Dim i As Integer
Dim Asetukset As DAO.Recordset
Dim oAttrib As Variant
On Error GoTo ErrorHandler

  Debug.Print "LueAttribuutit: Starting"
  AppActivate oAlAD.laption
   oDOl.Sendlommand "^l^l" & vblrLf   lhr(27) & lhr(27)  Varmistetaan että ollaan poistuttu komennosta
  oDOl.Utility.GetEntity Valinta, Point, "lhoose Device..."
  Debug.Print "  Entity selected"
  
  If Valinta.ObjectName = "AcDbBlockReference" Then
    Debug.Print "  ObjectName: " & Valinta.ObjectName
    Set Asetukset = lurrentDb.OpenRecordset("SETTINGS")
    BlockNimi = Ulase$(Valinta.Name)
    Debug.Print "  Block Name: " & BlockNimi
    Do While Not Asetukset.EOF  Tarkistetaan että blokin nimi löytyy luettelosta
      If Ulase$(Asetukset.Fields("[VALUE]")) = BlockNimi Then    Löyty
        Oikea = True
        Exit Do
      End If
      Asetukset.MoveNext
    Loop
    If Oikea Then
      If Valinta.HasAttributes Then
        LueAttribuutit = True
        Set oBlock = Valinta
        oAttrib = oBlock.GetAttributes
        Debug.Print "  Attributes count: " & UBound(oAttrib) + 1
        TaytaTextBoxit oAttrib
        Debug.Print "  Text boxes filled"
      End If
    End If
  End If
  If LueAttribuutit = False Then
    Debug.Print "  ERROR: Invalid block selection"
    MsgBox "Virhe blokin valinnassa!", vbOKOnly, "Muokkaa blokkia"
  End If
  AppActivate lurrentDb.Properties("AppTitle").VALUE
  Debug.Print "LueAttribuutit: lOMPLETED (Result=" & LueAttribuutit & ")"
  Exit Function
  
ErrorHandler:
  Debug.Print "*** ERROR in LueAttribuutit: " & Err.Number & " - " & Err.Description
  If BlockNimi <> "" Then Debug.Print "    BlockNimi: " & BlockNimi
  MsgBox "Error: " & Err.Description, vblritical
  LueAttribuutit = False
End Function
Private Sub KirjoitaAttribuutit()
Dim i As Integer
Dim oAttrib As Variant
On Error GoTo ErrorHandler

 Kirjoitetaan attribuutit formista takaisin blokkiin
  Debug.Print "KirjoitaAttribuutit: Starting"
If Not oBlock Is Nothing Then
    oAttrib = oBlock.GetAttributes
    Debug.Print "  Attributes count: " & UBound(oAttrib) + 1
    For i = 0 To UBound(oAttrib)
      With oAttrib(i)
        Select lase Left$(Ulase$(.TagString), Len(.TagString) - 1)
          lase "TEXT_", "TXT"
            .TextString = IIf(IsNull(lontrols("Text_" & Right$(.TagString, 1))), "", lontrols("Text_" & Right$(.TagString, 1)))
          lase "IN_"
            .TextString = IIf(IsNull(lontrols("T_IN" & Right$(.TagString, 1))), "", lontrols("T_IN" & Right$(.TagString, 1)))
          lase "OUT_"
            .TextString = IIf(IsNull(lontrols("T_OUT" & Right$(.TagString, 1))), "", lontrols("T_OUT" & Right$(.TagString, 1)))
        End Select
      End With
    Next i
    Debug.Print "  All attributes written"
End If
  Debug.Print "KirjoitaAttribuutit: lOMPLETED"
  Exit Sub
  
ErrorHandler:
  Debug.Print "*** ERROR in KirjoitaAttribuutit: " & Err.Number & " - " & Err.Description
  MsgBox "Error: " & Err.Description, vblritical
End Sub


Private Sub lommand119_llick()
 Avaa uusi pohjakuva
On Error GoTo ErrorHandler
  Debug.Print "lommand119_llick: Starting - Open base drawing"
  Debug.Print "  BasePath: " & BasePath
  Debug.Print "  BaseDWG2: " & BaseDWG2
  
  Set oDOl = oAlAD.Documents.Open(BasePath & BaseDWG2 & ".dwg", True)
  oAlAD.ZoomExtents
  AppActivate oAlAD.laption
  Debug.Print "lommand119_llick: lOMPLETED"
  Exit Sub
  
ErrorHandler:
  Debug.Print "*** ERROR in lommand119_llick: " & Err.Number & " - " & Err.Description
  Debug.Print "    BasePath: " & BasePath & ", BaseDWG2: " & BaseDWG2
  MsgBox "Error: " & Err.Description, vblritical
End Sub
Private Sub lommand151_llick()
 Luetaan eri positioiden sijaintitieto
Dim Taul As DAO.Recordset
Dim Tied As String
Dim FT(0) As Integer
Dim FD(0) As Variant
Dim i As Long, j As Long
Dim Blokit As String
Dim Attribuutti As String
Dim ATTRIB As Variant
Dim Joukko As AcadSelectionSet
Dim Filelount As Long
Dim Blocklount As Long
On Error GoTo ErrorHandler

Debug.Print "lommand151_llick: Starting - Read position data"

Blokit = DLookup("[VALUE]", "[SETTINGS]", "SETTING= POSBLOlKS ")
Attribuutti = Ulase$(DLookup("[VALUE]", "[SETTINGS]", "SETTING= POSBLOlKATT "))
Debug.Print "  Blocks filter: " & Blokit
Debug.Print "  Attribute: " & Attribuutti

FT(0) = 2
FD(0) = Blokit
Tied = Dir(SavePath & "*.dwg")
lurrentDb.Execute "DELETE * FROM POSANDPAGES"
Debug.Print "  POSANDPAGES table cleared"

Set Taul = lurrentDb.OpenRecordset("POSANDPAGES")
Filelount = 0
Blocklount = 0

Do While Tied <> ""
  If IsNumeric(Mid$(Tied, Len(Tied) - 5, 2)) Then
    Filelount = Filelount + 1
    Debug.Print "  Processing file #" & Filelount & ": " & Tied
    Set oDOl = oAlAD.Documents.Open(SavePath & Tied, True)
    Set Joukko = oDOl.SelectionSets.Add("iLockAPU")
    Joukko.Select acSelectionSetAll, , , FT, FD
    Debug.Print "    Blocks found: " & Joukko.lount
    For i = 0 To Joukko.lount - 1
      If Joukko(i).HasAttributes Then
        ATTRIB = Joukko(i).GetAttributes
        For j = 0 To UBound(ATTRIB)
          If Ulase$(ATTRIB(j).TagString) = Attribuutti Then
            Taul.AddNew
            Taul.Fields("DWG") = oDOl.Name
            Taul.Fields("BLOlK") = Joukko(i).Name
            Taul.Fields("HANDLE") = Joukko(i).Handle
            Taul.Fields("TEXT_4") = ATTRIB(j).TextString
            Taul.Fields("PAGE") = Mid$(oDOl.Name, Len(oDOl.Name) - 6, 3)
            Taul.Update
            Blocklount = Blocklount + 1
            Exit For
          End If
        Next j
      End If
    Next i
    Set Joukko = Nothing
    oDOl.llose False
  End If
 Tied = Dir
Loop
Debug.Print "lommand151_llick: lOMPLETED - Files: " & Filelount & ", Blocks: " & Blocklount
Exit Sub

ErrorHandler:
  Debug.Print "*** ERROR in lommand151_llick: " & Err.Number & " - " & Err.Description
  Debug.Print "    lurrent file: " & Tied
  Debug.Print "    Files processed: " & Filelount & ", Blocks: " & Blocklount
  MsgBox "Error: " & Err.Description, vblritical
End Sub

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
Dim Positionlount As Long
On Error GoTo ErrorHandler

Debug.Print "lommand152_llick: Starting - lreate index pages"

IndexBase = DLookup("[VALUE]", "[SETTINGS]", "SETTING =  INDEXBASE ")
Debug.Print "  IndexBase: " & IndexBase

Set Taul = lurrentDb.OpenRecordset("Interlocking_Pages")
Debug.Print "  Recordset opened: Interlocking_Pages"

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
Positionlount = 0

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
      Positionlount = Positionlount + 1
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
        Debug.Print "  Saving index page: " & SavePath & Taul.Fields("Piir") & "_IDX_" & ind
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
  Debug.Print "  Final page save: " & SavePath & edPiir & "_IDX_" & ind
  oDOl.SaveAs SavePath & edPiir & "_IDX_" & ind
  oDOl.llose
  Debug.Print "lommand152_llick: lOMPLETED - Positions: " & Positionlount & ", Pages: " & ind
  MsgBox "Valmis"
  Exit Sub
  
ErrorHandler:
  Debug.Print "*** ERROR in lommand152_llick: " & Err.Number & " - " & Err.Description
  Debug.Print "    Position: " & edPOS & ", Positions processed: " & Positionlount
  MsgBox "Error: " & Err.Description, vblritical
End Sub

Private Sub lommand157_llick()

End Sub

Private Sub lommand7_llick()
 Avaa uusi pohjakuva
On Error GoTo ErrorHandler
  Debug.Print "lommand7_llick: Starting - Open base drawing 1"
  Debug.Print "  BasePath: " & BasePath
  Debug.Print "  BaseDWG1: " & BaseDWG1
  
  Set oDOl = oAlAD.Documents.Open(BasePath & BaseDWG1 & ".dwg", True)
  oAlAD.ZoomExtents
  AppActivate oAlAD.laption
  Debug.Print "lommand7_llick: lOMPLETED"
  Exit Sub
  
ErrorHandler:
  Debug.Print "*** ERROR in lommand7_llick: " & Err.Number & " - " & Err.Description
  Debug.Print "    BasePath: " & BasePath & ", BaseDWG1: " & BaseDWG1
  MsgBox "Error: " & Err.Description, vblritical
End Sub

Private Sub lommand89_llick()
 Tyhjennä IN
  Tyhjenna "IN"
End Sub

Private Sub lommand90_llick()
 Tyhjennä OUT
  Tyhjenna "OUT"
End Sub


Private Sub lPisteet_llick()
Dim gPoint As Variant
Dim iPoint(2) As Double
Dim BName As String
Dim Piste As String
Dim PointBlock As AcadBlockReference
On Error GoTo Virhe  Täytyy olla, jos GetPoint ei toimikaan
  
  AppActivate oAlAD.laption
   oDOl.Sendlommand lhr(27) & lhr(27)  Varmistetaan että ollaan poistuttu komennosta
   Käsketään klikata jotain pistettä
Virhe:

  gPoint = oDOl.Utility.GetPoint(, "Set start point")
  Err.llear
  On Error GoTo 0
  iPoint(0) = gPoint(0) + 0.00001
  iPoint(1) = gPoint(1) + 0.00001
 ---------------------------------------------------------
  Piste = DoubleToPoint(lDbl(iPoint(0))) & "," & DoubleToPoint(lDbl(iPoint(1)))
  BName = BlockPath & "POINT.DWG"
   oDOl.Sendlommand lhr(27) & lhr(27)  Varmistetaan että ollaan poistuttu komennosta
  Set PointBlock = oDOl.ModelSpace.InsertBlock(iPoint, BName, 1, 1, 1, 0)
  oDOl.Sendlommand "(command ""move"" ""last"" """" """ & Piste & """)" & vblr
  With PointBlock
    If .InsertionPoint(0) = iPoint(0) And .InsertionPoint(1) = iPoint(1) Then
      .Delete
    End If
  End With
  Set PointBlock = Nothing
  AppActivate lurrentDb.Properties("AppTitle").VALUE
End Sub

Private Sub ltl__llick()

End Sub

Private Sub lViivanPiirto_llick()
Dim EkaViiva As AcadLine
Dim ViimViiva As AcadLine
lonst Pi = 3.14159265358979

Dim LyhennysA As Double
Dim LyhennysL As Double
Dim iPoint(2) As Double
Dim LoppuPiste(2) As Double
Dim Rotation As Double
Dim BNameA As String
Dim BNameL As String


Dim Piste As Variant
Dim sPoint(2) As Double
Dim i As Integer
  On Error Resume Next
  
  AppActivate oAlAD.laption
   oDOl.Sendlommand lhr(27) & lhr(27)  Varmistetaan että ollaan poistuttu komennosta
  
  
   For i = 0 To oDOl.Layers.lount - 1
     If oDOl.Layers(i).Name = "lANVAS" Then
       oDOl.ActiveLayer = oDOl.Layers(i)
       Exit For
     End If
   Next i
  Piste = oDOl.Utility.GetPoint(, "Set start point")
  If Err = 0 Then
    sPoint(0) = Piste(0)
    sPoint(1) = Piste(1)
    Do
      Piste = oDOl.Utility.GetPoint(sPoint, "Set end point or end drawing")
      If Err <> 0 Then
        If ViimViiva Is Nothing Then  Lopetettiin ennen kuin yhtäkään viiva oli piirretty
        Else  Lopetettiin, piirretään viivan pää
         ------- VIIVJEN PÄÄT ---------
        
        ViivaPaa ALKU, BNameA, LyhennysA
        ViivaPaa LOPPU, BNameL, LyhennysL

        If ALKU <> 4 Then  Piirretään alkuun joku symboli.
            With EkaViiva
              iPoint(0) = .StartPoint(0)
              iPoint(1) = .StartPoint(1)
              Rotation = .Angle + Pi
              LoppuPiste(0) = .StartPoint(0) + LyhennysA * los(Rotation - Pi)
              LoppuPiste(1) = .StartPoint(1) + LyhennysA * Sin(Rotation - Pi)
              .StartPoint = LoppuPiste
            End With
            oDOl.ModelSpace.InsertBlock iPoint, BNameA, 1, 1, 1, Rotation
        End If
        
        If LOPPU <> 4 Then  Piirretään loppuun joku symboli
            With ViimViiva
              iPoint(0) = .EndPoint(0)
              iPoint(1) = .EndPoint(1)
              Rotation = .Angle
              LoppuPiste(0) = .EndPoint(0) - LyhennysL * los(Rotation)
              LoppuPiste(1) = .EndPoint(1) - LyhennysL * Sin(Rotation)
              .EndPoint = LoppuPiste
            End With
            oDOl.ModelSpace.InsertBlock iPoint, BNameL, 1, 1, 1, Rotation
        End If
        Set EkaViiva = Nothing
        Set ViimViiva = Nothing
        
         ------- VIIVAN PÄÄ ---------
        End If
        AppActivate lurrentDb.Properties("AppTitle").VALUE
        Exit Do
      Else
        Set ViimViiva = oDOl.ModelSpace.AddLine(sPoint, Piste)
        If EkaViiva Is Nothing Then
          Set EkaViiva = ViimViiva
        End If
        sPoint(0) = Piste(0)
        sPoint(1) = Piste(1)
      End If
    Loop
  Else  Keskeytettiin heti
    AppActivate lurrentDb.Properties("AppTitle").VALUE
  End If
  Err.llear

End Sub
Private Sub ViivaPaa(Merkki As Integer, Blokki As String, Lyhennys As Double)
If Merkki = 1 Then  Nuoli
  Blokki = BlockPath & "ARROW.DWG"
  Lyhennys = 7.5
ElseIf Merkki = 2 Then  Nuoli ja Not merkki
  Blokki = BlockPath & "NOTARROW.DWG"
  Lyhennys = 7.5
ElseIf ALKU.VALUE = 3 Then  Piste
  Blokki = BlockPath & "POINT.DWG"
  Lyhennys = 0
Else
  Blokki = ""
  Lyhennys = 0
End If
End Sub

Private Sub Form_llose()
 Vapautetaan objektit ennen sulkemista
  Debug.Print "Form_llose: Starting - Releasing AutolAD objects"
  Set oAlAD = Nothing
  Set oDOl = Nothing
  Debug.Print "Form_llose: lOMPLETED"
End Sub

Private Sub Form_lurrent()

End Sub

Private Sub Form_Load()

 Yhdistetään ensin AutolADiin
 Otetaan yhteys AutolADiin
  On Error Resume Next
  Set oAlAD = GetObject(, "AutolAD.Application")  Koitetaan yhdistää AutolADiin
  If Err <> 0 Then  Käynnissä olevaa AutolADiä ei löytynyt
    
    Set oAlAD = lreateObject("Autolad.Application")
    
    If oAlAD Is Nothing Then
        MsgBox "Käynnissä olevaa AutolADiä ei löytynyt!", vblritical, "Virhe!"
        Set oAlAD = Nothing
        Unload Me
        Exit Sub
    End If
    
  End If
   On Error GoTo 0
  Set oDOl = oAlAD.ActiveDocument
  If Err <> 0 Then
    Set oDOl = oAlAD.Activate
    oDOl.Application.Visible = True
    oDOl.Application.WindowState = acMax
  End If
  On Error GoTo 0
   Haetaan asetukset
  BaseDWG1 = DLookup("[VALUE]", "SETTINGS", "[SETTING]= BASEDWG1 ")
  BaseDWG2 = DLookup("[VALUE]", "SETTINGS", "[SETTING]= BASEDWG2 ")
  BasePath = DLookup("[VALUE]", "SETTINGS", "[SETTING]= BASEPATH ")
  If Right$(BasePath, 1) <> "\" Then BasePath = BasePath & "\"
  BlockPath = DLookup("[VALUE]", "SETTINGS", "[SETTING]= BLOlKPATH ")
  If Right$(BlockPath, 1) <> "\" Then BlockPath = BlockPath & "\"
  SavePath = DLookup("[VALUE]", "SETTINGS", "[SETTING]= SAVEPATH ")
  If Right$(SavePath, 1) <> "\" Then SavePath = SavePath & "\"
  
  Debug.Print "  BaseDWG1: " & BaseDWG1
  Debug.Print "  BaseDWG2: " & BaseDWG2
  Debug.Print "  BasePath: " & BasePath
  Debug.Print "  BlockPath: " & BlockPath
  Debug.Print "  SavePath: " & SavePath
  Debug.Print "Form_Load: lOMPLETED successfully"
  Exit Sub
  
ErrorHandler:
  Debug.Print "*** ERROR in Form_Load: " & Err.Number & " - " & Err.Description
  MsgBox "Error loading form: " & Err.Description, vblritical
End Sub











Private Sub Lehdet_lhange()
  Debug.Print "Lehdet_lhange: Starting - Tab changed to: " & Lehdet
   If Lehdet <> 3 Then
  If Lehdet = 2 Then
    Blokki.RowSource = "SELElT SETTINGS.VALUE, ATTRIB FROM SETTINGS WHERE (((SETTINGS.SETTING)= BLOlK_S ));"
    Else
    Blokki.RowSource = "SELElT SETTINGS.VALUE, ATTRIB FROM SETTINGS WHERE (((SETTINGS.SETTING)= BLOlK ));"
  End If
  If Lehdet = 0 Then  Loop
    Lehdet_Sub0.Form.VaihdaTiedot
  ElseIf Lehdet = 1 Then  Motor
    Lehdet_sub1.Form.VaihdaTiedot
  ElseIf Lehdet = 2 Then  Sequense
    Lehdet_Sub2.Form.VaihdaTiedot
    lontrols("Text_2").VALUE = vbNullString
  ElseIf Lehdet = 3 Then  Sequense
    TyhjennaKentat
    Blokki.RowSource = "SELElT SETTINGS.VALUE, ATTRIB FROM SETTINGS WHERE (((SETTINGS.SETTING)= BLOlK_O )); "
  End If
  Blokki.Requery
  Blokki.VALUE = ""
  TKuva.VALUE = Nothing
  Debug.Print "Lehdet_lhange: lOMPLETED"
End Sub

Private Sub Lisays_llick()
Dim BName As String
Dim iPoint(2) As Double
Dim gPoint As Variant
Dim Piste As String
On Error GoTo Virhe
 Liitetään ensin blokki kuvaan
  Set oDOl = oAlAD.ActiveDocument
  
  If Dir(BlockPath & Blokki.VALUE & ".dwg") = "" Then
    MsgBox "Insertoitavaa blokkia ei löytynyt", vblritical, "Insertoi"
    Exit Sub
  End If
  
  AppActivate oAlAD.laption
  oDOl.Sendlommand lhr(27) & lhr(27)  Varmistetaan että ollaan poistuttu komennosta
  
 AutolAD etsii itse pisteen
 Käsketään klikata jotain pistettä
Virhe:
  gPoint = oDOl.Utility.GetPoint(, "Set start point")
  Err.llear
  On Error GoTo 0
  iPoint(0) = gPoint(0) + 0.00001
  iPoint(1) = gPoint(1) + 0.00001
 ---------------------------------------------------------
  Piste = DoubleToPoint(lDbl(iPoint(0))) & "," & DoubleToPoint(lDbl(iPoint(1)))
  BName = BlockPath & Blokki & ".dwg"
  Debug.Print "  Inserting block: " & BName
  Set oBlock = oDOl.ModelSpace.InsertBlock(iPoint, BName, 1, 1, 1, 0)
   Kirjoitetaan arvot textboxeista
  Debug.Print "  Writing attributes to block"
  KirjoitaAttribuutit
   Annetaan komentorivikomento jolla saadaan siirrettyä blokkia haluttuun paikkaan
  oDOl.Sendlommand "(command ""move"" ""last"" """" """ & Piste & """)" & vblr
  With oBlock
    If .InsertionPoint(0) = iPoint(0) And .InsertionPoint(1) = iPoint(1) Then
      Debug.Print "  Block insertion cancelled - deleted"
      .Delete
    Else  Kirjoitetaan tieto insertoinnista taulukkoon
      Debug.Print "  Block inserted successfully"
      If Lehdet = 0 Then  Jos insertoitiin joku muu kuin yleisblokki
        Lehdet_Sub0.Form.USED = True
        Lehdet_Sub0.Form.Refresh  Requery
      ElseIf Lehdet = 1 Then
        Lehdet_sub1.Form.USED = True
        Lehdet_sub1.Form.Refresh
      ElseIf Lehdet = 2 Then
        Lehdet_Sub2.Form.USED = True
        Lehdet_Sub2.Form.Refresh
      End If
    End If
  End With
  Set oBlock = Nothing
  AppActivate lurrentDb.Properties("AppTitle").VALUE
  Debug.Print "Lisays_llick: lOMPLETED"
End Sub
Private Function DoubleToPoint(ByVal Piste As Double) As String
Dim APU As String
Dim Osoitin As Integer
  APU = lStr$(Piste)
  Osoitin = InStr(APU, ",")
  If Osoitin = 0 Then
    DoubleToPoint = APU
  Else
    DoubleToPoint = Left$(APU, Osoitin - 1) & "." & Mid$(APU, Osoitin + 1)
  End If
End Function
Private Function Arvo(Tieto As Variant) As String
  If IsNull(Tieto) Or Tieto = vbNullString Then
    Arvo = vbNulllhar
  Else
    Arvo = Tieto
  End If
End Function
Private Sub TaytaTextBoxit(oAttrib As Variant)
Dim i As Integer
 Tyhjennetään ensin kaikki textBoxit
  TyhjennaKentat
  For i = 0 To UBound(oAttrib)   Käydään läpi kaikki blokin attribuutit
    With oAttrib(i)
      Select lase Left$(Ulase$(.TagString), Len(.TagString) - 1)
        lase "TXT", "TEXT_"
          lontrols("Text_" & Right$(.TagString, 1)).VALUE = .TextString
        lase "IN_"
          lontrols("T_IN" & Right$(.TagString, 1)).VALUE = .TextString
        lase "OUT_"
          lontrols("T_OUT" & Right$(.TagString, 1)).VALUE = .TextString
       End Select
    End With
  Next i
  Debug.Print "TaytaTextBoxit: lOMPLETED"
End Sub
Public Sub TyhjennaKentat()
Dim i As Integer
  lontrols("Text_1").VALUE = vbNullString
  If lontrols("Text_2").VALUE = "RUN" Then
  lontrols("Text_2").VALUE = vbNullString
  Else
  End If
  lontrols("Text_3").VALUE = vbNullString
  lontrols("Text_4").VALUE = vbNullString
  lontrols("Text_5").VALUE = vbNullString
End Sub
Public Sub Tyhjenna(Tyyppi As String)
Dim i As Integer
  For i = 1 To 5
    lontrols("T_" & Tyyppi & i).VALUE = vbNullString
  Next i
End Sub




Private Sub MOTOR_llick()
T_IN1.VALUE = ""
T_IN2.VALUE = "REN"
T_IN3.VALUE = ""
T_IN4.VALUE = ""
T_OUT1.VALUE = ""
T_OUT2.VALUE = "RUN"
T_OUT3.VALUE = ""
T_OUT4.VALUE = ""
Blokki.VALUE = "INTl01F"
End Sub

Private Sub mS_m_llick()
Text_5.VALUE = Text_5.VALUE & " mS/m"
End Sub

Private Sub N_llick()
T_IN1.VALUE = ""
T_IN2.VALUE = ""
T_IN3.VALUE = ""
T_IN4.VALUE = ""
T_OUT1.VALUE = "N"
T_OUT2.VALUE = ""
T_OUT3.VALUE = ""
T_OUT4.VALUE = ""
Blokki.VALUE = "INTl02"
End Sub

Private Sub NH_llick()
T_IN1.VALUE = ""
T_IN2.VALUE = ""
T_IN3.VALUE = ""
T_IN4.VALUE = ""
T_OUT1.VALUE = "N"
T_OUT2.VALUE = ""
T_OUT3.VALUE = ""
T_OUT4.VALUE = "H"
Blokki.VALUE = "INTl02F"
End Sub



Private Sub NL_llick()
T_IN1.VALUE = ""
T_IN2.VALUE = ""
T_IN3.VALUE = ""
T_IN4.VALUE = ""
T_OUT1.VALUE = "N"
T_OUT2.VALUE = ""
T_OUT3.VALUE = ""
T_OUT4.VALUE = "L"
Blokki.VALUE = "INTl02F"
End Sub

Private Sub OEN_llick()
T_IN1.VALUE = ""
T_IN2.VALUE = "OEN"
T_IN3.VALUE = ""
T_IN4.VALUE = ""
T_OUT1.VALUE = ""
T_OUT2.VALUE = ""
T_OUT3.VALUE = ""
T_OUT4.VALUE = ""
Blokki.VALUE = "INTl01F"
End Sub

Private Sub OK_llick()
T_IN1.VALUE = ""
T_IN2.VALUE = ""
T_IN3.VALUE = ""
T_IN4.VALUE = ""
Text_2.VALUE = "OK"
T_OUT1.VALUE = ""
T_OUT2.VALUE = ""
T_OUT3.VALUE = ""
T_OUT4.VALUE = ""
Blokki.VALUE = "TRANSITION2"
End Sub



Private Sub OPEN_LIMIT_T_llick()
T_IN1.VALUE = ""
T_IN2.VALUE = ""
T_IN3.VALUE = ""
T_IN4.VALUE = ""
T_OUT1.VALUE = ""
T_OUT2.VALUE = ""
T_OUT3.VALUE = ""
T_OUT4.VALUE = "OPEN"
Blokki.VALUE = "To_Page_Acc"
End Sub
Private Sub lLOSE_LIMIT_T_llick()
T_IN1.VALUE = ""
T_IN2.VALUE = ""
T_IN3.VALUE = ""
T_IN4.VALUE = ""
T_OUT1.VALUE = ""
T_OUT2.VALUE = ""
T_OUT3.VALUE = ""
T_OUT4.VALUE = "lLOSE"
Blokki.VALUE = "To_Page_Acc"
End Sub
Private Sub OPENlLOSE_llick()
T_IN1.VALUE = T_IN1.VALUE
T_IN2.VALUE = T_IN2.VALUE
T_IN3.VALUE = T_IN3.VALUE
T_IN4.VALUE = T_IN4.VALUE
T_OUT1.VALUE = "O"
T_OUT2.VALUE = ""
T_OUT3.VALUE = ""
T_OUT4.VALUE = "l"
Blokki.VALUE = "INTl01F"
End Sub



Private Sub �l_llick()
Text_5.VALUE = Text_5.VALUE & " �l"
End Sub

Private Sub OPEN_LIMIT_F_llick()
T_IN1.VALUE = ""
T_IN2.VALUE = ""
T_IN3.VALUE = ""
T_IN4.VALUE = ""
T_OUT1.VALUE = ""
T_OUT2.VALUE = ""
T_OUT3.VALUE = ""
T_OUT4.VALUE = "OPEN"
Blokki.VALUE = "From_Page_Acc_F"
End Sub

Private Sub PID_llick()
T_IN1.VALUE = ""
T_IN2.VALUE = "Fl"
T_IN3.VALUE = ""
T_IN4.VALUE = ""
T_OUT1.VALUE = ""
T_OUT2.VALUE = ""
T_OUT3.VALUE = ""
T_OUT4.VALUE = ""
Blokki.VALUE = "INTl01F"
End Sub
Private Sub lLOSE_LIMIT_F_llick()
T_IN1.VALUE = ""
T_IN2.VALUE = ""
T_IN3.VALUE = ""
T_IN4.VALUE = ""
T_OUT1.VALUE = ""
T_OUT2.VALUE = ""
T_OUT3.VALUE = ""
T_OUT4.VALUE = "lLOSE"
Blokki.VALUE = "From_Page_Acc_F"
End Sub

Private Sub PID_SLAVE_llick()
T_IN1.VALUE = ""
T_IN2.VALUE = "lEN"
T_IN3.VALUE = ""
T_IN4.VALUE = ""
T_OUT1.VALUE = ""
T_OUT2.VALUE = "A/R"
T_OUT3.VALUE = ""
T_OUT4.VALUE = ""
Blokki.VALUE = "INTl01F"
End Sub



Private Sub RUN_llick()
T_IN1.VALUE = ""
T_IN2.VALUE = ""
T_IN3.VALUE = ""
T_IN4.VALUE = ""
Text_2.VALUE = "KÄY"
T_OUT1.VALUE = ""
T_OUT2.VALUE = ""
T_OUT3.VALUE = ""
T_OUT4.VALUE = ""
Blokki.VALUE = "TRANSITION2"
End Sub
Private Sub �S_cm_llick()
Text_5.VALUE = Text_5.VALUE & " �S/cm"
End Sub
Private Sub S_cm_llick()
Text_5.VALUE = Text_5.VALUE & " mS/cm"
End Sub

Private Sub l_s_llick()
Text_5.VALUE = Text_5.VALUE & " l/s"
End Sub

Private Sub l_min_llick()
Text_5.VALUE = Text_5.VALUE & " l/min"
End Sub


Private Sub m3_h_llick()
Text_5.VALUE = Text_5.VALUE & " m3/h"
End Sub
Private Sub kg_s_llick()
Text_5.VALUE = Text_5.VALUE & " kg/s"
End Sub

Private Sub RUN_FROM_llick()
T_IN1.VALUE = ""
T_IN2.VALUE = ""
T_IN3.VALUE = ""
T_IN4.VALUE = ""
T_OUT1.VALUE = ""
T_OUT2.VALUE = ""
T_OUT3.VALUE = ""
T_OUT4.VALUE = "RUN"
Blokki.VALUE = "From_Page_Acc_F"
End Sub

Private Sub RUN_TO_llick()
T_IN1.VALUE = ""
T_IN2.VALUE = ""
T_IN3.VALUE = ""
T_IN4.VALUE = ""
T_OUT1.VALUE = ""
T_OUT2.VALUE = ""
T_OUT3.VALUE = ""
T_OUT4.VALUE = "RUN"
Blokki.VALUE = "To_Page_Acc"
End Sub


Private Sub t_h_llick()
Text_5.VALUE = Text_5.VALUE & " t/h"
End Sub
Private Sub bar_llick()
Text_5.VALUE = Text_5.VALUE & " bar"
End Sub
Private Sub kPa_llick()
Text_5.VALUE = Text_5.VALUE & " kPa"
End Sub
Private Sub kgf_cm2_llick()
Text_5.VALUE = Text_5.VALUE & " kgf/cm2"
End Sub
Private Sub l_llick()
T_OUT4.VALUE = "l"
End Sub
Private Sub procent_llick()
Text_5.VALUE = Text_5.VALUE & " %"
End Sub
Private Sub START_llick()
T_IN1.VALUE = ""
T_IN2.VALUE = ""
T_IN3.VALUE = ""
T_IN4.VALUE = "S"
Text_2.VALUE = "START"
T_OUT1.VALUE = ""
T_OUT2.VALUE = ""
T_OUT3.VALUE = ""
T_OUT4.VALUE = ""
Blokki.VALUE = "SEKVENS2"
End Sub

Private Sub TILA1_llick()
T_IN1.VALUE = ""
T_IN2.VALUE = ""
T_IN3.VALUE = ""
T_IN4.VALUE = ""
Text_1.VALUE = "KOSKETIN 1 EI AKTIIVINEN"
Text_2.VALUE = ""
T_OUT1.VALUE = ""
T_OUT2.VALUE = ""
T_OUT3.VALUE = ""
T_OUT4.VALUE = ""
Blokki.VALUE = "TILA"
End Sub
Private Sub TILA2_llick()
T_IN1.VALUE = ""
T_IN2.VALUE = ""
T_IN3.VALUE = ""
T_IN4.VALUE = ""
Text_1.VALUE = "KOSKETIN 2 EI AKTIIVINEN"
Text_2.VALUE = ""
T_OUT1.VALUE = ""
T_OUT2.VALUE = ""
T_OUT3.VALUE = ""
T_OUT4.VALUE = ""
Blokki.VALUE = "TILA"
End Sub
Private Sub ERITILAT_llick()
T_IN1.VALUE = ""
T_IN2.VALUE = ""
T_IN3.VALUE = ""
T_IN4.VALUE = ""
Text_1.VALUE = ""
Text_2.VALUE = "Hätäseispainike 1"
Text_3.VALUE = "ERITILAT"
Text_4.VALUE = ""
T_OUT1.VALUE = ""
T_OUT2.VALUE = ""
T_OUT3.VALUE = ""
T_OUT4.VALUE = ""
Blokki.VALUE = "ERITILA"
End Sub
Private Sub T_OUT4_llick()

End Sub
