Option Compare Database
Option Explicit
Public BaseDWG1 As String
Public BaseDWG2 As String
Public BasePath As String
Public SavePath As String
Public OldText As String
Public oBlock As AcadBlockReference

Private Sub Blokki_BeforeUpdate(Cancel As Integer)
Dim TAttrib As String
Dim i As Integer
  TKuva.VALUE = DLookup("[IMAGE]", "SETTINGS", "[VALUE]='" & Blokki.VALUE & "'")
  TAttrib = Blokki.Column(1, Blokki.ListIndex)
  For i = 1 To 5
    Controls("Text_" & i).BackColor = RGB(255, 0, 0)
    Controls("T_IN" & i).BackColor = RGB(255, 0, 0)
    Controls("T_OUT" & i).BackColor = RGB(255, 0, 0)
    
    If InStr(TAttrib, "T" & i) Then
      Controls("Text_" & i).BackColor = RGB(60, 255, 60)
    End If
    If InStr(TAttrib, "I" & i) Then
      Controls("T_IN" & i).BackColor = RGB(60, 255, 60)
    End If
    If InStr(TAttrib, "O" & i) Then
      Controls("T_OUT" & i).BackColor = RGB(60, 255, 60)
    End If
  
  Next i
End Sub

Private Sub Blokki_Change()

'If Blokki.VALUE = "transition2KKS" Then
'    If Text_2.VALUE = "open" Then
'    Text_2.VALUE = Text_2.VALUE & " LIMIT"
'    Else
'    If Text_2.VALUE = "close" Then
'    Text_2.VALUE = Text_2.VALUE & " LIMIT"
'    End If
'    End If
'End If

'If Blokki.VALUE = "transition2kks" And Text_4.VALUE Like "*M*" Then
'    Text_2.VALUE = "RUN"
'Else
'    Text_2.VALUE = Text_2.VALUE
'End If

'If Blokki.VALUE = "To_Sec_V" And T_OUT4.VALUE Like "*reset*" Then
'    Text_5.VALUE = "R:"
'Else
'    Text_5.VALUE = ""
'End If
'If Blokki.VALUE = "transition2" And Left(Right(Text_4.VALUE, 2), 1) = "M" Then

'If Blokki.VALUE = "From_Page_Acc_N" Then
'    OldText = Text_1.VALUE
'    Text_1.VALUE = Loop_text.VALUE
'Else
'    Text_1.VALUE = OldText
'    OldText = Text_1.VALUE
'End If

End Sub

Private Sub Blokki_Click()
'If Blokki.VALUE = "transition2" And Text_4.VALUE Like "*M*" Then
'Text_2.VALUE = "RUN"
'Else
'Text_2.VALUE = Text_2.VALUE
'End If
End Sub

Private Sub BMuokkaa_Click()
'Muokkaa blokkia
If BMuokkaa.Caption = "Muokkaa blokkia" Then
  If LueAttribuutit Then
    BMuokkaa.Caption = "Vie blokkiin"
    BMuokkaa.FontBold = True
  End If
Else
  KirjoitaAttribuutit
  Set oBlock = Nothing
  BMuokkaa.Caption = "Muokkaa blokkia"
  BMuokkaa.FontBold = False
End If
End Sub
Private Function LueAttribuutit() As Boolean
'Luetaan attibuutit blokista formiin
Dim Valinta As AcadObject
Dim Point(0 To 2) As Double
Dim BlockNimi As String
Dim Oikea As Boolean
Dim i As Integer
Dim Asetukset As Recordset
Dim oAttrib As Variant

  AppActivate oACAD.Caption
  'oDOC.SendCommand "^C^C" & vbCrLf  'Chr(27) & Chr(27) 'Varmistetaan ett� ollaan poistuttu komennosta
  oDOC.Utility.GetEntity Valinta, Point, "Choose Device..."
  If Valinta.ObjectName = "AcDbBlockReference" Then
    Set Asetukset = CurrentDb.OpenRecordset("SETTINGS")
    BlockNimi = UCase(Valinta.Name)
    Do While Not Asetukset.EOF 'Tarkistetaan ett� blokin nimi l�ytyy luettelosta
      If UCase(Asetukset.Fields("[VALUE]")) = BlockNimi Then   'L�yty
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
        TaytaTextBoxit oAttrib
      End If
    End If
  End If
  If LueAttribuutit = False Then
    MsgBox "Virhe blokin valinnassa!", vbOKOnly, "Muokkaa blokkia"
  End If
  AppActivate CurrentDb.Properties("AppTitle").VALUE
End Function
Private Sub KirjoitaAttribuutit()
Dim i As Integer
Dim oAttrib As Variant
'Kirjoitetaan attribuutit formista takaisin blokkiin
If Not oBlock Is Nothing Then
    oAttrib = oBlock.GetAttributes
    For i = 0 To UBound(oAttrib)
      With oAttrib(i)
        Select Case Left(UCase(.TagString), Len(.TagString) - 1)
          Case "TEXT_", "TXT"
            .TextString = Nz(Controls("Text_" & Right(.TagString, 1)))
          Case "IN_"
            .TextString = Nz(Controls("T_IN" & Right(.TagString, 1)))
          Case "OUT_"
            .TextString = Nz(Controls("T_OUT" & Right(.TagString, 1)))
        End Select
      End With
    Next i
End If
End Sub


Private Sub Command119_Click()
'Avaa uusi pohjakuva
  Set oDOC = oACAD.Documents.Open(BasePath & BaseDWG2 & ".dwg", True)
  oACAD.ZoomExtents
  AppActivate oACAD.Caption
End Sub
Private Sub Command151_Click()
'Luetaan eri positioiden sijaintitieto
Dim Taul As Recordset
Dim Tied As String
Dim FT(0) As Integer
Dim FD(0) As Variant
Dim i As Long, j As Long
Dim Blokit As String
Dim Attribuutti As String
Dim ATTRIB As Variant
Dim Joukko As AcadSelectionSet



Blokit = DLookup("[VALUE]", "[SETTINGS]", "SETTING='POSBLOCKS'")
Attribuutti = UCase(DLookup("[VALUE]", "[SETTINGS]", "SETTING='POSBLOCKATT'"))

FT(0) = 2
FD(0) = Blokit
Tied = Dir(SavePath & "*.dwg")
CurrentDb.Execute "DELETE * FROM POSANDPAGES"
Set Taul = CurrentDb.OpenRecordset("POSANDPAGES")
Do While Tied <> ""
  If IsNumeric(Mid(Tied, Len(Tied) - 5, 2)) Then
    Set oDOC = oACAD.Documents.Open(SavePath & Tied, True)
    Set Joukko = oDOC.SelectionSets.Add("iLockAPU")
    Joukko.Select acSelectionSetAll, , , FT, FD
    For i = 0 To Joukko.Count - 1
      If Joukko(i).HasAttributes Then
        ATTRIB = Joukko(i).GetAttributes
        For j = 0 To UBound(ATTRIB)
          If UCase(ATTRIB(j).TagString) = Attribuutti Then
            Taul.AddNew
            Taul.Fields("DWG") = oDOC.Name
            Taul.Fields("BLOCK") = Joukko(i).Name
            Taul.Fields("HANDLE") = Joukko(i).Handle
            Taul.Fields("TEXT_4") = ATTRIB(j).TextString
            Taul.Fields("PAGE") = Mid(oDOC.Name, Len(oDOC.Name) - 6, 3)
            Taul.Update
            Exit For
          End If
        Next j
      End If
    Next i
    Set Joukko = Nothing
    oDOC.Close False
  End If
 Tied = Dir
Loop
End Sub

Private Sub Command152_Click()
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
Set oDOC = oACAD.Documents.Open(BasePath & IndexBase & ".dwg")

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

Private Sub Command157_Click()

End Sub

Private Sub Command7_Click()
'Avaa uusi pohjakuva
  Set oDOC = oACAD.Documents.Open(BasePath & BaseDWG1 & ".dwg", True)
  oACAD.ZoomExtents
  AppActivate oACAD.Caption
  
End Sub

Private Sub Command89_Click()
'Tyhjenn� IN
  Tyhjenna "IN"
End Sub

Private Sub Command90_Click()
'Tyhjenn� OUT
  Tyhjenna "OUT"
End Sub


Private Sub CPisteet_Click()
Dim gPoint As Variant
Dim iPoint(2) As Double
Dim BName As String
Dim Piste As String
Dim PointBlock As AcadBlockReference
On Error GoTo Virhe 'T�ytyy olla, jos GetPoint ei toimikaan
  
  AppActivate oACAD.Caption
  'oDOC.SendCommand Chr(27) & Chr(27) 'Varmistetaan ett� ollaan poistuttu komennosta
  'K�sket��n klikata jotain pistett�
Virhe:

  gPoint = oDOC.Utility.GetPoint(, "Set start point")
  Err.Clear
  On Error GoTo 0
  iPoint(0) = gPoint(0) + 0.00001
  iPoint(1) = gPoint(1) + 0.00001
'---------------------------------------------------------
  Piste = DoubleToPoint(CDbl(iPoint(0))) & "," & DoubleToPoint(CDbl(iPoint(1)))
  BName = BlockPath & "POINT.DWG"
  'oDOC.SendCommand Chr(27) & Chr(27) 'Varmistetaan ett� ollaan poistuttu komennosta
  Set PointBlock = oDOC.ModelSpace.InsertBlock(iPoint, BName, 1, 1, 1, 0)
  oDOC.SendCommand "(command ""move"" ""last"" """" """ & Piste & """)" & vbCr
  With PointBlock
    If .InsertionPoint(0) = iPoint(0) And .InsertionPoint(1) = iPoint(1) Then
      .Delete
    End If
  End With
  Set PointBlock = Nothing
  AppActivate CurrentDb.Properties("AppTitle").VALUE
End Sub

Private Sub Ctl__Click()

End Sub

Private Sub CViivanPiirto_Click()
Dim EkaViiva As AcadLine
Dim ViimViiva As AcadLine
Const Pi = 3.14159265358979

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
  
  AppActivate oACAD.Caption
'  oDOC.SendCommand Chr(27) & Chr(27) 'Varmistetaan ett� ollaan poistuttu komennosta
  
  
'  For i = 0 To oDOC.Layers.Count - 1
'    If oDOC.Layers(i).Name = "CANVAS" Then
'      oDOC.ActiveLayer = oDOC.Layers(i)
'      Exit For
'    End If
'  Next i
  Piste = oDOC.Utility.GetPoint(, "Set start point")
  If Err = 0 Then
    sPoint(0) = Piste(0)
    sPoint(1) = Piste(1)
    Do
      Piste = oDOC.Utility.GetPoint(sPoint, "Set end point or end drawing")
      If Err <> 0 Then
        If ViimViiva Is Nothing Then 'Lopetettiin ennen kuin yht�k��n viiva oli piirretty
        Else 'Lopetettiin, piirret��n viivan p��
        '------- VIIVJEN P��T ---------
        
        ViivaPaa ALKU, BNameA, LyhennysA
        ViivaPaa LOPPU, BNameL, LyhennysL

        If ALKU <> 4 Then 'Piirret��n alkuun joku symboli.
            With EkaViiva
              iPoint(0) = .StartPoint(0)
              iPoint(1) = .StartPoint(1)
              Rotation = .Angle + Pi
              LoppuPiste(0) = .StartPoint(0) + LyhennysA * Cos(Rotation - Pi)
              LoppuPiste(1) = .StartPoint(1) + LyhennysA * Sin(Rotation - Pi)
              .StartPoint = LoppuPiste
            End With
            oDOC.ModelSpace.InsertBlock iPoint, BNameA, 1, 1, 1, Rotation
        End If
        
        If LOPPU <> 4 Then 'Piirret��n loppuun joku symboli
            With ViimViiva
              iPoint(0) = .EndPoint(0)
              iPoint(1) = .EndPoint(1)
              Rotation = .Angle
              LoppuPiste(0) = .EndPoint(0) - LyhennysL * Cos(Rotation)
              LoppuPiste(1) = .EndPoint(1) - LyhennysL * Sin(Rotation)
              .EndPoint = LoppuPiste
            End With
            oDOC.ModelSpace.InsertBlock iPoint, BNameL, 1, 1, 1, Rotation
        End If
        Set EkaViiva = Nothing
        Set ViimViiva = Nothing
        
        '------- VIIVAN P�� ---------
        End If
        AppActivate CurrentDb.Properties("AppTitle").VALUE
        Exit Do
      Else
        Set ViimViiva = oDOC.ModelSpace.AddLine(sPoint, Piste)
        If EkaViiva Is Nothing Then
          Set EkaViiva = ViimViiva
        End If
        sPoint(0) = Piste(0)
        sPoint(1) = Piste(1)
      End If
    Loop
  Else 'Keskeytettiin heti
    AppActivate CurrentDb.Properties("AppTitle").VALUE
  End If
  Err.Clear

End Sub
Private Sub ViivaPaa(Merkki As Integer, Blokki As String, Lyhennys As Double)
If Merkki = 1 Then 'Nuoli
  Blokki = BlockPath & "ARROW.DWG"
  Lyhennys = 7.5
ElseIf Merkki = 2 Then 'Nuoli ja Not merkki
  Blokki = BlockPath & "NOTARROW.DWG"
  Lyhennys = 7.5
ElseIf ALKU.VALUE = 3 Then 'Piste
  Blokki = BlockPath & "POINT.DWG"
  Lyhennys = 0
Else
  Blokki = ""
  Lyhennys = 0
End If
End Sub

Private Sub Form_Close()
'Vapautetaan objektit ennen sulkemista
  Set oACAD = Nothing
  Set oDOC = Nothing
End Sub

Private Sub Form_Current()

End Sub

Private Sub Form_Load()

'Yhdistet��n ensin AutoCADiin
'Otetaan yhteys AutoCADiin
  On Error Resume Next
  Set oACAD = GetObject(, "AutoCAD.Application") 'Koitetaan yhdist�� AutoCADiin
  If Err <> 0 Then 'K�ynniss� olevaa AutoCADi� ei l�ytynyt
    
    Set oACAD = CreateObject("AutoCad.Application")
    
    If oACAD Is Nothing Then
        MsgBox "K�ynniss� olevaa AutoCADi� ei l�ytynyt!", vbCritical, "Virhe!"
        Set oACAD = Nothing
        Unload Me
        Exit Sub
    End If
    
  End If
  'On Error GoTo 0
  Set oDOC = oACAD.ActiveDocument
  If Err <> 0 Then
    Set oDOC = oACAD.Activate
    oDOC.Application.Visible = True
    oDOC.Application.WindowState = acMax
  End If
  On Error GoTo 0
  'Haetaan asetukset
  BaseDWG1 = DLookup("[VALUE]", "SETTINGS", "[SETTING]='BASEDWG1'")
  BaseDWG2 = DLookup("[VALUE]", "SETTINGS", "[SETTING]='BASEDWG2'")
  BasePath = DLookup("[VALUE]", "SETTINGS", "[SETTING]='BASEPATH'")
  If Right(BasePath, 1) <> "\" Then BasePath = BasePath & "\"
  BlockPath = DLookup("[VALUE]", "SETTINGS", "[SETTING]='BLOCKPATH'")
  If Right(BlockPath, 1) <> "\" Then BlockPath = BlockPath & "\"
  SavePath = DLookup("[VALUE]", "SETTINGS", "[SETTING]='SAVEPATH'")
  If Right(SavePath, 1) <> "\" Then SavePath = SavePath & "\"
   
  
End Sub











Private Sub Lehdet_Change()
  'If Lehdet <> 3 Then
  If Lehdet = 2 Then
    Blokki.RowSource = "SELECT SETTINGS.VALUE, ATTRIB FROM SETTINGS WHERE (((SETTINGS.SETTING)='BLOCK_S'));"
    Else
    Blokki.RowSource = "SELECT SETTINGS.VALUE, ATTRIB FROM SETTINGS WHERE (((SETTINGS.SETTING)='BLOCK'));"
  End If
  If Lehdet = 0 Then 'Loop
    Lehdet_Sub0.Form.VaihdaTiedot
  ElseIf Lehdet = 1 Then 'Motor
    Lehdet_sub1.Form.VaihdaTiedot
  ElseIf Lehdet = 2 Then 'Sequense
    Lehdet_Sub2.Form.VaihdaTiedot
    Controls("Text_2").VALUE = vbNullString
  ElseIf Lehdet = 3 Then 'Sequense
    TyhjennaKentat
    Blokki.RowSource = "SELECT SETTINGS.VALUE, ATTRIB FROM SETTINGS WHERE (((SETTINGS.SETTING)='BLOCK_O')); "
  End If
  Blokki.Requery
  Blokki.VALUE = ""
  TKuva.VALUE = Nothing
End Sub

Private Sub Lisays_Click()
Dim BName As String
Dim iPoint(2) As Double
Dim gPoint As Variant
Dim Piste As String
On Error GoTo Virhe
'Liitet��n ensin blokki kuvaan
  Set oDOC = oACAD.ActiveDocument
  
  If Dir(BlockPath & Blokki.VALUE & ".dwg") = "" Then
    MsgBox "Insertoitavaa blokkia ei l�ytynyt", vbCritical, "Insertoi"
    Exit Sub
  End If
  
  AppActivate oACAD.Caption
  oDOC.SendCommand Chr(27) & Chr(27) 'Varmistetaan ett� ollaan poistuttu komennosta
  
'AutoCAD etsii itse pisteen
'K�sket��n klikata jotain pistett�
Virhe:
  gPoint = oDOC.Utility.GetPoint(, "Set start point")
  Err.Clear
  On Error GoTo 0
  iPoint(0) = gPoint(0) + 0.00001
  iPoint(1) = gPoint(1) + 0.00001
'---------------------------------------------------------
  Piste = DoubleToPoint(CDbl(iPoint(0))) & "," & DoubleToPoint(CDbl(iPoint(1)))
  BName = BlockPath & Blokki & ".dwg"
  Set oBlock = oDOC.ModelSpace.InsertBlock(iPoint, BName, 1, 1, 1, 0)
  'Kirjoitetaan arvot textboxeista
  KirjoitaAttribuutit
  'Annetaan komentorivikomento jolla saadaan siirretty� blokkia haluttuun paikkaan
  oDOC.SendCommand "(command ""move"" ""last"" """" """ & Piste & """)" & vbCr
  With oBlock
    If .InsertionPoint(0) = iPoint(0) And .InsertionPoint(1) = iPoint(1) Then
      .Delete
    Else 'Kirjoitetaan tieto insertoinnista taulukkoon
      If Lehdet = 0 Then 'Jos insertoitiin joku muu kuin yleisblokki
        Lehdet_Sub0.Form.USED = True
        Lehdet_Sub0.Form.Refresh 'Requery
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
  AppActivate CurrentDb.Properties("AppTitle").VALUE

End Sub
Private Function DoubleToPoint(ByVal Piste As Double) As String
Dim APU As String
Dim Osoitin As Integer
  APU = CStr(Piste)
  Osoitin = InStr(APU, ",")
  If Osoitin = 0 Then
    DoubleToPoint = APU
  Else
    DoubleToPoint = Left(APU, Osoitin - 1) & "." & Mid(APU, Osoitin + 1)
  End If
End Function
Private Function Arvo(Tieto As Variant) As String
  If IsNull(Tieto) Or Tieto = vbNullString Then
    Arvo = vbNullChar
  Else
    Arvo = Tieto
  End If
End Function
Private Sub TaytaTextBoxit(oAttrib As Variant)
Dim i As Integer
'Tyhjennet��n ensin kaikki textBoxit
  TyhjennaKentat
  For i = 0 To UBound(oAttrib)  'K�yd��n l�pi kaikki blokin attribuutit
    With oAttrib(i)
      Select Case Left(UCase(.TagString), Len(.TagString) - 1)
        Case "TXT", "TEXT_"
          Controls("Text_" & Right(.TagString, 1)).VALUE = .TextString
        Case "IN_"
          Controls("T_IN" & Right(.TagString, 1)).VALUE = .TextString
        Case "OUT_"
          Controls("T_OUT" & Right(.TagString, 1)).VALUE = .TextString
       End Select
    End With
  Next i
  
End Sub
Public Sub TyhjennaKentat()
Dim i As Integer
  Controls("Text_1").VALUE = vbNullString
  If Controls("Text_2").VALUE = "RUN" Then
  Controls("Text_2").VALUE = vbNullString
  Else
  End If
  Controls("Text_3").VALUE = vbNullString
  Controls("Text_4").VALUE = vbNullString
  Controls("Text_5").VALUE = vbNullString
End Sub
Public Sub Tyhjenna(Tyyppi As String)
Dim i As Integer
  For i = 1 To 5
    Controls("T_" & Tyyppi & i).VALUE = vbNullString
  Next i
End Sub




Private Sub MOTOR_Click()
T_IN1.VALUE = ""
T_IN2.VALUE = "REN"
T_IN3.VALUE = ""
T_IN4.VALUE = ""
T_OUT1.VALUE = ""
T_OUT2.VALUE = "RUN"
T_OUT3.VALUE = ""
T_OUT4.VALUE = ""
Blokki.VALUE = "INTC01F"
End Sub

Private Sub mS_m_Click()
Text_5.VALUE = Text_5.VALUE & " mS/m"
End Sub

Private Sub N_Click()
T_IN1.VALUE = ""
T_IN2.VALUE = ""
T_IN3.VALUE = ""
T_IN4.VALUE = ""
T_OUT1.VALUE = "N"
T_OUT2.VALUE = ""
T_OUT3.VALUE = ""
T_OUT4.VALUE = ""
Blokki.VALUE = "INTC02"
End Sub

Private Sub NH_Click()
T_IN1.VALUE = ""
T_IN2.VALUE = ""
T_IN3.VALUE = ""
T_IN4.VALUE = ""
T_OUT1.VALUE = "N"
T_OUT2.VALUE = ""
T_OUT3.VALUE = ""
T_OUT4.VALUE = "H"
Blokki.VALUE = "INTC02F"
End Sub



Private Sub NL_Click()
T_IN1.VALUE = ""
T_IN2.VALUE = ""
T_IN3.VALUE = ""
T_IN4.VALUE = ""
T_OUT1.VALUE = "N"
T_OUT2.VALUE = ""
T_OUT3.VALUE = ""
T_OUT4.VALUE = "L"
Blokki.VALUE = "INTC02F"
End Sub

Private Sub OEN_Click()
T_IN1.VALUE = ""
T_IN2.VALUE = "OEN"
T_IN3.VALUE = ""
T_IN4.VALUE = ""
T_OUT1.VALUE = ""
T_OUT2.VALUE = ""
T_OUT3.VALUE = ""
T_OUT4.VALUE = ""
Blokki.VALUE = "INTC01F"
End Sub

Private Sub OK_Click()
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



Private Sub OPEN_LIMIT_T_Click()
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
Private Sub CLOSE_LIMIT_T_Click()
T_IN1.VALUE = ""
T_IN2.VALUE = ""
T_IN3.VALUE = ""
T_IN4.VALUE = ""
T_OUT1.VALUE = ""
T_OUT2.VALUE = ""
T_OUT3.VALUE = ""
T_OUT4.VALUE = "CLOSE"
Blokki.VALUE = "To_Page_Acc"
End Sub
Private Sub OPENCLOSE_Click()
T_IN1.VALUE = T_IN1.VALUE
T_IN2.VALUE = T_IN2.VALUE
T_IN3.VALUE = T_IN3.VALUE
T_IN4.VALUE = T_IN4.VALUE
T_OUT1.VALUE = "O"
T_OUT2.VALUE = ""
T_OUT3.VALUE = ""
T_OUT4.VALUE = "C"
Blokki.VALUE = "INTC01F"
End Sub



Private Sub �C_Click()
Text_5.VALUE = Text_5.VALUE & " �C"
End Sub

Private Sub OPEN_LIMIT_F_Click()
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

Private Sub PID_Click()
T_IN1.VALUE = ""
T_IN2.VALUE = "FC"
T_IN3.VALUE = ""
T_IN4.VALUE = ""
T_OUT1.VALUE = ""
T_OUT2.VALUE = ""
T_OUT3.VALUE = ""
T_OUT4.VALUE = ""
Blokki.VALUE = "INTC01F"
End Sub
Private Sub CLOSE_LIMIT_F_Click()
T_IN1.VALUE = ""
T_IN2.VALUE = ""
T_IN3.VALUE = ""
T_IN4.VALUE = ""
T_OUT1.VALUE = ""
T_OUT2.VALUE = ""
T_OUT3.VALUE = ""
T_OUT4.VALUE = "CLOSE"
Blokki.VALUE = "From_Page_Acc_F"
End Sub

Private Sub PID_SLAVE_Click()
T_IN1.VALUE = ""
T_IN2.VALUE = "CEN"
T_IN3.VALUE = ""
T_IN4.VALUE = ""
T_OUT1.VALUE = ""
T_OUT2.VALUE = "A/R"
T_OUT3.VALUE = ""
T_OUT4.VALUE = ""
Blokki.VALUE = "INTC01F"
End Sub



Private Sub RUN_Click()
T_IN1.VALUE = ""
T_IN2.VALUE = ""
T_IN3.VALUE = ""
T_IN4.VALUE = ""
Text_2.VALUE = "K�Y"
T_OUT1.VALUE = ""
T_OUT2.VALUE = ""
T_OUT3.VALUE = ""
T_OUT4.VALUE = ""
Blokki.VALUE = "TRANSITION2"
End Sub
Private Sub �S_cm_Click()
Text_5.VALUE = Text_5.VALUE & " �S/cm"
End Sub
Private Sub S_cm_Click()
Text_5.VALUE = Text_5.VALUE & " mS/cm"
End Sub

Private Sub l_s_Click()
Text_5.VALUE = Text_5.VALUE & " l/s"
End Sub

Private Sub l_min_Click()
Text_5.VALUE = Text_5.VALUE & " l/min"
End Sub


Private Sub m3_h_Click()
Text_5.VALUE = Text_5.VALUE & " m3/h"
End Sub
Private Sub kg_s_Click()
Text_5.VALUE = Text_5.VALUE & " kg/s"
End Sub

Private Sub RUN_FROM_Click()
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

Private Sub RUN_TO_Click()
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


Private Sub t_h_Click()
Text_5.VALUE = Text_5.VALUE & " t/h"
End Sub
Private Sub bar_Click()
Text_5.VALUE = Text_5.VALUE & " bar"
End Sub
Private Sub kPa_Click()
Text_5.VALUE = Text_5.VALUE & " kPa"
End Sub
Private Sub kgf_cm2_Click()
Text_5.VALUE = Text_5.VALUE & " kgf/cm2"
End Sub
Private Sub C_Click()
T_OUT4.VALUE = "C"
End Sub
Private Sub procent_Click()
Text_5.VALUE = Text_5.VALUE & " %"
End Sub
Private Sub START_Click()
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

Private Sub TILA1_Click()
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
Private Sub TILA2_Click()
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
Private Sub ERITILAT_Click()
T_IN1.VALUE = ""
T_IN2.VALUE = ""
T_IN3.VALUE = ""
T_IN4.VALUE = ""
Text_1.VALUE = ""
Text_2.VALUE = "H�t�seispainike 1"
Text_3.VALUE = "ERITILAT"
Text_4.VALUE = ""
T_OUT1.VALUE = ""
T_OUT2.VALUE = ""
T_OUT3.VALUE = ""
T_OUT4.VALUE = ""
Blokki.VALUE = "ERITILA"
End Sub
Private Sub T_OUT4_Click()

End Sub
