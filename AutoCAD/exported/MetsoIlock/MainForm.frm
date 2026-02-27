VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} MainForm 
   Caption         =   "Interlocking"
   ClientHeight    =   7905
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8295.001
   OleObjectBlob   =   "MainForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "MainForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim BASEPATH  As String
Dim SAVEPATH As String
Dim BASEDWG   As String
Dim BHNDL As String
Dim EkaPoint(2) As Double
Dim TokaPoint(2) As Double

Private Sub Blokit_Change()
On Error Resume Next
  Preview.src = BLOCKPATH & Blokit.Value & ".dwg"
  Preview.Appearance = 1
  Preview.BackgroundColor = 2
  Preview.ZoomExtents
  Err.Clear
End Sub
Private Sub CommandButton3_Click()
  Application.Documents.Open BASEPATH & BASEDWG & ".dwg", True
  Application.ZoomExtents
End Sub
Private Sub DrawLine_Click()
  MainForm.Hide
  LineForm.Show
End Sub
Private Sub InsertCrossPoint_Click()
Dim gPoint As Variant
Dim iPoint(2) As Double
Dim BName As String
Dim Piste As String
Dim PointBlock As AcadBlockReference
  MainForm.Hide
  ThisDrawing.SendCommand Chr(27) & Chr(27) 'Varmistetaan ett‰ ollaan poistuttu komennosta

'K‰sket‰‰n klikata jotain pistett‰
  gPoint = ThisDrawing.Utility.GetPoint(, "Set point")
  BName = BLOCKPATH & "POINT.DWG"
  ThisDrawing.ModelSpace.InsertBlock gPoint, BName, 1, 1, 1, 0
  MainForm.Show
End Sub
Private Sub Lisays_Click() 'Laitteen insertointi
Dim oBlock As AcadBlockReference
Dim BName As String
Dim iPoint(2) As Double
Dim gPoint As Variant
Dim Attribuutit As Variant
Dim i As Integer
Dim Piste As String
'Liitet‰‰n ensin blokki kuvaan
  MainForm.Hide
  ThisDrawing.SendCommand Chr(27) & Chr(27) 'Varmistetaan ett‰ ollaan poistuttu komennosta
'Snap tila ja grid p‰‰lle
  ThisDrawing.ActiveViewport.SetSnapSpacing 1.25, 1.25
  ThisDrawing.ActiveViewport.SetGridSpacing 2.5, 2.5
  ThisDrawing.ActiveViewport.SnapOn = True
  ThisDrawing.ActiveViewport.GridOn = True
    
  gPoint = ThisDrawing.Utility.GetPoint(, "Set start point")
  iPoint(0) = gPoint(0) + 0.00001
  iPoint(1) = gPoint(1) + 0.00001
'---------------------------------------------------------
  Piste = DoubleToPoint(CDbl(iPoint(0))) & "," & DoubleToPoint(CDbl(iPoint(1)))
  BName = BLOCKPATH & Blokit.Value & ".dwg"
  Set oBlock = ThisDrawing.ModelSpace.InsertBlock(iPoint, BName, 1, 1, 1, 0)
  Attribuutit = oBlock.GetAttributes
  TaytaAttrib Attribuutit
  'Annetaan komentorivikomento jolla saadaan siirretty‰ blokkia haluttuun paikkaan
  ThisDrawing.SendCommand "(command ""move"" ""last"" """" """ & Piste & """)" & vbCr
  With oBlock
    If .InsertionPoint(0) = iPoint(0) And .InsertionPoint(1) = iPoint(1) Then
      .Delete
    End If
  End With
  Set oBlock = Nothing
  MainForm.Show
End Sub
Private Sub LoopLaitteet_Change()
Dim Lista As Variant
Dim i As Integer
  TyhjennaBoxit
  Muokkaa.Caption = "Edit Block"
  If IsNull(LoopLaitteet.Value) = False Then
    With LoopLaitteet
      TDB_INDEX.Value = .Column(0, .ListIndex)
'      TPOS.Value = .Column(1, .ListIndex)
      TXT1.Value = .Column(3, .ListIndex)
      TXT2.Value = .Column(4, .ListIndex)
      TXT3.Value = .Column(5, .ListIndex)
      Set TauluINSTTAG = DB.Execute("SELECT EqPOS FROM TOACADLoopEq WHERE AreaCode='" & .Column(6, .ListIndex) & "' AND LoopNo='" & .Column(7, .ListIndex) & "'")
      TXT4.Clear
      If TauluINSTTAG.EOF Then
        TXT4.AddItem .Column(2, .ListIndex)
      Else
        Lista = TauluINSTTAG.GetRows()
        For i = 0 To UBound(Lista, 2)
          TXT4.AddItem Lista(0, i)
        Next i
      End If
      TXT4.ListIndex = 0
      TXT4.Value = .Column(1, .ListIndex)
    End With
  End If
End Sub
Private Sub MotorLaitteet_Change()
  TyhjennaBoxit
  Muokkaa.Caption = "Edit Block"
  If IsNull(MotorLaitteet.Value) = False Then
    With MotorLaitteet
      TDB_INDEX.Value = .Column(0, .ListIndex)
'      TPOS.Value = .Column(1, .ListIndex)
      TXT1.Value = .Column(3, .ListIndex)
      TXT2.Value = .Column(4, .ListIndex)
      TXT3.Value = .Column(5, .ListIndex)
      TXT4.Clear
      TXT4.AddItem .Column(1, .ListIndex)
    End With
    TXT4.ListIndex = 0
    TXT5.Value = vbNullString
  End If
End Sub
Private Sub MultiPage1_Change()
  NaytaBlokki
  Muokkaa.Caption = "Edit Block"
End Sub
Private Sub Muokkaa_Click() 'Laitteen muokkaus
Dim Valinta As AcadObject
Dim Point(0 To 2) As Double
Dim BlockNimi As String
Dim Oikea As Boolean
Dim Attribuutit As Variant
Dim i As Integer
Dim OK As Boolean
Dim oBlock As AcadBlockReference
  If Muokkaa.Caption = "Update Block" Then
    Muokkaa.Caption = "Edit Block"
    Set oBlock = ThisDrawing.HandleToObject(BHNDL)
    Attribuutit = oBlock.GetAttributes
    TaytaAttrib Attribuutit
    Set oBlock = Nothing
  Else
    TyhjennaBoxit
    MainForm.Hide
    ThisDrawing.SendCommand Chr(27) & Chr(27) 'Varmistetaan ett‰ ollaan poistuttu komennosta
    ThisDrawing.Utility.GetEntity Valinta, Point, "Choose Device..."
    If Valinta.ObjectName = "AcDbBlockReference" Then
      BHNDL = Valinta.Handle
      BlockNimi = UCase(Valinta.Name)
      For i = 0 To Blokit.ListCount - 1 'Tarkistetaan ett‰ blokin nimi lˆytyy luettelosta
        If Blokit.List(i) = BlockNimi Then   'Lˆyty
          Oikea = True
          Exit For
        End If
      Next i
      If Oikea Then
        If Valinta.HasAttributes Then
          OK = True
          Attribuutit = Valinta.GetAttributes
          TaytaTextBoxit Attribuutit
        End If
      End If
    End If
    If OK = False Then
      MsgBox "Error in selecting block!", vbOKOnly, "Edit Block"
    End If
    Muokkaa.Caption = "Update Block"
    MainForm.Show
  End If
End Sub

Private Sub OBLoop_Click()
  LoopLaitteet.Visible = True
  MotorLaitteet.Visible = False
  LoopLaitteet_Change
End Sub
Private Sub OBMotor_Click()
  LoopLaitteet.Visible = False
  MotorLaitteet.Visible = True
  MotorLaitteet_Change
End Sub
Private Sub T1_Change()
  TPULSE.Value = T1.Value
End Sub
Private Sub TPULSE_Change()
  T1.Value = TPULSE.Value
End Sub

Private Sub PoistaPisteet_Click()
Dim i As Integer
Dim FilterType(0) As Integer
Dim FilterData(0) As Variant
Dim Joukko As AcadSelectionSet
If MsgBox("Are you sure you want to delete all points?" & vbCrLf & "This can not be undone!", vbYesNo, "Delete All Points") = vbYes Then
  For i = 0 To ThisDrawing.SelectionSets.Count - 1
    If ThisDrawing.SelectionSets(i).Name = "SSET" Then
      ThisDrawing.SelectionSets(i).Delete
      Exit For
    End If
  Next i
  Set Joukko = ThisDrawing.SelectionSets.Add("SSET")
  FilterType(0) = 2       'ObjecName
  FilterData(0) = "POINT"
  Joukko.Select acSelectionSetAll, , , FilterType, FilterData
  Joukko.Erase
  Joukko.Delete
  Set Joukko = Nothing
End If
End Sub
Private Sub TeePisteet_Click()
Dim i As Integer, j As Integer
Dim Viivat As AcadSelectionSet
Dim Joukko As AcadSelectionSet
Dim FilterType(1) As Integer
Dim FilterData(1) As Variant
Dim Viiva As AcadLine
Dim aViiva As AcadLine
Dim BName As String
Dim PisteLoytyi As Boolean
Dim ViivaLoytyi As Boolean
Dim iPoint(2) As Double
Dim a As Integer
Application.ZoomExtents
BName = BLOCKPATH & "POINT.DWG"
ThisDrawing.ActiveLayer = ThisDrawing.Layers("0")
For i = 0 To ThisDrawing.Layers.Count - 1
  If ThisDrawing.Layers(i).Name = "CANVAS" Then
    ThisDrawing.ActiveLayer = ThisDrawing.Layers(i)
    Exit For
  End If
Next i
For i = 0 To ThisDrawing.SelectionSets.Count - 1
  If ThisDrawing.SelectionSets(i).Name = "VIIVAT" Then
    ThisDrawing.SelectionSets(i).Delete
    Exit For
  End If
Next i
For i = 0 To ThisDrawing.SelectionSets.Count - 1
  If ThisDrawing.SelectionSets(i).Name = "SSET" Then
    ThisDrawing.SelectionSets(i).Delete
    Exit For
  End If
Next i
FilterType(0) = 0      'Entity type
FilterData(0) = "Line" 'Entity type = Line
FilterType(1) = 8      'Layer Name
FilterData(1) = "0"    'Layer Name = 0
Set Viivat = ThisDrawing.SelectionSets.Add("VIIVAT")
Set Joukko = ThisDrawing.SelectionSets.Add("SSET")
Viivat.Select acSelectionSetAll, , , FilterType, FilterData
For i = 0 To Viivat.Count - 1 'Etsit‰‰n l‰pi kaikki viivat
  If Viivat(i).ObjectName = "AcDbLine" And Viivat(i).Layer <> "DEBUG" Then
    Set Viiva = Viivat(i)
'Etsit‰‰n viivan alkup‰‰st‰ riste‰vi‰ viivoja
    For a = 1 To 2
      If a = 1 Then 'Etsit‰‰n ristet‰vi‰ viivoja viivan alkup‰‰st‰
        iPoint(0) = Viiva.StartPoint(0)
        iPoint(1) = Viiva.StartPoint(1)
      Else
        iPoint(0) = Viiva.EndPoint(0)
        iPoint(1) = Viiva.EndPoint(1)
      End If
      EkaPoint(0) = iPoint(0) + 1.25
      EkaPoint(1) = iPoint(1) + 1.25
      TokaPoint(0) = iPoint(0) - 1.25
      TokaPoint(1) = iPoint(1) - 1.25
      PisteLoytyi = False
      ViivaLoytyi = False
      Joukko.Clear
      Joukko.Select acSelectionSetCrossing, EkaPoint, TokaPoint
      For j = 0 To Joukko.Count - 1
        If Joukko(j).Handle <> Viiva.Handle Then
          If Joukko(j).ObjectName = "AcDbBlockReference" Then
            PisteLoytyi = True
          ElseIf Joukko(j).ObjectName = "AcDbLine" And Joukko(j).Layer <> "DEBUG" Then
            Set aViiva = Joukko(j)
            If OnkoAlueella(aViiva.EndPoint) = False And OnkoAlueella(aViiva.StartPoint) = False Then
              ViivaLoytyi = True
            End If
          End If
        End If
      Next j
      If ViivaLoytyi And PisteLoytyi = False Then 'Riste‰v‰ viiva lˆytyi, mutta ei pistett‰ sen p‰‰st‰
        ThisDrawing.ModelSpace.InsertBlock iPoint, BName, 1, 1, 1, 0 'Lis‰t‰‰n piste
      End If
    Next a
  End If
Next i
Viivat.Delete
Joukko.Delete
Set Viiva = Nothing
Set aViiva = Nothing
Set Viivat = Nothing
Set Joukko = Nothing
End Sub
Private Sub UserForm_Initialize()
Dim Asetukset As New ADODB.Recordset
Dim i As Integer
'Varmistetaan ettei AutoCADiss‰ ole komentoja p‰‰ll‰
AppActivate Application.Caption 'Vain t‰ll‰ tavoin voidaan katkaista mahdollinen Zoom komento
SendKeys "{Esc}{Esc}", True

TauluLoops.Open "TOACAD_Loops", DB ', adOpenDynamic
TauluMotors.Open "TOACAD_Motors", DB ', adOpenDynamic
Asetukset.Open "SETTINGS", DB
MuutBlokit.AddItem "AND"
MuutBlokit.AddItem "OR"
For i = 1 To 9
  Koko.AddItem CStr(i)
Next i
Do While Not Asetukset.EOF
  Select Case UCase(Asetukset.Fields(0).Value)
    Case "BLOCK"
      Blokit.AddItem Asetukset.Fields(1).Value
    Case "BLOCK_O"
      MuutBlokit.AddItem Asetukset.Fields(1).Value
    Case "BASEPATH"
      BASEPATH = Asetukset.Fields(1).Value
    Case "BLOCKPATH"
      BLOCKPATH = Asetukset.Fields(1).Value
    Case "SAVEPATH"
      SAVEPATH = Asetukset.Fields(1).Value
    Case "BASEDWG"
      BASEDWG = Asetukset.Fields(1).Value
  End Select
  Asetukset.MoveNext
Loop
MuutBlokit.ListIndex = 0
Blokit.ListIndex = 0
Set Asetukset = Nothing
TaytaLaitteet
End Sub
Private Sub StopButton_Click()
'Ohjelman suorituksen keskeytys
  Unload Me
End Sub
Private Sub UserForm_Terminate()
  Set DB = Nothing
  Set TauluLoops = Nothing
  Set TauluMotors = Nothing
  Set TauluINSTTAG = Nothing
End Sub
Private Sub TaytaLaitteet()
Dim ApuKoe As Variant
LoopLaitteet.Clear
MotorLaitteet.Clear
If TauluMotors.EOF = False Then
  MotorLaitteet.Column() = TauluMotors.GetRows()
  MotorLaitteet.ListIndex = 0
End If
If TauluLoops.EOF = False Then
  LoopLaitteet.Column() = TauluLoops.GetRows()
  LoopLaitteet.ListIndex = 0
End If
End Sub
Private Sub CommandButton2_Click()
  MainForm.Hide
  ApuForm.Show
End Sub
Private Function Arvo(Tieto As Variant) As String
  If IsNull(Tieto) Or Tieto = vbNullString Then
    Arvo = vbNullChar
  Else
    Arvo = Tieto
  End If
End Function
Private Sub TaytaTextBoxit(Attribuutit As Variant)
Dim i As Integer
'Tyhjennet‰‰n ensin kaikki textBoxit
  TyhjennaBoxit
  For i = 0 To UBound(Attribuutit) 'K‰yd‰‰n l‰pi kaikki blokin attribuutit
    With Attribuutit(i)
      If Left(UCase(.TagString), 3) = "TXT" Or Left(UCase(.TagString), 1) = "R" Or Left(UCase(.TagString), 1) = "L" Then
          Controls(.TagString).Value = .TextString
      End If
    End With
  Next i
End Sub
Private Sub TyhjennaBoxit()
Dim i As Integer
  TXT4.Clear
  TXT1.Value = vbNullString
  TXT2.Value = vbNullString
  TXT3.Value = vbNullString
  TXT4.Value = vbNullString
  TXT5.Value = vbNullString
  TDB_INDEX.Value = vbNullString
  For i = 1 To 7
    Controls("L" & i).Value = vbNullString
    Controls("R" & i).Value = vbNullString
  Next i
End Sub
Private Sub TaytaAttrib(Attribuutit As Variant)
Dim i As Integer
  For i = 0 To UBound(Attribuutit) 'K‰yd‰‰n l‰pi kaikki attribuutit
    With Attribuutit(i)
      If Left(.TagString, 3) = "TXT" Or Left(.TagString, 1) = "R" Or Left(.TagString, 1) = "L" Then
        .TextString = Arvo(Controls(.TagString).Value)
      End If
    End With
  Next i
End Sub
'----------------------------------------------
Private Sub MuutBlokit_Change()
Dim Nimi As String
  If MuutBlokit.Value = "AND" Or MuutBlokit.Value = "OR" Then
    LKoko.Visible = True
    Koko.Visible = True
    Nimi = MuutBlokit.Value & Koko.Value
  Else
    LKoko.Visible = False
    Koko.Visible = False
    Nimi = MuutBlokit.Value
  End If
On Error Resume Next
  Preview.src = BLOCKPATH & Nimi & ".dwg"
  Preview.Appearance = 1
  Preview.BackgroundColor = 2
  Preview.ZoomExtents
  StopButton.SetFocus
  Err.Clear
End Sub
Private Sub Koko_Change()
  MuutBlokit_Change
End Sub
Private Sub LisaaMuu_Click()
Dim gPoint As Variant
Dim iPoint(2) As Double
Dim BName As String
Dim MuuBlock As AcadBlockReference
Dim Attribuutit As Variant
Dim i As Integer
Dim Piste As String
  
  MainForm.Hide
  ThisDrawing.SendCommand Chr(27) & Chr(27) 'Varmistetaan ett‰ ollaan poistuttu komennosta
'Snap tila ja grid p‰‰lle
  ThisDrawing.ActiveViewport.SetSnapSpacing 1.25, 1.25
  ThisDrawing.ActiveViewport.SetGridSpacing 2.5, 2.5
  ThisDrawing.ActiveViewport.SnapOn = True
  ThisDrawing.ActiveViewport.GridOn = True
  
  
  gPoint = ThisDrawing.Utility.GetPoint(, "Set start point")
  iPoint(0) = gPoint(0) + 0.00001
  iPoint(1) = gPoint(1) + 0.00001
'---------------------------------------------------------
  Piste = DoubleToPoint(CDbl(iPoint(0))) & "," & DoubleToPoint(CDbl(iPoint(1)))
  If MuutBlokit.Value = "AND" Or MuutBlokit.Value = "OR" Then
    BName = BLOCKPATH & MuutBlokit.Value & Koko.Value & ".DWG"
  Else
    BName = BLOCKPATH & MuutBlokit.Value & ".DWG"
  End If
  ThisDrawing.SendCommand Chr(27) & Chr(27) 'Varmistetaan ett‰ ollaan poistuttu komennosta
  Set MuuBlock = ThisDrawing.ModelSpace.InsertBlock(iPoint, BName, 1, 1, 1, 0)
  If MuuBlock.HasAttributes Then
    Attribuutit = MuuBlock.GetAttributes
    For i = 0 To UBound(Attribuutit) 'K‰yd‰‰n l‰pi kaikki attribuutit
      With Attribuutit(i)
        If Left(UCase(.TagString), 3) = "TXT" Then
            .TextString = Arvo(Controls("M" & .TagString).Value)
        End If
      End With
    Next i
  End If
  ThisDrawing.SendCommand "(command ""move"" ""last"" """" """ & Piste & """)" & vbCr
  With MuuBlock
    If .InsertionPoint(0) = iPoint(0) And .InsertionPoint(1) = iPoint(1) Then
      .Delete
    End If
  End With
  Set MuuBlock = Nothing
  MainForm.Show
End Sub
Private Sub NaytaBlokki()
If MultiPage1.Value = 0 Then
  Blokit_Change
ElseIf MultiPage1.Value = 1 Then
  MuutBlokit_Change
End If
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
Private Function OnkoAlueella(iPoint As Variant) As Boolean
Dim XOK As Boolean
Dim YOK As Boolean
  If EkaPoint(0) > TokaPoint(0) Then
    If iPoint(0) > TokaPoint(0) And iPoint(0) < EkaPoint(0) Then
      XOK = True
    End If
  Else
    If iPoint(0) < TokaPoint(0) And iPoint(0) > EkaPoint(0) Then
      XOK = True
    End If
  End If
  If EkaPoint(1) > TokaPoint(1) Then
    If iPoint(1) > TokaPoint(1) And iPoint(1) < EkaPoint(1) Then
      YOK = True
    End If
  Else
    If iPoint(1) < TokaPoint(1) And iPoint(1) > EkaPoint(1) Then
      YOK = True
    End If
  End If
  If XOK And YOK Then
    OnkoAlueella = True
  Else
    OnkoAlueella = False
  End If
End Function
