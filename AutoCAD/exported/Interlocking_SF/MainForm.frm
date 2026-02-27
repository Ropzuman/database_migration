VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} MainForm 
   Caption         =   "Interlocking"
   ClientHeight    =   8295.001
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

Private Sub FrameTIMER_Click()

End Sub

Private Sub InsertCrossPoint_Click()
Dim gPoint As Variant
Dim iPoint(2) As Double
Dim BName As String
Dim Piste As String
Dim PointBlock As AcadBlockReference
On Error GoTo Virhe 'T‰ytyy olla, jos GetPoint ei toimikaan
  MainForm.Hide
  ThisDrawing.SendCommand Chr(27) & Chr(27) 'Varmistetaan ett‰ ollaan poistuttu komennosta
'  ThisDrawing.ActiveSpace = acPaperSpace
'  ThisDrawing.ActiveSpace = acModelSpace
'  iPoint(0) = ThisDrawing.ActiveViewport.Center(0)
'  iPoint(1) = ThisDrawing.ActiveViewport.Center(1)
'K‰sket‰‰n klikata jotain pistett‰
Virhe:
  gPoint = ThisDrawing.Utility.GetPoint(, "Set start point")
  Err.Clear
  On Error GoTo 0
  iPoint(0) = gPoint(0) + 0.00001
  iPoint(1) = gPoint(1) + 0.00001
'---------------------------------------------------------
  Piste = DoubleToPoint(CDbl(iPoint(0))) & "," & DoubleToPoint(CDbl(iPoint(1)))
  BName = BLOCKPATH & "POINT.DWG"
  ThisDrawing.SendCommand Chr(27) & Chr(27) 'Varmistetaan ett‰ ollaan poistuttu komennosta
  Set PointBlock = ThisDrawing.ModelSpace.InsertBlock(iPoint, BName, 1, 1, 1, 0)
  ThisDrawing.SendCommand "(command ""move"" ""last"" """" """ & Piste & """)" & vbCr
  With PointBlock
    If .InsertionPoint(0) = iPoint(0) And .InsertionPoint(1) = iPoint(1) Then
      .Delete
    End If
  End With
  Set PointBlock = Nothing
  MainForm.Show
End Sub
Private Sub LisaaMuu_Click()
Dim gPoint As Variant
Dim iPoint(2) As Double
Dim BName As String
Dim MuuBlock As AcadBlockReference
Dim Attribuutit As Variant
Dim i As Integer
Dim Piste As String
Dim aPiste(2) As Double
Dim lPiste(2) As Double

On Error GoTo Virhe 'T‰ytyy olla, jos GetPoint ei toimikaan
  MainForm.Hide
  ThisDrawing.SendCommand Chr(27) & Chr(27) 'Varmistetaan ett‰ ollaan poistuttu komennosta
'K‰sket‰‰n klikata jotain pistett‰
Virhe:
  gPoint = ThisDrawing.Utility.GetPoint(, "Set start point")
  Err.Clear
  On Error GoTo 0
  iPoint(0) = gPoint(0) + 0.00001
  iPoint(1) = gPoint(1) + 0.00001
'---------------------------------------------------------

  Piste = DoubleToPoint(CDbl(iPoint(0))) & "," & DoubleToPoint(CDbl(iPoint(1)))
  BName = BLOCKPATH & MuutBlokit.Value & ".DWG"
  ThisDrawing.SendCommand Chr(27) & Chr(27) 'Varmistetaan ett‰ ollaan poistuttu komennosta
  Set MuuBlock = ThisDrawing.ModelSpace.InsertBlock(iPoint, BName, 1, 1, 1, 0)
  If MuuBlock.HasAttributes Then
    Attribuutit = MuuBlock.GetAttributes
    For i = 0 To UBound(Attribuutit) 'K‰yd‰‰n l‰pi kaikki attribuutit
      With Attribuutit(i)
        Select Case UCase(.TagString)
          Case "REF"
            .TextString = Arvo(TREF.Value)
          Case "INSTTAG"
            .TextString = Arvo(TREFITAG.Value)
          Case "DESC1"
            .TextString = Arvo(TREFDESC.Value)
          Case "DWG.SHEET.PAGE"
            .TextString = Arvo(TREFPAGE.Value)
          Case "TIME", "TIME1"
            .TextString = Arvo(T1.Value)
          Case "TIME2"
            .TextString = Arvo(T2.Value)
          Case "IBG"
            .TextString = Arvo(TIBG.Value)
        End Select
      End With
    Next i
  End If
  ThisDrawing.SendCommand "(command ""move"" ""last"" """" """ & Piste & """)" & vbCr
  With MuuBlock
    If .InsertionPoint(0) = iPoint(0) And .InsertionPoint(1) = iPoint(1) Then
      .Delete
    Else
      If MuutBlokit.Value = "AND" Or MuutBlokit.Value = "OR" Then
        aPiste(0) = MuuBlock.InsertionPoint(0) - 5
        aPiste(1) = MuuBlock.InsertionPoint(1) - 7.25
        lPiste(0) = MuuBlock.InsertionPoint(0) - 5
        lPiste(1) = MuuBlock.InsertionPoint(1) + 7.25
        ThisDrawing.ModelSpace.AddLine aPiste, lPiste
      End If
    End If
  End With
  Set MuuBlock = Nothing
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
On Error GoTo Virhe
'Liitet‰‰n ensin blokki kuvaan
  MainForm.Hide
  ThisDrawing.SendCommand Chr(27) & Chr(27) 'Varmistetaan ett‰ ollaan poistuttu komennosta
'AutoCAD etsii itse pisteen
'  ThisDrawing.ActiveSpace = acPaperSpace
'  ThisDrawing.ActiveSpace = acModelSpace
'  iPoint(0) = ThisDrawing.ActiveViewport.Center(0)
'  iPoint(1) = ThisDrawing.ActiveViewport.Center(1)
'K‰sket‰‰n klikata jotain pistett‰
Virhe:
  gPoint = ThisDrawing.Utility.GetPoint(, "Set start point")
  Err.Clear
  On Error GoTo 0
  iPoint(0) = gPoint(0) + 0.00001
  iPoint(1) = gPoint(1) + 0.00001
'---------------------------------------------------------
  Piste = DoubleToPoint(CDbl(iPoint(0))) & "," & DoubleToPoint(CDbl(iPoint(1)))
  BName = BLOCKPATH & Blokit.Value & ".DWG"
  Set oBlock = ThisDrawing.ModelSpace.InsertBlock(iPoint, BName, 1, 1, 1, 0)
  Attribuutit = oBlock.GetAttributes
  For i = 0 To UBound(Attribuutit) 'K‰yd‰‰n l‰pi kaikki attribuutit
    With Attribuutit(i)
      Select Case UCase(.TagString)
'        Case "POS"
'          .TextString = Arvo(TPOS.Value)
        Case "INSTTAG"
          .TextString = Arvo(TINSTTAG.Value)
        Case "DESC1"
          .TextString = Arvo(TDESC1.Value)
        Case "DESC2"
          .TextString = Arvo(TDESC2.Value)
'        Case "DESC3"
'          .TextString = Arvo(TDESC3.Value)
        Case "DB_INDEX"
          .TextString = Arvo(TDB_INDEX.Value)
'        Case "RIGHT1"
'          .TextString = Arvo(R1.Value)
'        Case "RIGHT2"
'          .TextString = Arvo(R2.Value)
'        Case "RIGHT3"
'          .TextString = Arvo(R3.Value)
'        Case "RIGHT4"
'          .TextString = Arvo(R4.Value)
'        Case "LEFT1"
'          .TextString = Arvo(L1.Value)
'        Case "LEFT2"
'          .TextString = Arvo(L2.Value)
'        Case "LEFT3"
'          .TextString = Arvo(L3.Value)
'        Case "LEFT4"
'          .TextString = Arvo(L4.Value)
'        Case "AUTO"
'          .TextString = Arvo(TAUTO.Value)
'        Case "MANUAL"
'          .TextString = Arvo(TMANUAL.Value)
      End Select
    End With
  Next i
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
  If IsNull(LoopLaitteet.Value) = False Then
    With LoopLaitteet
      TDB_INDEX.Value = .Column(0, .ListIndex)
'      TPOS.Value = .Column(1, .ListIndex)
      If VaihdaDesc.Value = True Then
        TDESC1.Value = .Column(4, .ListIndex)
        TDESC2.Value = .Column(3, .ListIndex)
      Else
        TDESC1.Value = .Column(3, .ListIndex)
        TDESC2.Value = .Column(4, .ListIndex)
      End If
'      TDESC3.Value = .Column(5, .ListIndex)
      Set TauluINSTTAG = DB.Execute("SELECT EqTAG FROM TOACADLoopEq WHERE AreaCode='" & .Column(5, .ListIndex) & "' AND LoopNo='" & .Column(6, .ListIndex) & "'")
      TINSTTAG.Clear
      If TauluINSTTAG.EOF Then
        TINSTTAG.AddItem .Column(2, .ListIndex)
      Else
        Lista = TauluINSTTAG.GetRows()
        For i = 0 To UBound(Lista, 2)
          TINSTTAG.AddItem Lista(0, i)
        Next i
      End If
      TINSTTAG.ListIndex = 0
      TINSTTAG.Value = .Column(2, .ListIndex)
    End With
  End If
End Sub


Private Sub MotorLaitteet_Change()
  If IsNull(MotorLaitteet.Value) = False Then
    With MotorLaitteet
      TDB_INDEX.Value = .Column(0, .ListIndex)
'      TPOS.Value = .Column(1, .ListIndex)
      If VaihdaDesc.Value = True Then
        TDESC2.Value = .Column(3, .ListIndex)
        TDESC1.Value = .Column(4, .ListIndex)
      Else
        TDESC1.Value = .Column(3, .ListIndex)
        TDESC2.Value = .Column(4, .ListIndex)
      End If
'      TDESC3.Value = .Column(5, .ListIndex)
      TINSTTAG.Clear
      TINSTTAG.AddItem .Column(2, .ListIndex)
    End With
    TINSTTAG.ListIndex = 0
  End If
End Sub
Private Sub MultiPage1_Change()
  NaytaBlokki
End Sub
Private Sub Muokkaa_Click() 'Laitteen muokkaus
Dim Valinta As AcadObject
Dim Point(0 To 2) As Double
Dim BlockNimi As String
Dim Oikea As Boolean
Dim Attribuutit As Variant
Dim i As Integer
Dim OK As Boolean
  MainForm.Hide
  ThisDrawing.SendCommand Chr(27) & Chr(27) 'Varmistetaan ett‰ ollaan poistuttu komennosta
  ThisDrawing.Utility.GetEntity Valinta, Point, "Choose Device..."
  If Valinta.ObjectName = "AcDbBlockReference" Then
    BlockNimi = UCase(Valinta.Name)
    For i = 1 To Blokit.ListRows  'Tarkistetaan ett‰ blokin nimi lˆytyy luettelosta
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
  MainForm.Show
End Sub
Private Sub MuutBlokit_Change()
  FrameREF.Visible = False
  FrameTIMER.Visible = False
  If MuutBlokit.Value = "REF_LOG1" Or MuutBlokit.Value = "REF_LOG2" Or MuutBlokit.Value = "REF_SAMA1" Or MuutBlokit.Value = "REF_SAMA2" Then
    FrameREF.Visible = True
  ElseIf MuutBlokit.Value = "ONOFFDELAY" Or MuutBlokit.Value = "ONDELAY" Or MuutBlokit.Value = "OFFDELAY" Or MuutBlokit.Value = "PULSE" Then
    FrameTIMER.Visible = True
  End If
On Error Resume Next
  Preview.src = BLOCKPATH & MuutBlokit.Value & ".dwg"
  Preview.Appearance = 1
  Preview.BackgroundColor = 2
  Preview.ZoomExtents
  StopButton.SetFocus
  Err.Clear
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
'  TPULSE.Value = T1.Value
End Sub
Private Sub TPULSE_Change()
  T1.Value = TPULSE.Value
End Sub

Private Sub UserForm_Initialize()
Dim Asetukset As New ADODB.Recordset
'Varmistetaan ettei AutoCADiss‰ ole komentoja p‰‰ll‰
AppActivate Application.Caption 'Vain t‰ll‰ tavoin voidaan katkaista mahdollinen Zoom komento
SendKeys "{Esc}{Esc}", True

TauluLoops.Open "TOACAD_Loops", DB ', adOpenDynamic
TauluMotors.Open "TOACAD_Motors", DB ', adOpenDynamic
Asetukset.Open "SETTINGS", DB
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
  TAUTO.Value = vbNullString
  TMANUAL.Value = vbNullString
  TPOS.Value = vbNullString
  TINSTTAG.Value = vbNullString
  TDESC1.Value = vbNullString
  TDESC2.Value = vbNullString
  TDESC3.Value = vbNullString
  TDB_INDEX.Value = vbNullString
  For i = 1 To 8
    Controls("L" & i).Value = vbNullString
    Controls("R" & i).Value = vbNullString
  Next i
  For i = 0 To UBound(Attribuutit) 'K‰yd‰‰n l‰pi kaikki blokin attribuutit
    With Attribuutit(i)
      Select Case UCase(.TagString)
        Case "POS", "INSTTAG", "DESC1", "DESC2", "DESC3", "DB_INDEX", "AUTO", "MANUAL"
          Controls("T" & .TagString).Value = .TextString
        Case "L1", "L2", "L3", "L4", "L5", "L6", "L7", "L8", "R1", "R2", "R3", "R4", "R5", "R6", "R7", "R8"
          Controls(.TagString).Value = .TextString
       End Select
    End With
  Next i
End Sub
Private Sub NaytaBlokki()
If MultiPage1.Value = 0 Then
  Blokit_Change
ElseIf MultiPage1.Value = 1 Then
  MuutBlokit_Change
End If
End Sub
Private Function DoubleToPoint(ByVal Piste As Double) As String
Dim apu As String
Dim Osoitin As Integer
  apu = CStr(Piste)
  Osoitin = InStr(apu, ",")
  If Osoitin = 0 Then
    DoubleToPoint = apu
  Else
    DoubleToPoint = Left(apu, Osoitin - 1) & "." & Mid(apu, Osoitin + 1)
  End If
End Function

Private Sub VaihdaDesc_Click()
Dim apu As String
apu = TDESC1.Value
TDESC1.Value = TDESC2.Value
TDESC2.Value = apu
End Sub
