VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} MainForm 
   Caption         =   "Location"
   ClientHeight    =   4590
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8250.001
   OleObjectBlob   =   "MainForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "MainForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim TFiltter As String
Private Sub AreaJarj_Click()
  Jarjesta
End Sub
Private Sub LoopJarj_Click()
  Jarjesta
End Sub
Private Sub Filtteri_Click()
  If Filtteri.Value = True Then
    TFiltter = InputBox("Give filtter for INSTTAG. Eg. ""*3102*""." & vbCrLf & "Or press Cancel to clear filtter.", "Give Filtter", TFiltter)
  End If
  Jarjesta
End Sub
Private Sub Jarjesta()
Dim Sortti As String
Dim Missa As String
If AreaJarj.Value = True And LoopJarj.Value = True Then
  Sortti = "AreaCode, LoopNo"
ElseIf AreaJarj.Value = True Then
  Sortti = "AreaCode"
ElseIf LoopJarj.Value = True Then
  Sortti = "LoopNo"
End If
If Filtteri.Value = True Then
  If TFiltter <> "" Then
    Missa = "INSTTAG Like '" & TFiltter & "' "
    Missa = Replace(Missa, "*", "%")
    Missa = Replace(Missa, "?", "#")
  End If
End If
  TauluLoops.Sort = Sortti
  TauluLoops.Filter = Missa
  TauluMotors.Sort = Sortti
  TauluMotors.Filter = Missa
  TaytaLaitteet
End Sub
Private Sub Lisays_Click() 'Laitteen insertointi
Dim oBlock As AcadBlockReference
Dim oDot As AcadLWPolyline
Dim Viiva As AcadLine
Dim ipoint As Variant
Dim ePoint As Variant
Dim BName As String
Dim Attribuutit As Variant
Dim i As Integer
Dim Piste As String
Dim Lisatty As Boolean
'Liitet‰‰n ensin piste
On Error Resume Next
  MainForm.Hide
  ThisDrawing.ActiveLayer = ThisDrawing.Layers(TLAYER.Value)
  ipoint = ThisDrawing.Utility.GetPoint(, "Set point")
  Set oDot = AddDonut(ThisDrawing.ModelSpace, 0, 330 * TScale.Value, ipoint)
  oDot.Color = acYellow
'  Set oDot = ThisDrawing.ModelSpace.InsertBlock(ipoint, BName, TScale.Value, TScale.Value, 1, 0)
  If Err <> 0 Then GoTo Loppu
'Liitet‰‰n sitten positioblokki
  If OBLoop.Value = True Then
    BName = "LOOPBLOCK.dwg"
  Else
    BName = "MOTORBLOCK.dwg"
  End If
  Set oBlock = ThisDrawing.ModelSpace.InsertBlock(ipoint, BName, TScale.Value, TScale.Value, 1, 0)
  If Err <> 0 Then GoTo Loppu
'Vaihdetaan positioblokin  attribuutit
  Attribuutit = oBlock.GetAttributes
  For i = 0 To UBound(Attribuutit) 'K‰yd‰‰n l‰pi kaikki attribuutit
    With Attribuutit(i)
      If .TagString = "INSTTAG" Or .TagString = "MOTORTAG" Then
        .TextString = Arvo(TINSTTAG.Value)
      ElseIf .TagString = "HEIGHT" Then
        .TextString = Arvo(THEIGHT1 & THEIGHT2 & " " & THEIGHT3)
      ElseIf .TagString = "BOX" Or .TagString = "CABLE" Then
        .TextString = Arvo(TBOX.Value)
      ElseIf UCase(.TagString) = "POSX" Then
        .TextString = CStr(ipoint(0))
      ElseIf UCase(.TagString) = "POSY" Then
        .TextString = CStr(ipoint(1))
      ElseIf UCase(.TagString) = "POINTHNDL" Then
        .TextString = oDot.Handle
      ElseIf Left(.TagString, 1) = "@" Then
        .TextString = Arvo(Controls("A" & Mid(.TagString, 2)).Value)
      End If
    End With
  Next i
  Piste = DoubleToPoint(CDbl(ipoint(0))) & "," & DoubleToPoint(CDbl(ipoint(1)))
'Annetaan komentorivikomento jolla saadaan siirretty‰ blokkia haluttuun paikkaan
  ThisDrawing.SendCommand "(command ""move"" ""last"" """" """ & Piste & """)" & vbCr
'Tarkistetaan ett‰ blokkia yleens‰ siirettiin
  With oBlock
    If .InsertionPoint(0) = ipoint(0) And .InsertionPoint(1) = ipoint(1) Then 'Ei siirretty joten poistetaan blokki
      .Delete
      oDot.Delete
      GoTo Loppu
    End If
  End With
'Piirret‰‰n viiva
  Lisatty = True
  Do
    ePoint = ThisDrawing.Utility.GetPoint(ipoint, "Draw line")
    If Err <> 0 Then
'      If Viiva Is Nothing Then 'Lopetettiin ennen kuin yht‰k‰‰n viiva oli piirretty
        'oDot.Delete   'Tuhotaan piste, koska sit‰ ei tarvita
'        GoTo Loppu
'      End If
      Err.Clear
      Exit Do
    Else
      Set Viiva = ThisDrawing.ModelSpace.AddLine(ipoint, ePoint)
      Viiva.Color = acYellow
      ipoint(0) = ePoint(0)
      ipoint(1) = ePoint(1)
    End If
  Loop
'Asetetaan indeksi yht‰ suuremmaksi
On Error GoTo 0
If Lisatty Then
  TauluINSTTAG.AddNew
    TauluINSTTAG.Fields("INSTTAG").Value = TINSTTAG.Value
    TauluINSTTAG.Fields("DWG").Value = ThisDrawing.Name
  TauluINSTTAG.Update
  TaytaLaitteet
End If
Loppu:
  Set oBlock = Nothing
  Set oDot = Nothing
  Err.Clear
  MainForm.Show
End Sub
Private Sub Lisays_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
  MsgBox KeyAscii
End Sub
Private Sub LoopLaitteet_Change()
Dim Lista As Variant
Dim i As Integer
  With LoopLaitteet
    If IsNull(.Value) = False Then
      If .ListIndex >= 0 Then
        TINSTTAG.Value = .Column(0, .ListIndex)
      End If
    End If
  End With
End Sub
Private Sub MotorLaitteet_Change()
  If IsNull(MotorLaitteet.Value) = False Then
    With MotorLaitteet
      If .ListIndex >= 0 Then
        TINSTTAG.Value = .Column(0, .ListIndex)
      End If
    End With
  End If
End Sub

Private Sub Nakymat_Change()
If Nakymat.Value <> "" Then
  Dim XDataOut As Variant
  Dim XTypeOut As Variant
  Dim i As Integer
  ThisDrawing.Views(Nakymat.Value).GetXData "POSITIONING", XTypeOut, XDataOut
  If IsEmpty(XTypeOut) = False Then
    For i = LBound(XTypeOut) To UBound(XTypeOut)
      If XTypeOut(i) = 1040 Then
        VOffset.Value = Replace(XDataOut(i), ".", ",")
      End If
    Next i
  Else
    VOffset.Value = ""
  End If
End If
End Sub

Private Sub OBLoop_Click()
  LoopLaitteet.Visible = True
  MotorLaitteet.Visible = False
  LBoxCabel.Caption = "Box:"
  LoopLaitteet_Change
 
End Sub
Private Sub OBMotor_Click()
  LoopLaitteet.Visible = False
  MotorLaitteet.Visible = True
  LBoxCabel.Caption = "Cabel:"
  MotorLaitteet_Change
End Sub
Private Sub CTScale_Change()
  TScale.Value = Replace(CTScale.Value * 0.001, ".", ",")
End Sub
Private Sub PoimiKork_Click()
Dim Piste As Variant
Dim Nakyma As AcadView
Dim ViewportObj As AcadViewport
Set ViewportObj = ThisDrawing.ActiveViewport

If VOffset.Value = "" Then
  MsgBox "Set offset first!", vbInformation, "Pick up point"
Else
  'Piilotetaan formi jotta p‰‰st‰‰n k‰siksi AutoCADiin
  MainForm.Hide
  'Talletetaan nykyinen n‰kym‰
  ThisDrawing.SendCommand "-view s apu_curv" & vbCrLf
  'Zoomatan haluttuun n‰kym‰‰n
  Set Nakyma = ThisDrawing.Views(Nakymat.Value)
  ViewportObj.SetView Nakyma
  ThisDrawing.ActiveViewport = ViewportObj
  On Error GoTo Virhe
  'Pyydet‰‰n poimimaan piste
  Piste = ThisDrawing.Utility.GetPoint(, "Pick up point")
'Asetetaan korkeus poimitun pisteen mukaan
  THEIGHT2.Value = Round((VOffset.Value + Piste(1) * CDbl(Replace(TRatio.Value, ".", ","))), 2)
Virhe:
'Zoomataan aloitusn‰kym‰‰n
  Set Nakyma = ThisDrawing.Views("apu_curv")
  ViewportObj.SetView Nakyma
  ThisDrawing.ActiveViewport = ViewportObj
  Nakyma.Delete
  Err.Clear
  MainForm.Show
End If
Set ViewportObj = Nothing
Set Nakyma = Nothing
End Sub
Private Sub TRatio_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
Dim xtype(0 To 1) As Integer
Dim XData(0 To 1) As Variant
Dim Ohjelma As AcadRegisteredApplication
  
  Dim Ratio As String
  Dim RatioArvo As Double
  Ratio = TRatio.Value
'Pyydet‰‰n ratio k‰ytt‰j‰lt‰
  Ratio = InputBox("Give Ratio for current drawing.", "Set Ratio", Ratio)
  If Ratio <> "" Then
    RatioArvo = CDbl(Replace(Ratio, ".", ","))
    TRatio.Value = CStr(OffsetRatio)
    xtype(0) = 1001: XData(0) = "POSITIONING"
    xtype(1) = 1040: XData(1) = RatioArvo
    On Error GoTo Virhe
'Haetaan ohjelma, johon tieto liitet‰‰n
    Set Ohjelma = ThisDrawing.RegisteredApplications("POSITIONING")
Virhe:
    If Err <> 0 Then 'Ohjelmaa ei ollut, joten lis‰t‰‰n se
      Err.Clear
      Set Ohjelma = ThisDrawing.RegisteredApplications.Add("POSITIONING")
    End If
'Talletaan Ration kuvaan
    Ohjelma.SetXData xtype, XData
    TRatio.Value = CStr(RatioArvo)
  End If
End Sub
Private Sub TScale_AfterUpdate()
  If TScale.Value = 0 Or TScale.Value = "" Then
    TScale.Value = 1
  ElseIf CDbl(TScale.Value) > 4 Then
    TScale.Value = 4
  End If
  CTScale.Value = TScale.Value * 1000
End Sub
Private Sub TScale_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
  If KeyAscii > 47 And KeyAscii < 58 Then
  ElseIf KeyAscii = 44 Or 46 Then   'pilkku = 44, piste=46
    If InStr(TScale.Value, ",") Then
      KeyAscii = 0
    Else
      KeyAscii = 44
    End If
  Else
    KeyAscii = 0
  End If
End Sub
Private Sub UserForm_Activate()
Dim i As Integer
If ThisDrawing.Views.Count > Nakymat.ListCount Then
  Nakymat.Clear
  For i = 0 To ThisDrawing.Views.Count - 1
    Nakymat.AddItem ThisDrawing.Views(i).Name
  Next i
  Nakymat.ListIndex = 0
End If
End Sub
Private Sub UserForm_Initialize()
Dim i As Integer
Dim Ohjelma As AcadRegisteredApplication
Dim XDataType As Variant
Dim XDataValue As Variant
  On Error GoTo Virhe
  Set Ohjelma = ThisDrawing.RegisteredApplications("POSITIONING")
  Ohjelma.GetXData "POSITIONING", XDataType, XDataValue
  TRatio.Value = Replace(XDataValue(1), ".", ",")
  Set Ohjelma = Nothing
  GoTo OK
Virhe:
  Err.Clear
  TRatio.Value = "1"
OK:
  TauluLoops.Open "SELECT INSTTAG, DESCRIPTION, DWG, AreaCode, LoopNo FROM _ForAcadLoops", DB ', adOpenDynamic
  TauluMotors.Open "SELECT  INSTTAG, DESCRIPTION, DWG, AreaCode, LoopNo FROM _ForAcadMotors", DB ', adOpenDynamic
  TauluINSTTAG.Open "_InsertedToAcad", DB, adOpenStatic, adLockOptimistic
  For i = 0 To ThisDrawing.Layers.Count - 1
    TLAYER.AddItem ThisDrawing.Layers(i).Name
  Next i
  TLAYER.Value = ThisDrawing.ActiveLayer.Name
  TaytaLaitteet
End Sub
Private Sub StopButton_Click() 'Ohjelman suorituksen keskeytys
  Unload Me
End Sub
Private Sub UserForm_Terminate()
  Set DB = Nothing
  Set TauluLoops = Nothing
  Set TauluMotors = Nothing
  Set TauluINSTTAG = Nothing
End Sub
Private Sub TaytaLaitteet()
LoopLaitteet.Clear
MotorLaitteet.Clear
TauluMotors.Requery
TauluLoops.Requery
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
Private Function AddDonut(space As AcadBlock, inRad As Double, outRad As Double, cenPt As Variant) As AcadLWPolyline

    Dim width As Double, radius As Double
    Dim tmp, v(3) As Double, pl As AcadLWPolyline
    Dim PI As Double
    
    PI = Atn(1) * 4
    width = (outRad - inRad) / 2
    radius = (inRad + width) / 2
    tmp = ThisDrawing.Utility.PolarPoint(cenPt, PI, radius)
    v(0) = tmp(0): v(1) = tmp(1)
    tmp = ThisDrawing.Utility.PolarPoint(cenPt, 0, radius)
    v(2) = tmp(0): v(3) = tmp(1)
    Set pl = space.AddLightWeightPolyline(v)
    With pl
        .Closed = True
        .SetWidth 0, width, width
        .SetBulge 0, -1
        .SetWidth 1, width, width
        .SetBulge 1, -1
    End With
    
    Set AddDonut = pl

End Function
Private Sub VOffset_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
Dim xtype(0 To 1) As Integer
Dim XData(0 To 1) As Variant

If Nakymat.Value <> "" Then
  Dim Offset As String
  Dim OffsetArvo As Double
  Offset = VOffset.Value
  Offset = InputBox("Give Offset for the selected view.", "Set Offset", Offset)
  If Offset <> "" Then
    OffsetArvo = CDbl(Replace(Offset, ".", ","))
    VOffset.Value = CStr(OffsetArvo)
    xtype(0) = 1001: XData(0) = "POSITIONING"
    xtype(1) = 1040: XData(1) = OffsetArvo
    ThisDrawing.Views(Nakymat.Value).SetXData xtype, XData
  End If
Else
  MsgBox "Select named view first!", vbInformation, "Set Offset"
End If
End Sub
