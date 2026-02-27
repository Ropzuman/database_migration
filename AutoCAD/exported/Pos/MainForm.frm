VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} MainForm 
   Caption         =   "Interlocking"
   ClientHeight    =   4605
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8280.001
   OleObjectBlob   =   "MainForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "MainForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Lisays_Click() 'Laitteen insertointi
Dim oBlock As AcadBlockReference
Dim oDot As AcadBlockReference
Dim Viiva As AcadLine
Dim iPoint As Variant
Dim ePoint As Variant
Dim BName As String
Dim Attribuutit As Variant
Dim i As Integer
Dim Piste As String
'Liitet‰‰n ensin piste
On Error Resume Next
  MainForm.Hide
  ThisDrawing.ActiveLayer = ThisDrawing.Layers(TLAYER.Value)
  iPoint = ThisDrawing.Utility.GetPoint(, "Set point")
  If Err <> 0 Then GoTo Loppu
  BName = BLOCKPATH & "DOT.DWG"
  Set oDot = ThisDrawing.ModelSpace.InsertBlock(iPoint, BName, TScale.Value, TScale.Value, 1, 0)
  If Err <> 0 Then GoTo Loppu
'Liitet‰‰n sitten positioblokki
  If OBLoop Then
    BName = BLOCKPATH & "LOOPPOS.dwg"
  Else
    BName = BLOCKPATH & "MOTPOS.dwg"
  End If
  Set oBlock = ThisDrawing.ModelSpace.InsertBlock(iPoint, BName, TScale.Value, TScale.Value, 1, 0)
  If Err <> 0 Then GoTo Loppu
'Vaihdetaan positioblokin  attribuutit
  Attribuutit = oBlock.GetAttributes
  For i = 0 To UBound(Attribuutit) 'K‰yd‰‰n l‰pi kaikki attribuutit
    With Attribuutit(i)
      Select Case UCase(.TagString)
        Case "INSTTAG"
          .TextString = Arvo(TINSTTAG.Value)
        Case "HDIR"
          .TextString = Arvo(THDIR.Value)
        Case "POWER"
          .TextString = Arvo(TPOWER.Value)
        Case "PUNIT"
          .TextString = Arvo(TPUNIT.Value)
        Case "HEIGHT"
          .TextString = Arvo(THEIGHT.Value)
        Case "HUNIT"
          .TextString = Arvo(THUNIT.Value)
        Case "NOTE1"
          .TextString = Arvo(TNOTE1.Value)
        Case "NOTE2"
          .TextString = Arvo(TNOTE2.Value)
        Case "NOTE3"
          .TextString = Arvo(TNOTE3.Value)
        Case "POSX"
          .TextString = DoubleToPoint(CDbl(iPoint(0)))
        Case "POSY"
          .TextString = DoubleToPoint(CDbl(iPoint(1)))
        Case "POS"
          .TextString = Arvo(TPOS.Value)
        Case "DB_INDEX"
          .TextString = Arvo(TDB_INDEX.Value)
      End Select
    End With
  Next i
  Piste = DoubleToPoint(CDbl(iPoint(0))) & "," & DoubleToPoint(CDbl(iPoint(1)))
'Annetaan komentorivikomento jolla saadaan siirretty‰ blokkia haluttuun paikkaan
  ThisDrawing.SendCommand "(command ""move"" ""last"" """" """ & Piste & """)" & vbCr
'Tarkistetaan ett‰ blokkia yleens‰ siirettiin
  With oBlock
    If .InsertionPoint(0) = iPoint(0) And .InsertionPoint(1) = iPoint(1) Then 'Ei siirretty joten poistetaan blokki
      .Delete
      oDot.Delete
      GoTo Loppu
    End If
  End With
'Piirret‰‰n viiva
  Do
    ePoint = ThisDrawing.Utility.GetPoint(iPoint, "Draw line")
    If Err <> 0 Then
      If Viiva Is Nothing Then 'Lopetettiin ennen kuin yht‰k‰‰n viiva oli piirretty
        oDot.Delete   'Tuhotaan piste, koska sit‰ ei tarvita
        GoTo Loppu
      End If
      Exit Do
    Else
      Set Viiva = ThisDrawing.ModelSpace.AddLine(iPoint, ePoint)
      iPoint(0) = ePoint(0)
      iPoint(1) = ePoint(1)
    End If
  Loop
'Asetetaan indeksi yht‰ suuremmaksi
If OBLoop Then
  LoopLaitteet.ListIndex = LoopLaitteet.ListIndex + 1
Else
  MotorLaitteet.ListIndex = MotorLaitteet.ListIndex + 1
End If
Loppu:
  Set oBlock = Nothing
  Set oDot = Nothing
  Err.Clear
  MainForm.Show
End Sub
Private Sub LoopLaitteet_Change()
Dim Lista As Variant
Dim i As Integer
  If IsNull(LoopLaitteet.Value) = False Then
    With LoopLaitteet
      TDB_INDEX.Value = .Column(0, .ListIndex)
      TPOS.Value = .Column(1, .ListIndex)
      TNOTE1.Value = .Column(11, .ListIndex)
      TNOTE2.Value = .Column(12, .ListIndex)
      TNOTE3.Value = .Column(13, .ListIndex)
      Set TauluINSTTAG = DB.Execute("SELECT EqTAG, NOTE_1, NOTE_2, NOTE_3 FROM TOACADLoopEq WHERE AreaCode='" & .Column(6, .ListIndex) & "' AND LoopNo='" & .Column(7, .ListIndex) & "'")
      TINSTTAG.Clear
      If TauluINSTTAG.EOF Then
        TINSTTAG.AddItem .Column(2, .ListIndex)
      Else
        TINSTTAG.Column() = TauluINSTTAG.GetRows()
      End If
      TINSTTAG.ListIndex = 0
    End With
  End If
End Sub
Private Sub MotorLaitteet_Change()
  If IsNull(MotorLaitteet.Value) = False Then
    With MotorLaitteet
      TDB_INDEX.Value = .Column(0, .ListIndex)
      TPOS.Value = .Column(1, .ListIndex)
      TPOWER.Value = .Column(10, .ListIndex)
      TNOTE1.Value = .Column(11, .ListIndex)
      TNOTE2.Value = .Column(12, .ListIndex)
      TNOTE3.Value = .Column(13, .ListIndex)
      TINSTTAG.Clear
      TINSTTAG.AddItem .Column(2, .ListIndex)
    End With
    TINSTTAG.ListIndex = 0
  End If
End Sub
Private Sub OBLoop_Click()
  LoopLaitteet.Visible = True
  MotorLaitteet.Visible = False
  MotFrame.Visible = False
  LoopLaitteet_Change
End Sub
Private Sub OBMotor_Click()
  LoopLaitteet.Visible = False
  MotorLaitteet.Visible = True
  MotFrame.Visible = True
  MotorLaitteet_Change
End Sub
Private Sub CTScale_Change()
  TScale.Value = Replace(CTScale.Value * 0.001, ".", ",")
End Sub

Private Sub TINSTTAG_Change()
  If OBLoop And TINSTTAG.ListIndex > -1 Then
    With TINSTTAG
      TNOTE1.Value = .Column(1, .ListIndex)
      TNOTE2.Value = .Column(2, .ListIndex)
      TNOTE3.Value = .Column(3, .ListIndex)
    End With
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

Private Sub UserForm_Initialize()
Dim Asetukset As New ADODB.Recordset
Dim i As Integer
TauluLoops.Open "POSIT_Loops", DB ', adOpenDynamic
TauluMotors.Open "POSIT_Motors", DB ', adOpenDynamic
Asetukset.Open "SELECT * FROM SETTINGS WHERE SETTING='BLOCKPATH'", DB
BLOCKPATH = Asetukset.Fields(1)
THDIR.AddItem "+"
THDIR.AddItem "-"
THDIR.ListIndex = 0
Set Asetukset = Nothing
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
