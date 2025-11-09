Option Compare Database
Option Explicit
Dim EkaViiva As AcadLine
Dim ViimViiva As AcadLine
Dim Valmis As Boolean
Dim Sulje As Boolean
Const Pi = 3.14159265358979

Private Sub COK_Click()
Dim Lyhennys As Double
Dim iPoint(2) As Double
Dim LoppuPiste(2) As Double
Dim Rotation As Double
Dim BName As String
If Tyyppi.VALUE = 1 Then 'Nuoli
  BName = BlockPath & "ARROW.DWG"
  Lyhennys = 7.5
ElseIf Tyyppi.VALUE = 2 Then 'Nuoli ja Not merkki
  BName = BlockPath & "NOTARROW.DWG"
  Lyhennys = 7.5
Else 'Piste
  BName = BlockPath & "POINT.DWG"
  Lyhennys = 0
End If
If Sijainti.VALUE = 1 Then
  'Ei piirret� viivaa
ElseIf Sijainti.VALUE = 2 Then 'Piirret��n nuoli alkuun
  With EkaViiva
    iPoint(0) = .StartPoint(0)
    iPoint(1) = .StartPoint(1)
    Rotation = .Angle + Pi
    LoppuPiste(0) = .StartPoint(0) + Lyhennys * Cos(Rotation - Pi)
    LoppuPiste(1) = .StartPoint(1) + Lyhennys * Sin(Rotation - Pi)
    .StartPoint = LoppuPiste
  End With
  oDOC.ModelSpace.InsertBlock iPoint, BName, 1, 1, 1, Rotation
Else 'Piirret��n nuoli loppuun
  With ViimViiva
    iPoint(0) = .EndPoint(0)
    iPoint(1) = .EndPoint(1)
    Rotation = .Angle
    LoppuPiste(0) = .EndPoint(0) - Lyhennys * Cos(Rotation)
    LoppuPiste(1) = .EndPoint(1) - Lyhennys * Sin(Rotation)
    .EndPoint = LoppuPiste
  End With
  oDOC.ModelSpace.InsertBlock iPoint, BName, 1, 1, 1, Rotation
End If
Valmis = False
Set EkaViiva = Nothing
Set ViimViiva = Nothing
  DoCmd.Close acForm, "LineForm"
End Sub


Private Sub Form_Load()
'Viivan piirto aloitetaan heti kun t�m� formi aukeaa

Dim iPoint As Variant
Dim sPoint(2) As Double
Dim i As Integer
If Valmis = False Then
  On Error Resume Next
  
  AppActivate oACAD.Caption
  oDOC.SendCommand "^C^C" & vbCr 'Chr(27) & Chr(27) 'Varmistetaan ett� ollaan poistuttu komennosta
  
  
'  For i = 0 To oDOC.Layers.Count - 1
'    If oDOC.Layers(i).Name = "CANVAS" Then
'      oDOC.ActiveLayer = oDOC.Layers(i)
'      Exit For
'    End If
'  Next i
  iPoint = oDOC.Utility.GetPoint(, "Set start point")
  If Err = 0 Then
    sPoint(0) = iPoint(0)
    sPoint(1) = iPoint(1)
    Do
      iPoint = oDOC.Utility.GetPoint(sPoint, "Set end point or end drawing")
      If Err <> 0 Then
        If ViimViiva Is Nothing Then 'Lopetettiin ennen kuin yht�k��n viiva oli piirretty
          DoCmd.Close acForm
        Else
          Valmis = True
        End If
        AppActivate CurrentDb.Properties("AppTitle").VALUE
        Exit Do
      Else
        Set ViimViiva = oDOC.ModelSpace.AddLine(sPoint, iPoint)
        If EkaViiva Is Nothing Then
          Set EkaViiva = ViimViiva
        End If
        sPoint(0) = iPoint(0)
        sPoint(1) = iPoint(1)
      End If
    Loop
  Else 'Keskeytettiin heti
    DoCmd.Close acForm
    AppActivate CurrentDb.Properties("AppTitle").VALUE
  End If
  Err.Clear
End If
End Sub

Private Sub Form_Open(Cancel As Integer)
'Jos AutoCAD linkki� ei ole alustettu (t�t� formia ei avattu lukituskaavio formista, niin poistutaan)
If oACAD Is Nothing Then
  Cancel = True
End If
End Sub
