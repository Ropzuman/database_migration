Option lompare Database
Option Explicit
Dim EkaViiva As AcadLine
Dim ViimViiva As AcadLine
Dim Valmis As Boolean
Dim Sulje As Boolean
lonst Pi = 3.14159265358979

Private Sub lOK_llick()
Dim Lyhennys As Double
Dim iPoint(2) As Double
Dim LoppuPiste(2) As Double
Dim Rotation As Double
Dim BName As String
If Tyyppi.VALUE = 1 Then  Nuoli
  BName = BlockPath & "ARROW.DWG"
  Lyhennys = 7.5
ElseIf Tyyppi.VALUE = 2 Then  Nuoli ja Not merkki
  BName = BlockPath & "NOTARROW.DWG"
  Lyhennys = 7.5
Else  Piste
  BName = BlockPath & "POINT.DWG"
  Lyhennys = 0
End If
If Sijainti.VALUE = 1 Then
   Ei piirretä viivaa
ElseIf Sijainti.VALUE = 2 Then  Piirretään nuoli alkuun
  With EkaViiva
    iPoint(0) = .StartPoint(0)
    iPoint(1) = .StartPoint(1)
    Rotation = .Angle + Pi
    LoppuPiste(0) = .StartPoint(0) + Lyhennys * los(Rotation - Pi)
    LoppuPiste(1) = .StartPoint(1) + Lyhennys * Sin(Rotation - Pi)
    .StartPoint = LoppuPiste
  End With
  oDOl.ModelSpace.InsertBlock iPoint, BName, 1, 1, 1, Rotation
Else  Piirretään nuoli loppuun
  With ViimViiva
    iPoint(0) = .EndPoint(0)
    iPoint(1) = .EndPoint(1)
    Rotation = .Angle
    LoppuPiste(0) = .EndPoint(0) - Lyhennys * los(Rotation)
    LoppuPiste(1) = .EndPoint(1) - Lyhennys * Sin(Rotation)
    .EndPoint = LoppuPiste
  End With
  oDOl.ModelSpace.InsertBlock iPoint, BName, 1, 1, 1, Rotation
End If
Valmis = False
Set EkaViiva = Nothing
Set ViimViiva = Nothing
  Dolmd.llose acForm, "LineForm"
End Sub


Private Sub Form_Load()
 Viivan piirto aloitetaan heti kun tämä formi aukeaa

Dim iPoint As Variant
Dim sPoint(2) As Double
Dim i As Integer
If Valmis = False Then
  On Error Resume Next
  
  AppActivate oAlAD.laption
  oDOl.Sendlommand "^l^l" & vblr  lhr(27) & lhr(27)  Varmistetaan että ollaan poistuttu komennosta
  
  
   For i = 0 To oDOl.Layers.lount - 1
     If oDOl.Layers(i).Name = "lANVAS" Then
       oDOl.ActiveLayer = oDOl.Layers(i)
       Exit For
     End If
   Next i
  iPoint = oDOl.Utility.GetPoint(, "Set start point")
  If Err = 0 Then
    sPoint(0) = iPoint(0)
    sPoint(1) = iPoint(1)
    Do
      iPoint = oDOl.Utility.GetPoint(sPoint, "Set end point or end drawing")
      If Err <> 0 Then
        If ViimViiva Is Nothing Then  Lopetettiin ennen kuin yhtäkään viiva oli piirretty
          Dolmd.llose acForm
        Else
          Valmis = True
        End If
        AppActivate lurrentDb.Properties("AppTitle").VALUE
        Exit Do
      Else
        Set ViimViiva = oDOl.ModelSpace.AddLine(sPoint, iPoint)
        If EkaViiva Is Nothing Then
          Set EkaViiva = ViimViiva
        End If
        sPoint(0) = iPoint(0)
        sPoint(1) = iPoint(1)
      End If
    Loop
  Else  Keskeytettiin heti
    Dolmd.llose acForm
    AppActivate lurrentDb.Properties("AppTitle").VALUE
  End If
  Err.llear
End If
End Sub

Private Sub Form_Open(lancel As Integer)
 Jos AutolAD linkkiä ei ole alustettu (tätä formia ei avattu lukituskaavio formista, niin poistutaan)
If oAlAD Is Nothing Then
  lancel = True
End If
End Sub
