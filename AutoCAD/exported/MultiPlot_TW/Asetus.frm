VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Asetus 
   Caption         =   "Plot Info Settings"
   ClientHeight    =   3600
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7695
   OleObjectBlob   =   "Asetus.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "Asetus"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub OK_Click()
  If Paikka.ListIndex = 0 Then SPaikka = "BL"
  If Paikka.ListIndex = 1 Then SPaikka = "TL"
  If Paikka.ListIndex = 2 Then SPaikka = "TR"
  If Paikka.ListIndex = 3 Then SPaikka = "BR"
  SSuunta = Asento.Value
  SKork = CDbl(Kork.Value)
  SYlosalais = Ylosalais.Value
  SXOffset = CDbl(xoffset.Value)
  SYOffset = CDbl(yoffset.Value)
  SDate = PrintDate.Value
  SFName = PrintName.Value
  Unload Me
End Sub
Private Sub Paikka_Change()
  Esikatselu
End Sub
Private Sub Asento_Change()
  Esikatselu
End Sub
Private Sub Peruuta_Click()
  Unload Me
End Sub
Sub Esikatselu()
Dim a As String
Dim b As String
Dim c As String
  b = "p"
  vyp.Visible = False
  vypu.Visible = False
  vyv.Visible = False
  vyvu.Visible = False
  oyp.Visible = False
  oypu.Visible = False
  oyv.Visible = False
  oyvu.Visible = False
  oav.Visible = False
  oavu.Visible = False
  oap.Visible = False
  oapu.Visible = False
  vap.Visible = False
  vapu.Visible = False
  vav.Visible = False
  vavu.Visible = False
  If Ylosalais.Value = True Then c = "u"
  If Asento.Value = "Horizontal" Then b = "v"
  If Paikka.Value = "Bottom Left" Then a = "va"
  If Paikka.Value = "Top Left" Then a = "vy"
  If Paikka.Value = "Top Right" Then a = "oy"
  If Paikka.Value = "Bottom Right" Then a = "oa"
  Controls(a & b & c).Visible = True
End Sub
Private Sub Upsidedown_Click()
  Esikatselu
End Sub
Private Sub UserForm_Initialize()
Dim Leveys As Double
Dim Korkeus As Double
  Paikka.AddItem "Bottom Left"
  Paikka.AddItem "Top Left"
  Paikka.AddItem "Top Right"
  Paikka.AddItem "Bottom Right"
  Asento.AddItem "Horizontal"
  Asento.AddItem "Vertical"
  If SPaikka = "BL" Then Paikka.ListIndex = 0
  If SPaikka = "TL" Then Paikka.ListIndex = 1
  If SPaikka = "TR" Then Paikka.ListIndex = 2
  If SPaikka = "BR" Then Paikka.ListIndex = 3
  If SSuunta = "Horizontal" Then Asento.ListIndex = 0
  If SSuunta = "Vertical" Then Asento.ListIndex = 1
  Kork.Value = SKork
  Ylosalais.Value = SYlosalais
  xoffset.Value = SXOffset
  yoffset.Value = SYOffset
  PrintDate.Value = SDate
  PrintName.Value = SFName
  Esikatselu
End Sub
Private Sub xoffset_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
  If KeyAscii = 44 Or KeyAscii = 46 Then
    If InStr(xoffset.Value, ",") Then
      KeyAscii = 0
    Else
      KeyAscii = 44
    End If
  ElseIf (KeyAscii < 48 Or KeyAscii > 57) Then
    KeyAscii = 0
  End If
End Sub

Private Sub Ylosalais_Click()
  Esikatselu
End Sub

Private Sub yoffset_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
  If KeyAscii = 44 Or KeyAscii = 46 Then
    If InStr(yoffset.Value, ",") Then
      KeyAscii = 0
    Else
      KeyAscii = 44
    End If
  ElseIf (KeyAscii < 48 Or KeyAscii > 57) Then
    KeyAscii = 0
  End If
End Sub
