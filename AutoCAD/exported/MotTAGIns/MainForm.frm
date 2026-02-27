VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} MainForm 
   Caption         =   "Motor TAG Insertion"
   ClientHeight    =   3270
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6720
   OleObjectBlob   =   "MainForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "MainForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Lisays_Click()
Dim oBlock As AcadBlockReference
Dim BName As String
Dim iPoint(2) As Double
Dim gPoint As Variant
Dim Attribuutit As Variant
Dim i As Integer
Dim Piste As String
Dim aTAG As AcadAttributeReference
Dim aNAME As AcadAttributeReference
Dim aRPM As AcadAttributeReference
Dim aKW As AcadAttributeReference
Dim P1 As Variant
Dim P2 As Variant
Dim Pisteet(9) As Double
Dim ValiTeksti As AcadText
'Liitet‰‰n ensin blokki kuvaan
  MainForm.Hide
  ThisDrawing.SendCommand Chr(27) & Chr(27) 'Varmistetaan ett‰ ollaan poistuttu komennosta
  ThisDrawing.ActiveLayer = ThisDrawing.Layers(TLAYER.Value)
'Snap tila ja grid p‰‰lle
  ThisDrawing.ActiveViewport.SetSnapSpacing 1.25, 1.25
  ThisDrawing.ActiveViewport.SetGridSpacing 2.5, 2.5
  ThisDrawing.ActiveViewport.SnapOn = True
  ThisDrawing.ActiveViewport.GridOn = True
    
  gPoint = ThisDrawing.Utility.GetPoint(, "Set Motor position")
  iPoint(0) = gPoint(0) + 0.00001
  iPoint(1) = gPoint(1) + 0.00001
'---------------------------------------------------------
  Piste = DoubleToPoint(CDbl(iPoint(0))) & "," & DoubleToPoint(CDbl(iPoint(1)))
  BName = "MOTORTAG"
'Insertoidaan blokki
  Set oBlock = ThisDrawing.ModelSpace.InsertBlock(iPoint, BName, 1, 1, 1, 0)
'Pyydet‰‰n siirt‰m‰‰n blokki haluttuun paikkaa
  ThisDrawing.SendCommand "(command ""move"" ""last"" """" """ & Piste & """)" & vbCr
  With oBlock
    If .InsertionPoint(0) = iPoint(0) And .InsertionPoint(1) = iPoint(1) Then
      .Delete
      GoTo Loppu
    End If
  End With
'Haetaan Siirrett‰v‰t attribuutit muistiin
  Attribuutit = oBlock.GetAttributes
  For i = 0 To UBound(Attribuutit)
    Select Case UCase(Attribuutit(i).TagString)
      Case "TAG"
        Set aTAG = Attribuutit(i)
      Case "NAME"
        Set aNAME = Attribuutit(i)
      Case "RPM"
        Set aRPM = Attribuutit(i)
      Case "KW"
        Set aKW = Attribuutit(i)
      Case "REV"
    End Select
  Next i
'Lis‰t‰‰n "valeTAG" kuvaan
  Set ValiTeksti = ThisDrawing.ModelSpace.AddText(Arvo(Moottorit.Column(0, Moottorit.ListIndex)), aTAG.InsertionPoint, 2.5)
  Piste = DoubleToPoint(CDbl(ValiTeksti.InsertionPoint(0))) & "," & DoubleToPoint(CDbl(ValiTeksti.InsertionPoint(1)))
'Pyydet‰‰n siirt‰m‰‰n se oikeaan paikkaan
  ThisDrawing.Utility.Prompt "Set TAG postion"
  ThisDrawing.SendCommand "(command ""move"" ""last"" """" """ & Piste & """)" & vbCr
'T‰ytet‰‰n Attribuutit valituilla arvoilla
  For i = 0 To UBound(Attribuutit)
    Select Case UCase(Attribuutit(i).TagString)
      Case "TAG"
        Set aTAG = Attribuutit(i)
        aTAG.TextString = Arvo(Moottorit.Column(0, Moottorit.ListIndex))
      Case "NAME"
        Set aNAME = Attribuutit(i)
        aNAME.TextString = Arvo(Moottorit.Column(1, Moottorit.ListIndex))
      Case "RPM"
        Set aRPM = Attribuutit(i)
        aRPM.TextString = Arvo(Moottorit.Column(2, Moottorit.ListIndex))
      Case "KW"
        Set aKW = Attribuutit(i)
        aKW.TextString = Arvo(Moottorit.Column(3, Moottorit.ListIndex))
      Case "REV"
        Attribuutit(i).TextString = Arvo(Moottorit.Column(4, Moottorit.ListIndex))
      Case "FC"
        Attribuutit(i).TextString = Arvo(Moottorit.Column(5, Moottorit.ListIndex))
    End Select
  Next i
'Siirret‰‰n Tag ja muut Attribuutit Tagin alle sovitulle et‰isyydelle
  iPoint(0) = ValiTeksti.InsertionPoint(0)
  iPoint(1) = ValiTeksti.InsertionPoint(1)
  aTAG.InsertionPoint = iPoint
  iPoint(1) = iPoint(1) - 3.75
  aNAME.InsertionPoint = iPoint
  iPoint(1) = iPoint(1) - 2.5
  aKW.InsertionPoint = iPoint
  iPoint(1) = iPoint(1) - 2.5
  aRPM.InsertionPoint = iPoint
  ValiTeksti.Delete
'Piirret‰‰n laatikko tagin ymp‰rille, sek‰ viiva
  aTAG.GetBoundingBox P1, P2
  P1(0) = P1(0) - 1
  P1(1) = P1(1) - 1
  P2(0) = P2(0) + 1
  P2(1) = P2(1) + 1
  Pisteet(0) = P1(0): Pisteet(1) = P1(1)
  Pisteet(2) = P1(0): Pisteet(3) = P2(1)
  Pisteet(4) = P2(0): Pisteet(5) = P2(1)
  Pisteet(6) = P2(0): Pisteet(7) = P1(1)
  Pisteet(8) = P1(0): Pisteet(9) = P1(1)
  ThisDrawing.ModelSpace.AddLightWeightPolyline Pisteet
  If P1(1) < oBlock.InsertionPoint(1) + 5 Then 'Alapuolella
    P1(1) = P2(1)
  End If
  If P2(0) < oBlock.InsertionPoint(0) - 5 Then 'Vasemmalla
    P1(0) = P2(0)
    P2(0) = oBlock.InsertionPoint(0) - 5
  ElseIf P1(0) > oBlock.InsertionPoint(0) + 5 Then 'Oikealla puolella
    P2(0) = oBlock.InsertionPoint(0) + 5
  ElseIf P1(0) < oBlock.InsertionPoint(0) - 5 Then 'Vasemmalla puolella v‰h‰n
    P2(0) = oBlock.InsertionPoint(0) - 5
  Else
    P1(0) = P2(0)
    P2(0) = oBlock.InsertionPoint(0) + 5
  End If
  P2(1) = oBlock.InsertionPoint(1) + 5
  ThisDrawing.ModelSpace.AddLine P1, P2
'Piirret‰‰n viiva blokin ja n‰ytetyn moottorin paikan v‰lille
  If gPoint(1) > oBlock.InsertionPoint(1) Then  'Yl‰puolella
    P2(0) = oBlock.InsertionPoint(0)
    P2(1) = oBlock.InsertionPoint(1) + 10
  Else
    P2(0) = oBlock.InsertionPoint(0)
    P2(1) = oBlock.InsertionPoint(1)
  End If
  ThisDrawing.ModelSpace.AddLine gPoint, P2
  ThisDrawing.Regen acActiveViewport
Loppu:
  Set oBlock = Nothing
  MainForm.Show
End Sub

Private Sub Moottorit_Click()

End Sub

Private Sub MuokkaaKuvaa_Click()
  MainForm.Hide
  ApuForm.Show
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
Private Sub UserForm_Initialize()
Dim i As Integer
  For i = 0 To ThisDrawing.Layers.Count - 1
    TLAYER.AddItem ThisDrawing.Layers(i).Name
  Next i
  TLAYER.Value = ThisDrawing.ActiveLayer.Name
  TaytaLaitteet
End Sub
Private Sub TaytaLaitteet()
Moottorit.Clear
TauluMotors.Requery
If TauluMotors.EOF = False Then
  Moottorit.Column() = TauluMotors.GetRows()
  Moottorit.ListIndex = 0
End If
End Sub

