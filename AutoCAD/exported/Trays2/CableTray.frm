VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} CableTray 
   Caption         =   "Hyllyn piirto"
   ClientHeight    =   5160
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6780
   OleObjectBlob   =   "CableTray.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "CableTray"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Joukko() As AcadEntity
Dim hPituus As Double
Dim MerenPinta As Long
Private Sub BPiirra_Click()
Dim Komento As String
Dim i As Integer
If BPiirra.Caption = "Piirr‰" Then
  Kulmia = 0
  Piirto = True
  Me.Height = 276
  Me.Width = 105
  If Lev150.Value Then
    Lev = 150
  ElseIf Lev300.Value Then
    Lev = 300
  ElseIf Lev450.Value Then
    Lev = 450
  ElseIf Lev600.Value Then
    Lev = 600
  ElseIf Lev750.Value Then
    Lev = 750
  ElseIf Lev800.Value Then
    Lev = 800
  End If
  Komento = "MLINE ST STANDARD S " & Lev
  If OOikealle.Value = True Then
    Komento = Komento & " J T "
    Puoli = 0
  ElseIf OVasemmalle.Value = True Then
    Komento = Komento & " J B "
    Puoli = 1
  Else
    Komento = Komento & " J Z "
    Puoli = 2
  End If
  BPiirra.Caption = "OK"
  Me.Hide
  ActiveDocument.SendCommand Chr(3) & Chr(3)
  ActiveDocument.SendCommand Komento
Else
  ActiveDocument.SendCommand vbCr & Chr(3)
  BPiirra.Caption = "Piirr‰"
  Piirto = False
  If Kulmia > 0 Then
    TLopetus.Value = TAloitus.Value
  End If
  If Not EdViiva Is Nothing Then
    MuutaViiva
  End If
  Kulmia = 0
  Me.Height = 124
  Me.Width = 265
End If
End Sub
Private Sub Image1_Click()
  OOikealle.Value = True
End Sub
Private Sub Image2_Click()
  OVasemmalle.Value = True
End Sub
Private Sub Image3_Click()
  OKeskelle.Value = True
End Sub
Private Sub TAloitus_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
Dim Arvo As String
  Arvo = InputBox("Anna loppukorkeus (mm).", "Hyllyn loppukorkeus", TAloitus.Value)
  If Arvo <> "" Then
    TAloitus.Value = CLng(Arvo)
    If Kulmia > 0 Then TLopetus.Value = TAloitus.Value
    LaskeMLinePituus Abs(CLng(TAloitus.Value) - CLng(TLopetus.Value))
  End If
End Sub
Private Sub TLopetus_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
Dim Arvo As String
  Arvo = InputBox("Anna loppukorkeus (mm).", "Hyllyn loppukorkeus", TLopetus.Value)
  If Arvo <> "" Then
    TLopetus.Value = CLng(Arvo)
    If Kulmia > 0 Then TAloitus.Value = TLopetus.Value
    LaskeMLinePituus Abs(CLng(TAloitus.Value) - CLng(TLopetus.Value))
  End If
End Sub
Private Sub TArea_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
Dim Arvo As String
  Arvo = InputBox("Annan aluekoodi.", "Aluekoodi", TArea.Value)
  If Arvo <> "" Then TArea.Value = Arvo
End Sub
Private Sub TNumber_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
Dim Arvo As String
Dim Oletus As String
  Oletus = CStr(CLng(TNumber.Value) + 1)
  If Len(Oletus) = 1 Then
    Oletus = "00" & Oletus
  ElseIf Len(Oletus) = 2 Then
    Oletus = "0" & Oletus
  End If
  Arvo = InputBox("Annan numero.", "Aluekoodi", Oletus)
  If Arvo <> "" Then
    If Len(Arvo) = 1 Then
      Arvo = "00" & Arvo
    ElseIf Len(Arvo) = 2 Then
      Arvo = "0" & Arvo
    End If
    TNumber.Value = Arvo
  End If
End Sub
Private Sub UserForm_Initialize()
  Me.Height = 124
  Me.Width = 265
  TService.AddItem "A":   TService.List(0, 1) = "INSTRUMENTATION(ANALOG)"
  TService.AddItem "B":   TService.List(1, 1) = "CONTROL OR INSTR. (DIGITAL)"
  TService.AddItem "C":   TService.List(2, 1) = "<1KV POWER-LIGHTING-SERVICE"
  TService.AddItem "D":   TService.List(3, 1) = "3.3kV POWER"
  TService.AddItem "E":   TService.List(4, 1) = "6.6KV & 13.2KV POWER"
  TService.Value = "A"
  TBarrier.AddItem "0":   TBarrier.List(0, 1) = "NO BARRIER"
  TBarrier.AddItem "B":   TBarrier.List(1, 1) = "BARRIER"
  TBarrier.Value = "0"
  AsetaTaso
  PoistaTyhjatJoukot
End Sub
Private Sub UserForm_Terminate()
  Piirto = False
End Sub
Private Sub MuutaViiva()
Dim i As Long, j As Integer, ii As Integer, a As Integer

Dim Piste1(2) As Double
Dim Piste2(2) As Double
Dim Piste3(2) As Double

Dim Oik As Integer
Dim oPline As Acad3DPolyline
Dim Ryhma As AcadGroup

Dim iPoint(2) As Double
Dim iPoint2(2) As Double
Dim BName As String
Dim oBlock As AcadBlockReference
Dim Attrib As Variant

Dim ApuPiste1 As Variant
Dim ApuPiste2 As Variant
Dim ApuPiste3 As Variant
Dim ApuPiste4 As Variant
Dim ApuPisteX1 As Variant
Dim ApuPisteX2 As Variant

Dim UusiPiste As Variant
Dim Eta As Double
    
    BName = "K:\INSEDATA\Tools\ForAutoCAD\BLOCKS\TRAYINFO.dwg" 'muutettu
'    BName = "trayinfo"
    j = 1
    ReDim Joukko(0)
    Do
      Loytyi = True
      For i = 0 To ActiveDocument.Groups.Count - 1
        If ActiveDocument.Groups(i).Name = "T" & j Then
          Loytyi = False
          Exit For
        End If
      Next i
      If Loytyi Then
        Set Ryhma = ActiveDocument.Groups.Add("T" & j)
      End If
      j = j + 1
    Loop Until Loytyi
    ReDim kPisteet(UBound(EdViiva.Coordinates))
    ReDim oPisteet(UBound(EdViiva.Coordinates))
    ReDim vPisteet(UBound(EdViiva.Coordinates))
'------- Etsit‰‰n keskiviivan pisteet---------------
    If Puoli = 0 Or 1 Then
      If Puoli = 0 Then Eta = -1 * Lev / 2
      If Puoli = 1 Then Eta = Lev / 2
      For i = 0 To (UBound(EdViiva.Coordinates) - 3) Step 3
        Piste1(0) = EdViiva.Coordinates(i)
        Piste1(1) = EdViiva.Coordinates(i + 1)
        Piste2(0) = EdViiva.Coordinates(i + 3)
        Piste2(1) = EdViiva.Coordinates(i + 4)
        ApuPiste1 = ActiveDocument.Utility.PolarPoint(Piste1, ActiveDocument.Utility.AngleFromXAxis(Piste1, Piste2) + PI / 2, Eta)
        ApuPiste2 = ActiveDocument.Utility.PolarPoint(Piste2, ActiveDocument.Utility.AngleFromXAxis(Piste1, Piste2) + PI / 2, Eta)
        If i = 0 Then
          kPisteet(0) = ApuPiste1(0)
          kPisteet(1) = ApuPiste1(1)
        End If
        kPisteet(i + 3) = ApuPiste2(0)
        kPisteet(i + 4) = ApuPiste2(1)
        If UBound(EdViiva.Coordinates) > i + 5 Then
          Piste3(0) = EdViiva.Coordinates(i + 6)
          Piste3(1) = EdViiva.Coordinates(i + 7)
          'Lasketaan keskiviivan paikka
          ApuPiste3 = ActiveDocument.Utility.PolarPoint(Piste2, ActiveDocument.Utility.AngleFromXAxis(Piste2, Piste3) + PI / 2, Eta)
          ApuPiste4 = ActiveDocument.Utility.PolarPoint(Piste3, ActiveDocument.Utility.AngleFromXAxis(Piste2, Piste3) + PI / 2, Eta)
          UusiPiste = LaskeLeikkauspiste(ApuPiste1, ApuPiste2, ApuPiste3, ApuPiste4)
          kPisteet(i + 3) = UusiPiste(0)
          kPisteet(i + 4) = UusiPiste(1)
        End If
      Next i
    Else
      For i = 0 To UBound(EdViiva.Coordinates)
        kPisteet(i) = EdViiva.Coordinates(i)
      Next i
    End If
    LaskeOVPisteet
'----------------------------
    ReDim kPisteet2(UBound(kPisteet) + (UBound(kPisteet) - 5))
    ReDim oPisteet2(UBound(kPisteet) + (UBound(kPisteet) - 5))
    ReDim vPisteet2(UBound(kPisteet) + (UBound(kPisteet) - 5))
    
    For j = 0 To (UBound(kPisteet) - 3) Step 3
      kPisteet(j + 2) = CLng(TAloitus.Value)
      oPisteet(j + 2) = CLng(TAloitus.Value)
      vPisteet(j + 2) = CLng(TAloitus.Value)
      
      kPisteet(j + 5) = CLng(TLopetus.Value)
      oPisteet(j + 5) = CLng(TLopetus.Value)
      vPisteet(j + 5) = CLng(TLopetus.Value)
    Next j
    
    For i = 0 To 2
      kPisteet2(i) = kPisteet(i)
      oPisteet2(i) = oPisteet(i)
      vPisteet2(i) = vPisteet(i)
      kPisteet2(UBound(kPisteet2) - i) = kPisteet(UBound(kPisteet) - i)
      oPisteet2(UBound(oPisteet2) - i) = oPisteet(UBound(oPisteet) - i)
      vPisteet2(UBound(vPisteet2) - i) = vPisteet(UBound(vPisteet) - i)
    Next i
  
    ii = 3
    For j = 3 To (UBound(kPisteet) - 5) Step 3
      For i = 0 To 2
        kPisteet2(ii + i) = kPisteet(j + i)
        oPisteet2(ii + i) = oPisteet(j + i)
        vPisteet2(ii + i) = vPisteet(j + i)
        kPisteet2(ii + 3 + i) = kPisteet(j + i)
        oPisteet2(ii + 3 + i) = oPisteet(j + i)
        vPisteet2(ii + 3 + i) = vPisteet(j + i)
      Next i
      ii = ii + 6
    Next j
'---------------
    For j = 0 To (UBound(kPisteet) - 3) Step 3
      If KaaretYES.Value = True Then
        If UBound(kPisteet) > j + 5 Then
          TeeKaaret j / 3
        End If
      End If
      
      Piste1(0) = kPisteet2(j / 3 * 6)
      Piste1(1) = kPisteet2(j / 3 * 6 + 1)
      Piste1(2) = kPisteet2(j / 3 * 6 + 2)
      Piste2(0) = kPisteet2(j / 3 * 6 + 3)
      Piste2(1) = kPisteet2(j / 3 * 6 + 4)
      Piste2(2) = kPisteet2(j / 3 * 6 + 5)
      
      hPituus = LaskePituus(Piste1, Piste2)
     
      Set oPline = PiirraPolyLine(j / 3, 0)
      Ryhmaan oPline
      Set oPline = PiirraPolyLine(j / 3, 1)
      Ryhmaan oPline
      Set oPline = PiirraPolyLine(j / 3, 2)
      Ryhmaan oPline
      
      Maara = hPituus \ 250
      a = 1
      If UBound(kPisteet) > j + 5 Then
        a = 0
      End If
      If j = 0 Then a = 1
      For i = a To Maara
        Set oPline = PiirraPViiva(Piste1, Piste2, i * 250)
        Ryhmaan oPline
        Set oPline = PiirraPViiva(Piste1, Piste2, i * 250 + 30)
        Ryhmaan oPline
      Next i
    Next j
    
    'Blokin Paikka
    iPoint(0) = kPisteet(0) + (kPisteet(3) - kPisteet(0)) / 2
    iPoint(1) = kPisteet(1) + (kPisteet(4) - kPisteet(1)) / 2
    iPoint(2) = kPisteet(2) + (kPisteet(5) - kPisteet(2)) / 2
    iPoint2(0) = iPoint(0) + 1200
    iPoint2(1) = iPoint(1) + 1200
    iPoint2(2) = iPoint(2)
    
    Set oPline = Piirra3dPoly(iPoint, iPoint2)
    oPline.Color = acRed
    Ryhmaan oPline
    
    Set oBlock = ActiveDocument.ModelSpace.InsertBlock(iPoint2, BName, 10, 10, 1, 0)
    Attrib = oBlock.GetAttributes
    For i = 0 To UBound(Attrib)
      Select Case Attrib(i).TagString
        Case "ID"
          Attrib(i).TextString = KokoaID
        Case "ENDHEIGHTS"
          Attrib(i).TextString = CStr(CLng(TLopetus.Value) + MerenPinta)
        Case "ENDHEIGHT"
          Attrib(i).TextString = TLopetus.Value
        Case "STARTHEIGHTS"
          If TLopetus.Value <> TAloitus.Value Then
            Attrib(i).TextString = CStr(CLng(TAloitus.Value))
          End If
        Case "STARTHEIGHT"
          If TLopetus.Value <> TAloitus.Value Then
            Attrib(i).TextString = CStr(CLng(TAloitus.Value))
          End If
        Case "LENGTH"
            Attrib(i).TextString = CStr(CLng(VPituus.Value))
        Case "WIDTH"
            Attrib(i).TextString = CStr(Lev)
        Case "CORNERS"
            Attrib(i).TextString = CStr(Kulmia)
      End Select
    Next i
    Set Joukko(UBound(Joukko)) = oBlock
    Ryhma.AppendItems Joukko
    Set Ryhma = Nothing
    EdViiva.Delete
    ActiveDocument.Regen acActiveViewport
    Set EdViiva = Nothing
End Sub
Private Sub AsetaTaso()
Dim Taso As AcadLayer
Dim i As Integer
Dim Loytyi As Boolean
For i = 0 To ActiveDocument.Layers.Count - 1
  If ActiveDocument.Layers(i).Name = "CableTrays" Then
    Set Taso = ActiveDocument.Layers(i)
    Loytyi = True
    Exit For
  End If
Next i
If Loytyi = False Then
  Set Taso = ActiveDocument.Layers.Add("CableTrays")
  Taso.Color = acBlue
End If
Loytyi = False
ActiveDocument.ActiveLayer = Taso
For i = 0 To ActiveDocument.Layers.Count - 1
  If ActiveDocument.Layers(i).Name = "CableTraysCenter" Then
    Loytyi = True
    Exit For
  End If
Next i
If Loytyi = False Then
  Set Taso = ActiveDocument.Layers.Add("CableTraysCenter")
  Taso.Color = acRed
End If

End Sub
Sub Ryhmaan(Objekti As AcadEntity)
  Set Joukko(UBound(Joukko)) = Objekti
  ReDim Preserve Joukko(UBound(Joukko) + 1)
End Sub
Function PiirraPolyLine(Jar As Integer, Viiva As Integer) As Acad3DPolyline
Dim viivaPisteet(5) As Double
Dim i As Integer
If Viiva = 0 Then ' Piirret‰‰n vasen viiva
  For i = 0 To 5
    viivaPisteet(i) = vPisteet2(Jar * 6 + i)
  Next i
  Set PiirraPolyLine = ActiveDocument.ModelSpace.Add3DPoly(viivaPisteet)
ElseIf Viiva = 1 Then ' Piirret‰‰n keskiviiva
  ActiveDocument.ActiveLayer = ActiveDocument.Layers("CableTraysCenter")
  For i = 0 To 5
    viivaPisteet(i) = kPisteet2(Jar * 6 + i)
    Next i
  Set PiirraPolyLine = ActiveDocument.ModelSpace.Add3DPoly(viivaPisteet)
  ActiveDocument.ActiveLayer = ActiveDocument.Layers("CableTrays")
Else 'Piirret‰‰n Oikea viiva
  For i = 0 To 5
    viivaPisteet(i) = oPisteet2(Jar * 6 + i)
  Next i
  Set PiirraPolyLine = ActiveDocument.ModelSpace.Add3DPoly(viivaPisteet)
End If

End Function
Sub TeeKaaret(Kaari As Integer)
'T‰m‰ Piirt‰‰ halutun kaaren
Dim Piste1(2) As Double
Dim Piste2(2) As Double
Dim Piste3(2) As Double
Dim Piste4(2) As Double
Dim Kulma1 As Double

Dim aKulma As Double
Dim lKulma As Double
Dim apuKulma As Double

Dim Pituus As Double
Dim Pituus2 As Double
Dim KP As Variant
Dim oArc1 As AcadArc
Dim oArc2 As AcadArc
Dim oArc3 As AcadArc

Dim p1 As Integer
Dim p2 As Integer

Dim i As Integer

Piste1(0) = vPisteet(Kaari * 3 + 3) 'X
Piste1(1) = vPisteet(Kaari * 3 + 4) 'Y
Piste1(2) = vPisteet(Kaari * 3 + 5) 'Z

Piste2(0) = kPisteet(Kaari * 3 + 3) 'X
Piste2(1) = kPisteet(Kaari * 3 + 4) 'Y
Piste2(2) = kPisteet(Kaari * 3 + 5) 'Z

Piste3(0) = kPisteet(Kaari * 3)     'X
Piste3(1) = kPisteet(Kaari * 3 + 1) 'Y
Piste3(2) = kPisteet(Kaari * 3 + 2) 'Z

Piste4(0) = kPisteet(Kaari * 3 + 6) 'X
Piste4(1) = kPisteet(Kaari * 3 + 7) 'Y
Piste4(2) = kPisteet(Kaari * 3 + 8) 'Z

aKulma = ActiveDocument.Utility.AngleFromXAxis(Piste2, Piste3) - PI / 2
lKulma = ActiveDocument.Utility.AngleFromXAxis(Piste2, Piste4) + PI / 2

Pituus1 = LaskePituus(Piste1, Piste2)
KP = LaskeUusiPiste(Piste1, Piste2, Pituus1 * 3)

Set oArc1 = ActiveDocument.ModelSpace.AddArc(KP, Lev / 2, lKulma, aKulma)
p1 = 3
If oArc1.TotalAngle > PI Then
  oArc1.Delete
  KP = LaskeUusiPiste(Piste1, Piste2, Pituus1 * -1)
  apuKulma = lKulma
  lKulma = aKulma
  aKulma = apuKulma
  aKulma = aKulma + PI
  lKulma = lKulma + PI
  Set oArc1 = ActiveDocument.ModelSpace.AddArc(KP, Lev / 2, lKulma, aKulma)
  p1 = 0
  p2 = 3
End If
'Keskipisteen kaari
ActiveDocument.ActiveLayer = ActiveDocument.Layers("CableTraysCenter")
Set oArc2 = ActiveDocument.ModelSpace.AddArc(KP, Lev, lKulma, aKulma)
ActiveDocument.ActiveLayer = ActiveDocument.Layers("CableTrays")
Set oArc3 = ActiveDocument.ModelSpace.AddArc(KP, Lev + Lev / 2, lKulma, aKulma)

Ryhmaan oArc1
Ryhmaan oArc2
Ryhmaan oArc3

For i = 0 To 2
  kPisteet2(p1 + i + 3 + Kaari * 6) = oArc2.StartPoint(i)
  kPisteet2(p2 + i + 3 + Kaari * 6) = oArc2.EndPoint(i)

  If p1 = 3 Then
    oPisteet2(p1 + i + 3 + Kaari * 6) = oArc1.StartPoint(i)
    oPisteet2(p2 + i + 3 + Kaari * 6) = oArc1.EndPoint(i)
  
    vPisteet2(p1 + i + 3 + Kaari * 6) = oArc3.StartPoint(i)
    vPisteet2(p2 + i + 3 + Kaari * 6) = oArc3.EndPoint(i)
  Else
    vPisteet2(p1 + i + 3 + Kaari * 6) = oArc1.StartPoint(i)
    vPisteet2(p2 + i + 3 + Kaari * 6) = oArc1.EndPoint(i)
  
    oPisteet2(p1 + i + 3 + Kaari * 6) = oArc3.StartPoint(i)
    oPisteet2(p2 + i + 3 + Kaari * 6) = oArc3.EndPoint(i)
  End If
Next i

End Sub
Function PiirraPViiva(Piste1 As Variant, Piste2 As Variant, Etaisyys As Double) As Acad3DPolyline
'Piirt‰‰ poikkiviivan viivalle tietylle et‰isyydelle alkuper‰isen viivan suuntaisesti.

Dim ApuViiva As Acad3DPolyline
Dim APisteet(5) As Double
Dim Kulma As Double

Dim AlkuP As Variant
Dim LoppuP As Variant
Dim aPiste1 As Variant
Dim aPiste2 As Variant
Dim bPiste1 As Variant
Dim bPiste2 As Variant
     
  Kulma = ActiveDocument.Utility.AngleFromXAxis(Piste1, Piste2)

  aPiste1 = ActiveDocument.Utility.PolarPoint(Piste1, Kulma + PI / 2, Lev / 2)
  aPiste2 = ActiveDocument.Utility.PolarPoint(Piste2, Kulma + PI / 2, Lev / 2)
  bPiste1 = ActiveDocument.Utility.PolarPoint(Piste1, Kulma - PI / 2, Lev / 2)
  bPiste2 = ActiveDocument.Utility.PolarPoint(Piste2, Kulma - PI / 2, Lev / 2)
   
  AlkuP = LaskeUusiPiste(aPiste1, aPiste2, Etaisyys)
  LoppuP = LaskeUusiPiste(bPiste1, bPiste2, Etaisyys)
    
    
  APisteet(0) = AlkuP(0)
  APisteet(1) = AlkuP(1)
  APisteet(2) = AlkuP(2)
  APisteet(3) = LoppuP(0)
  APisteet(4) = LoppuP(1)
  APisteet(5) = LoppuP(2)
  Set PiirraPViiva = ActiveDocument.ModelSpace.Add3DPoly(APisteet)
    
    
End Function
Function LaskeUusiPiste(Piste1 As Variant, Piste2 As Variant, Etaisyys As Double) As Variant

Dim xlen As Double
Dim ylen As Double
Dim zlen As Double

Dim Kerroin As Double
Dim Pointti(2) As Double

Kerroin = Etaisyys / LaskePituus(Piste1, Piste2)

xlen = (Piste2(0) - Piste1(0)) * Kerroin
ylen = (Piste2(1) - Piste1(1)) * Kerroin
zlen = (Piste2(2) - Piste1(2)) * Kerroin

Pointti(0) = Piste1(0) + xlen
Pointti(1) = Piste1(1) + ylen
Pointti(2) = Piste1(2) + zlen

LaskeUusiPiste = Pointti

End Function
Sub PoistaTyhjatJoukot()
Dim Loytyi As Boolean
Dim i As Integer
Do
  Loytyi = False
  For i = 0 To ActiveDocument.Groups.Count - 1
    If ActiveDocument.Groups(i).Count = 0 Then
      Loytyi = True
      ActiveDocument.Groups(i).Delete
      Exit For
    End If
  Next i
Loop Until Loytyi = False
End Sub
Function Piirra3dPoly(Piste1 As Variant, Piste2 As Variant) As Acad3DPolyline
Dim ApuP(5) As Double
ApuP(0) = Piste1(0)
ApuP(1) = Piste1(1)
ApuP(2) = Piste1(2)
ApuP(3) = Piste2(0)
ApuP(4) = Piste2(1)
ApuP(5) = Piste2(2)
Set Piirra3dPoly = ActiveDocument.ModelSpace.Add3DPoly(ApuP)
End Function
Function KokoaID() As String
'T‰ss‰ funktiossa kootaan TrayID
  KokoaID = TArea.Value & "-" & TService.Value & TNumber.Value & TBarrier.Value
End Function
Function LaskeLeikkauspiste(Piste1 As Variant, Piste2 As Variant, Piste3 As Variant, Piste4 As Variant) As Variant
Dim ApuViiva1 As AcadLine
Dim ApuViiva2 As AcadLine

  Set ApuViiva1 = ActiveDocument.ModelSpace.AddLine(Piste1, Piste2)
  Set ApuViiva2 = ActiveDocument.ModelSpace.AddLine(Piste3, Piste4)
  LaskeLeikkauspiste = ApuViiva1.IntersectWith(ApuViiva2, acExtendBoth)
  ApuViiva1.Delete
  ApuViiva2.Delete
End Function
Sub LaskeOVPisteet()
'Lasketaan keskiviivan perusteella vasemman ja oikean viivan pisteet
Dim Piste1(2) As Double
Dim Piste2(2) As Double
Dim Piste3(2) As Double

Dim oPiste1 As Variant
Dim oPiste2 As Variant

Dim vPiste1 As Variant
Dim vPiste2 As Variant

Dim ApuPiste1 As Variant
Dim ApuPiste2 As Variant

Dim UusiPiste As Variant
      
      For i = 0 To (UBound(kPisteet) - 3) Step 3
        Piste1(0) = kPisteet(i)
        Piste1(1) = kPisteet(i + 1)
        Piste2(0) = kPisteet(i + 3)
        Piste2(1) = kPisteet(i + 4)
        
        'Lasketaan vasemmanpuoleisen viivan p‰‰tepisteet keskiviivan suhteen
        oPiste1 = ActiveDocument.Utility.PolarPoint(Piste1, ActiveDocument.Utility.AngleFromXAxis(Piste1, Piste2) + PI / 2, Lev / 2)
        oPiste2 = ActiveDocument.Utility.PolarPoint(Piste2, ActiveDocument.Utility.AngleFromXAxis(Piste1, Piste2) + PI / 2, Lev / 2)
        
        'Lasketaan oikeanpuoleisen viivan p‰‰tepisteet keskiviivan suhteen
        vPiste1 = ActiveDocument.Utility.PolarPoint(Piste1, ActiveDocument.Utility.AngleFromXAxis(Piste1, Piste2) + PI / 2, Lev / -2)
        vPiste2 = ActiveDocument.Utility.PolarPoint(Piste2, ActiveDocument.Utility.AngleFromXAxis(Piste1, Piste2) + PI / 2, Lev / -2)
        
        If i = 0 Then 'Ensimm‰isell‰ kerralla asetetaan alkupisteiksi juuri lasketut pisteet
          vPisteet(0) = oPiste1(0)
          vPisteet(1) = oPiste1(1)
          oPisteet(0) = vPiste1(0)
          oPisteet(1) = vPiste1(1)
        End If
        
        vPisteet(i + 3) = oPiste2(0)
        vPisteet(i + 4) = oPiste2(1)
        oPisteet(i + 3) = vPiste2(0)
        oPisteet(i + 4) = vPiste2(1)
        
        If UBound(kPisteet) > i + 5 Then
          Piste3(0) = kPisteet(i + 6)
          Piste3(1) = kPisteet(i + 7)
          
          '------ Lasketaa vasemmanpuoleisen viivan p‰‰tepiste ----
          ApuPiste1 = ActiveDocument.Utility.PolarPoint(Piste2, ActiveDocument.Utility.AngleFromXAxis(Piste2, Piste3) + PI / 2, Lev / 2)
          ApuPiste2 = ActiveDocument.Utility.PolarPoint(Piste3, ActiveDocument.Utility.AngleFromXAxis(Piste2, Piste3) + PI / 2, Lev / 2)
          UusiPiste = LaskeLeikkauspiste(oPiste1, oPiste2, ApuPiste1, ApuPiste2)
          vPisteet(i + 3) = UusiPiste(0)
          vPisteet(i + 4) = UusiPiste(1)
          
          '------ Lasketaa oikeanpuoleisen viivan p‰‰tepiste ----
          ApuPiste1 = ActiveDocument.Utility.PolarPoint(Piste2, ActiveDocument.Utility.AngleFromXAxis(Piste2, Piste3) + PI / 2, Lev / -2)
          ApuPiste2 = ActiveDocument.Utility.PolarPoint(Piste3, ActiveDocument.Utility.AngleFromXAxis(Piste2, Piste3) + PI / 2, Lev / -2)
          UusiPiste = LaskeLeikkauspiste(vPiste1, vPiste2, ApuPiste1, ApuPiste2)
          oPisteet(i + 3) = UusiPiste(0)
          oPisteet(i + 4) = UusiPiste(1)
        End If
      Next i
End Sub
