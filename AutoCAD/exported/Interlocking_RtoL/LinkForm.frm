VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} LinkForm 
   Caption         =   "Interlocking Linking"
   ClientHeight    =   3840
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   2280
   OleObjectBlob   =   "LinkForm.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "LinkForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim KohdeHandle As String     'Kohteen handle
Dim ARROWs() As Integer       'Muutuja kun tutkitan tulevaa signaalia (0=viiva, 1=nuoli, 2=Not nuoli,9=virhe)
Dim KohdeHandles() As String  'Kohteen handle
Dim EkaPoint(2) As Double
Dim TokaPoint(2) As Double
Dim Piirrasuunta As Boolean
Dim Nuolityyppi As Integer
Const Pi = 3.14159265358979
Private Sub MakeLinks_Click()
'T‰m‰ yritt‰‰ tehd‰ linkit lukituksiin
  TeeLinkit
  MsgBox "Links made!", vbInformation, "Make Links"
End Sub
Private Sub ShowLinks_Click()
'T‰m‰ aliohjelma etsii XDatasta linkit ja piirt‰‰ viivat linkkien v‰lille DEBUG Layerille
  If MsgBox("Do you want arrows pointing for direction?", vbYesNo, "Show links") = vbYes Then
    Piirrasuunta = True
  Else
    Piirrasuunta = False
  End If
  PiirraLinkit
  MsgBox "Links drawn to drawing!", vbInformation, "Draw Links"
End Sub
Private Sub ShowHandles_Click()
'T‰m‰ etsii yhteydet ja merkitsee yhteyksien p‰iss‰ olevien pisteiden handlekset
  PiirraHandlet
  MsgBox "Linked handles marked to drawing!", vbInformation, "Mark Handles"
End Sub
Private Sub ShowErrors_Click()
'T‰m‰ etsii ne attribuutit, joihin ei ole linkitetty mit‰‰n ja piirt‰‰ niiden kohdalle merkin DEBUG Layerille
  MarkUnknown
  MsgBox "Unknown attributes marked in drawing!", vbInformation, "Mark Unknown"
End Sub
Private Sub ClearButton_Click()
'T‰m‰ tyhjent‰‰ DEBUG layerin
Dim Vastaus As VbMsgBoxResult
  Vastaus = MsgBox("Do you want to also delete DEBUG layer?", vbYesNoCancel, "Clear DEBUG layer")
  If Vastaus = vbYes Then
    ClearDEBUG True
  ElseIf Vastaus = vbNo Then
    ClearDEBUG
  End If
End Sub
Private Sub ClearX_Click()
'T‰m‰ tuhoaa kaiken XDatan joka on liitettyn‰ ILOCK application nimen alle
  If MsgBox("This will destroy all XData under ILOCK application name!" & vbCrLf & " Are you sure you want to do this?", vbOKCancel, "Clear XData") = vbOK Then
    TuhoaXData
  End If
End Sub
Private Sub MakeDots_Click()
'Etsii T-risteykset ja piirt‰‰ pisteen niiden leikkaukseen
  TeePisteet
End Sub
Private Sub UserForm_Initialize()
  BLOCKPATH = "L:\Projekti\Stendal\Tyo\Interlocking\Blokit\"
End Sub
Private Sub Vaihdablokit_Click()
'Insertoidaan kaikki referenssiblokit uudelleen, koska ne ovat muuttuneet
Dim oBlock As AcadBlockReference
Dim uBlock As AcadBlockReference
Dim BNames(3) As String
Dim iPoint(2) As Double
Dim i As Integer, j As Integer, ii As Integer
Dim Joukko As AcadSelectionSet
Dim FilterType(0) As Integer
Dim FilterData(0) As Variant
Dim AttAlkup As Variant
Dim AttUusi As Variant
Dim Nimi As Integer
BNames(0) = BLOCKPATH & "REF_LOG1.DWG"
BNames(1) = BLOCKPATH & "REF_LOG2.DWG"
BNames(2) = BLOCKPATH & "REF_SAMA1.DWG"
BNames(3) = BLOCKPATH & "REF_SAMA2.DWG"
For i = 0 To 3
  Set oBlock = ThisDrawing.ModelSpace.InsertBlock(iPoint, BNames(i), 1, 1, 1, 0)
  oBlock.Delete
Next i
For i = 0 To ThisDrawing.SelectionSets.Count - 1
  If ThisDrawing.SelectionSets(i).Name = "REFBLOKIT" Then
    ThisDrawing.SelectionSets(i).Delete
    Exit For
  End If
Next i
Set Joukko = ThisDrawing.SelectionSets.Add("REFBLOKIT")
FilterType(0) = 2
FilterData(0) = "REF_LOG1,REF_LOG2,REF_SAMA1,REF_SAMA2"
Joukko.Select acSelectionSetAll, , , FilterType, FilterData
For i = 0 To Joukko.Count - 1
  Set oBlock = Joukko(i)
  AttAlkup = oBlock.GetAttributes
  Set uBlock = ThisDrawing.ModelSpace.InsertBlock(oBlock.InsertionPoint, oBlock.Name, 1, 1, 1, 0)
  AttUusi = uBlock.GetAttributes
  For j = 0 To UBound(AttUusi)
    For ii = 0 To UBound(AttAlkup)
      If AttUusi(j).TagString = AttAlkup(ii).TagString Then
        AttUusi(j).TextString = AttAlkup(ii).TextString
      End If
    Next ii
  Next j
  oBlock.Delete
Next i
Joukko.Delete
ThisDrawing.Regen acActiveViewport
Set oBlock = Nothing
End Sub
Private Sub PiirraViiva_Click()
  LinkkiFormi = True
  LinkForm.Hide
  LineForm.Show
End Sub
Private Sub TeeLinkit()
Dim i As Integer
Dim oEntity As AcadEntity
Dim oBlock As AcadBlockReference
Dim Attribuutit As Variant
Dim oAttribute As AcadAttributeReference
Application.ZoomExtents
ClearDEBUG
TuhoaXData
For Each oEntity In ThisDrawing.ModelSpace   'Kierret‰‰n l‰pi kaikki piirustuksen elementit (viivat, blockit jne.)
  If oEntity.EntityName = "AcDbBlockReference" Then
    Set oBlock = oEntity
    If OnkoLaite(oBlock.Name) Then
      Attribuutit = oBlock.GetAttributes
      For i = 0 To UBound(Attribuutit)
        Set oAttribute = Attribuutit(i)
        With oAttribute
          If .TextString <> "" Then
            Select Case .TagString
              Case "LEFT1", "LEFT2", "LEFT3", "LEFT4", "LEFT5", "LEFT6", "LEFT7", "LEFT8"
                EkaPoint(0) = .TextAlignmentPoint(0)
                EkaPoint(1) = .TextAlignmentPoint(1) + 2.5
                TokaPoint(0) = .TextAlignmentPoint(0) - 5
                TokaPoint(1) = .TextAlignmentPoint(1) - 2.5
                EtsiLahteet
                LaitaXData oAttribute
              Case "AUTO", "MANUAL"
                EkaPoint(0) = .TextAlignmentPoint(0) - 2.5
                EkaPoint(1) = .TextAlignmentPoint(1)
                TokaPoint(0) = .TextAlignmentPoint(0) + 2.5
                TokaPoint(1) = .TextAlignmentPoint(1) - 5
                EtsiLahteet
                LaitaXData oAttribute
              Case "RIGHT1", "RIGHT2", "RIGHT3", "RIGHT4"
                EkaPoint(0) = .TextAlignmentPoint(0)
                EkaPoint(1) = .TextAlignmentPoint(1) + 2.5
                TokaPoint(0) = .TextAlignmentPoint(0) + 5
                TokaPoint(1) = .TextAlignmentPoint(1) - 2.5
                EtsiLahteet
                LaitaXData oAttribute
            End Select
          End If
        End With
      Next i
    ElseIf OnkoRefOut(oBlock.Name) Then
      SeuraaViivaa , oBlock.InsertionPoint
      ReDim ARROWs(1)
      ARROWs(0) = 0
      ReDim KohdeHandles(1)
      KohdeHandles(0) = KohdeHandle
      LaitaXData oBlock
    ElseIf OnkoJaTai(oBlock.Name) Then
      EkaPoint(0) = oBlock.InsertionPoint(0)
      EkaPoint(1) = oBlock.InsertionPoint(1) + 10
      TokaPoint(0) = oBlock.InsertionPoint(0) - 5
      TokaPoint(1) = oBlock.InsertionPoint(1)
      EtsiLahteet
      LaitaXData oBlock
    End If
  End If
Next
Set oAttribute = Nothing
Set oEntity = Nothing
Set oBlock = Nothing
End Sub
Private Function OnkoLaite(Nimi As String) As Boolean
  Select Case UCase(Nimi)
    Case "SWITCH1", "SWITCH2", "SWITCH3", "LIMITSW1", "LIMITSW2", "LOOP1", "LOOP2", "MOTOR1", "MOTOR2", "MOTOR3", "VALVE1"
      OnkoLaite = True
    Case Else
      OnkoLaite = False
  End Select
End Function
Private Function OnkoRefIn(Nimi As String) As Boolean
  Select Case UCase(Nimi)
    Case "REF_LOG1", "REF_SAMA1", "IBG"
      OnkoRefIn = True
    Case Else
      OnkoRefIn = False
  End Select
End Function
Private Function OnkoRefOut(Nimi As String) As Boolean
  Select Case UCase(Nimi)
    Case "REF_LOG2", "REF_SAMA2"
      OnkoRefOut = True
    Case Else
      OnkoRefOut = False
  End Select
End Function
Private Function OnkoJaTai(Nimi As String) As Boolean
  Select Case Nimi
    Case "AND", "OR", "OFFDELAY", "ONDELAY", "ONOFFDELAY", "PULSE1", "SR", "RS"
      OnkoJaTai = True
    Case Else
      OnkoJaTai = False
  End Select
End Function
Private Sub EtsiLahteet()
Dim Joukko As AcadSelectionSet
Dim oEntity As AcadEntity
Dim oBlock As AcadBlockReference
Dim i As Integer
ReDim ARROWs(0)
ReDim KohdeHandles(0)
  For i = 0 To ThisDrawing.SelectionSets.Count - 1
    If ThisDrawing.SelectionSets(i).Name = "SSET" Then
      ThisDrawing.SelectionSets(i).Delete
      Exit For
    End If
  Next i
  Set Joukko = ThisDrawing.SelectionSets.Add("SSET")
  Joukko.Select acSelectionSetCrossing, EkaPoint, TokaPoint
  PiirraNelio
  For Each oEntity In Joukko
    If oEntity.EntityName = "AcDbBlockReference" Then
      Set oBlock = oEntity
      If oBlock.Name = "ARROW" Then
        ARROWs(UBound(ARROWs)) = 1
        ReDim Preserve ARROWs(UBound(ARROWs) + 1)
        SeuraaViivaa oBlock
        KohdeHandles(UBound(KohdeHandles)) = KohdeHandle
        ReDim Preserve KohdeHandles(UBound(KohdeHandles) + 1)
      ElseIf oBlock.Name = "NOTARROW" Then
        ARROWs(UBound(ARROWs)) = 2
        ReDim Preserve ARROWs(UBound(ARROWs) + 1)
        SeuraaViivaa oBlock
        KohdeHandles(UBound(KohdeHandles)) = KohdeHandle
        ReDim Preserve KohdeHandles(UBound(KohdeHandles) + 1)
      End If
    ElseIf oEntity.EntityName = "AcDbLine" Then
      ARROWs(UBound(ARROWs)) = 0
      ReDim Preserve ARROWs(UBound(ARROWs) + 1)
      KohdeHandles(UBound(KohdeHandles)) = ""
      ReDim Preserve KohdeHandles(UBound(KohdeHandles) + 1)
    End If
  Next
  Joukko.Delete
  Set oBlock = Nothing
  Set Joukko = Nothing
End Sub
Private Sub LaitaXData(Objekti As AcadEntity)
Dim SIGNAALI(2) As String
Dim XData() As Variant
Dim XType() As Integer
Dim Koko As Integer
Dim i As Integer, j As Integer
If UBound(ARROWs) > 0 Then
  Koko = (UBound(ARROWs)) * 2
  SIGNAALI(0) = "LINE"
  SIGNAALI(1) = "ARROW"
  SIGNAALI(2) = "NOTARROW"
  ReDim XType(Koko)
  ReDim XData(Koko)
  XType(0) = 1001: XData(0) = "ILOCK"      'Sovellusnimi
  j = 1
  For i = 0 To UBound(ARROWs) - 1
    XType(j) = 1000: XData(j) = SIGNAALI(ARROWs(i))       'Signaalin tyyppi joka kyseiseen attribuuttiin tulee
    XType(j + 1) = 1005: XData(j + 1) = KohdeHandles(i)   'Sen blokin/attribuutin handles, joka lukitsee
    j = j + 2
  Next i
  Objekti.SetXData XType, XData
End If
End Sub
Private Sub ClearDEBUG(Optional Tuhoa As Boolean) 'Muodostaa tai tyhjent‰‰ DEBUG layerin tai tuhoaa sen
Dim Joukko As AcadSelectionSet
Dim FilterType(0) As Integer
Dim FilterData(0) As Variant
Dim i As Integer
Application.ZoomExtents
TeeDEBUG
For i = 0 To ThisDrawing.SelectionSets.Count - 1
  If ThisDrawing.SelectionSets(i).Name = "DELETE" Then
    ThisDrawing.SelectionSets(i).Delete
  End If
Next i
Set Joukko = ThisDrawing.SelectionSets.Add("DELETE")
FilterType(0) = 8       'Type 8 = Layer
FilterData(0) = "DEBUG" 'Layer name
Joukko.Select acSelectionSetAll, , , FilterType, FilterData
Joukko.Erase
Joukko.Delete
If Tuhoa Then
  ThisDrawing.ActiveLayer = ThisDrawing.Layers("0")
  ThisDrawing.Layers("DEBUG").Delete
End If
Set Joukko = Nothing
End Sub
Private Sub TeeDEBUG()
Dim i As Integer
Dim Kerros As AcadLayer
For i = 0 To ThisDrawing.Layers.Count - 1
  If ThisDrawing.Layers(i).Name = "DEBUG" Then
    Set Kerros = ThisDrawing.Layers(i)
    Exit For
  End If
Next i
If Kerros Is Nothing Then
  Set Kerros = ThisDrawing.Layers.Add("DEBUG")
  Kerros.Color = acGreen
End If
ThisDrawing.ActiveLayer = Kerros
Set Kerros = Nothing
End Sub
Private Sub PiirraNelio()
'Piirt‰‰ polylinella nelion jossa kulmina on EkaPoin ja TokaPoint
Dim Points(14) As Double
Dim oLine As AcadPolyline
  Points(0) = EkaPoint(0): Points(1) = EkaPoint(1): Points(2) = 0
  Points(3) = EkaPoint(0): Points(4) = TokaPoint(1): Points(5) = 0
  Points(6) = TokaPoint(0): Points(7) = TokaPoint(1): Points(8) = 0
  Points(9) = TokaPoint(0): Points(10) = EkaPoint(1): Points(11) = 0
  Points(12) = EkaPoint(0): Points(13) = EkaPoint(1): Points(14) = 0
  Set oLine = ThisDrawing.ModelSpace.AddPolyline(Points)
  oLine.Color = acRed
  Set oLine = Nothing
End Sub
Private Sub SeuraaViivaa(Optional Blokki As AcadBlockReference, Optional AlkuPiste As Variant)
'T‰m‰n aliohjelman tarkoituksena on seurata nuoliblokkiin kiinnitetty‰ viivaa sen l‰hteeseen
Dim Joukko As AcadSelectionSet
Dim Viiva As AcadLine
Dim ApuViiva As AcadLine
Dim oBlock As AcadBlockReference
Dim i As Integer, j As Integer
Dim ViivaLoytyi As Boolean
Dim BlockLoytyi As Boolean
Dim NuoliLoytyi As Boolean
Dim Attribuutit As Variant
Dim oAttribute As AcadAttributeReference
Dim TPoint(2) As Double
Dim sPoint(2) As Double
Dim Lyhin As Double
Dim ApuVali As Double
Dim Alkuun As Boolean
Dim VanhaHandle As String
KohdeHandle = ""
NuoliLoytyi = False
If Not Blokki Is Nothing Then
With Blokki
    sPoint(0) = .InsertionPoint(0) - Cos(.Rotation) * 7.5
    sPoint(1) = .InsertionPoint(1) - Sin(.Rotation) * 7.5
  End With
Else
  sPoint(0) = AlkuPiste(0)
  sPoint(1) = AlkuPiste(1)
End If
EkaPoint(0) = sPoint(0) - 1.24 '1.25
EkaPoint(1) = sPoint(1) - 1.24 '1.25
TokaPoint(0) = sPoint(0) + 1.24 '1.25
TokaPoint(1) = sPoint(1) + 1.24 '1.25
For i = 0 To ThisDrawing.SelectionSets.Count - 1
  If ThisDrawing.SelectionSets(i).Name = "SEURAA" Then
     ThisDrawing.SelectionSets(i).Delete
     Exit For
  End If
Next i
Set Joukko = ThisDrawing.SelectionSets.Add("SEURAA")
Joukko.Select acSelectionSetCrossing, EkaPoint, TokaPoint
For i = 0 To Joukko.Count - 1 'Etsit‰‰n joukosta ensimm‰inen viiva
  If Joukko(i).EntityName = "AcDbLine" Then
    Set Viiva = Joukko(i)
    Alkuun = OnkoAlueella(Viiva.EndPoint)
    Exit For
  End If
Next i
Joukko.Delete
PiirraNelio
Do 'Seurataan sitten viivaa sen l‰htˆsuuntaan
  If Alkuun Then 'Seurataan viivaa sen alkuun p‰in
    TPoint(0) = Viiva.StartPoint(0)
    TPoint(1) = Viiva.StartPoint(1)
  Else
    TPoint(0) = Viiva.EndPoint(0)
    TPoint(1) = Viiva.EndPoint(1)
  End If
  EkaPoint(0) = TPoint(0) + 1.24 '1.25
  EkaPoint(1) = TPoint(1) + 1.24 '1.25
  TokaPoint(0) = TPoint(0) - 1.24 '1.25
  TokaPoint(1) = TPoint(1) - 1.24 '1.25
  PiirraNelio
  Set Joukko = ThisDrawing.SelectionSets.Add("SEURAA")
  Joukko.Select acSelectionSetCrossing, EkaPoint, TokaPoint
  ViivaLoytyi = False
  BlockLoytyi = False
  VanhaHandle = Viiva.Handle
  For i = 0 To Joukko.Count - 1 'Tsekataan kaikki entityt viivan p‰‰sti
    If Joukko(i).Layer <> "DEBUG" Then 'Debug tasaolla olevia entityj‰ ei huomioida
      If Joukko(i).EntityName = "AcDbLine" Then 'Joukosta lˆytyi Viiva
        If Joukko(i).Handle <> VanhaHandle Then 'Viiva ei ole sama kuin mit‰ tutkitaan
          Set Viiva = Joukko(i)
          If OnkoAlueella(Viiva.EndPoint) Then
            Alkuun = True
          ElseIf OnkoAlueella(Viiva.StartPoint) Then
            Alkuun = False
          End If
'          Alkuun = OnkoAlueella(Viiva.EndPoint)
          ViivaLoytyi = True
        End If
      ElseIf Joukko(i).EntityName = "AcDbBlockReference" Then 'Lˆytyi blokki
        If BlockLoytyi = False Then 'Ei ole viel‰ lˆytynyt laite tai ja/tai blokkia t‰st‰ joukosta
          Set oBlock = Joukko(i)
          If OnkoLaite(oBlock.Name) Or OnkoRefIn(oBlock.Name) Or OnkoJaTai(oBlock.Name) Then 'Oikean tyyppinen blokki lˆytyi
            BlockLoytyi = True
          ElseIf oBlock.Name = "ARROW" Or oBlock.Name = "NOTARROW" Or OnkoRefOut(oBlock.Name) Then
            If NuoliLoytyi Then 'Jo toinen nuoli loytyi, joten jossain on virhe, lopetetaan looppi
              Exit Do
            Else
              NuoliLoytyi = True
            End If
          End If
        End If
      End If
    End If
  Next i
  Joukko.Delete
  If BlockLoytyi Then 'Lˆydettiin blokki, josta etsit‰‰n Handle
    If OnkoLaite(oBlock.Name) Then 'Blokki josta etsit‰‰n l‰hin attribuutti
      EkaPoint(0) = TPoint(0)
      EkaPoint(1) = TPoint(1)
      KohdeHandle = ""
      Attribuutit = oBlock.GetAttributes
      Lyhin = 10000
      For j = 0 To UBound(Attribuutit)
        Set oAttribute = Attribuutit(j)
        If oAttribute.TextString <> "" Then
          TokaPoint(0) = oAttribute.TextAlignmentPoint(0)
          TokaPoint(1) = oAttribute.TextAlignmentPoint(1)
          ApuVali = LaskeVali
          If ApuVali < Lyhin Then
            Lyhin = ApuVali
            KohdeHandle = oAttribute.Handle
          End If
        End If
      Next j
    ElseIf OnkoJaTai(oBlock.Name) Or OnkoRefIn(oBlock.Name) Then
      KohdeHandle = oBlock.Handle
    End If
    Exit Do 'Finaali kohde lˆytyi, joten ei jatketa enemp‰‰ viivan seuraamista
  ElseIf ViivaLoytyi Then 'Ei lˆytynyt blokki, mutta lˆytyi viiva, joten Looppi saa jatkua
  ElseIf NuoliLoytyi Then 'Ei lˆytynyt blokkia tai viivaa, mutta nuoli lˆytyi, joten kokeillaan toiseen suuntaan
    Alkuun = Not Alkuun
  Else 'Toisesta p‰‰st‰ ei lˆytynyt viivaa eik‰ blokkia, joten poistutaan
    Exit Do
  End If
Loop
Set oAttribute = Nothing
Set oBlock = Nothing
Set Joukko = Nothing
Set Viiva = Nothing
End Sub
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
Private Sub PiirraLinkit()
Dim i As Integer
Dim oEntity As AcadEntity
Dim oKohde As AcadEntity
Dim oBlock As AcadBlockReference
Dim oAttribute As AcadAttributeReference
Dim Attribuutit As Variant
Application.ZoomExtents
TeeDEBUG 'Varmistetaan ett‰ DEBUG on olemassa ja ett‰ se on valittu
For Each oEntity In ThisDrawing.ModelSpace   'Kierret‰‰n l‰pi kaikki piirustuksen elementit (viivat, blockit jne.)
  If oEntity.EntityName = "AcDbBlockReference" Then
    Set oBlock = oEntity
    If OnkoLaite(oBlock.Name) Then
      Attribuutit = oBlock.GetAttributes
      For i = 0 To UBound(Attribuutit)
        Set oAttribute = Attribuutit(i)
        With oAttribute
          If .TextString <> "" Then
            PiirraViivat oAttribute
          End If
        End With
      Next i
    ElseIf OnkoJaTai(oBlock.Name) Or OnkoRefOut(oBlock.Name) Then
      PiirraViivat , oBlock
    End If
  End If
Next
Set oAttribute = Nothing
Set oEntity = Nothing
Set oBlock = Nothing
End Sub
Private Sub PiirraHandlet()
Dim i As Integer
Dim oEntity As AcadEntity
Dim oKohde As AcadEntity
Dim oBlock As AcadBlockReference
Dim oAttribute As AcadAttributeReference
Dim Attribuutit As Variant
Application.ZoomExtents
TeeDEBUG 'Varmistetaan ett‰ DEBUG on olemassa ja ett‰ se on valittu
For Each oEntity In ThisDrawing.ModelSpace   'Kierret‰‰n l‰pi kaikki piirustuksen elementit (viivat, blockit jne.)
  If oEntity.EntityName = "AcDbBlockReference" Then
    Set oBlock = oEntity
    If OnkoLaite(oBlock.Name) Then
      Attribuutit = oBlock.GetAttributes
      For i = 0 To UBound(Attribuutit)
        Set oAttribute = Attribuutit(i)
        With oAttribute
          If .TextString <> "" Then
            PiirraHandle oAttribute
          End If
        End With
      Next i
    ElseIf OnkoJaTai(oBlock.Name) Then
      PiirraHandle , oBlock
    End If
  End If
Next
Set oAttribute = Nothing
Set oEntity = Nothing
Set oBlock = Nothing
End Sub
Private Sub PiirraHandle(Optional Attribuutti As AcadAttributeReference, Optional Blokki As AcadBlockReference)
Dim XDataType As Variant
Dim XDataValue As Variant
Dim oKohde As AcadEntity
Dim oAttribute As AcadAttributeReference
Dim oBlock As AcadBlockReference
Dim Ympyra As AcadCircle
Dim Kahva As String
Dim i  As Integer
On Error Resume Next
  If Blokki Is Nothing Then
    EkaPoint(0) = Attribuutti.TextAlignmentPoint(0)
    EkaPoint(1) = Attribuutti.TextAlignmentPoint(1)
    Attribuutti.GetXData "ILOCK", XDataType, XDataValue
    Kahva = Attribuutti.Handle
  Else
    EkaPoint(0) = Blokki.InsertionPoint(0)
    EkaPoint(1) = Blokki.InsertionPoint(1)
    Blokki.GetXData "ILOCK", XDataType, XDataValue
    Kahva = Blokki.Handle
  End If
  If Not IsEmpty(XDataValue) Then
    For i = 0 To UBound(XDataValue)
      If XDataType(i) = 1005 Then
        If XDataValue(i) <> 0 Then
          HandleMerkki EkaPoint, Kahva
          Set oKohde = ThisDrawing.HandleToObject(XDataValue(i)) 'Haetaan Entity sen handleksen perusteella
          If oKohde.ObjectName = "AcDbBlockReference" Then
            Set oBlock = oKohde
            HandleMerkki oBlock.InsertionPoint, oBlock.Handle
          ElseIf oKohde.ObjectName = "AcDbAttribute" Then
            Set oAttribute = oKohde
            HandleMerkki oAttribute.TextAlignmentPoint, oAttribute.Handle
          End If
        End If
      End If
    Next i
  End If
  XDataType = Empty
  XDataValue = Empty
  Set oBlock = Nothing
  Set oKohde = Nothing
  Set oAttribute = Nothing
End Sub
Private Sub HandleMerkki(iPoint As Variant, Kahva As String)
Dim sPoint(2) As Double
Dim ePoint(2) As Double
Dim Viiva As AcadLine
Dim Teksti As AcadText
 sPoint(0) = iPoint(0)
 sPoint(1) = iPoint(1)
 ePoint(0) = iPoint(0) + 10
 ePoint(1) = iPoint(1) + 20
 Set Viiva = ThisDrawing.ModelSpace.AddLine(sPoint, ePoint)
 Viiva.Color = acYellow
 sPoint(0) = ePoint(0) + 20
 sPoint(1) = ePoint(1)
 Set Viiva = ThisDrawing.ModelSpace.AddLine(sPoint, ePoint)
 Viiva.Color = acYellow
 ePoint(1) = ePoint(1) + 0.5
 Set Teksti = ThisDrawing.ModelSpace.AddText(Kahva, ePoint, 3)
 Teksti.Color = acYellow
 Teksti.StyleName = "STANDARD"
 Set Viiva = Nothing
 Set Teksti = Nothing
End Sub
Private Sub PiirraViivat(Optional Attribuutti As AcadAttributeReference, Optional Blokki As AcadBlockReference)
Dim XDataType As Variant
Dim XDataValue As Variant
Dim oKohde As AcadEntity
Dim oAttribute As AcadAttributeReference
Dim oBlock As AcadBlockReference
Dim Ympyra As AcadCircle
Dim i  As Integer
On Error GoTo Virhe
  If Blokki Is Nothing Then
    EkaPoint(0) = Attribuutti.TextAlignmentPoint(0)
    EkaPoint(1) = Attribuutti.TextAlignmentPoint(1)
    Attribuutti.GetXData "ILOCK", XDataType, XDataValue
  Else
    EkaPoint(0) = Blokki.InsertionPoint(0)
    EkaPoint(1) = Blokki.InsertionPoint(1)
    Blokki.GetXData "ILOCK", XDataType, XDataValue
  End If
  If Not IsEmpty(XDataValue) Then
    For i = 0 To UBound(XDataValue)
      If XDataType(i) = 1000 Then
        If XDataValue(i) = "ARROW" Then
          Nuolityyppi = 1
        Else
          Nuolityyppi = 0
        End If
      ElseIf XDataType(i) = 1005 Then
        If XDataValue(i) <> 0 Then
          Set oKohde = ThisDrawing.HandleToObject(XDataValue(i)) 'Haetaan Entity sen handleksen perusteella
          If oKohde.ObjectName = "AcDbBlockReference" Then
            Set oBlock = oKohde
            TokaPoint(0) = oBlock.InsertionPoint(0)
            TokaPoint(1) = oBlock.InsertionPoint(1)
          Else
            Set oAttribute = oKohde
            TokaPoint(0) = oAttribute.TextAlignmentPoint(0)
            TokaPoint(1) = oAttribute.TextAlignmentPoint(1)
          End If
          PiirraArc
        End If
      End If
    Next i
  End If
Ulos:
  XDataType = Empty
  XDataValue = Empty
  Set oBlock = Nothing
  Set oKohde = Nothing
  Set oAttribute = Nothing
  Exit Sub
Virhe:
  Err.Clear
  Set Ympyra = ThisDrawing.ModelSpace.AddCircle(EkaPoint, 8)
  Ympyra.Color = acRed
  Resume Ulos
End Sub
Private Sub PiirraArc()
'Piirt‰‰ kaaren kahden pisteen v‰lille
Dim cPoint(2) As Double
Dim Sade As Double
Dim Viiva As AcadLine
Dim aViiva As AcadLine
Dim AlkuKulma As Double
Dim LoppuKulma As Double
Dim Kaari As AcadArc
Dim kPiste As Variant
Dim oText As AcadText
If EkaPoint(0) > TokaPoint(0) Then
  cPoint(0) = TokaPoint(0) + (EkaPoint(0) - TokaPoint(0)) / 2
Else
  cPoint(0) = EkaPoint(0) + (TokaPoint(0) - EkaPoint(0)) / 2
End If
If EkaPoint(1) > TokaPoint(1) Then
  cPoint(1) = TokaPoint(1) + (EkaPoint(1) - TokaPoint(1)) / 2
Else
  cPoint(1) = EkaPoint(1) + (TokaPoint(1) - EkaPoint(1)) / 2
End If
  Set Viiva = ThisDrawing.ModelSpace.AddLine(EkaPoint, TokaPoint)
  Viiva.Rotate cPoint, Pi / 2
  Set aViiva = ThisDrawing.ModelSpace.AddLine(Viiva.EndPoint, TokaPoint)
  LoppuKulma = aViiva.Angle
  aViiva.Delete
  Set aViiva = ThisDrawing.ModelSpace.AddLine(Viiva.EndPoint, EkaPoint)
  AlkuKulma = aViiva.Angle
  Sade = aViiva.Length
  aViiva.Delete
  Set Kaari = ThisDrawing.ModelSpace.AddArc(Viiva.EndPoint, Sade, AlkuKulma, LoppuKulma)
  If Piirrasuunta Then
    kPiste = Kaari.IntersectWith(Viiva, acExtendNone)
    Set oText = ThisDrawing.ModelSpace.AddText(Nuolityyppi & ">", kPiste, 3)
    oText.Rotate kPiste, Viiva.Angle + Pi / 2
    oText.StyleName = "STANDARD"
  End If
  Viiva.Delete
  Set oText = Nothing
  Set Kaari = Nothing
  Set Viiva = Nothing
  Set aViiva = Nothing
End Sub
Private Sub MarkUnknown()
Dim i As Integer, j As Integer
Dim oEntity As AcadEntity
Dim oKohde As AcadEntity
Dim oBlock As AcadBlockReference
Dim oAttribute As AcadAttributeReference
Dim Ympyra As AcadCircle
Dim Attribuutit As Variant
Dim XDataType As Variant
Dim XDataValue As Variant
Dim Virhe As Boolean
Application.ZoomExtents
TeeDEBUG 'Varmistetaan ett‰ DEBUG on olemassa ja ett‰ se on valittu
For Each oEntity In ThisDrawing.ModelSpace   'Kierret‰‰n l‰pi kaikki piirustuksen elementit (viivat, blockit jne.)
  If oEntity.EntityName = "AcDbBlockReference" Then
    Set oBlock = oEntity
    If OnkoLaite(oBlock.Name) Then
      Attribuutit = oBlock.GetAttributes
      For i = 0 To UBound(Attribuutit)
        Set oAttribute = Attribuutit(i)
        With oAttribute
          Select Case .TagString
            Case "LEFT1", "LEFT2", "LEFT3", "LEFT4", "AUTO", "MANUAL", "RIGHT1", "RIGHT2", "RIGHT3", "RIGHT4"
              If .TextString <> "" Then
                oAttribute.GetXData "ILOCK", XDataType, XDataValue
                Virhe = False
                If IsEmpty(XDataValue) Then
                  Virhe = True
                End If
                If Virhe Then
                  Set Ympyra = ThisDrawing.ModelSpace.AddCircle(oAttribute.TextAlignmentPoint, 8)
                  Ympyra.Color = acYellow
                End If
              End If
          End Select
        End With
      Next i
    End If
  End If
Next
Set Ympyra = Nothing
Set oAttribute = Nothing
Set oEntity = Nothing
Set oBlock = Nothing
End Sub
Private Sub TeePisteet()
Dim i As Integer, j As Integer
Dim Viivat As AcadSelectionSet
Dim Joukko As AcadSelectionSet
Dim FilterType(0) As Integer
Dim FilterData(0) As Variant
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
FilterType(0) = 0
FilterData(0) = "Line"
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
Private Function LaskeVali() As Double
'Laskee kahden pisteen v‰lisen et‰isyyden
  LaskeVali = Sqr(Abs(EkaPoint(0) - TokaPoint(0)) ^ 2 + Abs(EkaPoint(1) - TokaPoint(1)) ^ 2)
End Function
Private Sub TuhoaXData()
Dim i As Integer
Dim oEntity As AcadEntity
Dim oBlock As AcadBlockReference
Dim Attribuutit As Variant
For Each oEntity In ThisDrawing.ModelSpace   'Kierret‰‰n l‰pi kaikki piirustuksen elementit (viivat, blockit jne.)
  TuhoaX oEntity
  If oEntity.ObjectName = "AcDbBlockReference" Then
    Set oBlock = oEntity
    If oBlock.HasAttributes Then
      Attribuutit = oBlock.GetAttributes
      For i = 0 To UBound(Attribuutit)
        TuhoaX Attribuutit(i)
      Next i
    End If
  End If
Next
Set oEntity = Nothing
Set oBlock = Nothing
End Sub
Private Sub TuhoaX(Objekti As Variant)
Dim XDataType As Variant
Dim XDataValue As Variant
Dim DataType(0) As Integer
Dim Data(0) As Variant
Dim i  As Integer
  DataType(0) = 1001 'Type = Application name
  Data(0) = "ILOCK"  'Application name
  Objekti.GetXData "", XDataType, XDataValue
  If Not IsEmpty(XDataValue) Then
    Objekti.SetXData DataType, Data
  End If
  XDataType = Empty
  XDataValue = Empty
End Sub

