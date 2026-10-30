Option Compare Database
Option Explicit

#If VBA7 Then
  Private Declare PtrSafe Function SetForegroundWindow Lib "user32" (ByVal hwnd As LongPtr) As Long
#Else
  Private Declare Function SetForegroundWindow Lib "user32" (ByVal hwnd As Long) As Long
#End If

'================================================================================
' Lomake: USysPipeFromTo
' Tarkoitus: AutoCAD-putkilinjablokkien FROM/TO-attribuuttien muokkaustyˆkalu
' P‰ivitetty: 2025-11-12 - VBA7/64-bit-tuki ja optimointi lis‰tty
'
' Kuvaus:
'   Interaktiivinen lomake putkilinjablokkien FROM/TO-attribuuttien muokkaamiseen:
'   - Yhdistyy k‰ynniss‰ olevaan AutoCAD-instanssiin
'   - Suodattaa ja n‰ytt‰‰ vain putkilinjablokit (PIPELINE, PIPELINE_DATA, METSO_PIPE)
'   - Mahdollistaa navigoinnin blokkien v‰lill‰ (Seuraava, Edellinen, Etsi tyhj‰)
'   - Poimi FROM/TO piirustuksesta tai syˆt‰ k‰sin
'   - P‰ivitt‰‰ blokin attribuutit reaaliajassa
'
' Riippuvuudet:
'   - AutoCAD-sovellus (COM-automaatio)
'   - Putkilinjablokit, joissa on FROM/TO-attribuutit
'
' Kontrollit:
'   - TFrom, TTo: FROM/TO-tekstikent‰t
'   - BNext, BPrev: Navigointipainikkeet
'   - BFind: Etsi seuraava tyhj‰ -painike
'   - BPickFrom, BPickTo: Poimi piirustuksesta -painikkeet
'   - BClearFrom, BClearTo: Tyhjenn‰ attribuutti -painikkeet
'   - LPipeLine: N‰ytt‰‰ nykyisen putkilinjaviittauksen
'================================================================================

' AutoCAD-oliot ovat late binding -muodossa, jotta sama koodi toimii 32- ja 64-bittisess‰ Accessissa
Private oACAD As Object       ' AcadApplication
Private oDOC As Object        ' AcadDocument
Private AttribFrom As Object  ' AcadAttributeReference
Private AttribTo As Object    ' AcadAttributeReference
Private oBlock As Object      ' AcadBlockReference
Private Joukko As Object      ' AcadSelectionSet
Private mIndeksi As Long      ' Nykyinen blokkiindeksi
Private mSuljeLataus As Boolean  ' Lataus ep‰onnistui, lomake suljetaan ajastimella

' AutoCAD-vakiot m‰‰ritell‰‰n paikallisesti, koska tyyppikirjastoviittausta ei k‰ytet‰
Private Const acSelectionSetAll As Long = 5
Private Const acZoomScaledRelative As Long = 1

'================================================================================
' Tapahtuma: BClearFrom_Click
' Tarkoitus: Tyhjent‰‰ nykyisen putkilinjablokin FROM-attribuutin
'================================================================================
Private Sub BClearFrom_Click()
  On Error GoTo ErrorHandler
  
  If AttribFrom Is Nothing Then
    MsgBox "Blokissa ei ole FROM-attribuuttia.", vbExclamation
    Exit Sub
  End If
  AttribFrom.TextString = ""
  TFrom.Value = ""
  
  Exit Sub
ErrorHandler:
  MsgBox "Virhe FROM-attribuutin tyhjennyksess‰: " & Err.Description, vbExclamation
End Sub

'================================================================================
' Tapahtuma: BClearTo_Click
' Tarkoitus: Tyhjent‰‰ nykyisen putkilinjablokin TO-attribuutin
'================================================================================
Private Sub BClearTo_Click()
  On Error GoTo ErrorHandler
  
  If AttribTo Is Nothing Then
    MsgBox "Blokissa ei ole TO-attribuuttia.", vbExclamation
    Exit Sub
  End If
  AttribTo.TextString = ""
  TTo.Value = ""
  
  Exit Sub
ErrorHandler:
  MsgBox "Virhe TO-attribuutin tyhjennyksess‰: " & Err.Description, vbExclamation
End Sub

'================================================================================
' Tapahtuma: BClose_Click
' Tarkoitus: Sulkee lomakkeen
'================================================================================
Private Sub BClose_Click()
  DoCmd.Close acForm, Me.Name
End Sub

'================================================================================
' Tapahtuma: Form_Timer
' Tarkoitus: Sulkee lomakkeen, jos lataus ep‰onnistui (sulkeminen Form_Load-tapahtumassa aiheuttaisi virheen)
'================================================================================
Private Sub Form_Timer()
  Me.TimerInterval = 0
  If mSuljeLataus Then DoCmd.Close acForm, Me.Name
End Sub

'================================================================================
' Aliohjelma: OpenSource
' Tarkoitus: Poimii putkilinjan, s‰iliˆn, pumpun tai laitteen piirroksesta ja palauttaa viittauksen
' Parametrit:
'   Teksti - Tekstikentt‰, johon viittaus asetetaan
'   oAttrib - P‰ivitett‰v‰ attribuuttiviittaus
'================================================================================
Private Sub OpenSource(Teksti As TextBox, oAttrib As Object)
  Dim Kohde As Object
  Dim Piste As Variant
  Dim oPipeBlock As Object  ' AcadBlockReference
  Dim Attrib As Variant
  Dim j As Long
  Dim Virhe As String
  
  If oAttrib Is Nothing Then
    MsgBox "Nykyisess‰ blokissa ei ole t‰h‰n kentt‰‰n liittyv‰‰ attribuuttia.", vbExclamation
    Exit Sub
  End If
  
  On Error Resume Next
  AppActivate oACAD.Caption, True
  If Err.Number <> 0 Then
    Virhe = Err.Description
    Err.Clear
    On Error GoTo 0
    MsgBox "AutoCAD-ikkunaa ei voitu aktivoida: " & Virhe, vbExclamation
    Exit Sub
  End If
  
  ' Virhe t‰ss‰ tarkoittaa yleens‰ sit‰, ett‰ k‰ytt‰j‰ perui valinnan, joten se ohitetaan hiljaa
  oDOC.Utility.GetEntity Kohde, Piste, "Poimi putkilinja, s‰iliˆ, pumppu tai laite: "
  If Err.Number <> 0 Then
    Err.Clear
    On Error GoTo 0
    AktivoiAccess
    Exit Sub
  End If
  On Error GoTo ErrorHandler
  
  If Kohde.ObjectName = "AcDbBlockReference" Then
    Set oPipeBlock = Kohde
    Select Case LCase$(oPipeBlock.EffectiveName)
      Case "pipeline", "pipeline_f", "arapipel", "pipeline_data", "metso_pipe"
        Attrib = oPipeBlock.GetAttributes
        Teksti.Value = AttribToPipeline(Attrib)
      Case Else
        If oPipeBlock.HasAttributes Then
          Attrib = oPipeBlock.GetAttributes
          ' METSOPU001-pumppublokin tunniste on CUSTITEM; muilla blokeilla k‰ytet‰‰n paikkanumerotunnisteita
          For j = 0 To UBound(Attrib)
            Select Case Attrib(j).TagString
              Case "POS.NO", "CUSTPOS", "SDPOS", "CUSTITEM"
                Teksti.Value = Attrib(j).TextString
                Exit For
            End Select
          Next j
        End If
    End Select
    oAttrib.TextString = Nz(Teksti.Value, "")
  ElseIf Kohde.ObjectName = "AcDbText" Then
    Teksti.Value = Kohde.TextString
    oAttrib.TextString = Nz(Teksti.Value, "")
  End If
  
  AktivoiAccess
  Exit Sub
  
ErrorHandler:
  Virhe = Err.Description
  AktivoiAccess
  MsgBox "Virhe poimittaessa tietoa piirustuksesta: " & Virhe, vbExclamation
End Sub

' Nostetaan Access-ikkuna etualalle ikkunakahvan avulla, koska Accessin Application-oliolla ei ole Caption-ominaisuutta
Private Sub AktivoiAccess()
  On Error Resume Next
  SetForegroundWindow Application.hWndAccessApp
  Err.Clear
End Sub

'================================================================================
' Funktio: AttribToPipeline
' Tarkoitus: Poimii putkilinjaviittauksen blokin attribuuteista (DEP+FSB+LINE)
' Parametrit:
'   Attrib - Blokin attribuuttitaulukko
' Palautusarvo: Yhdistetty putkilinjaviittausmerkkijono
'================================================================================
Private Function AttribToPipeline(Attrib As Variant) As String
  Dim j As Long
  Dim DEP As String
  Dim FSB As String
  Dim LINE As String

  For j = 0 To UBound(Attrib)
    Select Case Attrib(j).TagString
      Case "DEP", "DEPA"
        DEP = Attrib(j).TextString
      Case "FSB", "FLOWSUBSTMAINGROUP"
        FSB = Attrib(j).TextString
      Case "LINE", "LINENO"
        LINE = Attrib(j).TextString
    End Select
  Next j
  AttribToPipeline = DEP & FSB & LINE
End Function

'================================================================================
' Tapahtuma: BFind_Click
' Tarkoitus: Etsii seuraavan putkilinjablokin, jolla FROM tai TO on tyhj‰
'================================================================================
Private Sub BFind_Click()
  On Error GoTo ErrorHandler
  
  Do
    mIndeksi = mIndeksi + 1
    If mIndeksi > Joukko.Count - 1 Then
      If MsgBox("Viimeinen blokki!" & vbCrLf & "Aloitetaanko alusta?", vbYesNo, "Etsi seuraava tyhj‰") = vbYes Then
        mIndeksi = 0
      Else
        mIndeksi = Joukko.Count - 1
        Exit Do
      End If
    End If
    AvaaBlokki
    ' Puuttuva attribuutti (Nothing) tulkitaan tyhj‰ksi
    If AttribFrom Is Nothing Or AttribTo Is Nothing Then Exit Do
    If AttribFrom.TextString = "" Or AttribTo.TextString = "" Then Exit Do
  Loop
  
  Exit Sub
ErrorHandler:
  MsgBox "Virhe seuraavan tyhj‰n blokin etsimisess‰: " & Err.Description, vbExclamation
End Sub

'================================================================================
' Tapahtuma: BNext_Click
' Tarkoitus: Navigoi seuraavaan putkilinjan blokkiin
'================================================================================
Private Sub BNext_Click()
  On Error GoTo ErrorHandler
  NextBlock
  Exit Sub
ErrorHandler:
  MsgBox "Virhe seuraavaan blokkiin navigoitaessa: " & Err.Description, vbExclamation
End Sub

'================================================================================
' Aliohjelma: NextBlock
' Tarkoitus: Siirtyy seuraavaan blokkiin valintajoukossa (kiert‰‰ alkuun)
'================================================================================
Private Sub NextBlock()
  On Error GoTo ErrorHandler
  
  mIndeksi = mIndeksi + 1
  If mIndeksi > Joukko.Count - 1 Then
    mIndeksi = 0
  End If
  AvaaBlokki
  
  Exit Sub
ErrorHandler:
  MsgBox "Virhe NextBlock-rutiinissa: " & Err.Description, vbExclamation
End Sub

'================================================================================
' Aliohjelma: KirjoitaAtt
' Tarkoitus: Kirjoittaa tekstin attribuuttiviittaukseen
' Parametrit:
'   Teksti - Kirjoitettava teksti
'   oAttrib - P‰ivitett‰v‰ attribuuttiviittaus
'================================================================================
Private Sub KirjoitaAtt(ByVal Teksti As String, oAttrib As Object)
  On Error GoTo ErrorHandler
  
  If oAttrib Is Nothing Then
    Exit Sub
  Else
    oAttrib.TextString = Teksti
  End If
  
  Exit Sub
ErrorHandler:
  MsgBox "Virhe attribuutin kirjoittamisessa: " & Err.Description, vbExclamation
End Sub

'================================================================================
' Tapahtuma: BPickFrom_Click
' Tarkoitus: Poimii FROM-viittauksen AutoCAD-piirustuksesta
'================================================================================
Private Sub BPickFrom_Click()
  On Error GoTo ErrorHandler
  
  OpenSource TFrom, AttribFrom
  AktivoiAccess
  
  Exit Sub
ErrorHandler:
  MsgBox "Virhe FROM-tiedon poiminnassa: " & Err.Description, vbExclamation
End Sub

'================================================================================
' Tapahtuma: BPickTo_Click
' Tarkoitus: Poimii TO-viittauksen AutoCAD-piirustuksesta
'================================================================================
Private Sub BPickTo_Click()
  On Error GoTo ErrorHandler
  
  OpenSource TTo, AttribTo
  AktivoiAccess
  
  Exit Sub
ErrorHandler:
  MsgBox "Virhe TO-tiedon poiminnassa: " & Err.Description, vbExclamation
End Sub

'================================================================================
' Tapahtuma: BPrev_Click
' Tarkoitus: Navigoi edelliseen putkilinjan blokkiin
'================================================================================
Private Sub BPrev_Click()
  On Error GoTo ErrorHandler
  PrevBlock
  Exit Sub
ErrorHandler:
  MsgBox "Virhe edelliseen blokkiin navigoitaessa: " & Err.Description, vbExclamation
End Sub

'================================================================================
' Aliohjelma: PrevBlock
' Tarkoitus: Siirtyy edelliseen blokkiin valintajoukossa (kiert‰‰ loppuun)
'================================================================================
Private Sub PrevBlock()
  On Error GoTo ErrorHandler
  
  mIndeksi = mIndeksi - 1
  If mIndeksi < 0 Then
    mIndeksi = Joukko.Count - 1
  End If
  AvaaBlokki
  
  Exit Sub
ErrorHandler:
  MsgBox "Virhe PrevBlock-rutiinissa: " & Err.Description, vbExclamation
End Sub

'================================================================================
' Tapahtuma: Form_Close
' Tarkoitus: Siivotaan AutoCAD-objektit lomakkeen sulkeutuessa
'================================================================================
Private Sub Form_Close()
  On Error Resume Next
  
  ' Poistetaan korostus ennen viittausten vapauttamista
  If Not oBlock Is Nothing Then oBlock.Highlight False
  
  ' Siivotaan valintajoukko
  If Not Joukko Is Nothing Then
    Joukko.Delete
    Set Joukko = Nothing
  End If
  
  ' Siivotaan AutoCAD-objektit
  Set AttribFrom = Nothing
  Set AttribTo = Nothing
  Set oBlock = Nothing
  Set oDOC = Nothing
  Set oACAD = Nothing
End Sub

'================================================================================
' Tapahtuma: Form_Load
' Tarkoitus: Alustetaan lomake, yhdistet‰‰n AutoCADiin, suodatetaan putkilinjablokit
'================================================================================
Private Sub Form_Load()
  Dim FT(0) As Integer  ' DXF-ryhm‰koodin on oltava Integer (16-bittinen), ei Long
  Dim FD(0) As Variant
  Dim Poista() As Object  ' AcadEntity
  Dim j As Long, L As Long
  Dim Nimi As String
  Dim Virhe As String
  
  On Error GoTo ErrorHandler
  
  If Not Yhdista() Then
    mSuljeLataus = True
    Me.TimerInterval = 10
    Exit Sub
  End If
  
  FT(0) = 0
  FD(0) = "INSERT"
  TeeJoukko "APUVALINTA", Joukko
  If Joukko Is Nothing Then Err.Raise vbObjectError + 1001, , "Valintajoukkoa ei voitu luoda."
  Joukko.Select acSelectionSetAll, , , FT, FD
  
  ' Suodatus: pidet‰‰n vain putkilinjablokit
  For j = 0 To Joukko.Count - 1
    Nimi = UCase$(Joukko(j).EffectiveName)
    Select Case Nimi
      Case "PIPELINE", "PIPELINE_F", "ARAPIPEL", "PIPELINE_DATA", "METSO_PIPE"
      Case Else
        ReDim Preserve Poista(L)
        Set Poista(L) = Joukko(j)
        L = L + 1
    End Select
  Next j
  
  If L > 0 Then Joukko.RemoveItems Poista
  
  If Joukko.Count = 0 Then
    MsgBox "Putkilinjojen blokkeja ei lˆytynyt.", vbExclamation
    mSuljeLataus = True
    Me.TimerInterval = 10
    Exit Sub
  End If
  
  mIndeksi = 0
  AvaaBlokki
  Exit Sub
  
ErrorHandler:
  Virhe = Err.Description
  MsgBox "Virhe lomakkeen latauksessa: " & Virhe, vbExclamation
  mSuljeLataus = True
  Me.TimerInterval = 10
End Sub

'================================================================================
' Aliohjelma: TeeJoukko
' Tarkoitus: Luo tai uudelleenluo nimetty valintajoukko
' Parametrit:
'   Nimi - Valintajoukon nimi
'   VJoukko - Luotava valintajoukko-objekti
'================================================================================
Private Sub TeeJoukko(Nimi As String, VJoukko As Object)  ' AcadSelectionSet
  Dim i As Long
  
  On Error GoTo ErrorHandler
  
  ' Poistetaan, jos on jo olemassa
  For i = 0 To oDOC.SelectionSets.Count - 1
    If oDOC.SelectionSets(i).Name = Nimi Then
      oDOC.SelectionSets(i).Delete
      Exit For
    End If
  Next i
  
  Set VJoukko = oDOC.SelectionSets.Add(Nimi)
  Exit Sub
  
ErrorHandler:
  MsgBox "Virhe valintajoukon luomisessa: " & Err.Description, vbExclamation
End Sub

'================================================================================
' Aliohjelma: AvaaBlokki
' Tarkoitus: N‰ytt‰‰ ja korostaa nykyisen putkilinjablokin ja lataa attribuutit
'================================================================================
Private Sub AvaaBlokki()
  Dim Attrib As Variant
  Dim MinPoint As Variant
  Dim MaxPoint As Variant
  Dim j As Long
  
  On Error GoTo ErrorHandler
  
  Me.Caption = (mIndeksi + 1) & "/" & Joukko.Count
  
  ' Poistetaan edellinen korostus; blokki on voitu poistaa piirustuksesta
  If Not oBlock Is Nothing Then
    On Error Resume Next
    oBlock.Highlight False
    Err.Clear
    On Error GoTo ErrorHandler
  End If
  
  ' Nollataan edellisen blokin attribuuttiviittaukset, jotta kirjoitus ei mene v‰‰r‰‰n blokkiin
  Set AttribFrom = Nothing
  Set AttribTo = Nothing
  TFrom.Value = Null
  TTo.Value = Null
  LPipeLine.Caption = ""
  
  ' Haetaan ja n‰ytet‰‰n nykyinen blokki
  Set oBlock = Joukko(mIndeksi)
  oBlock.GetBoundingBox MinPoint, MaxPoint
  oACAD.ZoomWindow MinPoint, MaxPoint
  oACAD.ZoomScaled 0.2, acZoomScaledRelative
  oBlock.Highlight True
  Attrib = oBlock.GetAttributes
  LPipeLine.Caption = AttribToPipeline(Attrib)
  For j = 0 To UBound(Attrib)
    If Attrib(j).TagString = "FROM" Then
      Set AttribFrom = Attrib(j)
      TFrom.Value = Attrib(j).TextString
    ElseIf Attrib(j).TagString = "TO" Then
      Set AttribTo = Attrib(j)
      TTo.Value = Attrib(j).TextString
    End If
  Next j
  
  Exit Sub
ErrorHandler:
  MsgBox "Virhe blokin n‰ytt‰misess‰: " & Err.Description, vbExclamation
End Sub

'================================================================================
' Funktio: Yhdista
' Tarkoitus: Yhdist‰‰ k‰ynniss‰ olevaan AutoCAD-instanssiin
' Palautusarvo: True jos yhteys onnistui, False muuten
'================================================================================
Private Function Yhdista() As Boolean
  On Error Resume Next
  
  ' Yhdistet‰‰n k‰ynniss‰ olevaan AutoCADiin
  Set oACAD = GetObject(, "AutoCAD.Application")
  If Err.Number <> 0 Or oACAD Is Nothing Then
    Err.Clear
    On Error GoTo 0
    MsgBox "K‰ynniss‰ olevaa AutoCADi‰ ei lˆytynyt!", vbCritical, "Virhe!"
    Set oACAD = Nothing
    Exit Function
  End If
  
  If oACAD.Documents.Count = 0 Then
    Err.Clear
    On Error GoTo 0
    MsgBox "AutoCADissa ei ole avointa piirustusta!", vbCritical, "Virhe!"
    Set oACAD = Nothing
    Exit Function
  End If
  
  Set oDOC = oACAD.ActiveDocument
  If Err.Number <> 0 Or oDOC Is Nothing Then
    Err.Clear
    On Error GoTo 0
    MsgBox "AutoCADin aktiivista piirustusta ei voitu hakea.", vbCritical, "Virhe!"
    Set oDOC = Nothing
    Set oACAD = Nothing
    Exit Function
  End If
  
  On Error GoTo 0
  Yhdista = True
End Function

'================================================================================
' Tapahtuma: TFrom_AfterUpdate
' Tarkoitus: P‰ivitt‰‰ FROM-attribuutin tekstikent‰n arvon muuttuessa
'================================================================================
Private Sub TFrom_AfterUpdate()
  On Error GoTo ErrorHandler
  
  KirjoitaAtt Nz(TFrom.Value), AttribFrom
  
  Exit Sub
ErrorHandler:
  MsgBox "Virhe FROM-attribuutin p‰ivityksess‰: " & Err.Description, vbExclamation
End Sub

'================================================================================
' Tapahtuma: TTo_AfterUpdate
' Tarkoitus: P‰ivitt‰‰ TO-attribuutin tekstikent‰n arvon muuttuessa
'================================================================================
Private Sub TTo_AfterUpdate()
  On Error GoTo ErrorHandler
  
  KirjoitaAtt Nz(TTo.Value), AttribTo
  
  Exit Sub
ErrorHandler:
  MsgBox "Virhe TO-attribuutin p‰ivityksess‰: " & Err.Description, vbExclamation
End Sub

