Option Compare Database
Option Explicit

'================================================================================
' Lomake: USysPipeToOther
' Tarkoitus: Linkitä laiteblokit putkilinjoihin PIPELINE-attribuutin kautta
' Päivitetty: 2025-11-12 - VBA7/64-bit-tuki ja optimointi lisätty
'
' Kuvaus:
'   Interaktiivinen lomake putkilinjaviittausten asettamiseen laiteblokeille:
'   - Yhdistyy käynnissä olevaan AutoCAD-instanssiin
'   - Suodattaa blokit tyypin mukaan (OpenArgs-parametrin avulla)
'   - Mahdollistaa navigoinnin blokkien välillä (Seuraava, Edellinen, Etsi tyhjä)
'   - Poimi putkilinja piirustuksesta tai syötä käsin
'   - Päivittää PIPELINE/PIPELINENO-attribuutin reaaliajassa
'
' Riippuvuudet:
'   - AutoCAD-sovellus (COM-automaatio)
'   - Laiteblokit, joissa on PIPELINE- tai PIPELINENO-attribuutti
'
' OpenArgs:
'   - Blokin nimi (esim. "VALVE") - suodattaa yhteen blokkityyppiin
'   - Pilkuilla erotettu lista (esim. "VALVE,PUMP") - suodattaa useampaan tyyppiin
'
' Kontrollit:
'   - TPipeline: Putkilinjaviittauksen tekstikenttä
'   - BNext, BPrev: Navigointipainikkeet
'   - BFind: Etsi seuraava tyhjä -painike
'   - BPickPipeline: Poimi putkilinja piirustuksesta -painike
'================================================================================

Dim oACAD As Object  ' AcadApplication - muutettu late binding (64-bit)
Dim oDOC As Object   ' AcadDocument - muutettu late binding (64-bit)
Dim PipeAttrib As Object  ' AcadAttributeReference - muutettu late binding (64-bit)
Dim oBlock As Object  ' AcadBlockReference - muutettu late binding (64-bit)
Dim Joukko As Object  ' AcadSelectionSet - muutettu late binding (64-bit)
Dim i As Integer  ' Nykyinen blokkiindeksi

'================================================================================
' Tapahtuma: BClose_Click
' Tarkoitus: Sulkee lomakkeen
'================================================================================
Private Sub BClose_Click()
  DoCmd.Close acForm, "UsysPipeToOther"
End Sub

'================================================================================
' Aliohjelma: OpenPipeline
' Tarkoitus: Poimii putkilinja- tai levyblokin piirroksesta ja palauttaa viittauksen
'================================================================================
Private Sub OpenPipeline()
  Dim Kohde As Object
  Dim Piste As Variant
  Dim oPipeBlock As Object  ' AcadBlockReference - muutettu late binding (64-bit)
  Dim Attrib As Variant
  Dim Putkilinja As String
  Dim OK As Boolean
  Dim j As Integer
  Dim Nimi As String
  
  On Error Resume Next
  AppActivate oACAD.Caption, True
  If Err <> 0 Then
    MsgBox "AutoCAD-ikkunaa ei voitu aktivoida: " & Err.Description, vbExclamation
    Err.Clear
    Exit Sub
  End If
  
  OK = False
  Do
    oDOC.Utility.GetEntity Kohde, Piste, "Poimi putkilinja: "
    If Err <> 0 Then
      Err.Clear
      Exit Do
    Else
      If Kohde.ObjectName = "AcDbBlockReference" Then
        Set oPipeBlock = Kohde
        Nimi = UCase$(oPipeBlock.EffectiveName)
        If Nimi = "PIPELINE_F" Or Nimi = "PIPELINE" Or Nimi = "PIPELINE_DATA" Or Nimi = "METSO_PIPE" Or Nimi = "BAHPIPEL" Then
          Attrib = oPipeBlock.GetAttributes
          Putkilinja = AttribToPipeline(Attrib)
          TPipeline.Value = Putkilinja
          PipeAttrib.TextString = Putkilinja
          OK = True
        ElseIf Left$(Nimi, 6) = "_PLATE" Then
          Attrib = oPipeBlock.GetAttributes
          TPipeline.Value = Attrib(1).TextString
          PipeAttrib.TextString = Attrib(1).TextString
          OK = True
        End If
      End If
    End If
    If OK Then
      If i < Joukko.Count - 1 Then
        NextBlock
      Else
        Exit Do
      End If
    End If
  Loop
  On Error GoTo 0
  AppActivate CurrentDb.Properties("AppTitle").Value
End Sub

'================================================================================
' Funktio: AttribToPipeline
' Tarkoitus: Poimii putkilinjaviittauksen blokin attribuuteista (DEP+FSB+LINE)
' Parametrit:
'   Attrib - Blokin attribuuttitaulukko
' Palautusarvo: Yhdistetty putkilinjaviittausmerkkijono
'================================================================================
Private Function AttribToPipeline(Attrib As Variant) As String
  Dim j As Integer
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
      Case Else
      
    End Select
  Next j
  AttribToPipeline = DEP & FSB & LINE
End Function

'================================================================================
' Tapahtuma: BFind_Click
' Tarkoitus: Etsii seuraavan laiteblokin, jolla PIPELINE-attribuutti on tyhjÃ¤
'================================================================================
Private Sub BFind_Click()
  On Error GoTo ErrorHandler
  
  Do
    i = i + 1
    If i > Joukko.Count - 1 Then
      If MsgBox("Viimeinen blokki!" & vbCrLf & "Aloitetaanko alusta?", vbYesNo, "Etsi seuraava tyhjä") = vbYes Then
         i = 0
      Else
        Exit Do
      End If
    End If
    AvaaBlokki
    If TPipeline.Value = "" Then Exit Do
  Loop
  
  Exit Sub
ErrorHandler:
  MsgBox "Virhe seuraavan tyhjän blokin etsimisessä: " & Err.Description, vbExclamation
End Sub

'================================================================================
' Tapahtuma: BNext_Click
' Tarkoitus: Navigoi seuraavaan laitteen blokkiin
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
' Tarkoitus: Siirtyy seuraavaan blokkiin valintajoukossa (kiertÃ¤Ã¤ alkuun)
'================================================================================
Private Sub NextBlock()
  On Error GoTo ErrorHandler
  
  i = i + 1
  If i > Joukko.Count - 1 Then
    i = 0
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
'   Teksti - Text to write
'   oAttrib - PÃ¤ivitettÃ¤vÃ¤ attribuuttiviittaus
'================================================================================
Private Sub KirjoitaAtt(Teksti As String, oAttrib As Object)  ' AcadAttributeReference - muutettu late binding (64-bit)
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
' Tapahtuma: BPickPipeline_Click
' Tarkoitus: Poimii putkilinjaviittauksen AutoCAD-piirustuksesta
'================================================================================
Private Sub BPickPipeline_Click()
  On Error GoTo ErrorHandler
  
  OpenPipeline
  
  Exit Sub
ErrorHandler:
  MsgBox "Virhe putkilinjan poiminnassa: " & Err.Description, vbExclamation
End Sub

'================================================================================
' Tapahtuma: BPrev_Click
' Tarkoitus: Navigoi edelliseen laitteen blokkiin
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
' Tarkoitus: Siirtyy edelliseen blokkiin valintajoukossa (kiertÃ¤Ã¤ loppuun)
'================================================================================
Private Sub PrevBlock()
  On Error GoTo ErrorHandler
  
  i = i - 1
  If i < 0 Then
    i = Joukko.Count - 1
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
  
  ' Siivotaan valintajoukko
  If Not Joukko Is Nothing Then
    Joukko.Delete
    Set Joukko = Nothing
  End If
  
  ' Siivotaan AutoCAD-objektit
  Set PipeAttrib = Nothing
  Set oBlock = Nothing
  Set oDOC = Nothing
  Set oACAD = Nothing
End Sub

'================================================================================
' Tapahtuma: Form_Load
' Tarkoitus: Alustetaan lomake, yhdistetÃ¤Ã¤n AutoCADiin, suodatetaan laiteblokit
'================================================================================
Private Sub Form_Load()
  Dim FT(0) As Integer
  Dim FD(0) As Variant
  Dim Poista() As Object  ' AcadEntity - muutettu late binding (64-bit)
  Dim j As Integer, L As Integer
  Dim Nimi As String
  
  On Error GoTo ErrorHandler
  
  If Yhdista = False Then
    BClose_Click
    Exit Sub
  Else
    If Err = 0 Then
       TeeJoukko "APUVALINTA", Joukko
       
       ' Suodatus OpenArgs-parametrin mukaan: pilkuilla erotettu lista tai yksittäinen blokkityyppi
       If InStr(OpenArgs, ",") Then
          FT(0) = 2
          FD(0) = OpenArgs
          Joukko.Select acSelectionSetAll, , , FT, FD
       Else
          FT(0) = 0
          FD(0) = "INSERT"
          Joukko.Select acSelectionSetAll, , , FT, FD
          For j = 0 To Joukko.Count - 1
            Nimi = UCase$(Joukko(j).EffectiveName)
            If Nimi <> UCase$(OpenArgs) Then
              ReDim Preserve Poista(L)
              Set Poista(L) = Joukko(j)
              L = L + 1
            End If
          Next j
          If L > 0 Then Joukko.RemoveItems Poista
       End If
       
       If Joukko.Count = 0 Then
         MsgBox "Putkilinjojen blokkeja ei löytynyt.", vbExclamation
         Err.Raise vbObjectError + 1000
       Else
         AvaaBlokki
       End If
     End If
  End If
  Exit Sub
  
ErrorHandler:
  MsgBox "Virhe lomakkeen latauksessa: " & Err.Description, vbExclamation
  BClose_Click
End Sub

'================================================================================
' Aliohjelma: TeeJoukko
' Tarkoitus: Luo tai uudelleenluo nimetty valintajoukko
' Parametrit:
'   Nimi - Valintajoukon nimi
'   VJoukko - Luotava valintajoukko-objekti
'================================================================================
Private Sub TeeJoukko(Nimi As String, VJoukko As Object)  ' AcadSelectionSet - muutettu late binding (64-bit)
  Dim i As Integer
  
  On Error GoTo ErrorHandler
  
  ' Poistetaan, jos on jo olemassa
  For i = 0 To ActiveDocument.SelectionSets.Count - 1
    If ActiveDocument.SelectionSets(i).Name = Nimi Then
      ActiveDocument.SelectionSets(i).Delete
      Exit For
    End If
  Next i
  
  Set VJoukko = ActiveDocument.SelectionSets.Add(Nimi)
  Exit Sub
  
ErrorHandler:
  MsgBox "Virhe valintajoukon luomisessa: " & Err.Description, vbExclamation
End Sub

'================================================================================
' Aliohjelma: AvaaBlokki
' Tarkoitus: NÃ¤yttÃ¤Ã¤ ja korostaa nykyisen laiteblokin ja lataa PIPELINE-attribuutin
'================================================================================
Private Sub AvaaBlokki()
  Dim Attrib As Variant
  Dim MinPoint As Variant
  Dim MaxPoint As Variant
  Dim j As Integer
  
  On Error GoTo ErrorHandler
  
  Me.Caption = i + 1 & "/" & Joukko.Count
  
  ' Poistetaan edellinen korostus
  If Not oBlock Is Nothing Then
    oBlock.Highlight False
  End If
  
  ' Haetaan ja näytetään nykyinen blokki
  Set oBlock = Joukko(i)
  oBlock.GetBoundingBox MinPoint, MaxPoint
  oACAD.ZoomWindow MinPoint, MaxPoint
  oACAD.ZoomScaled 0.2, acZoomScaledRelative
  oBlock.Highlight True
  
  ' Haetaan PIPELINE-attribuutti
  Attrib = oBlock.GetAttributes
  Set PipeAttrib = Nothing
  For j = 0 To UBound(Attrib)
    If Attrib(j).TagString = "PIPELINE" Or Attrib(j).TagString = "PIPELINENO" Then
      Set PipeAttrib = Attrib(j)
      TPipeline.Value = PipeAttrib.TextString
    End If
  Next j
  
  If PipeAttrib Is Nothing Then
    MsgBox "Blokista ei löytynyt 'PIPELINE'-attribuuttia.", vbCritical, "VIRHE"
  End If
  
  Exit Sub
ErrorHandler:
  MsgBox "Virhe blokin näyttämisessä: " & Err.Description, vbExclamation
End Sub

'================================================================================
' Funktio: Yhdista
' Tarkoitus: YhdistÃ¤Ã¤ kÃ¤ynnissÃ¤ olevaan AutoCAD-instanssiin
' Palautusarvo: True jos yhteys onnistui, False muuten
'================================================================================
Private Function Yhdista() As Boolean
  On Error Resume Next
  
  ' Yhdistetään käynnissä olevaan AutoCADiin
  Set oACAD = GetObject(, "AutoCAD.Application")
  If Err <> 0 Then
    MsgBox "Käynnissä olevaa AutoCADiä ei löytynyt!", vbCritical, "Virhe!"
    Set oACAD = Nothing
    Yhdista = False
    Exit Function
  Else
    Set oDOC = oACAD.ActiveDocument
    Yhdista = True
  End If
  On Error GoTo 0
End Function

'================================================================================
' Tapahtuma: TPipeline_AfterUpdate
' Tarkoitus: PÃ¤ivittÃ¤Ã¤ PIPELINE-attribuutin tekstikentÃ¤n arvon muuttuessa
'================================================================================
Private Sub TPipeline_AfterUpdate()
  On Error GoTo ErrorHandler
  
  KirjoitaAtt Nz(TPipeline.Value), PipeAttrib
  
  Exit Sub
ErrorHandler:
  MsgBox "Virhe PIPELINE-attribuutin päivityksessä: " & Err.Description, vbExclamation
End Sub
