Option Compare Database
Option Explicit

'================================================================================
' Lomake: USysFlowPickNo
' Tarkoitus: Virtausblokkien putkilinjanumerovalitsin AutoCAD-piirustuksille
' P‰ivitetty: 2025-11-12 - VBA7/64-bit-tuki ja optimointi lis‰tty
'
' Kuvaus:
'   Interaktiivinen lomake tasearvoblokkien putkilinjanumeroiden asettamiseen:
'   - Yhdistyy k‰ynniss‰ olevaan AutoCAD-instanssiin
'   - Etsii kaikki virtausblokit (PDUFLOWDATA, ARAFLOW, FLOWBLOCK, BALANCE_DATA)
'   - Korostaa jokaisen blokin j‰rjestyksess‰
'   - K‰ytt‰j‰ voi poimia putkilinjan piirustuksesta tai syˆtt‰‰ k‰sin
'   - P‰ivitt‰‰ blokin PIPEREF/PIPELINENO/PIPE_LINE-attribuutin
'
' Riippuvuudet:
'   - AutoCAD-sovellus (COM-automaatio)
'   - Virtausblokit, joissa on putkilinjaviittausattribuutit
'
' Kontrollit:
'   - TRefNo: Putkilinjaviittauksen tekstikentt‰
'   - BSeuraava: Seuraava blokki -painike
'   - BEdellinen: Edellinen blokki -painike
'   - BSeurT: Etsi seuraava tyhj‰ -painike
'   - BPoimi: Poimi piirustuksesta -painike
'   - BAutohaku: Automaattinen haku -painike
'   - BSulje: Sulje lomake -painike
'================================================================================

Dim oACAD As Object  ' AcadApplication - muutettu late binding (64-bit)
Dim oDOC As Object   ' AcadDocument - muutettu late binding (64-bit)
Dim oAttrib As Object  ' AcadAttributeReference - muutettu late binding (64-bit)
Dim Joukko As Object   ' AcadSelectionSet - muutettu late binding (64-bit)
Dim j As Integer  ' Nykyinen blokkiindeksi
Dim EiLinjaa As Boolean  ' Flag if pipeline not found during auto-search

'--------------------------------------------------------------------------------
' Tapahtuma: Form_Close
' Tarkoitus: Siivotaan AutoCAD-valintajoukko lomakkeen sulkeutuessa
'--------------------------------------------------------------------------------
Private Sub Form_Close()
On Error Resume Next
  If Not Joukko Is Nothing Then
    Joukko.Delete
    Set Joukko = Nothing
  End If
  On Error GoTo 0
End Sub

'--------------------------------------------------------------------------------
' Tapahtuma: Form_Load
' Tarkoitus: Alustaa lomakkeen - yhdist‰‰ AutoCADiin ja luo virtausblokkien valintajoukko
'
' Toiminta:
'   1. Yhdistet‰‰n k‰ynniss‰ olevaan AutoCADiin (Yhdista-funktio)
'   2. Poistetaan olemassa oleva APUPICK-valintajoukko
'   3. Luodaan uusi valintajoukko virtausblokkiluodatilla
'   4. N‰ytet‰‰n ensimm‰inen blokki
'
' Huom:
'   - Suljetaan lomake, jos AutoCAD ei ole k‰ynniss‰
'   - Suljetaan lomake, jos virtausblokkeja ei lˆydy
'   - Suodatin: PDUFLOWDATA, ARAFLOW, FLOWBLOCK, BALANCE_DATA -blokit
'--------------------------------------------------------------------------------
Private Sub Form_Load()
On Error GoTo ErrorHandler
Dim i As Integer
Dim FilterType(0) As Integer
Dim FilterData(0) As Variant

' Yhdistet‰‰n AutoCADiin
If Yhdista = False Then
  DoCmd.Close acForm
  Exit Sub
End If

' Poistetaan olemassa oleva valintajoukko tarvittaessa
For i = 0 To oDOC.SelectionSets.Count - 1
  If oDOC.SelectionSets(i).Name = "APUPICK" Then
    oDOC.SelectionSets(i).Delete
    Exit For
  End If
Next i

' Luodaan valintajoukko virtausblokeille
FilterType(0) = 2  ' Blokin nimisuodatin
FilterData(0) = "PDUFLOWDATA,ARAFLOW,FLOWBLOCK,BALANCE_DATA"
Set Joukko = oDOC.SelectionSets.Add("APUPICK")
Joukko.Select acSelectionSetAll, , , FilterType, FilterData

If Joukko.Count = 0 Then
  MsgBox "Ei tasearvoblokkeja", vbInformation, "FlowPickNo"  ' "No flow blocks found"
  DoCmd.Close
  Exit Sub
End If

' Alustetaan ensimm‰isell‰ blokilla
j = 0
AvaaBlokki
Exit Sub

ErrorHandler:
    MsgBox "Virhe lomakkeen alustuksessa: " & Err.Description, vbCritical, "FlowPickNo"
    DoCmd.Close acForm
End Sub


'--------------------------------------------------------------------------------
' Tapahtuma: BAutohaku_Click
' Tarkoitus: Lˆyt‰‰ automaattisesti kaikki putkilinjat virtausblokeille
'
' Toiminta:
'   1. K‰yd‰‰n l‰pi jokainen virtausblokki valintajoukossa
'   2. Kutsutaan EtsiLinja etsim‰‰n l‰heisin putkilinja
'   3. Seurataan, lˆytyikˆ kaikille blokeille putkilinja
'   4. N‰ytet‰‰n yhteenviesti
'
' Huom:
'   - K‰y automaattisesti l‰pi kaikki blokit
'   - Asettaa EiLinjaa-lipun, jos jollekin blokille ei lˆydy putkilinjaa
'--------------------------------------------------------------------------------
Private Sub BAutohaku_Click()
On Error GoTo ErrorHandler
EiLinjaa = False
For j = 0 To Joukko.Count - 1
  AvaaBlokki
  EtsiLinja
Next j
If EiLinjaa Then
  MsgBox "Kaikki linjat on k‰yty l‰pi. Yksi tai useampia virtausarvoblokkeja j‰i ilman putkilinjanumeroa.", vbOKOnly, "Auto-haku"
  BSeurT_Click
Else
  MsgBox "Kaikki linjat on k‰yty l‰pi. Kaikkiin virtausarvoblokkeihin lˆytyi putkilinjanumero.", vbOKOnly, "Auto-haku"
End If
Exit Sub

ErrorHandler:
    MsgBox "Virhe automaattisessa haussa: " & Err.Description, vbCritical, "Auto-haku"
End Sub
'--------------------------------------------------------------------------------
' Tapahtuma: BPoimi_Click
' Tarkoitus: Poimii putkilinjan AutoCAD-piirustuksesta klikkaamalla blokkia
'
' Toiminta:
'   1. Aktivoidaan AutoCAD-ikkuna
'   2. Pyydet‰‰n k‰ytt‰j‰‰ poimiman putkilinjablokki
'   3. Poimitaan putkilinjaviittaus blokin attribuuteista
'   4. P‰ivitet‰‰n nykyisen virtausblokin putkilinja-attribuutti
'
' Huom:
'   - Toimii PIPELINE, PIPELINE_F, ARAPIPEL, BAHPIPEL, METSO_PIPE -blokkien kanssa
'   - K‰ytt‰‰ POS.NO, CUSTPOS, SDPOS -attribuutteja muille blokeille
'   - K‰ytt‰‰ AttribToPipeline-apufunktiota standardiputkilinjoille
'--------------------------------------------------------------------------------
Private Sub BPoimi_Click()
On Error GoTo ErrorHandler
Dim Kohde As Object
Dim Piste As Variant
Dim oBlock As Object  ' AcadBlockReference - muutettu late binding (64-bit)
Dim Attrib As Variant

AppActivate oACAD.Caption, True
On Error Resume Next
oDOC.Utility.GetEntity Kohde, Piste, "Poimi putkilinja..."
If Err <> 0 Then
  Err.Clear
  AppActivate CurrentDb.Properties("AppTitle").Value, True
  Exit Sub
End If
On Error GoTo ErrorHandler

If Kohde.ObjectName = "AcDbBlockReference" Then
  Set oBlock = Kohde
  Select Case LCase$(oBlock.EffectiveName)
    Case "pipeline", "pipeline_f", "arapipel", "bahpipel", "metso_pipe"
       Attrib = oBlock.GetAttributes
       TRefNo.Value = AttribToPipeline(Attrib)
    Case Else
      If oBlock.HasAttributes Then
        Attrib = oBlock.GetAttributes
        For j = 0 To UBound(Attrib)
          If Attrib(j).TagString = "POS.NO" Or Attrib(j).TagString = "CUSTPOS" Or Attrib(j).TagString = "SDPOS" Then
            TRefNo.Value = Attrib(j).TextString
            Exit For
          End If
        Next j
      End If
  End Select
  oAttrib.TextString = TRefNo.Value
End If
AppActivate CurrentDb.Properties("AppTitle").Value, True
Exit Sub

ErrorHandler:
    On Error Resume Next
    AppActivate CurrentDb.Properties("AppTitle").Value, True
    On Error GoTo 0
    MsgBox "Virhe putkilinjan poiminnassa: " & Err.Description, vbCritical, "BPoimi"
End Sub
'--------------------------------------------------------------------------------
' Funktio: AttribToPipeline
' Tarkoitus: Poimii putkilinjaviittauksen putkilinjablokin attribuuteista
'
' Parametrit:
'   Attrib - GetAttributes-funktion palauttama taulukko
'
' Palautusarvo: Merkkijono - muotoiltu putkilinjaviittaus (DEP & FSB & LINE)
'
' Toiminta:
'   1. Skannataan attribuutit: DEP/DEPA, FSB, LINE -tagit
'   2. Yhdistet‰‰n arvot ilman erotinmerkkej‰
'
' Huom:
'   - Palauttaa yhdistetyn merkkijonon (esim. "AREA01L001")
'   - K‰sittelee sek‰ DEP- ett‰ DEPA-attribuuttitagit
'--------------------------------------------------------------------------------
Private Function AttribToPipeline(Attrib As Variant) As String
On Error GoTo ErrorHandler
Dim j As Integer
Dim DEP As String
Dim FSB As String
Dim LINE As String

For j = 0 To UBound(Attrib)
  Select Case Attrib(j).TagString
    Case "DEP", "DEPA"
      DEP = Attrib(j).TextString
    Case "FSB"
      FSB = Attrib(j).TextString
    Case "LINE"
      LINE = Attrib(j).TextString
  End Select
Next j
AttribToPipeline = DEP & FSB & LINE
Exit Function

ErrorHandler:
    AttribToPipeline = ""
End Function
'--------------------------------------------------------------------------------
' Tapahtuma: BSeurT_Click
' Tarkoitus: Etsii seuraavan virtausblokin, jolla putkilinjaviittaus on tyhj‰
'
' Toiminta:
'   1. Poistetaan nykyisen blokin korostus
'   2. Siirryt‰‰n seuraavaan blokkiin
'   3. Tarkistetaan, onko putkilinjaviittaus tyhj‰
'   4. Toistetaan, kunnes tyhj‰ lˆytyy tai kaikki blokit on k‰yty
'
' Huom:
'   - Ohitetaan blokit, joilla on jo putkilinjaviittaus
'   - Pys‰hdyt‰‰n ensimm‰iseen tyhj‰‰n blokkiin
'--------------------------------------------------------------------------------
Private Sub BSeurT_Click()
On Error GoTo ErrorHandler
Dim i As Integer
  Do
    If j < Joukko.Count Then Joukko(j).Highlight False
    j = j + 1
    If j >= Joukko.Count Then j = 0
    AvaaBlokki
    If TRefNo.Value = "" Then
      Exit Do
    ElseIf j = Joukko.Count - 1 Then
      MsgBox "Sellaisia virtausarvoblokkeja ei lˆytynyt, joissa putkilinja olisi tyhj‰.", vbInformation, "Seur.Tyhj‰"
      Exit Do
    End If
  Loop
  Exit Sub

ErrorHandler:
    MsgBox "Virhe seuraavan tyhj‰n etsimisess‰: " & Err.Description, vbCritical, "Seur.Tyhj‰"
End Sub
'--------------------------------------------------------------------------------
' Tapahtuma: BEdellinen_Click
' Tarkoitus: Navigoi edelliseen virtausblokkiin
'
' Huom:
'   - Siirtyy viimeiseen blokkiin, jos ollaan alussa
'--------------------------------------------------------------------------------
Private Sub BEdellinen_Click()
On Error GoTo ErrorHandler
  Joukko(j).Highlight False
  If j > 0 Then
    j = j - 1
  Else
    j = Joukko.Count - 1
  End If
  AvaaBlokki
  Exit Sub

ErrorHandler:
    MsgBox "Virhe edelliseen navigoitaessa: " & Err.Description, vbCritical, "Edellinen"
End Sub

'--------------------------------------------------------------------------------
' Tapahtuma: BSeuraava_Click
' Tarkoitus: Navigoi seuraavaan virtausblokkiin
'
' Huom:
'   - Siirtyy ensimm‰iseen blokkiin, jos ollaan lopussa
'--------------------------------------------------------------------------------
Private Sub BSeuraava_Click()
On Error GoTo ErrorHandler
  Joukko(j).Highlight False
  If j < Joukko.Count - 1 Then
    j = j + 1
  Else
    j = 0
  End If
  AvaaBlokki
  Exit Sub

ErrorHandler:
    MsgBox "Virhe seuraavaan navigoitaessa: " & Err.Description, vbCritical, "Seuraava"
End Sub

'--------------------------------------------------------------------------------
' Tapahtuma: BSulje_Click
' Tarkoitus: Sulkee lomakkeen
'--------------------------------------------------------------------------------
Private Sub BSulje_Click()
  DoCmd.Close acForm, "USysFlowPickNo"
End Sub

'--------------------------------------------------------------------------------
' Aliohjelma: AvaaBlokki
' Tarkoitus: N‰ytt‰‰ ja zoomaa nykyisen virtausblokin AutoCADissa
'
' Toiminta:
'   1. Korostetaan nykyinen blokki valintajoukossa
'   2. P‰ivitet‰‰n lomakkeen otsikko blokkim‰‰r‰ll‰
'   3. Zoomataan blokin sijaintiin
'   4. Etsit‰‰n ja asetetaan putkilinja-attribuuttiviittaus
'   5. N‰ytet‰‰n nykyinen putkilinja-arvo TRefNo-kent‰ss‰
'
' Huom:
'   - Etsii PIPEREF-, PIPELINENO- tai PIPE_LINE-attribuutin
'   - N‰ytt‰‰ virheen, jos attribuuttia ei lˆydy
'   - P‰ivitt‰‰ otsikon muodossa "BlockName: X/Total"
'--------------------------------------------------------------------------------
Sub AvaaBlokki()
On Error GoTo ErrorHandler
Dim Attrib As Variant
Dim MinPoint As Variant
Dim MaxPoint As Variant
Dim Loytyi As Boolean
Dim ii As Integer

' Tarkistetaan, ett‰ indeksi on kelvollinen
If j >= Joukko.Count Or j < 0 Then j = 0

Joukko(j).Highlight True

' P‰ivitet‰‰n lomakkeen otsikko
Me.Caption = Joukko(j).Name & ":  " & j + 1 & "/" & Joukko.Count

' Zoomataan blokkiin
Joukko(j).GetBoundingBox MinPoint, MaxPoint
oACAD.ZoomCenter Joukko(j).InsertionPoint, 200

' Etsit‰‰n putkilinja-attribuutti
Attrib = Joukko(j).GetAttributes
Loytyi = False
For ii = 0 To UBound(Attrib)
  If UCase$(Attrib(ii).TagString) = "PIPEREF" Or UCase$(Attrib(ii).TagString) = "PIPELINENO" Or UCase$(Attrib(ii).TagString) = "PIPE_LINE" Then
    Set oAttrib = Attrib(ii)
    Loytyi = True
    Exit For
  End If
Next ii

If Loytyi = False Then
  MsgBox "Tasearvoblokki ei sis‰lt‰nyt attribuuttia putkilinjanumerolle.", vbCritical, "FlowPickNo"
  DoCmd.Close
  Exit Sub
End If

TRefNo.Value = oAttrib.TextString
Exit Sub

ErrorHandler:
    MsgBox "Virhe blokin avaamisessa: " & Err.Description, vbCritical, "AvaaBlokki"
    DoCmd.Close acForm
End Sub
'--------------------------------------------------------------------------------
' Funktio: Yhdista
' Tarkoitus: Yhdist‰‰ k‰ynniss‰ olevaan AutoCAD-instanssiin
'
' Palautusarvo: Boolean - True jos yhteys onnistui, False jos AutoCAD ei k‰ynniss‰
'
' Toiminta:
'   1. Kokeillaan GetObject-yhteys k‰ynniss‰ olevaan AutoCADiin
'   2. Jos onnistui, asetetaan oDOC aktiiviseen piirustukseen
'   3. Palautetaan yhteystila
'
' Huom:
'   - Kutsutaan Form_Load-tapahtumasta
'   - Asettaa moduulitason oACAD- ja oDOC-muuttujat
'--------------------------------------------------------------------------------
Private Function Yhdista() As Boolean
On Error Resume Next
  Set oACAD = GetObject(, "AutoCAD.Application")
  If Err <> 0 Then
    MsgBox "K‰ynniss‰ olevaa AutoCADi‰ ei lˆytynyt!", vbCritical, "Virhe!"
    Set oACAD = Nothing
    Yhdista = False
  Else
    Set oDOC = oACAD.ActiveDocument
    Yhdista = True
  End If
  On Error GoTo 0
End Function
'--------------------------------------------------------------------------------
' Aliohjelma: EtsiLinja
' Tarkoitus: Etsii automaattisesti putkilinjan l‰hell‰ nykyist‰ virtausblokkia
'
' Toiminta:
'   1. Haetaan nykyisen virtausblokin sijainti
'   2. M‰‰ritet‰‰n hakuikkuna (56 yks. oikealle, 13 yks. ylˆs)
'   3. Luodaan valintajoukko putkilinjablokkiluodatilla
'   4. Jos putkilinja lˆytyy, poimitaan viittaus ja p‰ivitet‰‰n blokki
'   5. Jos ei lˆydy, tyhjennet‰‰n attribuutti ja asetetaan EiLinjaa
'
' Huom:
'   - Etsii PIPELINE, BAHPIPEL, ARAPIPEL -blokkeja
'   - K‰ytt‰‰ leikkausikkunavalintaa
'   - Poistaa tilap‰isen valintajoukon k‰ytˆn j‰lkeen
'   - Asettaa EiLinjaa=True, jos putkilinjaa ei lˆydy
'--------------------------------------------------------------------------------
Private Sub EtsiLinja()
On Error GoTo ErrorHandler
Dim Poiminta As Object  ' AcadSelectionSet - muutettu late binding (64-bit)
Dim i As Integer
Dim aPiste(2) As Double
Dim lPiste(2) As Double
Dim Piste As Variant
Dim FilterType(0) As Integer
Dim FilterData(0) As Variant
Dim Attrib As Variant

' Haetaan tasearvoblokin sijainti
Piste = Joukko(j).InsertionPoint
aPiste(0) = Piste(0)
aPiste(1) = Piste(1)

' M‰‰ritet‰‰n hakuikkuna
lPiste(0) = aPiste(0) + 56
lPiste(1) = aPiste(1) + 13

' Poistetaan tilap‰inen valintajoukko
For i = 0 To oDOC.SelectionSets.Count - 1
  If oDOC.SelectionSets(i).Name = "APUPICK2" Then
    oDOC.SelectionSets(i).Delete
    Exit For
  End If
Next i

' Luodaan valintajoukko putkilinjoille
FilterType(0) = 2
FilterData(0) = "PIPELINE,BAHPIPEL,ARAPIPEL"

Set Poiminta = oDOC.SelectionSets.Add("APUPICK2")
Poiminta.Select acSelectionSetCrossing, aPiste, lPiste, FilterType, FilterData

If Poiminta.Count > 0 Then
  Attrib = Poiminta(0).GetAttributes
  TRefNo.Value = AttribToPipeline(Attrib)
  oAttrib.TextString = TRefNo.Value
Else
  oAttrib.TextString = ""
  EiLinjaa = True
End If

Poiminta.Delete
Set Poiminta = Nothing
Exit Sub

ErrorHandler:
    On Error Resume Next
    If Not Poiminta Is Nothing Then Poiminta.Delete
    Set Poiminta = Nothing
    On Error GoTo 0
    MsgBox "Virhe putkilinjan etsimisess‰: " & Err.Description, vbCritical, "EtsiLinja"
End Sub

'--------------------------------------------------------------------------------
' Tapahtuma: TRefNo_AfterUpdate
' Tarkoitus: P‰ivitt‰‰ AutoCAD-blokin attribuutin, kun k‰ytt‰j‰ muokkaa viittausta
'
' Huom:
'   - K‰ytt‰‰ Nz-funktiota Null-arvojen k‰sittelyyn
'--------------------------------------------------------------------------------
Private Sub TRefNo_AfterUpdate()
On Error GoTo ErrorHandler
If Not oAttrib Is Nothing Then
  oAttrib.TextString = Nz(TRefNo.Value, "")
End If
Exit Sub

ErrorHandler:
    ' Hiljainen virhe - ei keskeyt‰ k‰ytt‰j‰‰
End Sub

