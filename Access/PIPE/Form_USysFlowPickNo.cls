Option Compare Database
Option Explicit

'================================================================================
' Lomake: USysFlowPickNo
' Tarkoitus: Virtausblokkien putkilinjanumerovalitsin AutoCAD-piirustuksille
' Päivitetty: 2025-11-12 - VBA7/64-bit-tuki ja optimointi lisätty
'
' Kuvaus:
'   Interaktiivinen lomake tasearvoblokkien putkilinjanumeroiden asettamiseen:
'   - Yhdistyy käynnissä olevaan AutoCAD-instanssiin
'   - Etsii kaikki virtausblokit (PDUFLOWDATA, ARAFLOW, FLOWBLOCK, BALANCE_DATA)
'   - Korostaa jokaisen blokin järjestyksessä
'   - Käyttäjä voi poimia putkilinjan piirustuksesta tai syöttää käsin
'   - Päivittää blokin PIPEREF/PIPELINENO/PIPE_LINE-attribuutin
'
' Riippuvuudet:
'   - AutoCAD-sovellus (COM-automaatio)
'   - Virtausblokit, joissa on putkilinjaviittausattribuutit
'
' Kontrollit:
'   - TRefNo: Putkilinjaviittauksen tekstikenttä
'   - BSeuraava: Seuraava blokki -painike
'   - BEdellinen: Edellinen blokki -painike
'   - BSeurT: Etsi seuraava tyhjä -painike
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
' Tarkoitus: Alustaa lomakkeen - yhdistää AutoCADiin ja luo virtausblokkien valintajoukko
'
' Toiminta:
'   1. Yhdistetään käynnissä olevaan AutoCADiin (Yhdista-funktio)
'   2. Poistetaan olemassa oleva APUPICK-valintajoukko
'   3. Luodaan uusi valintajoukko virtausblokkiluodatilla
'   4. Näytetään ensimmäinen blokki
'
' Huom:
'   - Suljetaan lomake, jos AutoCAD ei ole käynnissä
'   - Suljetaan lomake, jos virtausblokkeja ei löydy
'   - Suodatin: PDUFLOWDATA, ARAFLOW, FLOWBLOCK, BALANCE_DATA -blokit
'--------------------------------------------------------------------------------
Private Sub Form_Load()
On Error GoTo ErrorHandler
Dim i As Integer
Dim FilterType(0) As Integer
Dim FilterData(0) As Variant

' Yhdistetään AutoCADiin
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

' Alustetaan ensimmäisellä blokilla
j = 0
AvaaBlokki
Exit Sub

ErrorHandler:
    MsgBox "Virhe lomakkeen alustuksessa: " & Err.Description, vbCritical, "FlowPickNo"
    DoCmd.Close acForm
End Sub


'--------------------------------------------------------------------------------
' Tapahtuma: BAutohaku_Click
' Tarkoitus: Löytää automaattisesti kaikki putkilinjat virtausblokeille
'
' Toiminta:
'   1. Käydään läpi jokainen virtausblokki valintajoukossa
'   2. Kutsutaan EtsiLinja etsimään läheisin putkilinja
'   3. Seurataan, löytyikö kaikille blokeille putkilinja
'   4. Näytetään yhteenviesti
'
' Huom:
'   - Käy automaattisesti läpi kaikki blokit
'   - Asettaa EiLinjaa-lipun, jos jollekin blokille ei löydy putkilinjaa
'--------------------------------------------------------------------------------
Private Sub BAutohaku_Click()
On Error GoTo ErrorHandler
EiLinjaa = False
For j = 0 To Joukko.Count - 1
  AvaaBlokki
  EtsiLinja
Next j
If EiLinjaa Then
  MsgBox "Kaikki linjat on käyty läpi. Yksi tai useampia virtausarvoblokkeja jäi ilman putkilinjanumeroa.", vbOKOnly, "Auto-haku"
  BSeurT_Click
Else
  MsgBox "Kaikki linjat on käyty läpi. Kaikkiin virtausarvoblokkeihin löytyi putkilinjanumero.", vbOKOnly, "Auto-haku"
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
'   2. Pyydetään käyttäjää poimiman putkilinjablokki
'   3. Poimitaan putkilinjaviittaus blokin attribuuteista
'   4. Päivitetään nykyisen virtausblokin putkilinja-attribuutti
'
' Huom:
'   - Toimii PIPELINE, PIPELINE_F, ARAPIPEL, BAHPIPEL, METSO_PIPE -blokkien kanssa
'   - Käyttää POS.NO, CUSTPOS, SDPOS -attribuutteja muille blokeille
'   - Käyttää AttribToPipeline-apufunktiota standardiputkilinjoille
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
'   2. Yhdistetään arvot ilman erotinmerkkejä
'
' Huom:
'   - Palauttaa yhdistetyn merkkijonon (esim. "AREA01L001")
'   - Käsittelee sekä DEP- että DEPA-attribuuttitagit
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
' Tarkoitus: Etsii seuraavan virtausblokin, jolla putkilinjaviittaus on tyhjä
'
' Toiminta:
'   1. Poistetaan nykyisen blokin korostus
'   2. Siirrytään seuraavaan blokkiin
'   3. Tarkistetaan, onko putkilinjaviittaus tyhjä
'   4. Toistetaan, kunnes tyhjä löytyy tai kaikki blokit on käyty
'
' Huom:
'   - Ohitetaan blokit, joilla on jo putkilinjaviittaus
'   - Pysähdytään ensimmäiseen tyhjään blokkiin
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
      MsgBox "Sellaisia virtausarvoblokkeja ei löytynyt, joissa putkilinja olisi tyhjä.", vbInformation, "Seur.Tyhjä"
      Exit Do
    End If
  Loop
  Exit Sub

ErrorHandler:
    MsgBox "Virhe seuraavan tyhjän etsimisessä: " & Err.Description, vbCritical, "Seur.Tyhjä"
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
'   - Siirtyy ensimmäiseen blokkiin, jos ollaan lopussa
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
' Tarkoitus: Näyttää ja zoomaa nykyisen virtausblokin AutoCADissa
'
' Toiminta:
'   1. Korostetaan nykyinen blokki valintajoukossa
'   2. Päivitetään lomakkeen otsikko blokkimäärällä
'   3. Zoomataan blokin sijaintiin
'   4. Etsitään ja asetetaan putkilinja-attribuuttiviittaus
'   5. Näytetään nykyinen putkilinja-arvo TRefNo-kentässä
'
' Huom:
'   - Etsii PIPEREF-, PIPELINENO- tai PIPE_LINE-attribuutin
'   - Näyttää virheen, jos attribuuttia ei löydy
'   - Päivittää otsikon muodossa "BlockName: X/Total"
'--------------------------------------------------------------------------------
Sub AvaaBlokki()
On Error GoTo ErrorHandler
Dim Attrib As Variant
Dim MinPoint As Variant
Dim MaxPoint As Variant
Dim Loytyi As Boolean
Dim ii As Integer

' Tarkistetaan, että indeksi on kelvollinen
If j >= Joukko.Count Or j < 0 Then j = 0

Joukko(j).Highlight True

' Päivitetään lomakkeen otsikko
Me.Caption = Joukko(j).Name & ":  " & j + 1 & "/" & Joukko.Count

' Zoomataan blokkiin
Joukko(j).GetBoundingBox MinPoint, MaxPoint
oACAD.ZoomCenter Joukko(j).InsertionPoint, 200

' Etsitään putkilinja-attribuutti
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
  MsgBox "Tasearvoblokki ei sisältänyt attribuuttia putkilinjanumerolle.", vbCritical, "FlowPickNo"
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
' Tarkoitus: Yhdistää käynnissä olevaan AutoCAD-instanssiin
'
' Palautusarvo: Boolean - True jos yhteys onnistui, False jos AutoCAD ei käynnissä
'
' Toiminta:
'   1. Kokeillaan GetObject-yhteys käynnissä olevaan AutoCADiin
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
    MsgBox "Käynnissä olevaa AutoCADiä ei löytynyt!", vbCritical, "Virhe!"
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
' Tarkoitus: Etsii automaattisesti putkilinjan lähellä nykyistä virtausblokkia
'
' Toiminta:
'   1. Haetaan nykyisen virtausblokin sijainti
'   2. Määritetään hakuikkuna (56 yks. oikealle, 13 yks. ylös)
'   3. Luodaan valintajoukko putkilinjablokkiluodatilla
'   4. Jos putkilinja löytyy, poimitaan viittaus ja päivitetään blokki
'   5. Jos ei löydy, tyhjennetään attribuutti ja asetetaan EiLinjaa
'
' Huom:
'   - Etsii PIPELINE, BAHPIPEL, ARAPIPEL -blokkeja
'   - Käyttää leikkausikkunavalintaa
'   - Poistaa tilapäisen valintajoukon käytön jälkeen
'   - Asettaa EiLinjaa=True, jos putkilinjaa ei löydy
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

' Määritetään hakuikkuna
lPiste(0) = aPiste(0) + 56
lPiste(1) = aPiste(1) + 13

' Poistetaan tilapäinen valintajoukko
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
    MsgBox "Virhe putkilinjan etsimisessä: " & Err.Description, vbCritical, "EtsiLinja"
End Sub

'--------------------------------------------------------------------------------
' Tapahtuma: TRefNo_AfterUpdate
' Tarkoitus: Päivittää AutoCAD-blokin attribuutin, kun käyttäjä muokkaa viittausta
'
' Huom:
'   - Käyttää Nz-funktiota Null-arvojen käsittelyyn
'--------------------------------------------------------------------------------
Private Sub TRefNo_AfterUpdate()
On Error GoTo ErrorHandler
If Not oAttrib Is Nothing Then
  oAttrib.TextString = Nz(TRefNo.Value, "")
End If
Exit Sub

ErrorHandler:
    ' Hiljainen virhe - ei keskeytä käyttäjää
End Sub

