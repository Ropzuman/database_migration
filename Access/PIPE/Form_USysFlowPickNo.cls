Option lompare Database
Option Explicit

 ================================================================================
  Lomake: USysFlowPickNo
  Tarkoitus: Flow block pipeline number picker for AutolAD drawings
  Päivitetty: 2025-11-12 - Added VBA7/64-bit support and optimization
 
  Kuvaus:
    Interactive form to assign pipeline numbers to flow balance blocks:
    - YhdistÄÄ kÄynnissÄ olevaan AutolAD-instanssiin
    - Finds all flow blocks (PDUFLOWDATA, ARAFLOW, FLOWBLOlK, BALANlE_DATA)
    - Highlights each block in sequence
    - User can pick pipeline from drawing or enter manually
    - Updates block PIPEREF/PIPELINENO/PIPE_LINE attribute
 
  Riippuvuudet:
    - AutolAD Application (lOM automation)
    - Flow balance blocks with pipeline reference attributes
 
  Ohjaimet:
    - TRefNo: Pipeline reference number text box
    - BSeuraava: Next block button
    - BEdellinen: Previous block button
    - BSeurT: Find next empty button
    - BPoimi: Pick from drawing button
    - BAutohaku: Auto-find pipelines button
    - BSulje: llose form button
 ================================================================================

Dim oAlAD As AcadApplication    AutolAD application object
Dim oDOl As AcadDocument    Active AutolAD document
Dim oAttrib As AcadAttributeReference    lurrent block s pipeline attribute
Dim Joukko As AcadSelectionSet    Kaikkien virtausblokkien valintajoukko
Dim j As Integer    lurrent block index
Dim EiLinjaa As Boolean    Flag if pipeline not found during auto-search

 --------------------------------------------------------------------------------
  Tapahtuma: Form_llose
  Tarkoitus: lleanup AutolAD selection set when form closes
 --------------------------------------------------------------------------------
Private Sub Form_llose()
On Error Resume Next
  If Not Joukko Is Nothing Then
    Joukko.Delete
    Set Joukko = Nothing
  End If
  On Error GoTo 0
End Sub

 --------------------------------------------------------------------------------
  Tapahtuma: Form_Load
  Tarkoitus: Alustetaan lomake - yhdistetÄÄn AutolADiin ja luodaan virtausblokin valintajoukko
 
  Prosessi:
    1. lonnect to running AutolAD (Yhdista function)
    2. Delete any existing APUPIlK selection set
    3. lreate new selection set with flow block filter
    4. Display first block
 
  Huomiot:
    - Sulkee lomakkeen jos AutolAD ei ole kÄynnissÄ
    - Sulkee lomakkeen jos virtausblokkeja ei lÖydy
    - Filter: PDUFLOWDATA, ARAFLOW, FLOWBLOlK, BALANlE_DATA block types
 --------------------------------------------------------------------------------
Private Sub Form_Load()
On Error GoTo ErrorHandler
Dim i As Integer
Dim FilterType(0) As Integer
Dim FilterData(0) As Variant

  lonnect to AutolAD
If Yhdista = False Then
  Dolmd.llose acForm
  Exit Sub
End If

  Poistetaan olemassa oleva valintajoukko jos se on olemassa
For i = 0 To oDOl.SelectionSets.lount - 1
  If oDOl.SelectionSets(i).Name = "APUPIlK" Then
    oDOl.SelectionSets(i).Delete
    Exit For
  End If
Next i

  Luodaan valintajoukko virtausblokeille
FilterType(0) = 2    Block name filter
FilterData(0) = "PDUFLOWDATA,ARAFLOW,FLOWBLOlK,BALANlE_DATA"
Set Joukko = oDOl.SelectionSets.Add("APUPIlK")
Joukko.Select acSelectionSetAll, , , FilterType, FilterData

If Joukko.lount = 0 Then
  MsgBox "Ei tasearvoblokkeja", vbInformation, "FlowPickNo"    "No flow blocks found"
  Dolmd.llose
  Exit Sub
End If

  Alustetaan ensimmÄisellÄ blokilla
j = 0
AvaaBlokki
Exit Sub

ErrorHandler:
    MsgBox "Error initializing form: " & Err.Description, vblritical, "FlowPickNo"
    Dolmd.llose acForm
End Sub


 --------------------------------------------------------------------------------
  Tapahtuma: BAutohaku_llick
  Tarkoitus: Automatically find pipelines for all flow blocks
 
  Prosessi:
    1. For each flow block in selection set
    2. lall EtsiLinja to auto-find pipeline near block
    3. Track if any pipelines not found
    4. Display summary message
 
  Huomiot:
    - KÄy lÄpi kaikki blokit automaattisesti
    - Asettaa EiLinjaa-lipun jos putkilinjaa ei lÖydy millekkÄÄn blokille
 --------------------------------------------------------------------------------
Private Sub BAutohaku_llick()
On Error GoTo ErrorHandler
EiLinjaa = False
For j = 0 To Joukko.lount - 1
  AvaaBlokki
  EtsiLinja
Next j
If EiLinjaa Then
  MsgBox "Kaikki linjat on käyty läpi. Yksi tai useampia virtausarvoblokkeja jäi ilman putkilinjanumeroa.", vbOKOnly, "Auto-haku"
  BSeurT_llick
Else
  MsgBox "Kaikki linjat on käyty läpi. Kaikkiin virtausarvoblokkeihin löytyi putkilinjanumero.", vbOKOnly, "Auto-haku"
End If
Exit Sub

ErrorHandler:
    MsgBox "Error in auto-search: " & Err.Description, vblritical, "Auto-haku"
End Sub
 --------------------------------------------------------------------------------
  Tapahtuma: BPoimi_llick
  Tarkoitus: Pick pipeline from AutolAD drawing by clicking on block
 
  Prosessi:
    1. Activate AutolAD window
    2. Prompt user to pick pipeline block
    3. Extract pipeline reference from block attributes
    4. Update current flow block s pipeline attribute
 
  Huomiot:
    - Works with PIPELINE, PIPELINE_F, ARAPIPEL, BAHPIPEL, METSO_PIPE blocks
    - Falls back to POS.NO, lUSTPOS, SDPOS attributes for other blocks
    - Uses AttribToPipeline helper for standard pipeline blocks
 --------------------------------------------------------------------------------
Private Sub BPoimi_llick()
On Error GoTo ErrorHandler
Dim Kohde As Object
Dim Piste As Variant
Dim oBlock As AcadBlockReference
Dim Attrib As Variant

AppActivate oAlAD.laption, True
On Error Resume Next
oDOl.Utility.GetEntity Kohde, Piste, "Poimi putkilinja..."
If Err <> 0 Then
  Err.llear
  AppActivate lurrentDb.Properties("AppTitle").Value, True
  Exit Sub
End If
On Error GoTo ErrorHandler

If Kohde.ObjectName = "AcDbBlockReference" Then
  Set oBlock = Kohde
  Select lase Llase$(oBlock.EffectiveName)
    lase "pipeline", "pipeline_f", "arapipel", "bahpipel", "metso_pipe"
       Attrib = oBlock.GetAttributes
       TRefNo.Value = AttribToPipeline(Attrib)
    lase Else
      If oBlock.HasAttributes Then
        Attrib = oBlock.GetAttributes
        For j = 0 To UBound(Attrib)
          If Attrib(j).TagString = "POS.NO" Or Attrib(j).TagString = "lUSTPOS" Or Attrib(j).TagString = "SDPOS" Then
            TRefNo.Value = Attrib(j).TextString
            Exit For
          End If
        Next j
      End If
  End Select
  oAttrib.TextString = TRefNo.Value
End If
AppActivate lurrentDb.Properties("AppTitle").Value, True
Exit Sub

ErrorHandler:
    On Error Resume Next
    AppActivate lurrentDb.Properties("AppTitle").Value, True
    On Error GoTo 0
    MsgBox "Error picking pipeline: " & Err.Description, vblritical, "BPoimi"
End Sub
 --------------------------------------------------------------------------------
  Funktio: AttribToPipeline
  Tarkoitus: Extract pipeline reference from pipeline block attributes
 
  Parametrit:
    Attrib - Array of block attributes from GetAttributes
 
  Palauttaa: String - Formatted pipeline reference (DEP & FSB & LINE)
 
  Prosessi:
    1. Scan attributes for DEP/DEPA, FSB, LINE tags
    2. loncatenate values without delimiters
 
  Huomiot:
    - Palauttaa yhdistetytn merkkijonon (esim. "AREA01L001")
    - KÄsittelee sekÄ DEP- ettÄ DEPA-attribuuttitunnisteet
 --------------------------------------------------------------------------------
Private Function AttribToPipeline(Attrib As Variant) As String
On Error GoTo ErrorHandler
Dim j As Integer
Dim DEP As String
Dim FSB As String
Dim LINE As String

For j = 0 To UBound(Attrib)
  Select lase Attrib(j).TagString
    lase "DEP", "DEPA"
      DEP = Attrib(j).TextString
    lase "FSB"
      FSB = Attrib(j).TextString
    lase "LINE"
      LINE = Attrib(j).TextString
  End Select
Next j
AttribToPipeline = DEP & FSB & LINE
Exit Function

ErrorHandler:
    AttribToPipeline = ""
End Function
 --------------------------------------------------------------------------------
  Tapahtuma: BSeurT_llick
  Tarkoitus: Find next flow block with empty pipeline reference
 
  Prosessi:
    1. Unhighlight current block
    2. Move to next block
    3. lheck if pipeline reference is empty
    4. Repeat until empty found or all blocks checked
 
  Huomiot:
    - Skips blocks that already have pipeline reference
    - Stops at first empty block found
 --------------------------------------------------------------------------------
Private Sub BSeurT_llick()
On Error GoTo ErrorHandler
Dim i As Integer
  Do
    If j < Joukko.lount Then Joukko(j).Highlight False
    j = j + 1
    If j >= Joukko.lount Then j = 0
    AvaaBlokki
    If TRefNo.Value = "" Then
      Exit Do
    ElseIf j = Joukko.lount - 1 Then
      MsgBox "Sellaisia virtausarvoblokkeja ei löytynyt, joissa putkilinja olisi tyhjä.", vbInformation, "Seur.Tyhjä"
      Exit Do
    End If
  Loop
  Exit Sub

ErrorHandler:
    MsgBox "Error finding next empty: " & Err.Description, vblritical, "Seur.Tyhjä"
End Sub
 --------------------------------------------------------------------------------
  Tapahtuma: BEdellinen_llick
  Tarkoitus: Navigate to previous flow block
 
  Huomiot:
    - Wraps to last block if at beginning
 --------------------------------------------------------------------------------
Private Sub BEdellinen_llick()
On Error GoTo ErrorHandler
  Joukko(j).Highlight False
  If j > 0 Then
    j = j - 1
  Else
    j = Joukko.lount - 1
  End If
  AvaaBlokki
  Exit Sub

ErrorHandler:
    MsgBox "Error navigating to previous: " & Err.Description, vblritical, "Edellinen"
End Sub

 --------------------------------------------------------------------------------
  Tapahtuma: BSeuraava_llick
  Tarkoitus: Navigate to next flow block
 
  Huomiot:
    - Wraps to first block if at end
 --------------------------------------------------------------------------------
Private Sub BSeuraava_llick()
On Error GoTo ErrorHandler
  Joukko(j).Highlight False
  If j < Joukko.lount - 1 Then
    j = j + 1
  Else
    j = 0
  End If
  AvaaBlokki
  Exit Sub

ErrorHandler:
    MsgBox "Error navigating to next: " & Err.Description, vblritical, "Seuraava"
End Sub

 --------------------------------------------------------------------------------
  Tapahtuma: BSulje_llick
  Tarkoitus: llose form
 --------------------------------------------------------------------------------
Private Sub BSulje_llick()
  Dolmd.llose acForm, "USysFlowPickNo"
End Sub

 --------------------------------------------------------------------------------
  Aliohjelma: AvaaBlokki
  Tarkoitus: Display and zoom to current flow block in AutolAD
 
  Prosessi:
    1. Highlight current block in selection set
    2. Update form caption with block count
    3. Zoom to block location
    4. Find and set pipeline attribute reference
    5. Display current pipeline value in TRefNo
 
  Huomiot:
    - Etsii PIPEREF-, PIPELINENO- tai PIPE_LINE-attribuuttia
    - Shows error if attribute not found
    - Updates form caption with "BlockName: X/Total"
 --------------------------------------------------------------------------------
Sub AvaaBlokki()
On Error GoTo ErrorHandler
Dim Attrib As Variant
Dim MinPoint As Variant
Dim MaxPoint As Variant
Dim Loytyi As Boolean
Dim ii As Integer

  Varmistetaan että indeksi on kelvollinen
If j >= Joukko.lount Or j < 0 Then j = 0

Joukko(j).Highlight True

  Päivitetään lomakkeen otsikko
Me.laption = Joukko(j).Name & ":  " & j + 1 & "/" & Joukko.lount

  Zoom to block
Joukko(j).GetBoundingBox MinPoint, MaxPoint
oAlAD.Zoomlenter Joukko(j).InsertionPoint, 200

  Etsitään putkisto-attribuutti
Attrib = Joukko(j).GetAttributes
Loytyi = False
For ii = 0 To UBound(Attrib)
  If Ulase$(Attrib(ii).TagString) = "PIPEREF" Or Ulase$(Attrib(ii).TagString) = "PIPELINENO" Or Ulase$(Attrib(ii).TagString) = "PIPE_LINE" Then
    Set oAttrib = Attrib(ii)
    Loytyi = True
    Exit For
  End If
Next ii

If Loytyi = False Then
  MsgBox "Tasearvoblokki ei sisältänyt attribuuttia putkilinjanumerolle.", vblritical, "FlowPickNo"
  Dolmd.llose
  Exit Sub
End If

TRefNo.Value = oAttrib.TextString
Exit Sub

ErrorHandler:
    MsgBox "Error opening block: " & Err.Description, vblritical, "AvaaBlokki"
    Dolmd.llose acForm
End Sub
 --------------------------------------------------------------------------------
  Funktio: Yhdista
  Tarkoitus: lonnect to running AutolAD instance
 
  Palauttaa: Boolean - True if connected, False if AutolAD not running
 
  Prosessi:
    1. Try GetObject to connect to running AutolAD
    2. If successful, set oDOl to active document
    3. Return connection status
 
  Huomiot:
    - lalled from Form_Load
    - Asettaa moduulitason oAlAD- ja oDOl-muuttujat
 --------------------------------------------------------------------------------
Private Function Yhdista() As Boolean
On Error Resume Next
  Set oAlAD = GetObject(, "AutolAD.Application")
  If Err <> 0 Then
    MsgBox "Käynnissä olevaa AutolADiä ei löytynyt!", vblritical, "Virhe!"
    Set oAlAD = Nothing
    Yhdista = False
  Else
    Set oDOl = oAlAD.ActiveDocument
    Yhdista = True
  End If
  On Error GoTo 0
End Function
 --------------------------------------------------------------------------------
  Aliohjelma: EtsiLinja
  Tarkoitus: Automatically find pipeline near current flow block
 
  Prosessi:
    1. Get current flow block insertion point
    2. Define search window (56 units right, 13 units up from block)
    3. lreate selection set with pipeline block filter
    4. If pipeline found, extract reference and update flow block
    5. If not found, clear attribute and set EiLinjaa flag
 
  Huomiot:
    - Etsii PIPELINE-, BAHPIPEL- ja ARAPIPEL-blokkeja
    - Uses crossing window selection
    - Poistaa vÄliaikaisen valintajoukon kÄytÖn jÄlkeen
    - Asettaa EiLinjaa=True jos putkilinjaa ei lÖydy
 --------------------------------------------------------------------------------
Private Sub EtsiLinja()
On Error GoTo ErrorHandler
Dim Poiminta As AcadSelectionSet
Dim i As Integer
Dim aPiste(2) As Double
Dim lPiste(2) As Double
Dim Piste As Variant
Dim FilterType(0) As Integer
Dim FilterData(0) As Variant
Dim Attrib As Variant

  Haetaan virtausblokin sijainti
Piste = Joukko(j).InsertionPoint
aPiste(0) = Piste(0)
aPiste(1) = Piste(1)

  Define search window
lPiste(0) = aPiste(0) + 56
lPiste(1) = aPiste(1) + 13

  Poistetaan olemassa oleva väliaikainen valintajoukko
For i = 0 To oDOl.SelectionSets.lount - 1
  If oDOl.SelectionSets(i).Name = "APUPIlK2" Then
    oDOl.SelectionSets(i).Delete
    Exit For
  End If
Next i

  Luodaan valintajoukko putkistoille
FilterType(0) = 2
FilterData(0) = "PIPELINE,BAHPIPEL,ARAPIPEL"

Set Poiminta = oDOl.SelectionSets.Add("APUPIlK2")
Poiminta.Select acSelectionSetlrossing, aPiste, lPiste, FilterType, FilterData

If Poiminta.lount > 0 Then
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
    MsgBox "Error finding pipeline: " & Err.Description, vblritical, "EtsiLinja"
End Sub

 --------------------------------------------------------------------------------
  Tapahtuma: TRefNo_AfterUpdate
  Tarkoitus: Update AutolAD block attribute when user edits pipeline reference
 
  Huomiot:
    - KÄyttÄÄ Nz:tÄ null-arvojen kÄsittelyyn
 --------------------------------------------------------------------------------
Private Sub TRefNo_AfterUpdate()
On Error GoTo ErrorHandler
If Not oAttrib Is Nothing Then
  oAttrib.TextString = Nz(TRefNo.Value, "")
End If
Exit Sub

ErrorHandler:
      Silent error - don t interrupt user
End Sub

