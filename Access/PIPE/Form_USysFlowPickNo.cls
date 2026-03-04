Option Compare Database
Option Explicit

'================================================================================
' Form: USysFlowPickNo
' Tarkoitus: Flow block pipeline number picker for AutoCAD drawings
' Päivitetty: 2025-11-12 - VBA7/64-bit-tuki ja optimointi lisätty
'
' Kuvaus:
'   Interactive form to assign pipeline numbers to flow balance blocks:
'   - Connects to running AutoCAD instance
'   - Finds all flow blocks (PDUFLOWDATA, ARAFLOW, FLOWBLOCK, BALANCE_DATA)
'   - Highlights each block in sequence
'   - User can pick pipeline from drawing or enter manually
'   - Updates block PIPEREF/PIPELINENO/PIPE_LINE attribute
'
' Riippuvuudet:
'   - AutoCAD Application (COM automation)
'   - Flow balance blocks with pipeline reference attributes
'
' Controls:
'   - TRefNo: Pipeline reference number text box
'   - BSeuraava: Next block button
'   - BEdellinen: Previous block button
'   - BSeurT: Find next empty button
'   - BPoimi: Pick from drawing button
'   - BAutohaku: Auto-find pipelines button
'   - BSulje: Close form button
'================================================================================

Dim oACAD As Object  ' AcadApplication - muutettu late binding (64-bit)
Dim oDOC As Object   ' AcadDocument - muutettu late binding (64-bit)
Dim oAttrib As Object  ' AcadAttributeReference - muutettu late binding (64-bit)
Dim Joukko As Object   ' AcadSelectionSet - muutettu late binding (64-bit)
Dim j As Integer  ' Current block index
Dim EiLinjaa As Boolean  ' Flag if pipeline not found during auto-search

'--------------------------------------------------------------------------------
' Event: Form_Close
' Tarkoitus: Cleanup AutoCAD selection set when form closes
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
' Event: Form_Load
' Tarkoitus: Initialize form - connect to AutoCAD and create flow block selection set
'
' Toiminta:
'   1. Connect to running AutoCAD (Yhdista function)
'   2. Delete any existing APUPICK selection set
'   3. Create new selection set with flow block filter
'   4. Display first block
'
' Huom:
'   - Closes form if AutoCAD not running
'   - Closes form if no flow blocks found
'   - Filter: PDUFLOWDATA, ARAFLOW, FLOWBLOCK, BALANCE_DATA block types
'--------------------------------------------------------------------------------
Private Sub Form_Load()
On Error GoTo ErrorHandler
Dim i As Integer
Dim FilterType(0) As Integer
Dim FilterData(0) As Variant

' Connect to AutoCAD
If Yhdista = False Then
  DoCmd.Close acForm
  Exit Sub
End If

' Delete existing selection set if present
For i = 0 To oDOC.SelectionSets.Count - 1
  If oDOC.SelectionSets(i).Name = "APUPICK" Then
    oDOC.SelectionSets(i).Delete
    Exit For
  End If
Next i

' Create selection set for flow blocks
FilterType(0) = 2  ' Block name filter
FilterData(0) = "PDUFLOWDATA,ARAFLOW,FLOWBLOCK,BALANCE_DATA"
Set Joukko = oDOC.SelectionSets.Add("APUPICK")
Joukko.Select acSelectionSetAll, , , FilterType, FilterData

If Joukko.Count = 0 Then
  MsgBox "Ei tasearvoblokkeja", vbInformation, "FlowPickNo"  ' "No flow blocks found"
  DoCmd.Close
  Exit Sub
End If

' Initialize with first block
j = 0
AvaaBlokki
Exit Sub

ErrorHandler:
    MsgBox "Virhe lomakkeen alustuksessa: " & Err.Description, vbCritical, "FlowPickNo"
    DoCmd.Close acForm
End Sub


'--------------------------------------------------------------------------------
' Event: BAutohaku_Click
' Tarkoitus: Automatically find pipelines for all flow blocks
'
' Toiminta:
'   1. For each flow block in selection set
'   2. Call EtsiLinja to auto-find pipeline near block
'   3. Track if any pipelines not found
'   4. Display summary message
'
' Huom:
'   - Iterates through all blocks automatically
'   - Sets EiLinjaa flag if pipeline not found for any block
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
' Event: BPoimi_Click
' Tarkoitus: Pick pipeline from AutoCAD drawing by clicking on block
'
' Toiminta:
'   1. Activate AutoCAD window
'   2. Prompt user to pick pipeline block
'   3. Extract pipeline reference from block attributes
'   4. Update current flow block's pipeline attribute
'
' Huom:
'   - Works with PIPELINE, PIPELINE_F, ARAPIPEL, BAHPIPEL, METSO_PIPE blocks
'   - Falls back to POS.NO, CUSTPOS, SDPOS attributes for other blocks
'   - Uses AttribToPipeline helper for standard pipeline blocks
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
' Function: AttribToPipeline
' Tarkoitus: Extract pipeline reference from pipeline block attributes
'
' Parametrit:
'   Attrib - Array of block attributes from GetAttributes
'
' Palautusarvo: String - Formatted pipeline reference (DEP & FSB & LINE)
'
' Toiminta:
'   1. Scan attributes for DEP/DEPA, FSB, LINE tags
'   2. Concatenate values without delimiters
'
' Huom:
'   - Returns concatenated string (e.g., "AREA01L001")
'   - Handles both DEP and DEPA attribute tags
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
' Event: BSeurT_Click
' Tarkoitus: Find next flow block with empty pipeline reference
'
' Toiminta:
'   1. Unhighlight current block
'   2. Move to next block
'   3. Check if pipeline reference is empty
'   4. Repeat until empty found or all blocks checked
'
' Huom:
'   - Skips blocks that already have pipeline reference
'   - Stops at first empty block found
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
' Event: BEdellinen_Click
' Tarkoitus: Navigate to previous flow block
'
' Huom:
'   - Wraps to last block if at beginning
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
' Event: BSeuraava_Click
' Tarkoitus: Navigate to next flow block
'
' Huom:
'   - Wraps to first block if at end
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
' Event: BSulje_Click
' Tarkoitus: Close form
'--------------------------------------------------------------------------------
Private Sub BSulje_Click()
  DoCmd.Close acForm, "USysFlowPickNo"
End Sub

'--------------------------------------------------------------------------------
' Sub: AvaaBlokki
' Tarkoitus: Display and zoom to current flow block in AutoCAD
'
' Toiminta:
'   1. Highlight current block in selection set
'   2. Update form caption with block count
'   3. Zoom to block location
'   4. Find and set pipeline attribute reference
'   5. Display current pipeline value in TRefNo
'
' Huom:
'   - Searches for PIPEREF, PIPELINENO, or PIPE_LINE attribute
'   - Shows error if attribute not found
'   - Updates form caption with "BlockName: X/Total"
'--------------------------------------------------------------------------------
Sub AvaaBlokki()
On Error GoTo ErrorHandler
Dim Attrib As Variant
Dim MinPoint As Variant
Dim MaxPoint As Variant
Dim Loytyi As Boolean
Dim ii As Integer

' Ensure index is valid
If j >= Joukko.Count Or j < 0 Then j = 0

Joukko(j).Highlight True

' Update form caption
Me.Caption = Joukko(j).Name & ":  " & j + 1 & "/" & Joukko.Count

' Zoom to block
Joukko(j).GetBoundingBox MinPoint, MaxPoint
oACAD.ZoomCenter Joukko(j).InsertionPoint, 200

' Find pipeline attribute
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
' Function: Yhdista
' Tarkoitus: Connect to running AutoCAD instance
'
' Palautusarvo: Boolean - True if connected, False if AutoCAD not running
'
' Toiminta:
'   1. Try GetObject to connect to running AutoCAD
'   2. If successful, set oDOC to active document
'   3. Return connection status
'
' Huom:
'   - Called from Form_Load
'   - Sets module-level oACAD and oDOC variables
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
' Sub: EtsiLinja
' Tarkoitus: Automatically find pipeline near current flow block
'
' Toiminta:
'   1. Get current flow block insertion point
'   2. Define search window (56 units right, 13 units up from block)
'   3. Create selection set with pipeline block filter
'   4. If pipeline found, extract reference and update flow block
'   5. If not found, clear attribute and set EiLinjaa flag
'
' Huom:
'   - Searches for PIPELINE, BAHPIPEL, ARAPIPEL blocks
'   - Uses crossing window selection
'   - Deletes temporary selection set after use
'   - Sets EiLinjaa=True if no pipeline found
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

' Get flow block position
Piste = Joukko(j).InsertionPoint
aPiste(0) = Piste(0)
aPiste(1) = Piste(1)

' Define search window
lPiste(0) = aPiste(0) + 56
lPiste(1) = aPiste(1) + 13

' Delete existing temp selection set
For i = 0 To oDOC.SelectionSets.Count - 1
  If oDOC.SelectionSets(i).Name = "APUPICK2" Then
    oDOC.SelectionSets(i).Delete
    Exit For
  End If
Next i

' Create selection set for pipelines
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
' Event: TRefNo_AfterUpdate
' Tarkoitus: Update AutoCAD block attribute when user edits pipeline reference
'
' Huom:
'   - Uses Nz to handle null values
'--------------------------------------------------------------------------------
Private Sub TRefNo_AfterUpdate()
On Error GoTo ErrorHandler
If Not oAttrib Is Nothing Then
  oAttrib.TextString = Nz(TRefNo.Value, "")
End If
Exit Sub

ErrorHandler:
    ' Silent error - don't interrupt user
End Sub

