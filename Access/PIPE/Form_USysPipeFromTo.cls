Option Compare Database
Option Explicit

'================================================================================
' Form: USysPipeFromTo
' Purpose: Pipeline FROM/TO attribute editor for AutoCAD
' Updated: 2025-11-12 - Added VBA7/64-bit support and optimization
'
' Description:
'   Interactive form to edit pipeline block FROM/TO attributes:
'   - Connects to running AutoCAD instance
'   - Filters and displays only pipeline blocks (PIPELINE, PIPELINE_DATA, METSO_PIPE)
'   - Allows navigation through blocks (Next, Previous, Find Next Empty)
'   - Pick FROM/TO from drawing or enter manually
'   - Updates block attributes in real-time
'
' Dependencies:
'   - AutoCAD Application (COM automation)
'   - Pipeline blocks with FROM/TO attributes
'
' Controls:
'   - TFrom, TTo: FROM/TO text boxes
'   - BNext, BPrev: Navigation buttons
'   - BFind: Find next empty button
'   - BPickFrom, BPickTo: Pick from drawing buttons
'   - BClearFrom, BClearTo: Clear attribute buttons
'   - LPipeLine: Display current pipeline reference
'================================================================================

Dim oACAD As Object  ' AcadApplication - muutettu late binding (64-bit)
Dim oDOC As Object   ' AcadDocument - muutettu late binding (64-bit)
Dim AttribFrom As Object  ' AcadAttributeReference - muutettu late binding (64-bit)
Dim AttribTo As Object    ' AcadAttributeReference - muutettu late binding (64-bit)
Dim oBlock As Object   ' AcadBlockReference - muutettu late binding (64-bit)
Dim Joukko As Object   ' AcadSelectionSet - muutettu late binding (64-bit)
Dim i As Integer  ' Current block index

'================================================================================
' Event: BClearFrom_Click
' Purpose: Clear FROM attribute of current pipeline block
'================================================================================
Private Sub BClearFrom_Click()
  On Error GoTo ErrorHandler
  
  AttribFrom.TextString = ""
  TFrom.Value = ""
  
  Exit Sub
ErrorHandler:
  MsgBox "Error clearing FROM attribute: " & Err.Description, vbExclamation
End Sub

'================================================================================
' Event: BClearTo_Click
' Purpose: Clear TO attribute of current pipeline block
'================================================================================
Private Sub BClearTo_Click()
  On Error GoTo ErrorHandler
  
  AttribTo.TextString = ""
  TTo.Value = ""
  
  Exit Sub
ErrorHandler:
  MsgBox "Error clearing TO attribute: " & Err.Description, vbExclamation
End Sub

'================================================================================
' Event: BClose_Click
' Purpose: Close the form
'================================================================================
Private Sub BClose_Click()
  DoCmd.Close acForm, "UsysPipeFromTo"
End Sub

'================================================================================
' Sub: OpenSource
' Purpose: Pick a pipeline, tank, pump, or equipment from drawing and extract reference
' Parameters:
'   Teksti - TextBox to populate with reference value
'   oAttrib - Attribute reference to update
'================================================================================
Private Sub OpenSource(Teksti As TextBox, oAttrib As Object)  ' AcadAttributeReference - muutettu late binding (64-bit)
  Dim Kohde As Object
  Dim Piste As Variant
  Dim oPipeBlock As Object  ' AcadBlockReference - muutettu late binding (64-bit)
  Dim Attrib As Variant
  Dim j As Integer
  
  On Error Resume Next
  AppActivate oACAD.Caption, True
  If Err <> 0 Then
    MsgBox "Cannot activate AutoCAD: " & Err.Description, vbExclamation
    Err.Clear
    Exit Sub
  End If
  
  oDOC.Utility.GetEntity Kohde, Piste, "Pick pipeline, tank, pump or equipment: "
  If Err <> 0 Then
    Err.Clear
  Else
    If Kohde.ObjectName = "AcDbBlockReference" Then
      Set oPipeBlock = Kohde
      Select Case LCase$(oPipeBlock.EffectiveName)
        Case "pipeline", "pipeline_f", "arapipel", "pipeline_data", "metso_pipe"
           Attrib = oPipeBlock.GetAttributes
           Teksti.Value = AttribToPipeline(Attrib)
        Case Else
          If oPipeBlock.HasAttributes Then
            Attrib = oPipeBlock.GetAttributes
            For j = 0 To UBound(Attrib)
              If Attrib(j).TagString = "POS.NO" Or Attrib(j).TagString = "CUSTPOS" Or Attrib(j).TagString = "SDPOS" Then
                Teksti.Value = Attrib(j).TextString
                Exit For
              End If
            Next j
          End If
      End Select
      oAttrib.TextString = Teksti.Value
    ElseIf Kohde.ObjectName = "AcDbText" Then
      Teksti.Value = Kohde.TextString
      oAttrib.TextString = Teksti.Value
    End If
  End If
  On Error GoTo 0
  AppActivate CurrentDb.Properties("AppTitle").Value
End Sub

'================================================================================
' Function: AttribToPipeline
' Purpose: Extract pipeline reference from block attributes (DEP+FSB+LINE)
' Parameters:
'   Attrib - Array of block attributes
' Returns: Concatenated pipeline reference string
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
' Event: BFind_Click
' Purpose: Find next pipeline block with empty FROM or TO attribute
'================================================================================
Private Sub BFind_Click()
  On Error GoTo ErrorHandler
  
  Do
    i = i + 1
    If i > Joukko.Count - 1 Then
      If MsgBox("Last block!" & vbCrLf & "Start from first?", vbYesNo, "Find next empty") = vbYes Then
         i = 0
      Else
        Exit Do
      End If
    End If
    AvaaBlokki
    If AttribFrom.TextString = "" Or AttribTo.TextString = "" Then Exit Do
  Loop
  
  Exit Sub
ErrorHandler:
  MsgBox "Error finding next empty block: " & Err.Description, vbExclamation
End Sub

'================================================================================
' Event: BNext_Click
' Purpose: Navigate to next pipeline block
'================================================================================
Private Sub BNext_Click()
  On Error GoTo ErrorHandler
  NextBlock
  Exit Sub
ErrorHandler:
  MsgBox "Error navigating to next block: " & Err.Description, vbExclamation
End Sub

'================================================================================
' Sub: NextBlock
' Purpose: Move to next block in selection set (wraps to first)
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
  MsgBox "Error in NextBlock: " & Err.Description, vbExclamation
End Sub

'================================================================================
' Sub: KirjoitaAtt
' Purpose: Write text to attribute reference
' Parameters:
'   Teksti - Text to write
'   oAttrib - Attribute reference to update
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
  MsgBox "Error writing attribute: " & Err.Description, vbExclamation
End Sub

'================================================================================
' Event: BPickFrom_Click
' Purpose: Pick FROM reference from AutoCAD drawing
'================================================================================
Private Sub BPickFrom_Click()
  On Error GoTo ErrorHandler
  
  OpenSource TFrom, AttribFrom
  AppActivate CurrentDb.Properties("AppTitle").Value
  
  Exit Sub
ErrorHandler:
  MsgBox "Error picking FROM: " & Err.Description, vbExclamation
End Sub

'================================================================================
' Event: BPickTo_Click
' Purpose: Pick TO reference from AutoCAD drawing
'================================================================================
Private Sub BPickTo_Click()
  On Error GoTo ErrorHandler
  
  OpenSource TTo, AttribTo
  AppActivate CurrentDb.Properties("AppTitle").Value
  
  Exit Sub
ErrorHandler:
  MsgBox "Error picking TO: " & Err.Description, vbExclamation
End Sub

'================================================================================
' Event: BPrev_Click
' Purpose: Navigate to previous pipeline block
'================================================================================
Private Sub BPrev_Click()
  On Error GoTo ErrorHandler
  PrevBlock
  Exit Sub
ErrorHandler:
  MsgBox "Error navigating to previous block: " & Err.Description, vbExclamation
End Sub

'================================================================================
' Sub: PrevBlock
' Purpose: Move to previous block in selection set (wraps to last)
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
  MsgBox "Error in PrevBlock: " & Err.Description, vbExclamation
End Sub

'================================================================================
' Event: Form_Close
' Purpose: Clean up AutoCAD objects when form closes
'================================================================================
Private Sub Form_Close()
  On Error Resume Next
  
  ' Clean up selection set
  If Not Joukko Is Nothing Then
    Joukko.Delete
    Set Joukko = Nothing
  End If
  
  ' Clean up AutoCAD objects
  Set AttribFrom = Nothing
  Set AttribTo = Nothing
  Set oBlock = Nothing
  Set oDOC = Nothing
  Set oACAD = Nothing
End Sub

'================================================================================
' Event: Form_Load
' Purpose: Initialize form, connect to AutoCAD, filter pipeline blocks
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
       FT(0) = 0
       FD(0) = "INSERT"
       TeeJoukko "APUVALINTA", Joukko
       Joukko.Select acSelectionSetAll, , , FT, FD
       
       ' Filter: Keep only pipeline blocks
       For j = 0 To Joukko.Count - 1
         Nimi = UCase$(Joukko(j).EffectiveName)
         If Nimi <> "PIPELINE" And Nimi <> "PIPELINE_DATA" And Nimi <> "METSO_PIPE" Then
           ReDim Preserve Poista(L)
           Set Poista(L) = Joukko(j)
           L = L + 1
         End If
       Next j
       
       If L > 0 Then Joukko.RemoveItems Poista
       
       If Joukko.Count = 0 Then
         MsgBox "No pipeline blocks found.", vbExclamation
         Err.Raise vbObjectError + 1000
       Else
         AvaaBlokki
       End If
     End If
  End If
  Exit Sub
  
ErrorHandler:
  MsgBox "Error loading form: " & Err.Description, vbExclamation
  BClose_Click
End Sub

'================================================================================
' Sub: TeeJoukko
' Purpose: Create or recreate a named selection set
' Parameters:
'   Nimi - Name of selection set
'   VJoukko - Selection set object to create
'================================================================================
Private Sub TeeJoukko(Nimi As String, VJoukko As Object)  ' AcadSelectionSet - muutettu late binding (64-bit)
  Dim i As Integer
  
  On Error GoTo ErrorHandler
  
  ' Delete if already exists
  For i = 0 To oDOC.SelectionSets.Count - 1
    If oDOC.SelectionSets(i).Name = Nimi Then
      oDOC.SelectionSets(i).Delete
      Exit For
    End If
  Next i
  
  Set VJoukko = ActiveDocument.SelectionSets.Add(Nimi)
  Exit Sub
  
ErrorHandler:
  MsgBox "Error creating selection set: " & Err.Description, vbExclamation
End Sub

'================================================================================
' Sub: AvaaBlokki
' Purpose: Display and highlight current pipeline block, load attributes
'================================================================================
Private Sub AvaaBlokki()
  Dim Attrib As Variant
  Dim MinPoint As Variant
  Dim MaxPoint As Variant
  Dim j As Integer
  
  On Error GoTo ErrorHandler
  
  Me.Caption = i + 1 & "/" & Joukko.Count
  
  ' Un-highlight previous block
  If Not oBlock Is Nothing Then
    oBlock.Highlight False
  End If
  
  ' Get and display current block
  Set oBlock = Joukko(i)
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
  MsgBox "Error displaying block: " & Err.Description, vbExclamation
End Sub

'================================================================================
' Function: Yhdista
' Purpose: Connect to running AutoCAD instance
' Returns: True if connection successful, False otherwise
'================================================================================
Private Function Yhdista() As Boolean
  On Error Resume Next
  
  ' Try to connect to running AutoCAD
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
' Event: TFrom_AfterUpdate
' Purpose: Update FROM attribute when text box value changes
'================================================================================
Private Sub TFrom_AfterUpdate()
  On Error GoTo ErrorHandler
  
  KirjoitaAtt Nz(TFrom.Value), AttribFrom
  
  Exit Sub
ErrorHandler:
  MsgBox "Error updating FROM: " & Err.Description, vbExclamation
End Sub

'================================================================================
' Event: TTo_AfterUpdate
' Purpose: Update TO attribute when text box value changes
'================================================================================
Private Sub TTo_AfterUpdate()
  On Error GoTo ErrorHandler
  
  KirjoitaAtt Nz(TTo.Value), AttribTo
  
  Exit Sub
ErrorHandler:
  MsgBox "Error updating TO: " & Err.Description, vbExclamation
End Sub

