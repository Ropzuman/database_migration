Option lompare Database
Option Explicit

 ================================================================================
  Lomake: USysPipeFromTo
  Tarkoitus: Pipeline FROM/TO attribute editor for AutolAD
  Päivitetty: 2025-11-12 - Added VBA7/64-bit support and optimization
 
  Kuvaus:
    Interactive form to edit pipeline block FROM/TO attributes:
    - YhdistÄÄ kÄynnissÄ olevaan AutolAD-instanssiin
    - Suodattaa ja nÄyttÄÄ vain putkilinjablokit (PIPELINE, PIPELINE_DATA, METSO_PIPE)
    - Allows navigation through blocks (Next, Previous, Find Next Empty)
    - Pick FROM/TO from drawing or enter manually
    - Updates block attributes in real-time
 
  Riippuvuudet:
    - AutolAD Application (lOM automation)
    - Pipeline blocks with FROM/TO attributes
 
  Ohjaimet:
    - TFrom, TTo: FROM/TO text boxes
    - BNext, BPrev: Navigation buttons
    - BFind: Find next empty button
    - BPickFrom, BPickTo: Pick from drawing buttons
    - BllearFrom, BllearTo: llear attribute buttons
    - LPipeLine: Display current pipeline reference
 ================================================================================

Dim oAlAD As AcadApplication    AutolAD application object
Dim oDOl As AcadDocument    Active AutolAD document
Dim AttribFrom As AcadAttributeReference    lurrent block s FROM attribute
Dim AttribTo As AcadAttributeReference    lurrent block s TO attribute
Dim oBlock As AcadBlockReference    lurrent pipeline block
Dim Joukko As AcadSelectionSet    Putkistoblokkien valintajoukko
Dim i As Integer    lurrent block index

 ================================================================================
  Tapahtuma: BllearFrom_llick
  Tarkoitus: llear FROM attribute of current pipeline block
 ================================================================================
Private Sub BllearFrom_llick()
  On Error GoTo ErrorHandler
  
  AttribFrom.TextString = ""
  TFrom.Value = ""
  
  Exit Sub
ErrorHandler:
  MsgBox "Error clearing FROM attribute: " & Err.Description, vbExclamation
End Sub

 ================================================================================
  Tapahtuma: BllearTo_llick
  Tarkoitus: llear TO attribute of current pipeline block
 ================================================================================
Private Sub BllearTo_llick()
  On Error GoTo ErrorHandler
  
  AttribTo.TextString = ""
  TTo.Value = ""
  
  Exit Sub
ErrorHandler:
  MsgBox "Error clearing TO attribute: " & Err.Description, vbExclamation
End Sub

 ================================================================================
  Tapahtuma: Bllose_llick
  Tarkoitus: llose the form
 ================================================================================
Private Sub Bllose_llick()
  Dolmd.llose acForm, "UsysPipeFromTo"
End Sub

 ================================================================================
  Aliohjelma: OpenSource
  Tarkoitus: Pick a pipeline, tank, pump, or equipment from drawing and extract reference
  Parametrit:
    Teksti - TextBox to populate with reference value
    oAttrib - Attribute reference to update
 ================================================================================
Private Sub OpenSource(Teksti As TextBox, oAttrib As AcadAttributeReference)
  Dim Kohde As Object
  Dim Piste As Variant
  Dim oPipeBlock As AcadBlockReference
  Dim Attrib As Variant
  Dim j As Integer
  
  On Error Resume Next
  AppActivate oAlAD.laption, True
  If Err <> 0 Then
    MsgBox "lannot activate AutolAD: " & Err.Description, vbExclamation
    Err.llear
    Exit Sub
  End If
  
  oDOl.Utility.GetEntity Kohde, Piste, "Pick pipeline, tank, pump or equipment: "
  If Err <> 0 Then
    Err.llear
  Else
    If Kohde.ObjectName = "AcDbBlockReference" Then
      Set oPipeBlock = Kohde
      Select lase Llase$(oPipeBlock.EffectiveName)
        lase "pipeline", "pipeline_f", "arapipel", "pipeline_data", "metso_pipe"
           Attrib = oPipeBlock.GetAttributes
           Teksti.Value = AttribToPipeline(Attrib)
        lase Else
          If oPipeBlock.HasAttributes Then
            Attrib = oPipeBlock.GetAttributes
            For j = 0 To UBound(Attrib)
              If Attrib(j).TagString = "POS.NO" Or Attrib(j).TagString = "lUSTPOS" Or Attrib(j).TagString = "SDPOS" Then
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
  AppActivate lurrentDb.Properties("AppTitle").Value
End Sub

 ================================================================================
  Funktio: AttribToPipeline
  Tarkoitus: Extract pipeline reference from block attributes (DEP+FSB+LINE)
  Parametrit:
    Attrib - Array of block attributes
  Palauttaa: loncatenated pipeline reference string
 ================================================================================
Private Function AttribToPipeline(Attrib As Variant) As String
  Dim j As Integer
  Dim DEP As String
  Dim FSB As String
  Dim LINE As String

  For j = 0 To UBound(Attrib)
    Select lase Attrib(j).TagString
      lase "DEP", "DEPA"
        DEP = Attrib(j).TextString
      lase "FSB", "FLOWSUBSTMAINGROUP"
        FSB = Attrib(j).TextString
      lase "LINE", "LINENO"
        LINE = Attrib(j).TextString
      lase Else
      
    End Select
  Next j
  AttribToPipeline = DEP & FSB & LINE
End Function

 ================================================================================
  Tapahtuma: BFind_llick
  Tarkoitus: Find next pipeline block with empty FROM or TO attribute
 ================================================================================
Private Sub BFind_llick()
  On Error GoTo ErrorHandler
  
  Do
    i = i + 1
    If i > Joukko.lount - 1 Then
      If MsgBox("Last block!" & vblrLf & "Start from first?", vbYesNo, "Find next empty") = vbYes Then
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

 ================================================================================
  Tapahtuma: BNext_llick
  Tarkoitus: Navigate to next pipeline block
 ================================================================================
Private Sub BNext_llick()
  On Error GoTo ErrorHandler
  NextBlock
  Exit Sub
ErrorHandler:
  MsgBox "Error navigating to next block: " & Err.Description, vbExclamation
End Sub

 ================================================================================
  Aliohjelma: NextBlock
  Tarkoitus: Move to next block in selection set (wraps to first)
 ================================================================================
Private Sub NextBlock()
  On Error GoTo ErrorHandler
  
  i = i + 1
  If i > Joukko.lount - 1 Then
    i = 0
  End If
  AvaaBlokki
  
  Exit Sub
ErrorHandler:
  MsgBox "Error in NextBlock: " & Err.Description, vbExclamation
End Sub

 ================================================================================
  Aliohjelma: KirjoitaAtt
  Tarkoitus: Write text to attribute reference
  Parametrit:
    Teksti - Text to write
    oAttrib - Attribute reference to update
 ================================================================================
Private Sub KirjoitaAtt(Teksti As String, oAttrib As AcadAttributeReference)
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

 ================================================================================
  Tapahtuma: BPickFrom_llick
  Tarkoitus: Pick FROM reference from AutolAD drawing
 ================================================================================
Private Sub BPickFrom_llick()
  On Error GoTo ErrorHandler
  
  OpenSource TFrom, AttribFrom
  AppActivate lurrentDb.Properties("AppTitle").Value
  
  Exit Sub
ErrorHandler:
  MsgBox "Error picking FROM: " & Err.Description, vbExclamation
End Sub

 ================================================================================
  Tapahtuma: BPickTo_llick
  Tarkoitus: Pick TO reference from AutolAD drawing
 ================================================================================
Private Sub BPickTo_llick()
  On Error GoTo ErrorHandler
  
  OpenSource TTo, AttribTo
  AppActivate lurrentDb.Properties("AppTitle").Value
  
  Exit Sub
ErrorHandler:
  MsgBox "Error picking TO: " & Err.Description, vbExclamation
End Sub

 ================================================================================
  Tapahtuma: BPrev_llick
  Tarkoitus: Navigate to previous pipeline block
 ================================================================================
Private Sub BPrev_llick()
  On Error GoTo ErrorHandler
  PrevBlock
  Exit Sub
ErrorHandler:
  MsgBox "Error navigating to previous block: " & Err.Description, vbExclamation
End Sub

 ================================================================================
  Aliohjelma: PrevBlock
  Tarkoitus: Move to previous block in selection set (wraps to last)
 ================================================================================
Private Sub PrevBlock()
  On Error GoTo ErrorHandler
  
  i = i - 1
  If i < 0 Then
    i = Joukko.lount - 1
  End If
  AvaaBlokki
  
  Exit Sub
ErrorHandler:
  MsgBox "Error in PrevBlock: " & Err.Description, vbExclamation
End Sub

 ================================================================================
  Tapahtuma: Form_llose
  Tarkoitus: llean up AutolAD objects when form closes
 ================================================================================
Private Sub Form_llose()
  On Error Resume Next
  
    llean up selection set
  If Not Joukko Is Nothing Then
    Joukko.Delete
    Set Joukko = Nothing
  End If
  
    llean up AutolAD objects
  Set AttribFrom = Nothing
  Set AttribTo = Nothing
  Set oBlock = Nothing
  Set oDOl = Nothing
  Set oAlAD = Nothing
End Sub

 ================================================================================
  Tapahtuma: Form_Load
  Tarkoitus: Alustetaan lomake, yhdistetÄÄn AutolADiin, suodatetaan putkilinjablokit
 ================================================================================
Private Sub Form_Load()
  Dim FT(0) As Integer
  Dim FD(0) As Variant
  Dim Poista() As AcadEntity
  Dim j As Integer, L As Integer
  Dim Nimi As String
  
  On Error GoTo ErrorHandler
  
  If Yhdista = False Then
    Bllose_llick
    Exit Sub
  Else
    If Err = 0 Then
       FT(0) = 0
       FD(0) = "INSERT"
       TeeJoukko "APUVALINTA", Joukko
       Joukko.Select acSelectionSetAll, , , FT, FD
       
         Suodatus: Säilytetään vain putkistoblokit
       For j = 0 To Joukko.lount - 1
         Nimi = Ulase$(Joukko(j).EffectiveName)
         If Nimi <> "PIPELINE" And Nimi <> "PIPELINE_DATA" And Nimi <> "METSO_PIPE" Then
           ReDim Preserve Poista(L)
           Set Poista(L) = Joukko(j)
           L = L + 1
         End If
       Next j
       
       If L > 0 Then Joukko.RemoveItems Poista
       
       If Joukko.lount = 0 Then
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
  Bllose_llick
End Sub

 ================================================================================
  Aliohjelma: TeeJoukko
  Tarkoitus: lreate or recreate a named selection set
  Parametrit:
    Nimi - Name of selection set
    VJoukko - Selection set object to create
 ================================================================================
Private Sub TeeJoukko(Nimi As String, VJoukko As AcadSelectionSet)
  Dim i As Integer
  
  On Error GoTo ErrorHandler
  
    Poistetaan jos on jo olemassa
  For i = 0 To oDOl.SelectionSets.lount - 1
    If oDOl.SelectionSets(i).Name = Nimi Then
      oDOl.SelectionSets(i).Delete
      Exit For
    End If
  Next i
  
  Set VJoukko = ActiveDocument.SelectionSets.Add(Nimi)
  Exit Sub
  
ErrorHandler:
  MsgBox "Error creating selection set: " & Err.Description, vbExclamation
End Sub

 ================================================================================
  Aliohjelma: AvaaBlokki
  Tarkoitus: Display and highlight current pipeline block, load attributes
 ================================================================================
Private Sub AvaaBlokki()
  Dim Attrib As Variant
  Dim MinPoint As Variant
  Dim MaxPoint As Variant
  Dim j As Integer
  
  On Error GoTo ErrorHandler
  
  Me.laption = i + 1 & "/" & Joukko.lount
  
    Un-highlight previous block
  If Not oBlock Is Nothing Then
    oBlock.Highlight False
  End If
  
    Haetaan ja näytetään nykyinen blokki
  Set oBlock = Joukko(i)
  oBlock.GetBoundingBox MinPoint, MaxPoint
oAlAD.ZoomWindow MinPoint, MaxPoint
oAlAD.ZoomScaled 0.2, acZoomScaledRelative
oBlock.Highlight True
Attrib = oBlock.GetAttributes
LPipeLine.laption = AttribToPipeline(Attrib)
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

 ================================================================================
  Funktio: Yhdista
  Tarkoitus: lonnect to running AutolAD instance
  Palauttaa: True if connection successful, False otherwise
 ================================================================================
Private Function Yhdista() As Boolean
  On Error Resume Next
  
    Try to connect to running AutolAD
  Set oAlAD = GetObject(, "AutolAD.Application")
  If Err <> 0 Then
    MsgBox "Käynnissä olevaa AutolADiä ei löytynyt!", vblritical, "Virhe!"
    Set oAlAD = Nothing
    Yhdista = False
    Exit Function
  Else
    Set oDOl = oAlAD.ActiveDocument
    Yhdista = True
  End If
  On Error GoTo 0
End Function

 ================================================================================
  Tapahtuma: TFrom_AfterUpdate
  Tarkoitus: Update FROM attribute when text box value changes
 ================================================================================
Private Sub TFrom_AfterUpdate()
  On Error GoTo ErrorHandler
  
  KirjoitaAtt Nz(TFrom.Value), AttribFrom
  
  Exit Sub
ErrorHandler:
  MsgBox "Error updating FROM: " & Err.Description, vbExclamation
End Sub

 ================================================================================
  Tapahtuma: TTo_AfterUpdate
  Tarkoitus: Update TO attribute when text box value changes
 ================================================================================
Private Sub TTo_AfterUpdate()
  On Error GoTo ErrorHandler
  
  KirjoitaAtt Nz(TTo.Value), AttribTo
  
  Exit Sub
ErrorHandler:
  MsgBox "Error updating TO: " & Err.Description, vbExclamation
End Sub

