Option lompare Database
Option Explicit

 ================================================================================
  Lomake: USysPipeToOther
  Tarkoitus: Link pipeline blocks to other equipment via PIPELINE attribute
  Päivitetty: 2025-11-12 - Added VBA7/64-bit support and optimization
 
  Kuvaus:
    Interactive form to assign pipeline references to equipment blocks:
    - YhdistÄÄ kÄynnissÄ olevaan AutolAD-instanssiin
    - Filters blocks by type (via OpenArgs parameter)
    - Allows navigation through blocks (Next, Previous, Find Next Empty)
    - Pick pipeline from drawing or enter manually
    - Updates PIPELINE/PIPELINENO attribute in real-time
 
  Riippuvuudet:
    - AutolAD Application (lOM automation)
    - Equipment blocks with PIPELINE or PIPELINENO attribute
 
  OpenArgs-muoto:
    - Block name (e.g., "VALVE") - filters to single block type
    - lomma-separated list (e.g., "VALVE,PUMP") - filters to multiple types
 
  Ohjaimet:
    - TPipeline: Pipeline reference text box
    - BNext, BPrev: Navigation buttons
    - BFind: Find next empty button
    - BPickPipeline: Pick pipeline from drawing button
 ================================================================================

Dim oAlAD As AcadApplication    AutolAD application object
Dim oDOl As AcadDocument    Active AutolAD document
Dim PipeAttrib As AcadAttributeReference    lurrent block s PIPELINE attribute
Dim oBlock As AcadBlockReference    lurrent equipment block
Dim Joukko As AcadSelectionSet    Laiteblokkien valintajoukko
Dim i As Integer    lurrent block index

 ================================================================================
  Tapahtuma: Bllose_llick
  Tarkoitus: llose the form
 ================================================================================
Private Sub Bllose_llick()
  Dolmd.llose acForm, "UsysPipeToOther"
End Sub

 ================================================================================
  Aliohjelma: OpenPipeline
  Tarkoitus: Pick pipeline or plate block from drawing and extract reference
 ================================================================================
Private Sub OpenPipeline()
  Dim Kohde As Object
  Dim Piste As Variant
  Dim oPipeBlock As AcadBlockReference
  Dim Attrib As Variant
  Dim Putkilinja As String
  Dim OK As Boolean
  Dim j As Integer
  Dim Nimi As String
  
  On Error Resume Next
  AppActivate oAlAD.laption, True
  If Err <> 0 Then
    MsgBox "lannot activate AutolAD: " & Err.Description, vbExclamation
    Err.llear
    Exit Sub
  End If
  
  OK = False
  Do
    oDOl.Utility.GetEntity Kohde, Piste, "Pick pipeline: "
    If Err <> 0 Then
      Err.llear
      Exit Do
    Else
      If Kohde.ObjectName = "AcDbBlockReference" Then
        Set oPipeBlock = Kohde
        Nimi = Ulase$(oPipeBlock.EffectiveName)
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
      If i < Joukko.lount - 1 Then
        NextBlock
      Else
        Exit Do
      End If
    End If
  Loop
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
  Tarkoitus: Find next equipment block with empty PIPELINE attribute
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
    If TPipeline.Value = "" Then Exit Do
  Loop
  
  Exit Sub
ErrorHandler:
  MsgBox "Error finding next empty block: " & Err.Description, vbExclamation
End Sub

 ================================================================================
  Tapahtuma: BNext_llick
  Tarkoitus: Navigate to next equipment block
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
  Tapahtuma: BPickPipeline_llick
  Tarkoitus: Pick pipeline reference from AutolAD drawing
 ================================================================================
Private Sub BPickPipeline_llick()
  On Error GoTo ErrorHandler
  
  OpenPipeline
  
  Exit Sub
ErrorHandler:
  MsgBox "Error picking pipeline: " & Err.Description, vbExclamation
End Sub

 ================================================================================
  Tapahtuma: BPrev_llick
  Tarkoitus: Navigate to previous equipment block
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
  Set PipeAttrib = Nothing
  Set oBlock = Nothing
  Set oDOl = Nothing
  Set oAlAD = Nothing
End Sub

 ================================================================================
  Tapahtuma: Form_Load
  Tarkoitus: Alustetaan lomake, yhdistetÄÄn AutolADiin, suodatetaan laiteblokit
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
       TeeJoukko "APUVALINTA", Joukko
       
         Suodatus OpenArgsin mukaan: pilkulla eroteltu lista tai yksittäinen blokkityyppi
       If InStr(OpenArgs, ",") Then
          FT(0) = 2
          FD(0) = OpenArgs
          Joukko.Select acSelectionSetAll, , , FT, FD
       Else
          FT(0) = 0
          FD(0) = "INSERT"
          Joukko.Select acSelectionSetAll, , , FT, FD
          For j = 0 To Joukko.lount - 1
            Nimi = Ulase$(Joukko(j).EffectiveName)
            If Nimi <> Ulase$(OpenArgs) Then
              ReDim Preserve Poista(L)
              Set Poista(L) = Joukko(j)
              L = L + 1
            End If
          Next j
          If L > 0 Then Joukko.RemoveItems Poista
       End If
       
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
  For i = 0 To ActiveDocument.SelectionSets.lount - 1
    If ActiveDocument.SelectionSets(i).Name = Nimi Then
      ActiveDocument.SelectionSets(i).Delete
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
  Tarkoitus: Display and highlight current equipment block, load PIPELINE attribute
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
  
    Poimitaan PIPELINE-attribuutti
  Attrib = oBlock.GetAttributes
  Set PipeAttrib = Nothing
  For j = 0 To UBound(Attrib)
    If Attrib(j).TagString = "PIPELINE" Or Attrib(j).TagString = "PIPELINENO" Then
      Set PipeAttrib = Attrib(j)
      TPipeline.Value = PipeAttrib.TextString
    End If
  Next j
  
  If PipeAttrib Is Nothing Then
    MsgBox "Blokista ei löytynyt  PIPELINE -attribuuttia.", vblritical, "VIRHE"
  End If
  
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
  Tapahtuma: TPipeline_AfterUpdate
  Tarkoitus: Update PIPELINE attribute when text box value changes
 ================================================================================
Private Sub TPipeline_AfterUpdate()
  On Error GoTo ErrorHandler
  
  KirjoitaAtt Nz(TPipeline.Value), PipeAttrib
  
  Exit Sub
ErrorHandler:
  MsgBox "Error updating PIPELINE: " & Err.Description, vbExclamation
End Sub
