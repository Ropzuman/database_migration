Option Compare Database
Option Explicit

'================================================================================
' Form: TYÖKALUT
' Purpose: Main AutoCAD P&I import/update toolbox for PIPE database
' Updated: 2025-11-12 - Added VBA7/64-bit support and optimization
'
' Description:
'   Central form for importing and updating P&I diagram data from AutoCAD:
'   - Batch processes .DWG files from specified folder
'   - Imports pipeline, valve, field instrument, and loop data
'   - Three-step workflow per data type:
'     1. Read from AutoCAD (Command1x_Click)
'     2. Append to staging table (Command2x_Click)
'     3. Update database (Command3x_Click)
'   - Opens specialized editor forms (FlowPickNo, PipeFromTo, PipeToOther)
'
' Dependencies:
'   - AutoCAD Application (COM automation)
'   - DAO.Recordset for database operations
'   - Query actions: *_APPEND, *_UPDATE, DELETE queries
'   - LPolku label control: Specifies subfolder for .DWG files
'
' Data Types:
'   - PIPELINEDATA: Pipeline blocks (PIPELINE_DATA)
'   - MANVALVEDATA: Manual valves (MANUALVALVE_DATA)
'   - INSTVALVEDATA: Instrument valves (INSTVALVE_DATA)
'   - FIELDINSTDATA: Field instruments (FIELDINST_DATA)
'   - INSTLOOPDATA: Instrument loops (UP079D, UP008C, UP085D, UP08C)
'
' Controls:
'   - Command11-13: Pipeline workflow
'   - Command21-23: Manual valve workflow
'   - Command31-33: Instrument valve workflow
'   - Command41-43: Field instrument workflow
'   - Command51-53: Instrument loop workflow
'   - Command18, 19, 40: Open editor forms
'   - BPipelineToManval: Link pipelines to manual valves
'================================================================================

Dim oACAD As AcadApplication  ' AutoCAD application object
Dim oDOC As AcadDocument  ' Current AutoCAD document being processed
Dim Tied As String  ' Current .DWG filename during batch processing
Dim Joukko As AcadSelectionSet  ' Selection set for filtered blocks

'===== Pipeline to Manual Valve =====
'================================================================================
' Event: BPipelineToManval_Click
' Purpose: Open form to link pipeline references to manual valves
'================================================================================
Private Sub BPipelineToManval_Click()
  On Error GoTo ErrorHandler
  
  If MsgBox("Onko haluttu P&I kaavio auki AutoCADissä?", vbYesNo, "Pipeline To Manualvalve") = vbYes Then
    DoCmd.OpenForm "USysPipeToOther", , , , , , "MANUALVALVE_DATA"
  End If
  
  Exit Sub
ErrorHandler:
  MsgBox "Error opening PipeToOther form: " & Err.Description, vbExclamation
End Sub

'===== Pipeline Workflow (Commands 11-13) =====
'================================================================================
' Event: Command11_Click
' Purpose: Read pipeline data from AutoCAD drawings
'================================================================================
Private Sub Command11_Click()
  On Error GoTo ErrorHandler
  
  Tied = ""
  LueTiedot "PIPELINEDATA", "PIPELINE_DATA"
  Application.SysCmd acSysCmdClearStatus
  MsgBox "Putkilinjojen luku valmis", vbInformation
  
  Exit Sub
ErrorHandler:
  Application.SysCmd acSysCmdClearStatus
  MsgBox "Error reading pipeline data: " & Err.Description, vbExclamation
End Sub

'================================================================================
' Event: Command12_Click
' Purpose: Append pipeline data to staging table
'================================================================================
Private Sub Command12_Click()
  On Error GoTo ErrorHandler
  
  CurrentDb.Execute "PIPELINES_APPEND"
  MsgBox "Putkilinjojen lisäys valmis", vbInformation
  
  Exit Sub
ErrorHandler:
  MsgBox "Error appending pipeline data: " & Err.Description, vbExclamation
End Sub

'================================================================================
' Event: Command13_Click
' Purpose: Update database with pipeline data from staging table
'================================================================================
Private Sub Command13_Click()
  On Error GoTo ErrorHandler
  
  If MsgBox("Varoitus! Tämä ajaa kaavioissa olevat tiedot kannassa olevien tietojen päälle." & vbCrLf & "Haluatko jatkaa?", vbYesNo, "Päivitä tiedot") = vbYes Then
    CurrentDb.Execute "DELETE * From UsysPIPEUPD"
    CurrentDb.Execute "PIPELINES_UPDATE_START"
    CurrentDb.Execute "PIPELINES_UPDATE"
    MsgBox "Putkilinjojen päivitys valmis", vbInformation
  End If
  
  Exit Sub
ErrorHandler:
  MsgBox "Error updating pipeline data: " & Err.Description, vbExclamation
End Sub

'===== Editor Forms (Commands 18-19, 40) =====
'================================================================================
' Event: Command18_Click
' Purpose: Open FlowPickNo editor form
'================================================================================
Private Sub Command18_Click()
  On Error GoTo ErrorHandler
  
  If MsgBox("Onko haluttu P&I kaavio auki AutoCADissä?", vbYesNo, "FlowPickNo") = vbYes Then
    DoCmd.OpenForm "USysFlowpickNo"
  End If
  
  Exit Sub
ErrorHandler:
  MsgBox "Error opening FlowPickNo form: " & Err.Description, vbExclamation
End Sub

'================================================================================
' Event: Command19_Click
' Purpose: Open Pipeline FROM/TO editor form
'================================================================================
Private Sub Command19_Click()
  On Error GoTo ErrorHandler
  
  If MsgBox("Onko haluttu P&I kaavio auki AutoCADissä?", vbYesNo, "Pipeline From/To") = vbYes Then
    DoCmd.OpenForm "USysPipeFromTo"
  End If
  
  Exit Sub
ErrorHandler:
  MsgBox "Error opening PipeFromTo form: " & Err.Description, vbExclamation
End Sub

'===== Manual Valve Workflow (Commands 21-23) =====
'================================================================================
' Event: Command21_Click
' Purpose: Read manual valve data from AutoCAD drawings
'================================================================================
Private Sub Command21_Click()
  On Error GoTo ErrorHandler
  
  Tied = ""
  LueTiedot "MANVALVEDATA", "MANUALVALVE_DATA"
  Application.SysCmd acSysCmdClearStatus
  MsgBox "Käsiventtiilien luku valmis", vbInformation
  
  Exit Sub
ErrorHandler:
  Application.SysCmd acSysCmdClearStatus
  MsgBox "Error reading manual valve data: " & Err.Description, vbExclamation
End Sub

'================================================================================
' Event: Command22_Click
' Purpose: Append manual valve data to staging table
'================================================================================
Private Sub Command22_Click()
  On Error GoTo ErrorHandler
  
  CurrentDb.Execute "MANVALVE_APPEND"
  MsgBox "Käsiventtiilien lisäys valmis", vbInformation
  
  Exit Sub
ErrorHandler:
  MsgBox "Error appending manual valve data: " & Err.Description, vbExclamation
End Sub

'================================================================================
' Event: Command23_Click
' Purpose: Update database with manual valve data from staging table
'================================================================================
Private Sub Command23_Click()
  On Error GoTo ErrorHandler
  
  If MsgBox("Varoitus! Tämä ajaa kaavioissa olevat tiedot kannassa olevien tietojen päälle." & vbCrLf _
         & "Uudet code ja pipeline tiedot otetaan kaaviosta tuoduista tiedoista." & vbCrLf _
         & "Jos olet tehnyt muutoksia näihin tietoihin kannassa, ne ylikirjoitetaan kaavion tiedoilla." & vbCrLf _
         & "Haluatko jatkaa?", vbYesNo, "Päivitä tiedot") = vbYes Then
    CurrentDb.Execute "MANVALVE_UPDATE"
    MsgBox "Käsiventtiilien päivitys valmis", vbInformation
  End If
  
  Exit Sub
ErrorHandler:
  MsgBox "Error updating manual valve data: " & Err.Description, vbExclamation
End Sub

'===== Instrument Valve Workflow (Commands 31-33) =====
'================================================================================
' Event: Command31_Click
' Purpose: Read instrument valve data from AutoCAD drawings
' Note: Searches for UP079D blocks where LOOPNO attribute contains "AA"
'================================================================================
Private Sub Command31_Click()
  On Error GoTo ErrorHandler
  
  Tied = ""
  LueTiedotByBlockAndAttribute "INSTVALVEDATA", "UP079D", "LOOPNO", "AA"
  Application.SysCmd acSysCmdClearStatus
  MsgBox "Instrumenttiventtiilien luku valmis", vbInformation
  
  Exit Sub
ErrorHandler:
  Application.SysCmd acSysCmdClearStatus
  MsgBox "Error reading instrument valve data: " & Err.Description, vbExclamation
End Sub

'================================================================================
' Event: Command32_Click
' Purpose: Append instrument valve data to staging table
'================================================================================
Private Sub Command32_Click()
  On Error GoTo ErrorHandler
  
  CurrentDb.Execute "INSTVALVE_APPEND"
  MsgBox "Instrumenttiventtiilien lisäys valmis", vbInformation
  
  Exit Sub
ErrorHandler:
  MsgBox "Error appending instrument valve data: " & Err.Description, vbExclamation
End Sub

'================================================================================
' Event: Command33_Click
' Purpose: Update database with instrument valve data from staging table
'================================================================================
Private Sub Command33_Click()
  On Error GoTo ErrorHandler
  
  If MsgBox("Varoitus! Tämä ajaa kaavioissa olevat tiedot kannassa olevien tietojen päälle." & vbCrLf _
         & "Uudet code ja pipeline tiedot otetaan kaaviosta tuoduista tiedoista." & vbCrLf _
         & "Jos olet tehnyt muutoksia näihin tietoihin kannassa, ne ylikirjoitetaan kaavion tiedoilla." & vbCrLf _
         & "Haluatko jatkaa?", vbYesNo, "Päivitä tiedot") = vbYes Then
    CurrentDb.Execute "INSTVALVE_UPDATE"
    MsgBox "Instrumenttiventtiilien päivitys valmis", vbInformation
  End If
  
  Exit Sub
ErrorHandler:
  MsgBox "Error updating instrument valve data: " & Err.Description, vbExclamation
End Sub

'================================================================================
' Event: Command40_Click
' Purpose: Open form to link pipeline references to instrument loops
'================================================================================
Private Sub Command40_Click()
  On Error GoTo ErrorHandler
  
  If MsgBox("Onko haluttu P&I kaavio auki AutoCADissä?", vbYesNo, "Pipeline To Loops") = vbYes Then
    DoCmd.OpenForm "USysPipeToOther", , , , , , "UP079D,UP008C,UP085D,UP08C,"
  End If
  
  Exit Sub
ErrorHandler:
  MsgBox "Error opening PipeToOther form: " & Err.Description, vbExclamation
End Sub

'===== Field Instrument Workflow (Commands 41-43) =====
'================================================================================
' Event: Command41_Click
' Purpose: Read field instrument data from AutoCAD drawings (UP080D blocks)
'================================================================================
Private Sub Command41_Click()
  On Error GoTo ErrorHandler
  
  Tied = ""
  LueTiedot "FIELDINSTDATA", "UP080D"
  Application.SysCmd acSysCmdClearStatus
  MsgBox "Kenttäinstrumenttien luku valmis", vbInformation
  
  Exit Sub
ErrorHandler:
  Application.SysCmd acSysCmdClearStatus
  MsgBox "Error reading field instrument data: " & Err.Description, vbExclamation
End Sub

'================================================================================
' Event: Command42_Click
' Purpose: Append field instrument data to staging table
'================================================================================
Private Sub Command42_Click()
  On Error GoTo ErrorHandler
  
  CurrentDb.Execute "FIELDINST_APPEND"
  MsgBox "Kenttäinstrumenttien lisäys valmis", vbInformation
  
  Exit Sub
ErrorHandler:
  MsgBox "Error appending field instrument data: " & Err.Description, vbExclamation
End Sub

'================================================================================
' Event: Command43_Click
' Purpose: Update database with field instrument data from staging table
'================================================================================
Private Sub Command43_Click()
  On Error GoTo ErrorHandler
  
  If MsgBox("Varoitus! Tämä ajaa kaavioissa olevat tiedot kannassa olevien tietojen päälle." & vbCrLf _
         & "Uudet code ja pipeline tiedot otetaan kaaviosta tuoduista tiedoista." & vbCrLf _
         & "Jos olet tehnyt muutoksia näihin tietoihin kannassa, ne ylikirjoitetaan kaavion tiedoilla." & vbCrLf _
         & "Haluatko jatkaa?", vbYesNo, "Päivitä tiedot") = vbYesNo Then
    CurrentDb.Execute "FIELDINST_UPDATE"
    MsgBox "Kenttäinstrumenttien päivitys valmis", vbInformation
  End If
  
  Exit Sub
ErrorHandler:
  MsgBox "Error updating field instrument data: " & Err.Description, vbExclamation
End Sub

'===== Instrument Loop Workflow (Commands 51-53) =====
'================================================================================
' Event: Command51_Click
' Purpose: Read instrument loop data from AutoCAD drawings
'================================================================================
Private Sub Command51_Click()
  On Error GoTo ErrorHandler
  
  Tied = ""
  LueTiedot "INSTLOOPDATA", "UP079D"
  Application.SysCmd acSysCmdClearStatus
  MsgBox "Instrumenttipiirien luku valmis", vbInformation
  
  Exit Sub
ErrorHandler:
  Application.SysCmd acSysCmdClearStatus
  MsgBox "Error reading instrument loop data: " & Err.Description, vbExclamation
End Sub

'================================================================================
' Event: Command52_Click
' Purpose: Append instrument loop data to staging table
'================================================================================
Private Sub Command52_Click()
  On Error GoTo ErrorHandler
  
  CurrentDb.Execute "INSTLOOP_APPEND"
  MsgBox "Instrumenttipiirien lisäys valmis", vbInformation
  
  Exit Sub
ErrorHandler:
  MsgBox "Error appending instrument loop data: " & Err.Description, vbExclamation
End Sub

'================================================================================
' Event: Command53_Click
' Purpose: Update database with instrument loop data from staging table
'================================================================================
Private Sub Command53_Click()
  On Error GoTo ErrorHandler
  
  If MsgBox("Varoitus! Tämä ajaa kaavioissa olevat tiedot kannassa olevien tietojen päälle." & vbCrLf _
         & "Uudet code ja pipeline tiedot otetaan kaaviosta tuoduista tiedoista." & vbCrLf _
         & "Jos olet tehnyt muutoksia näihin tietoihin kannassa, ne ylikirjoitetaan kaavion tiedoilla." & vbCrLf _
         & "Haluatko jatkaa?", vbYesNo, "Päivitä tiedot") = vbYes Then
    CurrentDb.Execute "INSTLOOP_UPDATE"
    MsgBox "Instrumenttipiirien päivitys valmis", vbInformation
  End If
  
  Exit Sub
ErrorHandler:
  MsgBox "Error updating instrument loop data: " & Err.Description, vbExclamation
End Sub

'===== Helper Functions =====
'================================================================================
' Function: FieldExists
' Purpose: Check if a field exists in a recordset
' Parameters:
'   RS - Recordset to check
'   FieldName - Field name to look for
' Returns: True if field exists, False otherwise
'================================================================================
Private Function FieldExists(RS As DAO.Recordset, FieldName As String) As Boolean
  Dim fld As DAO.Field
  On Error Resume Next
  Set fld = RS.Fields(FieldName)
  FieldExists = (Err.Number = 0)
  On Error GoTo 0
End Function

'================================================================================
' Sub: LueTiedot
' Purpose: Read block data from AutoCAD drawings and store in Access table
' Parameters:
'   Taulukko - Target Access table name
'   Blokki - Block type(s) to extract (comma-separated for multiple)
'================================================================================
Private Sub LueTiedot(Taulukko As String, Blokki As String)
  Dim i As Integer, j As Integer
  Dim oBlock As AcadBlockReference
  Dim Attrib As Variant
  Dim Tiedot As DAO.Recordset
  Dim fld As DAO.Field
  
  On Error GoTo ErrorHandler
  
  CurrentDb.Execute "DELETE * FROM " & Taulukko
  
  Do While SeuraavaKuva
    Set Tiedot = CurrentDb.OpenRecordset(Taulukko)
    PoimiJoukko Blokki
    
    For i = 0 To Joukko.Count - 1
      Application.SysCmd acSysCmdSetStatus, "Haetaan tietoja " & i + 1 & "/" & Joukko.Count
      Set oBlock = Joukko(i)
      Attrib = oBlock.GetAttributes
      
      Tiedot.AddNew
      Tiedot.Fields("PATH") = oDOC.Path
      Tiedot.Fields("ImpFileID") = oDOC.Name
      Tiedot.Fields("BLOCK") = oBlock.EffectiveName
      Tiedot.Fields("Handles") = oBlock.Handle
      
      For j = 0 To UBound(Attrib)
        If Attrib(j).TextString <> "" And FieldExists(Tiedot, "[" & Attrib(j).TagString & "]") Then
          Set fld = Tiedot.Fields("[" & Attrib(j).TagString & "]")
          fld.Value = IIf(Len(Attrib(j).TextString) > fld.Size, Left$(Attrib(j).TextString, fld.Size), Attrib(j).TextString)
        End If
      Next j
      
      Tiedot.Update
    Next i
    
    Tiedot.Close
    Set Tiedot = Nothing
    Joukko.Delete
    Set Joukko = Nothing
  Loop
  
  Application.SysCmd acSysCmdClearStatus
  Set oACAD = Nothing
  Set oDOC = Nothing
  Exit Sub
  
ErrorHandler:
  Application.SysCmd acSysCmdClearStatus
  If Not Tiedot Is Nothing Then
    Tiedot.Close
    Set Tiedot = Nothing
  End If
  MsgBox "Error reading data: " & Err.Description, vbExclamation
End Sub

'================================================================================
' Sub: LueTiedotByAttribute
' Purpose: Read block data from AutoCAD drawings filtered by attribute names
' Parameters:
'   Taulukko - Target Access table name
'   AttribNames - Attribute tag name(s) to filter by (comma-separated)
'================================================================================
Private Sub LueTiedotByAttribute(Taulukko As String, AttribNames As String)
  Dim i As Integer, j As Integer
  Dim oBlock As AcadBlockReference
  Dim Attrib As Variant
  Dim Tiedot As DAO.Recordset
  Dim fld As DAO.Field
  
  On Error GoTo ErrorHandler
  
  CurrentDb.Execute "DELETE * FROM " & Taulukko
  
  Do While SeuraavaKuva
    Set Tiedot = CurrentDb.OpenRecordset(Taulukko)
    PoimiJoukkoByAttribute AttribNames
    
    For i = 0 To Joukko.Count - 1
      Application.SysCmd acSysCmdSetStatus, "Haetaan tietoja " & i + 1 & "/" & Joukko.Count
      Set oBlock = Joukko(i)
      Attrib = oBlock.GetAttributes
      
      Tiedot.AddNew
      Tiedot.Fields("PATH") = oDOC.Path
      Tiedot.Fields("ImpFileID") = oDOC.Name
      Tiedot.Fields("BLOCK") = oBlock.EffectiveName
      Tiedot.Fields("Handles") = oBlock.Handle
      
      For j = 0 To UBound(Attrib)
        If Attrib(j).TextString <> "" And FieldExists(Tiedot, "[" & Attrib(j).TagString & "]") Then
          Set fld = Tiedot.Fields("[" & Attrib(j).TagString & "]")
          fld.Value = IIf(Len(Attrib(j).TextString) > fld.Size, Left$(Attrib(j).TextString, fld.Size), Attrib(j).TextString)
        End If
      Next j
      
      Tiedot.Update
    Next i
    
    Tiedot.Close
    Set Tiedot = Nothing
    Joukko.Delete
    Set Joukko = Nothing
  Loop
  
  Application.SysCmd acSysCmdClearStatus
  Set oACAD = Nothing
  Set oDOC = Nothing
  Exit Sub
  
ErrorHandler:
  Application.SysCmd acSysCmdClearStatus
  If Not Tiedot Is Nothing Then
    Tiedot.Close
    Set Tiedot = Nothing
  End If
  MsgBox "Error reading data by attribute: " & Err.Description, vbExclamation
End Sub

'================================================================================
' Sub: LueTiedotByBlockAndAttribute
' Purpose: Read blocks with specific name and attribute value
' Parameters:
'   Taulukko - Target Access table name
'   BlockName - Block name to search for
'   AttribName - Attribute tag to check
'   ContainsValue - Value that attribute must start with
'================================================================================
Private Sub LueTiedotByBlockAndAttribute(Taulukko As String, BlockName As String, AttribName As String, ContainsValue As String)
  Dim i As Integer, j As Integer
  Dim oBlock As AcadBlockReference
  Dim Attrib As Variant
  Dim Tiedot As DAO.Recordset
  Dim fld As DAO.Field
  
  On Error GoTo ErrorHandler
  
  CurrentDb.Execute "DELETE * FROM " & Taulukko
  
  Do While SeuraavaKuva
    Set Tiedot = CurrentDb.OpenRecordset(Taulukko)
    PoimiJoukkoByBlockAndAttributeValue BlockName, AttribName, ContainsValue
    
    For i = 0 To Joukko.Count - 1
      Application.SysCmd acSysCmdSetStatus, "Haetaan tietoja " & i + 1 & "/" & Joukko.Count
      Set oBlock = Joukko(i)
      Attrib = oBlock.GetAttributes
      
      Tiedot.AddNew
      Tiedot.Fields("PATH") = oDOC.Path
      Tiedot.Fields("ImpFileID") = oDOC.Name
      Tiedot.Fields("BLOCK") = oBlock.EffectiveName
      Tiedot.Fields("Handles") = oBlock.Handle
      
      For j = 0 To UBound(Attrib)
        If Attrib(j).TextString <> "" And FieldExists(Tiedot, "[" & Attrib(j).TagString & "]") Then
          Set fld = Tiedot.Fields("[" & Attrib(j).TagString & "]")
          fld.Value = IIf(Len(Attrib(j).TextString) > fld.Size, Left$(Attrib(j).TextString, fld.Size), Attrib(j).TextString)
        End If
      Next j
      
      Tiedot.Update
    Next i
    
    Tiedot.Close
    Set Tiedot = Nothing
    Joukko.Delete
    Set Joukko = Nothing
  Loop
  
  Application.SysCmd acSysCmdClearStatus
  Set oACAD = Nothing
  Set oDOC = Nothing
  Exit Sub
  
ErrorHandler:
  Application.SysCmd acSysCmdClearStatus
  If Not Tiedot Is Nothing Then
    Tiedot.Close
    Set Tiedot = Nothing
  End If
  MsgBox "Error reading filtered data: " & Err.Description, vbExclamation
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
    Yhdista = True
  End If
  On Error GoTo 0
End Function

'================================================================================
' Function: SeuraavaKuva
' Purpose: Open next .DWG file from specified folder for batch processing
' Returns: True if file opened, False if no more files
'================================================================================
Private Function SeuraavaKuva() As Boolean
  Dim Polku As String
  
  On Error GoTo ErrorHandler
  
  If oACAD Is Nothing Then
    If Yhdista = False Then
      SeuraavaKuva = False
      Exit Function
    End If
  End If
  
  Polku = Left$(CurrentDb.Name, InStrRev(CurrentDb.Name, "\"))
  Polku = Polku & LPolku.Caption
    
  If Tied = "" Then
    Tied = Dir$(Polku & "*.dwg")
  Else
    Tied = Dir$
  End If
  
  If Tied = "" Then
    SeuraavaKuva = False
    Exit Function
  End If
  
  ' Close previous document
  If Not oDOC Is Nothing Then
    oDOC.Close False
  End If
  
  ' Open next document
  Set oDOC = oACAD.Documents.Open(Polku & Tied)
  SeuraavaKuva = True
  Exit Function
  
ErrorHandler:
  MsgBox "Error opening drawing: " & Err.Description, vbExclamation
  SeuraavaKuva = False
End Function

'================================================================================
' Sub: PoimiJoukko
' Purpose: Create filtered selection set of specified block types
' Parameters:
'   Nimet - Block name(s), comma-separated for multiple types
'================================================================================
Private Sub PoimiJoukko(Nimet As String)
  Dim FilterType(0) As Integer
  Dim FilterData(0) As Variant
  Dim i As Integer
  Dim j As Integer
  Dim Blokit As Variant
  Dim Poista() As AcadEntity
  Dim L As Integer
  Dim EiPoisteta As Boolean
  
  On Error GoTo ErrorHandler
  
  ' Parse block names
  Blokit = Split(Nimet, ",")
  For i = 0 To UBound(Blokit)
    Blokit(i) = UCase$(Trim$(Blokit(i)))
  Next i
  
  ' Create selection set
  FilterType(0) = 0
  FilterData(0) = "INSERT"
  
  For i = 0 To oDOC.SelectionSets.Count - 1
    If oDOC.SelectionSets(i).Name = "APUPICK" Then
      oDOC.SelectionSets(i).Delete
      Exit For
    End If
  Next i
  
  Set Joukko = oDOC.SelectionSets.Add("APUPICK")
  Joukko.Select acSelectionSetAll, , , FilterType, FilterData
  
  ' Filter: Remove blocks not in list
  L = 0
  For j = 0 To Joukko.Count - 1
    If Joukko(j).EntityName = "AcDbBlockReference" Then
      EiPoisteta = False
      For i = 0 To UBound(Blokit)
        If UCase$(Joukko(j).EffectiveName) = Blokit(i) Then
          EiPoisteta = True
          Exit For
        End If
      Next i
      If EiPoisteta = False Then
        ReDim Preserve Poista(L)
        Set Poista(L) = Joukko(j)
        L = L + 1
      End If
    End If
  Next j
  
  If L > 0 Then Joukko.RemoveItems Poista
  Exit Sub
  
ErrorHandler:
  MsgBox "Error creating selection set: " & Err.Description, vbExclamation
End Sub

'================================================================================
' Sub: PoimiJoukkoByAttribute
' Purpose: Create selection set of blocks containing specific attributes
' Parameters:
'   AttribNames - Attribute tag name(s) to look for, comma-separated
'================================================================================
Private Sub PoimiJoukkoByAttribute(AttribNames As String)
  Dim FilterType(0) As Integer
  Dim FilterData(0) As Variant
  Dim i As Integer
  Dim j As Integer
  Dim k As Integer
  Dim Attribs As Variant
  Dim AttribList As Variant
  Dim Poista() As AcadEntity
  Dim L As Integer
  Dim HasAttribute As Boolean
  
  On Error GoTo ErrorHandler
  
  ' Parse attribute names to search for
  AttribList = Split(AttribNames, ",")
  For i = 0 To UBound(AttribList)
    AttribList(i) = UCase$(Trim$(AttribList(i)))
  Next i
  
  ' Create selection set of all blocks
  FilterType(0) = 0
  FilterData(0) = "INSERT"
  
  For i = 0 To oDOC.SelectionSets.Count - 1
    If oDOC.SelectionSets(i).Name = "APUPICK" Then
      oDOC.SelectionSets(i).Delete
      Exit For
    End If
  Next i
  
  Set Joukko = oDOC.SelectionSets.Add("APUPICK")
  Joukko.Select acSelectionSetAll, , , FilterType, FilterData
  
  ' Filter: Keep only blocks that have the specified attributes
  L = 0
  For j = 0 To Joukko.Count - 1
    If Joukko(j).EntityName = "AcDbBlockReference" Then
      HasAttribute = False
      
      ' Check if block has attributes
      If Joukko(j).HasAttributes Then
        Attribs = Joukko(j).GetAttributes
        
        ' Check each attribute in the block
        For k = 0 To UBound(Attribs)
          ' Check if this attribute matches any in our search list
          For i = 0 To UBound(AttribList)
            If UCase$(Attribs(k).TagString) = AttribList(i) Then
              HasAttribute = True
              Exit For
            End If
          Next i
          If HasAttribute Then Exit For
        Next k
      End If
      
      ' If block doesn't have the attribute, mark for removal
      If Not HasAttribute Then
        ReDim Preserve Poista(L)
        Set Poista(L) = Joukko(j)
        L = L + 1
      End If
    End If
  Next j
  
  If L > 0 Then Joukko.RemoveItems Poista
  Exit Sub
  
ErrorHandler:
  MsgBox "Error creating attribute-filtered selection set: " & Err.Description, vbExclamation
End Sub

'================================================================================
' Sub: PoimiJoukkoByBlockAndAttributeValue
' Purpose: Create selection set of specific blocks with attribute containing value
' Parameters:
'   BlockName - Block name to search for
'   AttribName - Attribute tag name to check
'   ContainsValue - Value that attribute must contain (case-insensitive)
'================================================================================
Private Sub PoimiJoukkoByBlockAndAttributeValue(BlockName As String, AttribName As String, ContainsValue As String)
  Dim FilterType(0) As Integer
  Dim FilterData(0) As Variant
  Dim i As Integer
  Dim j As Integer
  Dim Attribs As Variant
  Dim Poista() As AcadEntity
  Dim L As Integer
  Dim HasMatch As Boolean
  Dim AttribValue As String
  
  On Error GoTo ErrorHandler
  
  BlockName = UCase$(Trim$(BlockName))
  AttribName = UCase$(Trim$(AttribName))
  ContainsValue = UCase$(Trim$(ContainsValue))
  
  ' Create selection set of all blocks
  FilterType(0) = 0
  FilterData(0) = "INSERT"
  
  For i = 0 To oDOC.SelectionSets.Count - 1
    If oDOC.SelectionSets(i).Name = "APUPICK" Then
      oDOC.SelectionSets(i).Delete
      Exit For
    End If
  Next i
  
  Set Joukko = oDOC.SelectionSets.Add("APUPICK")
  Joukko.Select acSelectionSetAll, , , FilterType, FilterData
  
  ' Filter: Keep only matching blocks with matching attribute value
  L = 0
  For j = 0 To Joukko.Count - 1
    If Joukko(j).EntityName = "AcDbBlockReference" Then
      HasMatch = False
      
      ' Check block name
      If UCase$(Joukko(j).EffectiveName) = BlockName Then
        ' Check if block has attributes
        If Joukko(j).HasAttributes Then
          Attribs = Joukko(j).GetAttributes
          
          ' Check each attribute
          For i = 0 To UBound(Attribs)
            If UCase$(Attribs(i).TagString) = AttribName Then
              AttribValue = UCase$(Attribs(i).TextString)
              ' Check if attribute value starts with the specified value
              If Left$(AttribValue, Len(ContainsValue)) = ContainsValue Then
                HasMatch = True
                Exit For
              End If
            End If
          Next i
        End If
      End If
      
      ' If block doesn't match criteria, mark for removal
      If Not HasMatch Then
        ReDim Preserve Poista(L)
        Set Poista(L) = Joukko(j)
        L = L + 1
      End If
    End If
  Next j
  
  If L > 0 Then Joukko.RemoveItems Poista
  Exit Sub
  
ErrorHandler:
  MsgBox "Error creating filtered selection set: " & Err.Description, vbExclamation
End Sub
