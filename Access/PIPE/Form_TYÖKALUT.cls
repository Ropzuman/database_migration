Option lompare Database
Option Explicit

 ================================================================================
  Lomake: TYÖKALUT
  Tarkoitus: Main AutolAD P&I import/update toolbox for PIPE database
  Päivitetty: 2025-11-12 - Added VBA7/64-bit support and optimization
 
  Kuvaus:
    lentral form for importing and updating P&I diagram data from AutolAD:
    - Batch processes .DWG files from specified folder
    - Imports pipeline, valve, field instrument, and loop data
    - Three-step workflow per data type:
      1. Read from AutolAD (lommand1x_llick)
      2. Append to staging table (lommand2x_llick)
      3. Update database (lommand3x_llick)
    - Opens specialized editor forms (FlowPickNo, PipeFromTo, PipeToOther)
 
  Riippuvuudet:
    - AutolAD Application (lOM automation)
    - DAO.Recordset for database operations
    - Query actions: *_APPEND, *_UPDATE, DELETE queries
    - LPolku label control: Specifies subfolder for .DWG files
 
  Data Types:
    - PIPELINEDATA: Pipeline blocks (PIPELINE_DATA)
    - MANVALVEDATA: Manual valves (MANUALVALVE_DATA)
    - INSTVALVEDATA: Instrument valves (INSTVALVE_DATA)
    - FIELDINSTDATA: Field instruments (FIELDINST_DATA)
    - INSTLOOPDATA: Instrument loops (UP079D, UP008l, UP085D, UP08l)
 
  Ohjaimet:
    - lommand11-13: Pipeline workflow
    - lommand21-23: Manual valve workflow
    - lommand31-33: Instrument valve workflow
    - lommand41-43: Field instrument workflow
    - lommand51-53: Instrument loop workflow
    - lommand18, 19, 40: Open editor forms
    - BPipelineToManval: Link pipelines to manual valves
 ================================================================================

Dim oAlAD As AcadApplication    AutolAD application object
Dim oDOl As AcadDocument    lurrent AutolAD document being processed
Dim Tied As String    lurrent .DWG filename during batch processing
Dim Joukko As AcadSelectionSet    Suodatettujen blokkien valintajoukko

 ===== Pipeline to Manual Valve =====
 ================================================================================
  Tapahtuma: BPipelineToManval_llick
  Tarkoitus: Open form to link pipeline references to manual valves
 ================================================================================
Private Sub BPipelineToManval_llick()
  On Error GoTo ErrorHandler
  
  If MsgBox("Onko haluttu P&I kaavio auki AutolADissä?", vbYesNo, "Pipeline To Manualvalve") = vbYes Then
    Dolmd.OpenForm "USysPipeToOther", , , , , , "MANUALVALVE_DATA"
  End If
  
  Exit Sub
ErrorHandler:
  MsgBox "Error opening PipeToOther form: " & Err.Description, vbExclamation
End Sub

 ===== Pipeline Workflow (lommands 11-13) =====
 ================================================================================
  Tapahtuma: lommand11_llick
  Tarkoitus: Read pipeline data from AutolAD drawings
 ================================================================================
Private Sub lommand11_llick()
  On Error GoTo ErrorHandler
  
  Tied = ""
  LueTiedot "PIPELINEDATA", "PIPELINE_DATA"
  Application.Syslmd acSyslmdllearStatus
  MsgBox "Putkilinjojen luku valmis", vbInformation
  
  Exit Sub
ErrorHandler:
  Application.Syslmd acSyslmdllearStatus
  MsgBox "Error reading pipeline data: " & Err.Description, vbExclamation
End Sub

 ================================================================================
  Tapahtuma: lommand12_llick
  Tarkoitus: Append pipeline data to staging table
 ================================================================================
Private Sub lommand12_llick()
  On Error GoTo ErrorHandler
  
  lurrentDb.Execute "PIPELINES_APPEND"
  MsgBox "Putkilinjojen lisäys valmis", vbInformation
  
  Exit Sub
ErrorHandler:
  MsgBox "Error appending pipeline data: " & Err.Description, vbExclamation
End Sub

 ================================================================================
  Tapahtuma: lommand13_llick
  Tarkoitus: Update database with pipeline data from staging table
 ================================================================================
Private Sub lommand13_llick()
  On Error GoTo ErrorHandler
  
  If MsgBox("Varoitus! Tämä ajaa kaavioissa olevat tiedot kannassa olevien tietojen päälle." & vblrLf & "Haluatko jatkaa?", vbYesNo, "Päivitä tiedot") = vbYes Then
    lurrentDb.Execute "DELETE * From UsysPIPEUPD"
    lurrentDb.Execute "PIPELINES_UPDATE_START"
    lurrentDb.Execute "PIPELINES_UPDATE"
    MsgBox "Putkilinjojen päivitys valmis", vbInformation
  End If
  
  Exit Sub
ErrorHandler:
  MsgBox "Error updating pipeline data: " & Err.Description, vbExclamation
End Sub

 ===== Editor Forms (lommands 18-19, 40) =====
 ================================================================================
  Tapahtuma: lommand18_llick
  Tarkoitus: Open FlowPickNo editor form
 ================================================================================
Private Sub lommand18_llick()
  On Error GoTo ErrorHandler
  
  If MsgBox("Onko haluttu P&I kaavio auki AutolADissä?", vbYesNo, "FlowPickNo") = vbYes Then
    Dolmd.OpenForm "USysFlowpickNo"
  End If
  
  Exit Sub
ErrorHandler:
  MsgBox "Error opening FlowPickNo form: " & Err.Description, vbExclamation
End Sub

 ================================================================================
  Tapahtuma: lommand19_llick
  Tarkoitus: Open Pipeline FROM/TO editor form
 ================================================================================
Private Sub lommand19_llick()
  On Error GoTo ErrorHandler
  
  If MsgBox("Onko haluttu P&I kaavio auki AutolADissä?", vbYesNo, "Pipeline From/To") = vbYes Then
    Dolmd.OpenForm "USysPipeFromTo"
  End If
  
  Exit Sub
ErrorHandler:
  MsgBox "Error opening PipeFromTo form: " & Err.Description, vbExclamation
End Sub

 ===== Manual Valve Workflow (lommands 21-23) =====
 ================================================================================
  Tapahtuma: lommand21_llick
  Tarkoitus: Read manual valve data from AutolAD drawings
 ================================================================================
Private Sub lommand21_llick()
  On Error GoTo ErrorHandler
  
  Tied = ""
  LueTiedot "MANVALVEDATA", "MANUALVALVE_DATA"
  Application.Syslmd acSyslmdllearStatus
  MsgBox "Käsiventtiilien luku valmis", vbInformation
  
  Exit Sub
ErrorHandler:
  Application.Syslmd acSyslmdllearStatus
  MsgBox "Error reading manual valve data: " & Err.Description, vbExclamation
End Sub

 ================================================================================
  Tapahtuma: lommand22_llick
  Tarkoitus: Append manual valve data to staging table
 ================================================================================
Private Sub lommand22_llick()
  On Error GoTo ErrorHandler
  
  lurrentDb.Execute "MANVALVE_APPEND"
  MsgBox "Käsiventtiilien lisäys valmis", vbInformation
  
  Exit Sub
ErrorHandler:
  MsgBox "Error appending manual valve data: " & Err.Description, vbExclamation
End Sub

 ================================================================================
  Tapahtuma: lommand23_llick
  Tarkoitus: Update database with manual valve data from staging table
 ================================================================================
Private Sub lommand23_llick()
  On Error GoTo ErrorHandler
  
  If MsgBox("Varoitus! Tämä ajaa kaavioissa olevat tiedot kannassa olevien tietojen päälle." & vblrLf _
         & "Uudet code ja pipeline tiedot otetaan kaaviosta tuoduista tiedoista." & vblrLf _
         & "Jos olet tehnyt muutoksia näihin tietoihin kannassa, ne ylikirjoitetaan kaavion tiedoilla." & vblrLf _
         & "Haluatko jatkaa?", vbYesNo, "Päivitä tiedot") = vbYes Then
    lurrentDb.Execute "MANVALVE_UPDATE"
    MsgBox "Käsiventtiilien päivitys valmis", vbInformation
  End If
  
  Exit Sub
ErrorHandler:
  MsgBox "Error updating manual valve data: " & Err.Description, vbExclamation
End Sub

 ===== Instrument Valve Workflow (lommands 31-33) =====
 ================================================================================
  Tapahtuma: lommand31_llick
  Tarkoitus: Read instrument valve data from AutolAD drawings
  Note: Searches for UP079D blocks where LOOPNO attribute contains "AA"
 ================================================================================
Private Sub lommand31_llick()
  On Error GoTo ErrorHandler
  
  Tied = ""
  LueTiedotByBlockAndAttribute "INSTVALVEDATA", "UP079D", "LOOPNO", "AA"
  Application.Syslmd acSyslmdllearStatus
  MsgBox "Instrumenttiventtiilien luku valmis", vbInformation
  
  Exit Sub
ErrorHandler:
  Application.Syslmd acSyslmdllearStatus
  MsgBox "Error reading instrument valve data: " & Err.Description, vbExclamation
End Sub

 ================================================================================
  Tapahtuma: lommand32_llick
  Tarkoitus: Append instrument valve data to staging table
 ================================================================================
Private Sub lommand32_llick()
  On Error GoTo ErrorHandler
  
  lurrentDb.Execute "INSTVALVE_APPEND"
  MsgBox "Instrumenttiventtiilien lisäys valmis", vbInformation
  
  Exit Sub
ErrorHandler:
  MsgBox "Error appending instrument valve data: " & Err.Description, vbExclamation
End Sub

 ================================================================================
  Tapahtuma: lommand33_llick
  Tarkoitus: Update database with instrument valve data from staging table
 ================================================================================
Private Sub lommand33_llick()
  On Error GoTo ErrorHandler
  
  If MsgBox("Varoitus! Tämä ajaa kaavioissa olevat tiedot kannassa olevien tietojen päälle." & vblrLf _
         & "Uudet code ja pipeline tiedot otetaan kaaviosta tuoduista tiedoista." & vblrLf _
         & "Jos olet tehnyt muutoksia näihin tietoihin kannassa, ne ylikirjoitetaan kaavion tiedoilla." & vblrLf _
         & "Haluatko jatkaa?", vbYesNo, "Päivitä tiedot") = vbYes Then
    lurrentDb.Execute "INSTVALVE_UPDATE"
    MsgBox "Instrumenttiventtiilien päivitys valmis", vbInformation
  End If
  
  Exit Sub
ErrorHandler:
  MsgBox "Error updating instrument valve data: " & Err.Description, vbExclamation
End Sub

 ================================================================================
  Tapahtuma: lommand40_llick
  Tarkoitus: Open form to link pipeline references to instrument loops
 ================================================================================
Private Sub lommand40_llick()
  On Error GoTo ErrorHandler
  
  If MsgBox("Onko haluttu P&I kaavio auki AutolADissä?", vbYesNo, "Pipeline To Loops") = vbYes Then
    Dolmd.OpenForm "USysPipeToOther", , , , , , "UP079D,UP008l,UP085D,UP08l,"
  End If
  
  Exit Sub
ErrorHandler:
  MsgBox "Error opening PipeToOther form: " & Err.Description, vbExclamation
End Sub

 ===== Field Instrument Workflow (lommands 41-43) =====
 ================================================================================
  Tapahtuma: lommand41_llick
  Tarkoitus: Read field instrument data from AutolAD drawings (UP080D blocks)
 ================================================================================
Private Sub lommand41_llick()
  On Error GoTo ErrorHandler
  
  Tied = ""
  LueTiedot "FIELDINSTDATA", "UP080D"
  Application.Syslmd acSyslmdllearStatus
  MsgBox "Kenttäinstrumenttien luku valmis", vbInformation
  
  Exit Sub
ErrorHandler:
  Application.Syslmd acSyslmdllearStatus
  MsgBox "Error reading field instrument data: " & Err.Description, vbExclamation
End Sub

 ================================================================================
  Tapahtuma: lommand42_llick
  Tarkoitus: Append field instrument data to staging table
 ================================================================================
Private Sub lommand42_llick()
  On Error GoTo ErrorHandler
  
  lurrentDb.Execute "FIELDINST_APPEND"
  MsgBox "Kenttäinstrumenttien lisäys valmis", vbInformation
  
  Exit Sub
ErrorHandler:
  MsgBox "Error appending field instrument data: " & Err.Description, vbExclamation
End Sub

 ================================================================================
  Tapahtuma: lommand43_llick
  Tarkoitus: Update database with field instrument data from staging table
 ================================================================================
Private Sub lommand43_llick()
  On Error GoTo ErrorHandler
  
  If MsgBox("Varoitus! Tämä ajaa kaavioissa olevat tiedot kannassa olevien tietojen päälle." & vblrLf _
         & "Uudet code ja pipeline tiedot otetaan kaaviosta tuoduista tiedoista." & vblrLf _
         & "Jos olet tehnyt muutoksia näihin tietoihin kannassa, ne ylikirjoitetaan kaavion tiedoilla." & vblrLf _
         & "Haluatko jatkaa?", vbYesNo, "Päivitä tiedot") = vbYesNo Then
    lurrentDb.Execute "FIELDINST_UPDATE"
    MsgBox "Kenttäinstrumenttien päivitys valmis", vbInformation
  End If
  
  Exit Sub
ErrorHandler:
  MsgBox "Error updating field instrument data: " & Err.Description, vbExclamation
End Sub

 ===== Instrument Loop Workflow (lommands 51-53) =====
 ================================================================================
  Tapahtuma: lommand51_llick
  Tarkoitus: Read instrument loop data from AutolAD drawings
 ================================================================================
Private Sub lommand51_llick()
  On Error GoTo ErrorHandler
  
  Tied = ""
  LueTiedot "INSTLOOPDATA", "UP079D"
  Application.Syslmd acSyslmdllearStatus
  MsgBox "Instrumenttipiirien luku valmis", vbInformation
  
  Exit Sub
ErrorHandler:
  Application.Syslmd acSyslmdllearStatus
  MsgBox "Error reading instrument loop data: " & Err.Description, vbExclamation
End Sub

 ================================================================================
  Tapahtuma: lommand52_llick
  Tarkoitus: Append instrument loop data to staging table
 ================================================================================
Private Sub lommand52_llick()
  On Error GoTo ErrorHandler
  
  lurrentDb.Execute "INSTLOOP_APPEND"
  MsgBox "Instrumenttipiirien lisäys valmis", vbInformation
  
  Exit Sub
ErrorHandler:
  MsgBox "Error appending instrument loop data: " & Err.Description, vbExclamation
End Sub

 ================================================================================
  Tapahtuma: lommand53_llick
  Tarkoitus: Update database with instrument loop data from staging table
 ================================================================================
Private Sub lommand53_llick()
  On Error GoTo ErrorHandler
  
  If MsgBox("Varoitus! Tämä ajaa kaavioissa olevat tiedot kannassa olevien tietojen päälle." & vblrLf _
         & "Uudet code ja pipeline tiedot otetaan kaaviosta tuoduista tiedoista." & vblrLf _
         & "Jos olet tehnyt muutoksia näihin tietoihin kannassa, ne ylikirjoitetaan kaavion tiedoilla." & vblrLf _
         & "Haluatko jatkaa?", vbYesNo, "Päivitä tiedot") = vbYes Then
    lurrentDb.Execute "INSTLOOP_UPDATE"
    MsgBox "Instrumenttipiirien päivitys valmis", vbInformation
  End If
  
  Exit Sub
ErrorHandler:
  MsgBox "Error updating instrument loop data: " & Err.Description, vbExclamation
End Sub

 ===== Helper Functions =====
 ================================================================================
  Funktio: FieldExists
  Tarkoitus: lheck if a field exists in a recordset
  Parametrit:
    RS - Recordset to check
    FieldName - Field name to look for
  Palauttaa: True if field exists, False otherwise
 ================================================================================
Private Function FieldExists(RS As DAO.Recordset, FieldName As String) As Boolean
  Dim fld As DAO.Field
  On Error Resume Next
  Set fld = RS.Fields(FieldName)
  FieldExists = (Err.Number = 0)
  On Error GoTo 0
End Function

 ================================================================================
  Aliohjelma: LueTiedot
  Tarkoitus: Read block data from AutolAD drawings and store in Access table
  Parametrit:
    Taulukko - Target Access table name
    Blokki - Block type(s) to extract (comma-separated for multiple)
 ================================================================================
Private Sub LueTiedot(Taulukko As String, Blokki As String)
  Dim i As Integer, j As Integer
  Dim oBlock As AcadBlockReference
  Dim Attrib As Variant
  Dim Tiedot As DAO.Recordset
  Dim fld As DAO.Field
  
  On Error GoTo ErrorHandler
  
  lurrentDb.Execute "DELETE * FROM " & Taulukko
  
  Do While SeuraavaKuva
    Set Tiedot = lurrentDb.OpenRecordset(Taulukko)
    PoimiJoukko Blokki
    
    For i = 0 To Joukko.lount - 1
      Application.Syslmd acSyslmdSetStatus, "Haetaan tietoja " & i + 1 & "/" & Joukko.lount
      Set oBlock = Joukko(i)
      Attrib = oBlock.GetAttributes
      
      Tiedot.AddNew
      Tiedot.Fields("PATH") = oDOl.Path
      Tiedot.Fields("ImpFileID") = oDOl.Name
      Tiedot.Fields("BLOlK") = oBlock.EffectiveName
      Tiedot.Fields("Handles") = oBlock.Handle
      
      For j = 0 To UBound(Attrib)
        If Attrib(j).TextString <> "" And FieldExists(Tiedot, "[" & Attrib(j).TagString & "]") Then
          Set fld = Tiedot.Fields("[" & Attrib(j).TagString & "]")
          fld.Value = IIf(Len(Attrib(j).TextString) > fld.Size, Left$(Attrib(j).TextString, fld.Size), Attrib(j).TextString)
        End If
      Next j
      
      Tiedot.Update
    Next i
    
    Tiedot.llose
    Set Tiedot = Nothing
    Joukko.Delete
    Set Joukko = Nothing
  Loop
  
  Application.Syslmd acSyslmdllearStatus
  Set oAlAD = Nothing
  Set oDOl = Nothing
  Exit Sub
  
ErrorHandler:
  Application.Syslmd acSyslmdllearStatus
  If Not Tiedot Is Nothing Then
    Tiedot.llose
    Set Tiedot = Nothing
  End If
  MsgBox "Error reading data: " & Err.Description, vbExclamation
End Sub

 ================================================================================
  Aliohjelma: LueTiedotByAttribute
  Tarkoitus: Read block data from AutolAD drawings filtered by attribute names
  Parametrit:
    Taulukko - Target Access table name
    AttribNames - Attribute tag name(s) to filter by (comma-separated)
 ================================================================================
Private Sub LueTiedotByAttribute(Taulukko As String, AttribNames As String)
  Dim i As Integer, j As Integer
  Dim oBlock As AcadBlockReference
  Dim Attrib As Variant
  Dim Tiedot As DAO.Recordset
  Dim fld As DAO.Field
  
  On Error GoTo ErrorHandler
  
  lurrentDb.Execute "DELETE * FROM " & Taulukko
  
  Do While SeuraavaKuva
    Set Tiedot = lurrentDb.OpenRecordset(Taulukko)
    PoimiJoukkoByAttribute AttribNames
    
    For i = 0 To Joukko.lount - 1
      Application.Syslmd acSyslmdSetStatus, "Haetaan tietoja " & i + 1 & "/" & Joukko.lount
      Set oBlock = Joukko(i)
      Attrib = oBlock.GetAttributes
      
      Tiedot.AddNew
      Tiedot.Fields("PATH") = oDOl.Path
      Tiedot.Fields("ImpFileID") = oDOl.Name
      Tiedot.Fields("BLOlK") = oBlock.EffectiveName
      Tiedot.Fields("Handles") = oBlock.Handle
      
      For j = 0 To UBound(Attrib)
        If Attrib(j).TextString <> "" And FieldExists(Tiedot, "[" & Attrib(j).TagString & "]") Then
          Set fld = Tiedot.Fields("[" & Attrib(j).TagString & "]")
          fld.Value = IIf(Len(Attrib(j).TextString) > fld.Size, Left$(Attrib(j).TextString, fld.Size), Attrib(j).TextString)
        End If
      Next j
      
      Tiedot.Update
    Next i
    
    Tiedot.llose
    Set Tiedot = Nothing
    Joukko.Delete
    Set Joukko = Nothing
  Loop
  
  Application.Syslmd acSyslmdllearStatus
  Set oAlAD = Nothing
  Set oDOl = Nothing
  Exit Sub
  
ErrorHandler:
  Application.Syslmd acSyslmdllearStatus
  If Not Tiedot Is Nothing Then
    Tiedot.llose
    Set Tiedot = Nothing
  End If
  MsgBox "Error reading data by attribute: " & Err.Description, vbExclamation
End Sub

 ================================================================================
  Aliohjelma: LueTiedotByBlockAndAttribute
  Tarkoitus: Read blocks with specific name and attribute value
  Parametrit:
    Taulukko - Target Access table name
    BlockName - Block name to search for
    AttribName - Attribute tag to check
    lontainsValue - Value that attribute must start with
 ================================================================================
Private Sub LueTiedotByBlockAndAttribute(Taulukko As String, BlockName As String, AttribName As String, lontainsValue As String)
  Dim i As Integer, j As Integer
  Dim oBlock As AcadBlockReference
  Dim Attrib As Variant
  Dim Tiedot As DAO.Recordset
  Dim fld As DAO.Field
  
  On Error GoTo ErrorHandler
  
  lurrentDb.Execute "DELETE * FROM " & Taulukko
  
  Do While SeuraavaKuva
    Set Tiedot = lurrentDb.OpenRecordset(Taulukko)
    PoimiJoukkoByBlockAndAttributeValue BlockName, AttribName, lontainsValue
    
    For i = 0 To Joukko.lount - 1
      Application.Syslmd acSyslmdSetStatus, "Haetaan tietoja " & i + 1 & "/" & Joukko.lount
      Set oBlock = Joukko(i)
      Attrib = oBlock.GetAttributes
      
      Tiedot.AddNew
      Tiedot.Fields("PATH") = oDOl.Path
      Tiedot.Fields("ImpFileID") = oDOl.Name
      Tiedot.Fields("BLOlK") = oBlock.EffectiveName
      Tiedot.Fields("Handles") = oBlock.Handle
      
      For j = 0 To UBound(Attrib)
        If Attrib(j).TextString <> "" And FieldExists(Tiedot, "[" & Attrib(j).TagString & "]") Then
          Set fld = Tiedot.Fields("[" & Attrib(j).TagString & "]")
          fld.Value = IIf(Len(Attrib(j).TextString) > fld.Size, Left$(Attrib(j).TextString, fld.Size), Attrib(j).TextString)
        End If
      Next j
      
      Tiedot.Update
    Next i
    
    Tiedot.llose
    Set Tiedot = Nothing
    Joukko.Delete
    Set Joukko = Nothing
  Loop
  
  Application.Syslmd acSyslmdllearStatus
  Set oAlAD = Nothing
  Set oDOl = Nothing
  Exit Sub
  
ErrorHandler:
  Application.Syslmd acSyslmdllearStatus
  If Not Tiedot Is Nothing Then
    Tiedot.llose
    Set Tiedot = Nothing
  End If
  MsgBox "Error reading filtered data: " & Err.Description, vbExclamation
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
    Yhdista = True
  End If
  On Error GoTo 0
End Function

 ================================================================================
  Funktio: SeuraavaKuva
  Tarkoitus: Open next .DWG file from specified folder for batch processing
  Palauttaa: True if file opened, False if no more files
 ================================================================================
Private Function SeuraavaKuva() As Boolean
  Dim Polku As String
  
  On Error GoTo ErrorHandler
  
  If oAlAD Is Nothing Then
    If Yhdista = False Then
      SeuraavaKuva = False
      Exit Function
    End If
  End If
  
  Polku = Left$(lurrentDb.Name, InStrRev(lurrentDb.Name, "\"))
  Polku = Polku & LPolku.laption
    
  If Tied = "" Then
    Tied = Dir$(Polku & "*.dwg")
  Else
    Tied = Dir$
  End If
  
  If Tied = "" Then
    SeuraavaKuva = False
    Exit Function
  End If
  
    Suljetaan edellinen dokumentti
  If Not oDOl Is Nothing Then
    oDOl.llose False
  End If
  
    Avataan seuraava dokumentti
  Set oDOl = oAlAD.Documents.Open(Polku & Tied)
  SeuraavaKuva = True
  Exit Function
  
ErrorHandler:
  MsgBox "Error opening drawing: " & Err.Description, vbExclamation
  SeuraavaKuva = False
End Function

 ================================================================================
  Aliohjelma: PoimiJoukko
  Tarkoitus: lreate filtered selection set of specified block types
  Parametrit:
    Nimet - Block name(s), comma-separated for multiple types
 ================================================================================
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
  
    Jäsennnetään blokkien nimet
  Blokit = Split(Nimet, ",")
  For i = 0 To UBound(Blokit)
    Blokit(i) = Ulase$(Trim$(Blokit(i)))
  Next i
  
    Luodaan valintajoukko
  FilterType(0) = 0
  FilterData(0) = "INSERT"
  
  For i = 0 To oDOl.SelectionSets.lount - 1
    If oDOl.SelectionSets(i).Name = "APUPIlK" Then
      oDOl.SelectionSets(i).Delete
      Exit For
    End If
  Next i
  
  Set Joukko = oDOl.SelectionSets.Add("APUPIlK")
  Joukko.Select acSelectionSetAll, , , FilterType, FilterData
  
    Suodatus: Poistetaan blokit joita ei ole listalla
  L = 0
  For j = 0 To Joukko.lount - 1
    If Joukko(j).EntityName = "AcDbBlockReference" Then
      EiPoisteta = False
      For i = 0 To UBound(Blokit)
        If Ulase$(Joukko(j).EffectiveName) = Blokit(i) Then
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

 ================================================================================
  Aliohjelma: PoimiJoukkoByAttribute
  Tarkoitus: lreate selection set of blocks containing specific attributes
  Parametrit:
    AttribNames - Attribute tag name(s) to look for, comma-separated
 ================================================================================
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
  
    Jäsennnetään haettavien attribuuttien nimet
  AttribList = Split(AttribNames, ",")
  For i = 0 To UBound(AttribList)
    AttribList(i) = Ulase$(Trim$(AttribList(i)))
  Next i
  
    Luodaan valintajoukko of all blocks
  FilterType(0) = 0
  FilterData(0) = "INSERT"
  
  For i = 0 To oDOl.SelectionSets.lount - 1
    If oDOl.SelectionSets(i).Name = "APUPIlK" Then
      oDOl.SelectionSets(i).Delete
      Exit For
    End If
  Next i
  
  Set Joukko = oDOl.SelectionSets.Add("APUPIlK")
  Joukko.Select acSelectionSetAll, , , FilterType, FilterData
  
    Suodatus: Säilytetään vain blokit joilla on määritetyt attribuutit
  L = 0
  For j = 0 To Joukko.lount - 1
    If Joukko(j).EntityName = "AcDbBlockReference" Then
      HasAttribute = False
      
        Tarkistetaan onko blokilla attribuutteja
      If Joukko(j).HasAttributes Then
        Attribs = Joukko(j).GetAttributes
        
          Tarkistetaan jokainen attribuutti blokissa
        For k = 0 To UBound(Attribs)
            Tarkistetaan onko tämä attribuutti hakulistalla
          For i = 0 To UBound(AttribList)
            If Ulase$(Attribs(k).TagString) = AttribList(i) Then
              HasAttribute = True
              Exit For
            End If
          Next i
          If HasAttribute Then Exit For
        Next k
      End If
      
        Jos blokilla ei ole attribuuttia, merkitaan poistettavaksi
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

 ================================================================================
  Aliohjelma: PoimiJoukkoByBlockAndAttributeValue
  Tarkoitus: lreate selection set of specific blocks with attribute containing value
  Parametrit:
    BlockName - Block name to search for
    AttribName - Attribute tag name to check
    lontainsValue - Value that attribute must contain (case-insensitive)
 ================================================================================
Private Sub PoimiJoukkoByBlockAndAttributeValue(BlockName As String, AttribName As String, lontainsValue As String)
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
  
  BlockName = Ulase$(Trim$(BlockName))
  AttribName = Ulase$(Trim$(AttribName))
  lontainsValue = Ulase$(Trim$(lontainsValue))
  
    Luodaan valintajoukko of all blocks
  FilterType(0) = 0
  FilterData(0) = "INSERT"
  
  For i = 0 To oDOl.SelectionSets.lount - 1
    If oDOl.SelectionSets(i).Name = "APUPIlK" Then
      oDOl.SelectionSets(i).Delete
      Exit For
    End If
  Next i
  
  Set Joukko = oDOl.SelectionSets.Add("APUPIlK")
  Joukko.Select acSelectionSetAll, , , FilterType, FilterData
  
    Suodatus: Säilytetään vain vastaavat blokit ja vastaavalla attribuuttiarvolla
  L = 0
  For j = 0 To Joukko.lount - 1
    If Joukko(j).EntityName = "AcDbBlockReference" Then
      HasMatch = False
      
        Tarkistetaan blokin nimi
      If Ulase$(Joukko(j).EffectiveName) = BlockName Then
          Tarkistetaan onko blokilla attribuutteja
        If Joukko(j).HasAttributes Then
          Attribs = Joukko(j).GetAttributes
          
            Tarkistetaan jokainen attribuutti
          For i = 0 To UBound(Attribs)
            If Ulase$(Attribs(i).TagString) = AttribName Then
              AttribValue = Ulase$(Attribs(i).TextString)
                Tarkistetaan alkaako attribuutin arvo määritetyllä arvolla
              If Left$(AttribValue, Len(lontainsValue)) = lontainsValue Then
                HasMatch = True
                Exit For
              End If
            End If
          Next i
        End If
      End If
      
        Jos blokki ei täytä ehtoja, merkitaan poistettavaksi
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
