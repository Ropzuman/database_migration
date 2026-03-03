Option Compare Database
Option Explicit

'================================================================================
' Module: Koodit
' Purpose: Core utility functions and AutoCAD integration for PIPE database
' Updated: 2025-11-12 - Added VBA7/64-bit support
'
' Description:
'   Provides essential functionality for the PIPE database:
'   - AutoCAD document integration (block highlighting, zooming)
'   - User login tracking (network username, computer name)
'   - Block opening from database records (valves, pipelines)
'   - String parsing utilities
'
' Dependencies:
'   - AutoCAD Application (COM automation)
'   - UsysUsers table (login tracking)
'   - MANUALVALVES, PIPELINES, PIPELINEDATA, MANVALVEDATA tables
'   - advapi32.dll (GetUserName API)
'   - kernel32.dll (GetComputerName API)
'================================================================================

'--------------------------------------------------------------------------------
' Windows API Declarations
' Updated 2025-11-12: Added VBA7/64-bit support
'--------------------------------------------------------------------------------
#If VBA7 Then
    Private Declare PtrSafe Function api_GetUserName Lib "advapi32.dll" Alias "GetUserNameA" (ByVal lpBuffer As String, nSize As Long) As Long
    Private Declare PtrSafe Function api_GetComputerName Lib "kernel32" Alias "GetComputerNameA" (ByVal lpBuffer As String, nSize As Long) As Long
#Else
    Private Declare Function api_GetUserName Lib "advapi32.dll" Alias "GetUserNameA" (ByVal lpBuffer As String, nSize As Long) As Long
    Private Declare Function api_GetComputerName Lib "kernel32" Alias "GetComputerNameA" (ByVal lpBuffer As String, nSize As Long) As Long
#End If

'--------------------------------------------------------------------------------
' Sub: AvaaBlock
' Purpose: Open and highlight AutoCAD block from active database table
'
' Process:
'   1. Determine if called from MANUALVALVES or PIPELINES table
'   2. Query database for block information (path, drawing, handle)
'   3. Call AvaaKuvasta to open drawing and zoom to block
'
' Notes:
'   - Works with MANUALVALVES and PIPELINES datasheets
'   - For PIPELINES, delegates to frmOpenPIPELINE form
'   - Requires AutoCAD to be running
'   - Finnish: "Etsi kohde" = "Find object"
'--------------------------------------------------------------------------------
Sub AvaaBlock()
On Error GoTo ErrorHandler
    Dim DWG As String  ' Drawing filename
    Dim Polku As String  ' Path to drawings folder
    Dim Handle As String  ' AutoCAD block handle
    Dim Info As String  ' Block identification string
    Dim i As Integer  ' Loop counter
    Dim Doku As String  ' Document name (unused)
    Dim Tyyppi As Integer  ' Type: 1=ManualValve, 0=Pipeline
    Dim Taulu As DAO.Recordset  ' Database query results
    Dim Alue As String  ' Area code
    Dim Linja As String  ' Line number

    ' Get relative path to flowsheets folder
    Polku = Left$(CurrentDb.Name, Len(CurrentDb.Name) - Len(Dir(CurrentDb.Name))) & "..\..\R\FlowSheets\"

    ' Determine source table
    If UCase$(Application.CurrentObjectName) = "MANUALVALVES" Or UCase$(Application.CurrentObjectName) = "PIPELINES" Then
      If UCase$(Application.CurrentObjectName) = "MANUALVALVES" Then Tyyppi = 1
      If Tyyppi = 1 Then
        ' Get manual valve block information
        Set Taulu = CurrentDb.OpenRecordset("SELECT * FROM MANVALVEDATA WHERE AREACODE = '" & Screen.ActiveDatasheet("Area").Value & "' AND VAL_NO = '" & Screen.ActiveDatasheet("ValveNo").Value & "'")
        DWG = LCase$(Taulu.Fields("ImpFileID"))
        Handle = Taulu.Fields("Handles")
        Info = Taulu.Fields("AREACODE") & "-" & Taulu.Fields("VAL_NO")
      Else
        ' For pipelines, open selection form
        Alue = Nz(Screen.ActiveDatasheet("Area").Value)
        Linja = Nz(Screen.ActiveDatasheet("LineNo").Value)
        DoCmd.OpenForm "frmOpenPIPELINE", acNormal, , , , , Alue & "," & Linja
        Exit Sub
      End If
      If Polku = "" Then
        MsgBox "Ei ole tietoa missä kuvassa kohde " & Info & " on !", vbCritical, "Etsi kohde"  ' "No information where object is!"
        Exit Sub
      End If
      AvaaKuvasta Polku, DWG, Handle, Info
    ElseIf UCase$(Application.CurrentObjectName) = "PIPELINEDATA" Or UCase$(Application.CurrentObjectName) = "MANVALVEDATA" Then
        ' Open from detail table view
        Polku = Screen.ActiveDatasheet("PATH").Value
        If Right$(Polku, 1) <> "\" Then Polku = Polku & "\"
        
        DWG = LCase$(Screen.ActiveDatasheet("ImpFileID").Value)
        Handle = Screen.ActiveDatasheet("Handles").Value
        If UCase$(Application.CurrentObjectName) = "PIPELINEDATA" Then
          Info = Screen.ActiveDatasheet("DEP").Value & "-" & Screen.ActiveDatasheet("LINENO").Value
        Else
          Info = Screen.ActiveDatasheet("AREACODE").Value & "-" & Screen.ActiveDatasheet("VALVE_NO").Value
        End If
        AvaaKuvasta Polku, DWG, Handle, Info
    Else
      MsgBox "MANUALVALVES tai PIPELINES taulukon tulee avaoinna näytöllä!", vbCritical, "Etsi kohde"  ' "Table must be open!"
    End If
    Exit Sub

ErrorHandler:
    MsgBox "Error opening block: " & Err.Description, vbCritical, "AvaaBlock"
End Sub
'--------------------------------------------------------------------------------
' Sub: AvaaKuvasta
' Purpose: Open AutoCAD drawing and zoom/highlight specific block
'
' Parameters:
'   Polku - Directory path to drawing file
'   Nimi - Drawing filename (.dwg)
'   Handle - AutoCAD block handle (hex string)
'   Info - Block identification for error messages
'
' Process:
'   1. Connect to running AutoCAD instance
'   2. Check if drawing already open, otherwise open it
'   3. Find block by handle
'   4. Zoom to block and highlight it
'   5. Activate AutoCAD window
'
' Notes:
'   - Requires AutoCAD to be running
'   - Handles missing drawings gracefully
'   - Finnish: "Kohde" = "Object/target"
'   - Optimized: Uses lowercase comparison consistently
'--------------------------------------------------------------------------------
Sub AvaaKuvasta(Polku As String, Nimi As String, Handle As String, Info As String)
On Error GoTo ErrorHandler
    Dim oACAD As Object  ' AcadApplication - muutettu late binding (64-bit)
    Dim Entity As Object  ' AcadEntity - muutettu late binding (64-bit)
    Dim MinPoint As Variant  ' Bounding box minimum point
    Dim MaxPoint As Variant  ' Bounding box maximum point
    Dim i As Integer  ' Loop counter
    Dim OK As Boolean  ' Drawing open flag
    Dim LowerNimi As String  ' Lowercase drawing name for comparison
  
    ' Normalize drawing name for consistent comparison
    LowerNimi = LCase$(Nimi)
    
    ' Try to connect to running AutoCAD
    On Error Resume Next
    Set oACAD = GetObject(, "AutoCAD.Application")
    If Err <> 0 Then
      MsgBox "Käynnissä olevaa AutoCADiä ei löytynyt!" & vbCrLf & "Avaa Autocad ensin.", vbCritical, "Etsi Kohde"  ' "Running AutoCAD not found!"
      Exit Sub
    End If
    On Error GoTo ErrorHandler
    
    ' Check if drawing already open (use cached lowercase name)
    OK = False
    For i = 0 To oACAD.Documents.Count - 1
      If LCase$(oACAD.Documents(i).Name) = LowerNimi Then
        oACAD.Documents(i).Activate
        OK = True
        Exit For
      End If
    Next i
    
    ' Open drawing if not already open
    If Not OK Then
      On Error Resume Next
      oACAD.Documents.Open Polku & Nimi
      If Err = 0 Then
        OK = True
      Else
        MsgBox "Virhe avattaessa dokumenttia: " & vbCrLf & Nimi, vbCritical, "Etsi kohde"  ' "Error opening document"
        Err.Clear
      End If
      On Error GoTo ErrorHandler
    End If
    
    ' Find and highlight block if drawing opened successfully
    If OK Then
      If Handle = "" Then
        MsgBox "Kohteen  " & Info & " sijainti ei ole tiedossa, vain kuva avattiin.", vbCritical, "Etsi kohde"  ' "Block location unknown"
      Else
        oACAD.ActiveDocument.ActiveSpace = acModelSpace
        On Error Resume Next
        Set Entity = oACAD.ActiveDocument.HandleToObject(Handle)
        If Err <> 0 Then
          MsgBox "Kuvasta ei löytynyt kohdetta tietokannan tiedoilla (Handle oli väärä)!", vbCritical, "Etsi kohde"  ' "Block not found (wrong handle)"
          Err.Clear
          On Error GoTo ErrorHandler
        Else
          On Error GoTo ErrorHandler
          ' Zoom to block and highlight
          Entity.GetBoundingBox MinPoint, MaxPoint
          oACAD.ActiveDocument.WindowState = acMax
          oACAD.ZoomWindow MinPoint, MaxPoint
          oACAD.ZoomScaled 0.3, acZoomScaledRelative
          Entity.Highlight True
        End If
      End If
      AppActivate oACAD.Caption, True
    End If
    
    ' Cleanup
    Set Entity = Nothing
    Set oACAD = Nothing
    Exit Sub

ErrorHandler:
    On Error Resume Next
    Set Entity = Nothing
    Set oACAD = Nothing
    On Error GoTo 0
    MsgBox "Error in AvaaKuvasta: " & Err.Description, vbCritical, "Error"
End Sub
'--------------------------------------------------------------------------------
' Function: NetworkUserName
' Purpose: Get Windows network username for DOCUMENTS database forms
'
' Returns: String - Network username or "Unknown"
'
' Notes:
'   - Called by DOCUMENTS\Form_USysReserve and Form_USysAddDocument
'   - Uses api_GetUserName (advapi32.dll) declared in this module
'   - Updated 2026-03-03: Implemented - was previously a broken stub
'--------------------------------------------------------------------------------
Public Function NetworkUserName() As String
    ' Haetaan Windows-verkkokäyttäjänimi (käytetään DOCUMENTS-kannan lomakkeissa)
    Dim BuffSize As Long
    Dim NBuffer As String
    BuffSize = 256
    NBuffer = Space$(BuffSize)
    If api_GetUserName(NBuffer, BuffSize) Then
        NetworkUserName = Left$(NBuffer, InStr(NBuffer, Chr(0)) - 1)
    Else
        NetworkUserName = "Unknown"
    End If
End Function

'--------------------------------------------------------------------------------
' Function: SetStartup
' Purpose: Log user login information to UsysUsers table
'
' Process:
'   1. Get network username via Windows API (api_GetUserName)
'   2. Get computer name via Windows API (api_GetComputerName)
'   3. Get current Access database username
'   4. Write all information to UsysUsers table with timestamp
'
' Returns: Nothing (procedure performs logging silently)
'
' Notes:
'   - Typically called from AutoExec macro or startup form
'   - No error handling - assumes UsysUsers table exists
'   - Buffer size 256 characters for API calls
'--------------------------------------------------------------------------------
Function SetStartup()
On Error GoTo ErrorHandler
    Dim DB As DAO.Database  ' Current database
    Dim Taulu As DAO.Recordset  ' UsysUsers table
    Dim NWUserName As String  ' Network username from Windows
    Dim CName As String  ' Computer name from Windows
    Dim BuffSize As Long  ' Buffer size for API calls
    Dim NBuffer As String  ' Buffer string for API calls
    
    BuffSize = 256
    NBuffer = Space$(BuffSize)
    
    ' Get network username
    If api_GetUserName(NBuffer, BuffSize) Then
      NWUserName = Left$(NBuffer, InStr(NBuffer, Chr(0)) - 1)
    Else
      NWUserName = "Unknown"
    End If
    
    ' Get computer name
    BuffSize = 256
    NBuffer = Space$(BuffSize)
    If api_GetComputerName(NBuffer, BuffSize) Then
      CName = Left$(NBuffer, InStr(NBuffer, Chr(0)) - 1)
    Else
      CName = "Unknown"
    End If
       
    ' Write login record to UsysUsers table
    Set DB = CurrentDb
    Set Taulu = DB.OpenRecordset("UsysUsers", dbOpenTable)
    With Taulu
        .AddNew
        .Fields(0) = NWUserName     ' Network username
        .Fields(1) = CurrentUser()  ' Database username
        .Fields(2) = CName          ' Computer name
        .Fields(3) = Now            ' Login timestamp
        .Update
    End With
    
    ' Cleanup
    Taulu.Close
    Set Taulu = Nothing
    Set DB = Nothing
    Exit Function

ErrorHandler:
    ' Silent error handling - don't interrupt application startup
    On Error Resume Next
    If Not Taulu Is Nothing Then Taulu.Close
    Set Taulu = Nothing
    Set DB = Nothing
    On Error GoTo 0
End Function

'--------------------------------------------------------------------------------
' Function: POIMI
' Purpose: Extract part of hyphen-delimited string
'
' Parameters:
'   Tieto - String to parse (format: "part1-part2-part3")
'   osa - Part number to extract (1-based index)
'
' Returns: Variant - Extracted part or Null if input is null/empty
'
' Example:
'   POIMI("AREA-123-VALVE", 2) returns "123"
'
' Notes:
'   - Uses Split function with hyphen delimiter
'   - Returns Null for null or empty input
'   - Used for parsing valve/pipeline identifiers
'--------------------------------------------------------------------------------
Function POIMI(Tieto As Variant, osa As Integer) As Variant
On Error GoTo ErrorHandler
    Dim Osat As Variant  ' Array of string parts
    
    If IsNull(Tieto) Or Tieto = "" Then
       POIMI = Null
    Else
      Osat = Split(Tieto, "-")
      POIMI = Osat(osa - 1)  ' Convert 1-based to 0-based index
    End If
    Exit Function

ErrorHandler:
    POIMI = Null
End Function
