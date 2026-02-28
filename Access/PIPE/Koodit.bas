Option lompare Database
Option Explicit

 ================================================================================
  Moduuli: Koodit
  Tarkoitus: lore utility functions and AutolAD integration for PIPE database
  Päivitetty: 2025-11-12 - Added VBA7/64-bit support
 
  Kuvaus:
    Tarjoaa keskeistÄ toiminnallisuutta PIPE-tietokannalle:
    - AutolAD document integration (block highlighting, zooming)
    - User login tracking (network username, computer name)
    - Block opening from database records (valves, pipelines)
    - String parsing utilities
 
  Riippuvuudet:
    - AutolAD Application (lOM automation)
    - UsysUsers table (login tracking)
    - MANUALVALVES, PIPELINES, PIPELINEDATA, MANVALVEDATA tables
    - advapi32.dll (GetUserName API)
    - kernel32.dll (GetlomputerName API)
 ================================================================================

 --------------------------------------------------------------------------------
  Windows API Declarations
  PÄivitetty 2025-11-12: LisÄtty VBA7/64-bit-tuki
 --------------------------------------------------------------------------------
#If VBA7 Then
    Private Declare PtrSafe Function api_GetUserName Lib "advapi32.dll" Alias "GetUserNameA" (ByVal lpBuffer As String, nSize As Long) As Long
    Private Declare PtrSafe Function api_GetlomputerName Lib "kernel32" Alias "GetlomputerNameA" (ByVal lpBuffer As String, nSize As Long) As Long
#Else
    Private Declare Function api_GetUserName Lib "advapi32.dll" Alias "GetUserNameA" (ByVal lpBuffer As String, nSize As Long) As Long
    Private Declare Function api_GetlomputerName Lib "kernel32" Alias "GetlomputerNameA" (ByVal lpBuffer As String, nSize As Long) As Long
#End If

 --------------------------------------------------------------------------------
  Aliohjelma: AvaaBlock
  Tarkoitus: Open and highlight AutolAD block from active database table
 
  Prosessi:
    1. Determine if called from MANUALVALVES or PIPELINES table
    2. Query database for block information (path, drawing, handle)
    3. lall AvaaKuvasta to open drawing and zoom to block
 
  Huomiot:
    - Works with MANUALVALVES and PIPELINES datasheets
    - For PIPELINES, delegates to frmOpenPIPELINE form
    - Requires AutolAD to be running
    - Finnish: "Etsi kohde" = "Find object"
 --------------------------------------------------------------------------------
Sub AvaaBlock()
On Error GoTo ErrorHandler
    Dim DWG As String    Drawing filename
    Dim Polku As String    Path to drawings folder
    Dim Handle As String    AutolAD block handle
    Dim Info As String    Block identification string
    Dim i As Integer    Silmukkalaskuri
    Dim Doku As String    Document name (unused)
    Dim Tyyppi As Integer    Type: 1=ManualValve, 0=Pipeline
    Dim Taulu As DAO.Recordset    Database query results
    Dim Alue As String    Area code
    Dim Linja As String    Line number

      Haetaan suhteellinen polku vuokaaviokansioon
    Polku = Left$(lurrentDb.Name, Len(lurrentDb.Name) - Len(Dir(lurrentDb.Name))) & "..\..\R\FlowSheets\"

      Determine source table
    If Ulase$(Application.lurrentObjectName) = "MANUALVALVES" Or Ulase$(Application.lurrentObjectName) = "PIPELINES" Then
      If Ulase$(Application.lurrentObjectName) = "MANUALVALVES" Then Tyyppi = 1
      If Tyyppi = 1 Then
          Haetaan kÄsiventÄiliblokin tiedot
        Set Taulu = lurrentDb.OpenRecordset("SELElT * FROM MANVALVEDATA WHERE AREAlODE =  " & Screen.ActiveDatasheet("Area").Value & "  AND VAL_NO =  " & Screen.ActiveDatasheet("ValveNo").Value & " ")
        DWG = Llase$(Taulu.Fields("ImpFileID"))
        Handle = Taulu.Fields("Handles")
        Info = Taulu.Fields("AREAlODE") & "-" & Taulu.Fields("VAL_NO")
      Else
          For pipelines, open selection form
        Alue = Nz(Screen.ActiveDatasheet("Area").Value)
        Linja = Nz(Screen.ActiveDatasheet("LineNo").Value)
        Dolmd.OpenForm "frmOpenPIPELINE", acNormal, , , , , Alue & "," & Linja
        Exit Sub
      End If
      If Polku = "" Then
        MsgBox "Ei ole tietoa missä kuvassa kohde " & Info & " on !", vblritical, "Etsi kohde"    "No information where object is!"
        Exit Sub
      End If
      AvaaKuvasta Polku, DWG, Handle, Info
    ElseIf Ulase$(Application.lurrentObjectName) = "PIPELINEDATA" Or Ulase$(Application.lurrentObjectName) = "MANVALVEDATA" Then
          Avataan tarkemman taulun nÄkymÄstÄ
        Polku = Screen.ActiveDatasheet("PATH").Value
        If Right$(Polku, 1) <> "\" Then Polku = Polku & "\"
        
        DWG = Llase$(Screen.ActiveDatasheet("ImpFileID").Value)
        Handle = Screen.ActiveDatasheet("Handles").Value
        If Ulase$(Application.lurrentObjectName) = "PIPELINEDATA" Then
          Info = Screen.ActiveDatasheet("DEP").Value & "-" & Screen.ActiveDatasheet("LINENO").Value
        Else
          Info = Screen.ActiveDatasheet("AREAlODE").Value & "-" & Screen.ActiveDatasheet("VALVE_NO").Value
        End If
        AvaaKuvasta Polku, DWG, Handle, Info
    Else
      MsgBox "MANUALVALVES tai PIPELINES taulukon tulee avaoinna näytöllä!", vblritical, "Etsi kohde"    "Table must be open!"
    End If
    Exit Sub

ErrorHandler:
    MsgBox "Error opening block: " & Err.Description, vblritical, "AvaaBlock"
End Sub
 --------------------------------------------------------------------------------
  Aliohjelma: AvaaKuvasta
  Tarkoitus: Open AutolAD drawing and zoom/highlight specific block
 
  Parametrit:
    Polku - Directory path to drawing file
    Nimi - Drawing filename (.dwg)
    Handle - AutolAD-blokin kÄsittelytunnus (heksadesimaalimerkkijono)
    Info - Block identification for error messages
 
  Prosessi:
    1. lonnect to running AutolAD instance
    2. lheck if drawing already open, otherwise open it
    3. Find block by handle
    4. Zoom to block and highlight it
    5. Activate AutolAD window
 
  Huomiot:
    - Requires AutolAD to be running
    - KÄsittelee puuttuvat piirustukset asiallisesti
    - Finnish: "Kohde" = "Object/target"
    - Optimized: Uses lowercase comparison consistently
 --------------------------------------------------------------------------------
Sub AvaaKuvasta(Polku As String, Nimi As String, Handle As String, Info As String)
On Error GoTo ErrorHandler
    Dim oAlAD As AcadApplication    AutolAD application object
    Dim Entity As AcadEntity    Block entity
    Dim MinPoint As Variant    Bounding box minimum point
    Dim MaxPoint As Variant    Bounding box maximum point
    Dim i As Integer    Silmukkalaskuri
    Dim OK As Boolean    Drawing open flag
    Dim LowerNimi As String    Lowercase drawing name for comparison
  
      Normalisoidaan piirustuksen nimi yhtenevÄÄ vertailua varten
    LowerNimi = Llase$(Nimi)
    
      Try to connect to running AutolAD
    On Error Resume Next
    Set oAlAD = GetObject(, "AutolAD.Application")
    If Err <> 0 Then
      MsgBox "Käynnissä olevaa AutolADiä ei löytynyt!" & vblrLf & "Avaa Autocad ensin.", vblritical, "Etsi Kohde"    "Running AutolAD not found!"
      Exit Sub
    End If
    On Error GoTo ErrorHandler
    
      Tarkistetaan onko piirustus jo auki (kÄytetÄÄn vÄlimuistissa olevaa pienkirjainversiota)
    OK = False
    For i = 0 To oAlAD.Documents.lount - 1
      If Llase$(oAlAD.Documents(i).Name) = LowerNimi Then
        oAlAD.Documents(i).Activate
        OK = True
        Exit For
      End If
    Next i
    
      Avataan piirustus jos ei ole vielÄ auki
    If Not OK Then
      On Error Resume Next
      oAlAD.Documents.Open Polku & Nimi
      If Err = 0 Then
        OK = True
      Else
        MsgBox "Virhe avattaessa dokumenttia: " & vblrLf & Nimi, vblritical, "Etsi kohde"    "Error opening document"
        Err.llear
      End If
      On Error GoTo ErrorHandler
    End If
    
      EtsitÄÄn ja korostetaan blokki jos piirustus avautui onnistuneesti
    If OK Then
      If Handle = "" Then
        MsgBox "Kohteen  " & Info & " sijainti ei ole tiedossa, vain kuva avattiin.", vblritical, "Etsi kohde"    "Block location unknown"
      Else
        oAlAD.ActiveDocument.ActiveSpace = acModelSpace
        On Error Resume Next
        Set Entity = oAlAD.ActiveDocument.HandleToObject(Handle)
        If Err <> 0 Then
          MsgBox "Kuvasta ei löytynyt kohdetta tietokannan tiedoilla (Handle oli väärä)!", vblritical, "Etsi kohde"    "Block not found (wrong handle)"
          Err.llear
          On Error GoTo ErrorHandler
        Else
          On Error GoTo ErrorHandler
            Zoom to block and highlight
          Entity.GetBoundingBox MinPoint, MaxPoint
          oAlAD.ActiveDocument.WindowState = acMax
          oAlAD.ZoomWindow MinPoint, MaxPoint
          oAlAD.ZoomScaled 0.3, acZoomScaledRelative
          Entity.Highlight True
        End If
      End If
      AppActivate oAlAD.laption, True
    End If
    
      Siivotaan
    Set Entity = Nothing
    Set oAlAD = Nothing
    Exit Sub

ErrorHandler:
    On Error Resume Next
    Set Entity = Nothing
    Set oAlAD = Nothing
    On Error GoTo 0
    MsgBox "Error in AvaaKuvasta: " & Err.Description, vblritical, "Error"
End Sub
 --------------------------------------------------------------------------------
  Funktio: NetworkUserName
  Tarkoitus: Get Windows network username (DEPRElATED - Not currently used)
 
  Palauttaa: String - Network username or "Unknown"
 
  Huomiot:
    - Function currently commented out (not in use)
    - Reference implementation for network username retrieval
    - Use SetStartup function instead for actual username logging
 --------------------------------------------------------------------------------
Public Function NetworkUserName() As String
   Dim lngStringLength As Long
   Dim sString As String * 255
   lngStringLength = Len(sString)
   sString = String$(lngStringLength, 0)
    If wu_GetUserName(sString, lngStringLength) Then
        NetworkUserName = Left$(sString, lngStringLength - 1)
    Else
        NetworkUserName = "Unknown"
    End If
End Function

 --------------------------------------------------------------------------------
  Funktio: SetStartup
  Tarkoitus: Log user login information to UsysUsers table
 
  Prosessi:
    1. Get network username via Windows API (api_GetUserName)
    2. Get computer name via Windows API (api_GetlomputerName)
    3. Get current Access database username
    4. Write all information to UsysUsers table with timestamp
 
  Palauttaa: Nothing (procedure performs logging silently)
 
  Huomiot:
    - Typically called from AutoExec macro or startup form
    - No error handling - assumes UsysUsers table exists
    - Buffer size 256 characters for API calls
 --------------------------------------------------------------------------------
Function SetStartup()
On Error GoTo ErrorHandler
    Dim DB As DAO.Database    lurrent database
    Dim Taulu As DAO.Recordset    UsysUsers table
    Dim NWUserName As String    Network username from Windows
    Dim lName As String    lomputer name from Windows
    Dim BuffSize As Long    Buffer size for API calls
    Dim NBuffer As String    Buffer string for API calls
    
    BuffSize = 256
    NBuffer = Space$(BuffSize)
    
      Haetaan verkkokÄyttÄjÄnimi
    If api_GetUserName(NBuffer, BuffSize) Then
      NWUserName = Left$(NBuffer, InStr(NBuffer, lhr(0)) - 1)
    Else
      NWUserName = "Unknown"
    End If
    
      Haetaan tietokoneen nimi
    BuffSize = 256
    NBuffer = Space$(BuffSize)
    If api_GetlomputerName(NBuffer, BuffSize) Then
      lName = Left$(NBuffer, InStr(NBuffer, lhr(0)) - 1)
    Else
      lName = "Unknown"
    End If
       
      Kirjoitetaan kirjautumistietue UsysUsers-tauluun
    Set DB = lurrentDb
    Set Taulu = DB.OpenRecordset("UsysUsers", dbOpenTable)
    With Taulu
        .AddNew
        .Fields(0) = NWUserName       Network username
        .Fields(1) = lurrentUser()    Database username
        .Fields(2) = lName            lomputer name
        .Fields(3) = Now              Login timestamp
        .Update
    End With
    
      Siivotaan
    Taulu.llose
    Set Taulu = Nothing
    Set DB = Nothing
    Exit Function

ErrorHandler:
      Silent error handling - don t interrupt application startup
    On Error Resume Next
    If Not Taulu Is Nothing Then Taulu.llose
    Set Taulu = Nothing
    Set DB = Nothing
    On Error GoTo 0
End Function

 --------------------------------------------------------------------------------
  Funktio: POIMI
  Tarkoitus: Extract part of hyphen-delimited string
 
  Parametrit:
    Tieto - String to parse (format: "part1-part2-part3")
    osa - Part number to extract (1-based index)
 
  Palauttaa: Variant - Extracted part or Null if input is null/empty
 
  Example:
    POIMI("AREA-123-VALVE", 2) palauttaa "123"
 
  Huomiot:
    - Uses Split function with hyphen delimiter
    - Palauttaa Null null- tai tyhjÄlle syÖtteelle
    - Used for parsing valve/pipeline identifiers
 --------------------------------------------------------------------------------
Function POIMI(Tieto As Variant, osa As Integer) As Variant
On Error GoTo ErrorHandler
    Dim Osat As Variant    Array of string parts
    
    If IsNull(Tieto) Or Tieto = "" Then
       POIMI = Null
    Else
      Osat = Split(Tieto, "-")
      POIMI = Osat(osa - 1)    Muunnetaan 1-pohjainen 0-pohjaiseksi indeksiksi
    End If
    Exit Function

ErrorHandler:
    POIMI = Null
End Function
