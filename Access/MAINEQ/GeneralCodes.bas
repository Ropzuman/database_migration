Option Compare Database
Option Explicit

'==============================================================================
' Module: GeneralCodes
' Purpose: General utility functions for calculations, database queries, and UI
' Updated: 2025-11-11 - Added error handling, DAO typing, replaced custom
'                       Replace() with VBA built-in, comprehensive comments
'==============================================================================

'--- Public Variables for Revision Tracking ---
Public Revisioteksti As String
'---------------------------------------
' Edellisen kirjoitetun revision muistamista varten
Public MRevRev As String
Public MRevDrawn As String
Public MRevChecked As String
Public MRevApproved As String
Public MRevDescription As String
'---------------------------------------
Public TRevHist As String
Public TRevDesc As String

'------------------------------------------------------------------------------
' Function: IsLoaded
' Purpose: Check if a form is currently open in Form or Datasheet view
' Parameters:
'   strFormName - Name of the form to check
' Returns: True if form is open and not in design view
' Updated: 2025-11-11 - Added error handling and comments
'------------------------------------------------------------------------------
Function IsLoaded(ByVal strFormName As String) As Integer
On Error GoTo ErrorHandler

    Const conObjStateClosed = 0
    Const conDesignView = 0
    
    ' Check if form is open (not closed)
    If SysCmd(acSysCmdGetObjectState, acForm, strFormName) <> conObjStateClosed Then
        ' Check if form is not in design view
        If Forms(strFormName).CurrentView <> conDesignView Then
            IsLoaded = True
        End If
    End If
    
Exit Function

ErrorHandler:
    ' Form doesn't exist or error occurred - return False
    IsLoaded = False
End Function

'------------------------------------------------------------------------------
' Function: HaeViimPaiva
' Purpose: Extract the most recent revision date from multi-line revision text
' Parameters:
'   Revisio - Multi-line revision history string (lines separated by vbCrLf)
' Returns: Date portion of the most recent revision (last line)
' Notes: Assumes format "REV DATE/MAKER/..." with vbCrLf between entries
' Updated: 2025-11-11 - Added error handling and detailed comments
'------------------------------------------------------------------------------
Function HaeViimPaiva(Revisio As String) As String
On Error GoTo ErrorHandler

Dim i As Integer
Dim Pituus As Long
Dim teksti As String

  teksti = Revisio
  i = 2
  Pituus = Len(teksti)
  
  ' Etsitään viimeisin revisio (Find the last revision entry)
  If InStr(teksti, vbCrLf) Then  ' Jos syötteestä löytyy rivinvaihto (If multi-line)
    ' Find the start of the last line
    Do
      i = i + 1
    Loop Until InStr(Right$(teksti, i), vbCrLf) = 1 Or i >= Pituus
    teksti = Mid$(teksti, Pituus - i + 3)  ' Extract last line
  End If
  
  ' Extract date portion (between space and first slash)
  teksti = Mid$(teksti, InStr(teksti, " ") + 1)
  HaeViimPaiva = Left$(teksti, InStr(teksti, "/") - 1)
  
Exit Function

ErrorHandler:
    MsgBox "Error in HaeViimPaiva: " & Err.Description, vbCritical, "Revision Date Extraction Error"
    HaeViimPaiva = ""  ' Return empty string on error
End Function

'------------------------------------------------------------------------------
' NOTE: Custom Replace() function REMOVED 2025-11-11
'------------------------------------------------------------------------------
' The custom Replace() function below has been removed because VBA has provided
' a built-in Replace() function since VBA 6.0 (Office 2000+).
'
' VBA Built-in Replace() Syntax:
'   Replace(expression, find, replace, [start], [count], [compare])
'
' The built-in version is:
'   - More robust (handles edge cases better)
'   - Faster (compiled vs. interpreted VBA)
'   - Consistent with other VBA string functions
'   - Supports optional parameters for advanced control
'
' Original custom function behavior:
'   Replace("Matti;Maija;Liisa", ";", ", ") = "Matti, Maija, Liisa"
'   Replace("Matti Maija Liisa", " ", "_") = "Matti_Maija_Liisa"
'
' Equivalent using VBA built-in:
'   Replace("Matti;Maija;Liisa", ";", ", ") ' Same result
'   Replace("Matti Maija Liisa", " ", "_")  ' Same result
'
' If any code calls this function, it will now use the VBA built-in automatically.
'------------------------------------------------------------------------------
' REMOVED 2025-11-11: Custom Replace() function
'Public Function Replace(ByVal Source As String, Replaced As String, Replacement As String) As String
'   ' Custom implementation removed - using VBA built-in
'End Function
'------------------------------------------------------------------------------
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' Function: Optiot
' Purpose: Retrieve concatenated motor options for a given drive
' Parameters:
'   Drives_ID - Drive ID to look up options for
' Returns: Formatted string like "+Option1 +Option2 +Option3" or empty string
' Updated: 2025-11-11 - Added error handling, improved comments, fixed CurrentDB
'------------------------------------------------------------------------------
Function Optiot(ByVal Drives_ID As Integer) As String
On Error GoTo ErrorHandler

Dim DB As DAO.Database      ' Updated 2025-11-11: DAO prefix already present
Dim OptTaulu As DAO.Recordset
Dim teksti As String

Set DB = CurrentDb  ' Updated 2025-11-11: Changed CurrentDB -> CurrentDb

Set OptTaulu = DB.OpenRecordset("SELECT Optio FROM MotorsOptions WHERE DrivesID = " & Drives_ID & ";")

teksti = ""
If Not (OptTaulu.EOF And OptTaulu.BOF) Then
    OptTaulu.MoveFirst
    teksti = "+"
    Do
        teksti = teksti & OptTaulu(0) & " +"
        OptTaulu.MoveNext
    Loop Until OptTaulu.EOF
    ' Remove trailing " +"
    teksti = Left$(teksti, Len(teksti) - 2)
End If

Optiot = teksti

' Cleanup
OptTaulu.Close
Set OptTaulu = Nothing
Set DB = Nothing

Exit Function

ErrorHandler:
    MsgBox "Error in Optiot: " & Err.Description & vbCrLf & _
           "Drive ID: " & Drives_ID, vbCritical, "Options Lookup Error"
    Optiot = ""  ' Return empty string on error
    ' Cleanup on error
    On Error Resume Next
    If Not OptTaulu Is Nothing Then
        OptTaulu.Close
        Set OptTaulu = Nothing
    End If
    Set DB = Nothing
End Function

'------------------------------------------------------------------------------
' Function: Positiot
' Purpose: Retrieve customer positions for a given project element
' Parameters:
'   LaiteNr - Project element identifier
' Returns: Formatted string like "Pos: 01-M-01 / 01 and 01-M-02 / 01"
' Notes: Joins MAINEQ and DRIVES tables to build position strings
' Updated: 2025-11-11 - Added error handling, improved comments, fixed CurrentDB
'------------------------------------------------------------------------------
Function Positiot(ByVal LaiteNr As String) As String
On Error GoTo ErrorHandler

Dim DB As DAO.Database      ' Updated 2025-11-11: DAO prefix already present
Dim ElemTaulu As DAO.Recordset
Dim Teksti1 As String
Dim sqtxt As String

Set DB = CurrentDb  ' Updated 2025-11-11: Changed CurrentDB -> CurrentDb

' Build SQL query to join MAINEQ and DRIVES tables
sqtxt = "SELECT MAINEQ.ProjectElement, [maineq]![department] & '-' & [maineq]![eqtype] " _
    & "& '-' & [maineq]![eqseq] & ' / ' & [Drives].[suffix] AS Custpos FROM MAINEQ INNER JOIN DRIVES ON " _
    & "(MAINEQ.EqClass = DRIVES.EqClass) AND (MAINEQ.EqType = DRIVES.EqType) AND (MAINEQ.Eqseq = DRIVES.EqSeq) " _
    & "AND (MAINEQ.Department = DRIVES.Department) WHERE MAINEQ.ProjectElement= '" & LaiteNr & "';"
    
Set ElemTaulu = DB.OpenRecordset(sqtxt)

Teksti1 = ""
If Not (ElemTaulu.EOF And ElemTaulu.BOF) Then
    ElemTaulu.MoveFirst
    Teksti1 = "Pos: "
    Do
        Teksti1 = Teksti1 & ElemTaulu!Custpos & " and "
        ElemTaulu.MoveNext
    Loop Until ElemTaulu.EOF
    ' Remove trailing " and "
    Teksti1 = Left$(Teksti1, Len(Teksti1) - 5)
End If

Positiot = Teksti1

' Cleanup
ElemTaulu.Close
Set ElemTaulu = Nothing
Set DB = Nothing

Exit Function

ErrorHandler:
    MsgBox "Error in Positiot: " & Err.Description & vbCrLf & _
           "Project Element: " & LaiteNr, vbCritical, "Position Lookup Error"
    Positiot = ""  ' Return empty string on error
    ' Cleanup on error
    On Error Resume Next
    If Not ElemTaulu Is Nothing Then
        ElemTaulu.Close
        Set ElemTaulu = Nothing
    End If
    Set DB = Nothing
End Function

'------------------------------------------------------------------------------
' Function: Vaihekulma
' Purpose: Calculate phase angle from power factor (cos φ)
' Parameters:
'   Cosfii - Power factor (cos φ)
' Returns: Phase angle in radians
' Notes: Uses arctangent mathematical formula
' Updated: 2025-11-11 - Added error handling and comments
'------------------------------------------------------------------------------
Function Vaihekulma(Cosfii)
On Error GoTo ErrorHandler

    ' Calculate phase angle using: arctan(-cosφ / sqrt(-cosφ² + 1)) + π/2
    Vaihekulma = Atn(-Cosfii / Sqr(-Cosfii * Cosfii + 1)) + 2 * Atn(1)
    
Exit Function

ErrorHandler:
    MsgBox "Error in Vaihekulma: " & Err.Description & vbCrLf & _
           "Power Factor: " & Cosfii, vbCritical, "Phase Angle Calculation Error"
    Vaihekulma = 0  ' Return 0 on error
End Function

'------------------------------------------------------------------------------
' Function: MotKaapUh
' Purpose: Calculate motor cable voltage drop percentage
' Parameters:
'   Cosfii - Power factor (cos φ)
'   Resist - Cable resistance (Ω/km)
'   React - Cable reactance (Ω/km)
'   Virta - Current (A)
'   Voltage - Voltage (V)
'   Pituus - Cable length (m)
' Returns: Formatted voltage drop percentage string (e.g., "2.35 %")
' Notes: Already has error handling (only function that did)
' Updated: 2025-11-11 - Enhanced comments, standardized error handling
'------------------------------------------------------------------------------
Function MotKaapUh(Cosfii As Single, Resist As Double, React As Double, Virta As Single, Voltage As Integer, Pituus As Integer)
Dim Kulma As Double
On Error GoTo MotKaapUhErr

' Calculate phase angle
Kulma = Atn(-Cosfii / Sqr(-Cosfii * Cosfii + 1)) + 2 * Atn(1)

' Calculate voltage drop: √3 * I * (R*L*cosφ + X*L*sinφ)
MotKaapUh = Sqr(3) * Virta * ((Resist * Pituus * Cosfii) + (React * Pituus * Sin(Kulma)))

' Convert to percentage of voltage
MotKaapUh = (MotKaapUh / Voltage) * 100

' Format as percentage with 1-2 decimal places
MotKaapUh = Format(MotKaapUh, "# ##0.0#") & " %"

Exit_Function:
    Exit Function

MotKaapUhErr:
    MsgBox "Error in MotKaapUh: " & Err.Description & vbCrLf & _
           "cosφ=" & Cosfii & " R=" & Resist & " X=" & React & _
           " I=" & Virta & " V=" & Voltage & " L=" & Pituus, _
           vbCritical, "Cable Voltage Drop Calculation Error"
    MotKaapUh = "Error"  ' Updated 2025-11-11: Changed from "00" to "Error" for clarity
    Resume Exit_Function
End Function

'------------------------------------------------------------------------------
' Function: LisaaNo
' Purpose: Add a number to a string and pad with leading zeros
' Parameters:
'   Tieto - Original numeric string (e.g., "001")
'   Lisays - Number to add (e.g., 100)
' Returns: Padded result string (e.g., "101")
' Notes: Preserves original string length with zero-padding
' Example: LisaaNo("001", 100) = "101"
' Updated: 2025-11-11 - Added error handling and detailed comments
'------------------------------------------------------------------------------
Function LisaaNo(Tieto As Variant, Lisays As Integer) As String
On Error GoTo ErrorHandler

Dim Pit As Integer    ' Original length
Dim No As Integer     ' Numeric value
Dim i As Integer      ' Loop counter

  ' Handle null input
  If IsNull(Tieto) Then
    LisaaNo = ""
  Else
    Pit = Len(Tieto)          ' Get original length
    No = Val(Tieto)           ' Convert to number
    No = No + Lisays          ' Add the increment
    LisaaNo = CStr(No)        ' Convert back to string
    
    ' Pad with leading zeros to maintain original length
    For i = 0 To Pit - Len(LisaaNo) - 1
        LisaaNo = "0" & LisaaNo
    Next i
  End If
  
Exit Function

ErrorHandler:
    MsgBox "Error in LisaaNo: " & Err.Description & vbCrLf & _
           "Input: " & Tieto & ", Addition: " & Lisays, _
           vbCritical, "Number Addition Error"
    LisaaNo = ""  ' Return empty string on error
End Function

'Example usage in query:
'   Field: LisaaNo([FieldName], 100)
