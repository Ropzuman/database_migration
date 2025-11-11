Option Compare Database
Option Explicit
'================================================================================
' Module: general
' Purpose: General utility functions and file dialog support
' Updated: 2025-11-11 - Added VBA7/64-bit support
'
' Description:
'   Provides utility functions for:
'   - Number formatting (comma to period conversion)
'   - Revision tracking and date parsing
'   - Loop existence checking
'   - File open dialog (Windows Common Dialog)
'
' Dependencies:
'   - comdlg32.dll (Common Dialog API)
'   - _Revisions table (for revision tracking)
'   - qrysolvalve query (for loop checking)
'================================================================================

' Public module-level variables for page numbering (used by Sivunumerointi.bas)
Public Sivunro As Integer  ' Current page number
Public EdelArea As Integer  ' Previous area code
Public Sivuja As Integer  ' Page counter

'--------------------------------------------------------------------------------
' Windows Common Dialog API Declaration
' Updated 2025-11-11: Added VBA7/64-bit support for GetOpenFileName API
'--------------------------------------------------------------------------------
#If VBA7 Then
    Declare PtrSafe Function GetOpenFileName Lib "comdlg32.dll" Alias "GetOpenFileNameA" (pOpenfilename As OPENFILENAME) As Long
    Public Type OPENFILENAME
        lStructSize As Long
        hwndOwner As LongPtr  ' Updated for 64-bit (window handle)
        hInstance As LongPtr  ' Updated for 64-bit (instance handle)
        lpstrFilter As String
        lpstrCustomFilter As String
        nMaxCustFilter As Long
        nFilterIndex As Long
        lpstrFile As String
        nMaxFile As Long
        lpstrFileTitle As String
        nMaxFileTitle As Long
        lpstrInitialDir As String
        lpstrTitle As String
        flags As Long
        nFileOffset As Integer
        nFileExtension As Integer
        lpstrDefExt As String
        lCustData As LongPtr  ' Updated for 64-bit
        lpfnHook As LongPtr  ' Updated for 64-bit (callback pointer)
        lpTemplateName As String
    End Type
#Else
    Declare Function GetOpenFileName Lib "comdlg32.dll" Alias "GetOpenFileNameA" (pOpenfilename As OPENFILENAME) As Long
    Public Type OPENFILENAME
        lStructSize As Long
        hwndOwner As Long
        hInstance As Long
        lpstrFilter As String
        lpstrCustomFilter As String
        nMaxCustFilter As Long
        nFilterIndex As Long
        lpstrFile As String
        nMaxFile As Long
        lpstrFileTitle As String
        nMaxFileTitle As Long
        lpstrInitialDir As String
        lpstrTitle As String
        flags As Long
        nFileOffset As Integer
        nFileExtension As Integer
        lpstrDefExt As String
        lCustData As Long
        lpfnHook As Long
        lpTemplateName As String
    End Type
#End If

'--------------------------------------------------------------------------------
' Function: PilkkuPiste
' Purpose: Converts decimal comma to decimal point (Finnish to international format)
'
' Parameters:
'   Luku - Variant containing number with comma or point decimal separator
'
' Returns:
'   String with decimal point format (e.g., "3,14" becomes "3.14")
'
' Notes:
'   - Returns empty string if input is null or empty
'   - Used for international number format conversion
'   - Commonly used before exporting to CSV or external systems
'--------------------------------------------------------------------------------
Public Function PilkkuPiste(Luku As Variant) As String
On Error GoTo ErrorHandler
    Dim Osoitin As Long  ' Position of comma in string
    
    ' Handle null/empty input
    If Nz(Luku) = "" Then
        PilkkuPiste = ""
        Exit Function
    End If

    ' Find and replace comma with period
    Osoitin = InStr(Luku, ",")
    If Osoitin = 0 Then
        PilkkuPiste = Luku  ' No comma found, return as-is
    Else
        PilkkuPiste = Left(Luku, Osoitin - 1) & "." & Mid(Luku, Osoitin + 1)
    End If
    Exit Function

ErrorHandler:
    PilkkuPiste = ""
End Function

'--------------------------------------------------------------------------------
' Function: UdNoteToRev
' Purpose: Extracts revision number from user notes based on date
'
' Parameters:
'   UdNote - Variant containing user note string with format "text:date|moretext"
'
' Returns:
'   Variant - Revision code from _Revisions table or Null if not found
'
' Notes:
'   - Parses date from UdNote string (format: "something:MM/DD/YYYY|something")
'   - Looks up corresponding revision in _Revisions table
'   - Returns first revision where BeforeDate > parsed date
'   - Used for historical revision tracking
'--------------------------------------------------------------------------------
Public Function UdNoteToRev(UdNote As Variant) As Variant
On Error GoTo ErrorHandler
    Dim Paiva As String  ' Date string extracted from note
    Dim Os As Long  ' Position marker for string parsing
    Dim VP As Date  ' Parsed date value
    Dim RevTaul As DAO.Recordset  ' _Revisions table recordset
    
    ' Handle null input
    If IsNull(UdNote) Then
        UdNoteToRev = Null
        Exit Function
    End If
    
    ' Parse date from note string (format: "text:date|moretext")
    Os = InStr(UdNote, ":")
    If Os > 0 Then
        ' Extract date portion between : and |
        Paiva = Mid(UdNote, Os + 1)
        Paiva = Left(Paiva, InStr(Paiva, "|") - 1)
        VP = DateValue(Paiva)
        Paiva = Month(VP) & "/" & Day(VP) & "/" & Year(VP)   'Format: M/D/YYYY (e.g., 2/1/2007)
        
        ' Look up revision based on date
        Set RevTaul = CurrentDb.OpenRecordset("SELECT * FROM _Revisions WHERE (((BeforeDate) > #" & Paiva & "#)) ORDER BY BeforeDate ASC;")
        If RevTaul.RecordCount > 0 Then
            UdNoteToRev = RevTaul.Fields("Rev").Value
        Else
            UdNoteToRev = Null
        End If
        RevTaul.Close
        Set RevTaul = Nothing
    Else
        UdNoteToRev = Null
    End If
    Exit Function

ErrorHandler:
    On Error Resume Next
    If Not RevTaul Is Nothing Then RevTaul.Close
    Set RevTaul = Nothing
    On Error GoTo 0
    UdNoteToRev = Null
End Function

'--------------------------------------------------------------------------------
' Function: EtsiLoop
' Purpose: Checks if a loop exists in the system
'
' Parameters:
'   Alue - String containing area code
'   Looppi - String containing loop number
'
' Returns:
'   String - "1" if loop exists, "" (empty) if not found
'
' Notes:
'   - Queries qrysolvalve for matching AreaCode and LoopNo
'   - Returns simple existence flag (not boolean for backward compatibility)
'   - Used for validation before creating new loops
'--------------------------------------------------------------------------------
Function EtsiLoop(Alue As String, Looppi As String) As String
On Error GoTo ErrorHandler
    Dim Taul As DAO.Recordset  ' Query results recordset
    
    ' Query for matching loop
    Set Taul = CurrentDb.OpenRecordset("SELECT * From qrysolvalve WHERE AreaCode='" & Alue & "' AND LoopNo='" & Looppi & "'")
    If Taul.EOF Then
        EtsiLoop = ""  ' Not found
    Else
        EtsiLoop = "1"  ' Found
    End If
    Taul.Close
    Set Taul = Nothing
    Exit Function

ErrorHandler:
    On Error Resume Next
    If Not Taul Is Nothing Then Taul.Close
    Set Taul = Nothing
    On Error GoTo 0
    EtsiLoop = ""
End Function
