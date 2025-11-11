Option Compare Database
Option Explicit
'================================================================================
' Module: USysCheck
' Purpose: User logging and tracking functionality
' Author: VG Codes (2001)
' Updated: 2025-11-11 - Added VBA7/64-bit support
'
' Description:
'   Logs user login information to UsysUsers table including:
'   - Network username (from Windows API)
'   - Computer name (from Windows API)
'   - Access database username
'   - Login timestamp
'
' Dependencies:
'   - UsysUsers table in database
'   - advapi32.dll (GetUserName API)
'   - kernel32.dll (GetComputerName API)
'================================================================================

#If VBA7 Then
    Private Declare PtrSafe Function api_GetUserName _
                    Lib "advapi32.dll" _
                    Alias "GetUserNameA" _
                    (ByVal lpBuffer As String, nSize As LongPtr) As Long
    Private Declare PtrSafe Function api_GetComputerName _
                    Lib "kernel32" _
                    Alias "GetComputerNameA" _
                    (ByVal lpBuffer As String, nSize As LongPtr) As Long
#Else
    Private Declare Function api_GetUserName _
                    Lib "advapi32.dll" _
                    Alias "GetUserNameA" _
                    (ByVal lpBuffer As String, nSize As Long) As Long
    Private Declare Function api_GetComputerName _
                    Lib "kernel32" _
                    Alias "GetComputerNameA" _
                    (ByVal lpBuffer As String, nSize As Long) As Long
#End If

'--------------------------------------------------------------------------------
' Function: SniffUser
' Purpose: Logs current user's login information to tracking table
'
' Returns: Nothing (procedure performs logging silently)
'
' Notes:
'   - Errors are suppressed to prevent interruption of application startup
'   - Called typically from AutoExec macro or startup form
'   - Requires UsysUsers table with fields: NetworkUser, DBUser, ComputerName, LoginTime
'--------------------------------------------------------------------------------
Function SniffUser()
On Error GoTo ErrorHandler
    Dim DB As DAO.Database  ' Database reference
    Dim Taulu As DAO.Recordset  ' UsysUsers table recordset
    Dim NWUserName As String  ' Network username from Windows
    Dim CName As String  ' Computer name from Windows
    Dim BuffSize As Long  ' Buffer size for API calls
    Dim NBuffer As String  ' String buffer for API calls
    
    ' Get network username via Windows API
    BuffSize = 256
    NBuffer = Space$(BuffSize)
    
    If api_GetUserName(NBuffer, BuffSize) Then
      NWUserName = Left$(NBuffer, InStr(NBuffer, Chr(0)) - 1)
    Else
      NWUserName = "Unknown"
    End If
    
    ' Get computer name via Windows API
    BuffSize = 256
    NBuffer = Space$(BuffSize)
    If api_GetComputerName(NBuffer, BuffSize) Then
      CName = Left$(NBuffer, InStr(NBuffer, Chr(0)) - 1)
    Else
      CName = "Unknown"
    End If
       
    ' Write login record to tracking table
    Set DB = CurrentDb
    Set Taulu = DB.OpenRecordset("UsysUsers", dbOpenTable)
    With Taulu
        .AddNew
        .Fields(0) = NWUserName     'Users Name In Network
        .Fields(1) = CurrentUser()  'Users Name In This Database
        .Fields(2) = CName          'Users Computer Name
        .Fields(3) = Now            'Time At the Moment
        .Update
    End With
    
    ' Cleanup
    Taulu.Close
    Set Taulu = Nothing
    Set DB = Nothing
    Exit Function

ErrorHandler:
    ' Silent error handling - don't interrupt app startup
    On Error Resume Next
    If Not Taulu Is Nothing Then Taulu.Close
    Set Taulu = Nothing
    Set DB = Nothing
    On Error GoTo 0
End Function
