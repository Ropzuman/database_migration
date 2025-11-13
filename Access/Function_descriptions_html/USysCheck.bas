Option Compare Database
Option Explicit

'================================================================================
' Module: USysCheck
' Purpose: User tracking and validation utilities
' Updated: 2025-11-13 - Added VBA7/64-bit support
'
' Description:
'   Tracks users accessing the database by logging network username,
'   database username, computer name, and timestamp to UsysUsers table.
'
' Dependencies:
'   - Windows API (advapi32.dll, Kernel32)
'   - UsysUsers table
'   - DAO.Recordset
'================================================================================

'---------------------------------------------
' Windows API Declarations - 64-bit compatible
'---------------------------------------------
#If VBA7 Then
    Private Declare PtrSafe Function api_GetUserName Lib "advapi32.dll" Alias "GetUserNameA" (ByVal lpBuffer As String, nSize As Long) As Long
    Private Declare PtrSafe Function api_GetComputerName Lib "Kernel32" Alias "GetComputerNameA" (ByVal lpBuffer As String, nSize As Long) As Long
#Else
    Private Declare Function api_GetUserName Lib "advapi32.dll" Alias "GetUserNameA" (ByVal lpBuffer As String, nSize As Long) As Long
    Private Declare Function api_GetComputerName Lib "Kernel32" Alias "GetComputerNameA" (ByVal lpBuffer As String, nSize As Long) As Long
#End If

'================================================================================
' Function: SniffUser
' Purpose: Log current user information to UsysUsers table
' Returns: Nothing (implicit)
'
' Description:
'   Retrieves network username and computer name from Windows API,
'   combines with Access CurrentUser(), and logs to UsysUsers table with timestamp.
'================================================================================
Function SniffUser()
    Dim DB As DAO.Database
    Dim Taulu As DAO.Recordset
    Dim NWUserName As String
    Dim CName As String
    Dim BuffSize As Long
    Dim NBuffer As String
    
    On Error GoTo ErrorHandler
    
    BuffSize = 256
    NBuffer = Space$(BuffSize)
    
    ' Get Windows network username
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
       
    ' Log to database
    Set DB = CurrentDb
    Set Taulu = DB.OpenRecordset("UsysUsers", dbOpenTable)
    
    With Taulu
        .AddNew
        .Fields(0) = NWUserName     ' Network username
        .Fields(1) = CurrentUser()  ' Access database username
        .Fields(2) = CName          ' Computer name
        .Fields(3) = Now            ' Timestamp
        .Update
    End With
    
    Exit Function
    
ErrorHandler:
    ' Silent error handling - don't disrupt application flow
    On Error Resume Next
    If Not Taulu Is Nothing Then Taulu.Close
    Set Taulu = Nothing
    Set DB = Nothing
End Function
