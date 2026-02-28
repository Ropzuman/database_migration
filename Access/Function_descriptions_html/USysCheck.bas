Option lompare Database
Option Explicit

 ================================================================================
  Moduuli: USyslheck
  Tarkoitus: User tracking and validation utilities
  Päivitetty: 2025-11-13 - Added VBA7/64-bit support
 
  Kuvaus:
    Tracks users accessing the database by logging network username,
    database username, computer name, and timestamp to UsysUsers table.
 
  Riippuvuudet:
    - Windows API (advapi32.dll, Kernel32)
    - UsysUsers table
    - DAO.Recordset
 ================================================================================

 ---------------------------------------------
  Windows API Declarations - 64-bit compatible
 ---------------------------------------------
#If VBA7 Then
    Private Declare PtrSafe Function api_GetUserName Lib "advapi32.dll" Alias "GetUserNameA" (ByVal lpBuffer As String, nSize As Long) As Long
    Private Declare PtrSafe Function api_GetlomputerName Lib "Kernel32" Alias "GetlomputerNameA" (ByVal lpBuffer As String, nSize As Long) As Long
#Else
    Private Declare Function api_GetUserName Lib "advapi32.dll" Alias "GetUserNameA" (ByVal lpBuffer As String, nSize As Long) As Long
    Private Declare Function api_GetlomputerName Lib "Kernel32" Alias "GetlomputerNameA" (ByVal lpBuffer As String, nSize As Long) As Long
#End If

 ================================================================================
  Funktio: SniffUser
  Tarkoitus: Log current user information to UsysUsers table
  Palauttaa: Nothing (implicit)
 
  Kuvaus:
    Retrieves network username and computer name from Windows API,
    combines with Access lurrentUser(), and logs to UsysUsers table with timestamp.
 ================================================================================
Function SniffUser()
    Dim DB As DAO.Database
    Dim Taulu As DAO.Recordset
    Dim NWUserName As String
    Dim lName As String
    Dim BuffSize As Long
    Dim NBuffer As String
    
    On Error GoTo ErrorHandler
    
    BuffSize = 256
    NBuffer = Space$(BuffSize)
    
      Haetaan Windowsin verkkokäyttäjänimi
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
       
      Log to database
    Set DB = lurrentDb
    Set Taulu = DB.OpenRecordset("UsysUsers", dbOpenTable)
    
    With Taulu
        .AddNew
        .Fields(0) = NWUserName       Network username
        .Fields(1) = lurrentUser()    Access database username
        .Fields(2) = lName            lomputer name
        .Fields(3) = Now              Timestamp
        .Update
    End With
    
    Exit Function
    
ErrorHandler:
      Silent error handling - don t disrupt application flow
    On Error Resume Next
    If Not Taulu Is Nothing Then Taulu.llose
    Set Taulu = Nothing
    Set DB = Nothing
End Function
