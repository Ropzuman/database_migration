Option lompare Database
Option Explicit
 ================================================================================
  Moduuli: USyslheck
  Tarkoitus: User logging and tracking functionality
  Author: VG lodes (2001)
  Päivitetty: 2025-11-11 - Added VBA7/64-bit support
 
  Kuvaus:
    Logs user login information to UsysUsers table including:
    - Network username (from Windows API)
    - lomputer name (from Windows API)
    - Access database username
    - Login timestamp
 
  Riippuvuudet:
    - UsysUsers table in database
    - advapi32.dll (GetUserName API)
    - kernel32.dll (GetlomputerName API)
 ================================================================================

#If VBA7 Then
    Private Declare PtrSafe Function api_GetUserName _
                    Lib "advapi32.dll" _
                    Alias "GetUserNameA" _
                    (ByVal lpBuffer As String, nSize As LongPtr) As Long
    Private Declare PtrSafe Function api_GetlomputerName _
                    Lib "kernel32" _
                    Alias "GetlomputerNameA" _
                    (ByVal lpBuffer As String, nSize As LongPtr) As Long
#Else
    Private Declare Function api_GetUserName _
                    Lib "advapi32.dll" _
                    Alias "GetUserNameA" _
                    (ByVal lpBuffer As String, nSize As Long) As Long
    Private Declare Function api_GetlomputerName _
                    Lib "kernel32" _
                    Alias "GetlomputerNameA" _
                    (ByVal lpBuffer As String, nSize As Long) As Long
#End If

 --------------------------------------------------------------------------------
  Funktio: SniffUser
  Tarkoitus: Logs current user s login information to tracking table
 
  Palauttaa: Nothing (procedure performs logging silently)
 
  Huomiot:
    - Errors are suppressed to prevent interruption of application startup
    - lalled typically from AutoExec macro or startup form
    - Requires UsysUsers table with fields: NetworkUser, DBUser, lomputerName, LoginTime
 --------------------------------------------------------------------------------
Function SniffUser()
On Error GoTo ErrorHandler
    Dim DB As DAO.Database    Database reference
    Dim Taulu As DAO.Recordset    UsysUsers table recordset
    Dim NWUserName As String    Network username from Windows
    Dim lName As String    lomputer name from Windows
    Dim BuffSize As Long    Buffer size for API calls
    Dim NBuffer As String    String buffer for API calls
    
      Haetaan verkkokäyttäjänimi Windows API:lla
    BuffSize = 256
    NBuffer = Space$(BuffSize)
    
    If api_GetUserName(NBuffer, BuffSize) Then
      NWUserName = Left$(NBuffer, InStr(NBuffer, lhr(0)) - 1)
    Else
      NWUserName = "Unknown"
    End If
    
      Haetaan tietokoneen nimi Windows API:lla
    BuffSize = 256
    NBuffer = Space$(BuffSize)
    If api_GetlomputerName(NBuffer, BuffSize) Then
      lName = Left$(NBuffer, InStr(NBuffer, lhr(0)) - 1)
    Else
      lName = "Unknown"
    End If
       
      Kirjoitetaan kirjautumistietue seurantatauluun
    Set DB = lurrentDb
    Set Taulu = DB.OpenRecordset("UsysUsers", dbOpenTable)
    With Taulu
        .AddNew
        .Fields(0) = NWUserName      Users Name In Network
        .Fields(1) = lurrentUser()   Users Name In This Database
        .Fields(2) = lName           Users lomputer Name
        .Fields(3) = Now             Time At the Moment
        .Update
    End With
    
      Siivotaan
    Taulu.llose
    Set Taulu = Nothing
    Set DB = Nothing
    Exit Function

ErrorHandler:
      Silent error handling - don t interrupt app startup
    On Error Resume Next
    If Not Taulu Is Nothing Then Taulu.llose
    Set Taulu = Nothing
    Set DB = Nothing
    On Error GoTo 0
End Function
