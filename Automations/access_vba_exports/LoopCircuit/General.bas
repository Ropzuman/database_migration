Attribute VB_Name = "General"
Option Compare Database
Option Explicit
'---------------------------------------------
' 2001 VG Codes for checking current user
'---------------------------------------------
#If VBA7 Then
  Private Declare PtrSafe Function api_GetUserName Lib "advapi32.dll" Alias "GetUserNameA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
  Private Declare PtrSafe Function api_GetComputerName Lib "kernel32" Alias "GetComputerNameA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
#Else
  Private Declare Function api_GetUserName Lib "advapi32.dll" Alias "GetUserNameA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
  Private Declare Function api_GetComputerName Lib "kernel32" Alias "GetComputerNameA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
#End If
Function SniffUser()
  Dim db As DAO.Database
  Dim Taulu As DAO.Recordset
    Dim NWUserName As String
    Dim CName As String
    Dim BuffSize As Long
    Dim NBuffer As String
    
    BuffSize = 256
    NBuffer = Space$(BuffSize)
    
    If api_GetUserName(NBuffer, BuffSize) Then
      NWUserName = Left$(NBuffer, InStr(NBuffer, Chr(0)) - 1)
    Else
      NWUserName = "Unknown"
    End If
    BuffSize = 256
    NBuffer = Space$(BuffSize)
    If api_GetComputerName(NBuffer, BuffSize) Then
      CName = Left$(NBuffer, InStr(NBuffer, Chr(0)) - 1)
    Else
      CName = "Unknown"
    End If
       
  Set db = CurrentDb
  On Error GoTo ErrHandler
  ' Use a dynaset and transaction for portability and better performance
  DBEngine.BeginTrans
  Set Taulu = db.OpenRecordset("UsysUsers", dbOpenDynaset)
  With Taulu
    .AddNew
    .Fields(0) = NWUserName     'Users Name In Network
    .Fields(1) = CurrentUser()  'Users Name In This Database
    .Fields(2) = CName          'Users Computer Name
    .Fields(3) = Now            'Time At the Moment
    .Update
  End With
  DBEngine.CommitTrans
Cleanup:
  If Not Taulu Is Nothing Then
    On Error Resume Next
    Taulu.Close
    Set Taulu = Nothing
    On Error GoTo 0
  End If
  Set db = Nothing
  Exit Function

ErrHandler:
  On Error Resume Next
  DBEngine.Rollback
  Resume Cleanup
End Function

