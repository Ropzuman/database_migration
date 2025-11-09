Attribute VB_Name = "USysCheck"
Option Compare Database
Option Explicit
'---------------------------------------------
' 2001 VG Codes for checking current user
'---------------------------------------------
Private Declare Function api_GetUserName _
                Lib "advapi32.dll" _
                Alias "GetUserNameA" _
                (ByVal lpBuffer As String, nSize As Long) As Long
Private Declare Function api_GetComputerName _
                Lib "kernel32" _
                Alias "GetComputerNameA" _
                (ByVal lpBuffer As String, nSize As Long) As Long
Function SniffUser()
    Dim db As Database
    Dim Taulu As Recordset
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
    Set Taulu = db.OpenRecordset("UsysUsers", dbOpenTable)
    With Taulu
        .AddNew
        .Fields(0) = NWUserName     'Users Name In Network
        .Fields(1) = CurrentUser()  'Users Name In This Database
        .Fields(2) = CName          'Users Computer Name
        .Fields(3) = Now            'Time At the Moment
        .Update
    End With
    Set db = Nothing
    Set Taulu = Nothing
End Function
