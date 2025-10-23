Attribute VB_Name = "General"
Option Compare Database
Option Explicit

'---------------------------------------------
' VG 2001 - User tracking module
' Updated 2025-10-22: 64-bit compatibility, explicit DAO, transactions
'---------------------------------------------

#If VBA7 Then
  Private Declare PtrSafe Function api_GetUserName Lib "advapi32.dll" Alias "GetUserNameA" _
    (ByVal lpBuffer As String, nSize As LongPtr) As Long
  Private Declare PtrSafe Function api_GetComputerName Lib "kernel32" Alias "GetComputerNameA" _
    (ByVal lpBuffer As String, nSize As LongPtr) As Long
#Else
  Private Declare Function api_GetUserName Lib "advapi32.dll" Alias "GetUserNameA" _
    (ByVal lpBuffer As String, nSize As Long) As Long
  Private Declare Function api_GetComputerName Lib "kernel32" Alias "GetComputerNameA" _
    (ByVal lpBuffer As String, nSize As Long) As Long
#End If

' Logs current user and computer name to UsysUsers table
Function SniffUser()
  Dim db As DAO.Database
  Dim Taulu As DAO.Recordset
  Dim NWUserName As String
  Dim CName As String
  #If VBA7 Then
    Dim BuffSize As LongPtr
  #Else
    Dim BuffSize As Long
  #End If
  Dim NBuffer As String
    
  On Error GoTo ErrHandler
  
  ' Get network username
  BuffSize = 256
  NBuffer = Space$(BuffSize)
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
  
  ' Insert record with transaction
  Set db = CurrentDb
  DBEngine.BeginTrans
  
  Set Taulu = db.OpenRecordset("UsysUsers", dbOpenDynaset)
  With Taulu
    .AddNew
    .Fields(0) = NWUserName      ' Network username
    .Fields(1) = CurrentUser()   ' Database username
    .Fields(2) = CName           ' Computer name
    .Fields(3) = Now             ' Timestamp
    .Update
  End With
  
  DBEngine.CommitTrans

Cleanup:
  On Error Resume Next
  If Not Taulu Is Nothing Then
    Taulu.Close
    Set Taulu = Nothing
  End If
  Set db = Nothing
  On Error GoTo 0
  Exit Function

ErrHandler:
  On Error Resume Next
  DBEngine.Rollback
  ' Log error or handle silently
  Resume Cleanup
End Function
