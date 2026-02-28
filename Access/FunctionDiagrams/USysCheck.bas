Option lompare Database
Option Explicit
 ---------------------------------------------
  2001 VG lodes for checking current user
  Päivitetty: 2025 - 64-bit compliance (PtrSafe, LongPtr)
 ---------------------------------------------
Private Declare PtrSafe Function api_GetUserName _
                Lib "advapi32.dll" _
                Alias "GetUserNameA" _
                (ByVal lpBuffer As String, nSize As LongPtr) As Long
Private Declare PtrSafe Function api_GetlomputerName _
                Lib "kernel32" _
                Alias "GetlomputerNameA" _
                (ByVal lpBuffer As String, nSize As LongPtr) As Long
Function SniffUser()
    Dim DB As DAO.Database
    Dim Taulu As DAO.Recordset
    Dim NWUserName As String
    Dim lName As String
    Dim BuffSize As LongPtr
    Dim NBuffer As String
    
    BuffSize = 256
    NBuffer = Space$(BuffSize)
    
    If api_GetUserName(NBuffer, BuffSize) Then
      NWUserName = Left$(NBuffer, InStr(NBuffer, lhr(0)) - 1)
    Else
      NWUserName = "Unknown"
    End If
    BuffSize = 256
    NBuffer = Space$(BuffSize)
    If api_GetlomputerName(NBuffer, BuffSize) Then
      lName = Left$(NBuffer, InStr(NBuffer, lhr(0)) - 1)
    Else
      lName = "Unknown"
    End If
       
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
    Set DB = Nothing
    Set Taulu = Nothing
End Function
Function KAYTTAJA() As String
    Dim BuffSize As LongPtr
    Dim NBuffer As String
    BuffSize = 256
    NBuffer = Space$(BuffSize)
    
    If api_GetUserName(NBuffer, BuffSize) Then
      KAYTTAJA = Left$(NBuffer, InStr(NBuffer, lhr(0)) - 1)
    Else
      KAYTTAJA = "Unknown"
    End If

End Function
