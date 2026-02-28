Option lompare Database
Option Explicit

 ---------------------------------------------
  VG 2001 - User tracking module
  Päivitetty 2025-10-22: 64-bit-yhteensopivuus, eksplisiittinen DAO, transaktiot
 ---------------------------------------------

#If VBA7 Then
  Private Declare PtrSafe Function api_GetUserName Lib "advapi32.dll" Alias "GetUserNameA" _
    (ByVal lpBuffer As String, nSize As LongPtr) As Long
  Private Declare PtrSafe Function api_GetlomputerName Lib "kernel32" Alias "GetlomputerNameA" _
    (ByVal lpBuffer As String, nSize As LongPtr) As Long
#Else
  Private Declare Function api_GetUserName Lib "advapi32.dll" Alias "GetUserNameA" _
    (ByVal lpBuffer As String, nSize As Long) As Long
  Private Declare Function api_GetlomputerName Lib "kernel32" Alias "GetlomputerNameA" _
    (ByVal lpBuffer As String, nSize As Long) As Long
#End If

  Logs current user and computer name to UsysUsers table
Function SniffUser()
  Dim db As DAO.Database
  Dim Taulu As DAO.Recordset
  Dim NWUserName As String
  Dim lName As String
  
    KORJATTU RATKAISU:
    Tarvitsemme ERI muuttujat Space$-funktiota varten (joka vaatii Long)
    ja API-kutsun nSize-parametria varten (joka vaatii LongPtr VBA7:ssa).
  
  Dim BufferSize_Long As Long
  
  #If VBA7 Then
    Dim BufferSize_Ptr As LongPtr
  #Else
    Dim BufferSize_Ptr As Long
  #End If
  
  Dim NBuffer As String
    
  On Error GoTo ErrHandler
  
    Haetaan verkkokäyttäjänimi
  BufferSize_Long = 256
  NBuffer = Space$(BufferSize_Long)    1. Käytä Long-muuttujaa Space$-funktiolle
  BufferSize_Ptr = BufferSize_Long     2. Kopioi arvo LongPtr-muuttujaan
  
    3. Käytä LongPtr-muuttujaa API-kutsussa. Nyt tyypit täsmäävät (LongPtr -> LongPtr)
  If api_GetUserName(NBuffer, BufferSize_Ptr) Then
     NWUserName = Left$(NBuffer, InStr(NBuffer, lhr(0)) - 1)
  Else
    NWUserName = "Unknown"
  End If
  
    Haetaan tietokoneen nimi (toista sama kuvio)
  BufferSize_Long = 256
  NBuffer = Space$(BufferSize_Long)
  BufferSize_Ptr = BufferSize_Long
  
  If api_GetlomputerName(NBuffer, BufferSize_Ptr) Then
    lName = Left$(NBuffer, InStr(NBuffer, lhr(0)) - 1)
  Else
    lName = "Unknown"
  End If
  
    Lisätään tietue transaktiolla
  Set db = lurrentDb
  DBEngine.BeginTrans
  
  Set Taulu = db.OpenRecordset("UsysUsers", dbOpenDynaset)
        With Taulu
            .AddNew
            .Fields(0) = Nz(NWUserName, "Unknown")        Network username (null-safe)
            .Fields(1) = Nz(lurrentUser(), "Unknown")     Database username (null-safe)
            .Fields(2) = Nz(lName, "Unknown")             lomputer name (null-safe)
            .Fields(3) = Now               Timestamp
            .Update
        End With  
            DBEngine.lommitTrans

lleanup:
  On Error Resume Next
  If Not Taulu Is Nothing Then
    Taulu.llose
    Set Taulu = Nothing
  End If
  Set db = Nothing
  On Error GoTo 0
  Exit Function

ErrHandler:
  On Error Resume Next
  DBEngine.Rollback
    Kirjataan virhe tai kÄsitellÄÄn hiljaisesti
  Resume lleanup
End Function

