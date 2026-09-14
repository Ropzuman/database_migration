Option Compare Database
Option Explicit

'---------------------------------------------
' VG 2001 - K‰ytt‰j‰seurantamoduuli
' P‰ivitetty 2025-10-22: 64-bit-yhteensopivuus, eksplisiittinen DAO, tapahtumat
'---------------------------------------------

#If VBA7 Then
  ' KORJATTU: nSize on DWORD (32-bit) ó ByRef Long on oikea tyyppi, ei LongPtr
  Private Declare PtrSafe Function api_GetUserName Lib "advapi32.dll" Alias "GetUserNameA" _
    (ByVal lpBuffer As String, ByRef nSize As Long) As Long
  Private Declare PtrSafe Function api_GetComputerName Lib "kernel32" Alias "GetComputerNameA" _
    (ByVal lpBuffer As String, ByRef nSize As Long) As Long
#Else
  Private Declare Function api_GetUserName Lib "advapi32.dll" Alias "GetUserNameA" _
    (ByVal lpBuffer As String, ByRef nSize As Long) As Long
  Private Declare Function api_GetComputerName Lib "kernel32" Alias "GetComputerNameA" _
    (ByVal lpBuffer As String, ByRef nSize As Long) As Long
#End If

' Kirjaa nykyisen k‰ytt‰j‰n ja tietokoneen nimen UsysUsers-tauluun
Function SniffUser()
  Dim db As DAO.Database
  Dim Taulu As DAO.Recordset
  Dim NWUserName As String
  Dim CName As String
  
  ' Puskurikoko DWORD-yhteensopivana Long-muuttujana (riitt‰‰ molemmille, Space$- ja API-kutsulle)
  Dim BufferSize As Long
  Dim NBuffer As String
    
  On Error GoTo ErrHandler
  
  ' Haetaan verkkok‰ytt‰j‰nimi
  BufferSize = 256
  NBuffer = Space$(BufferSize)
  If api_GetUserName(NBuffer, BufferSize) Then
     NWUserName = Left$(NBuffer, InStr(NBuffer, Chr(0)) - 1)
  Else
    NWUserName = "Unknown"
  End If
  
  ' Haetaan tietokoneen nimi
  BufferSize = 256
  NBuffer = Space$(BufferSize)
  If api_GetComputerName(NBuffer, BufferSize) Then
    CName = Left$(NBuffer, InStr(NBuffer, Chr(0)) - 1)
  Else
    CName = "Unknown"
  End If
  
  ' Kirjataan sis‰‰nkirjautuminen tietokantaan tapahtumana
  Set db = CurrentDb
  Dim ws As DAO.Workspace
  Set ws = DBEngine.Workspaces(0) ' Nimenomainen tyˆtila ó ei koske muita avoimia yhteyksi‰
  
  ws.BeginTrans
  On Error GoTo ErrHandler ' Siirryt‰‰n transaktion peruutukseen virheess‰
  
  Set Taulu = db.OpenRecordset("UsysUsers", dbOpenDynaset)
        With Taulu
            .AddNew
            .Fields(0) = Nz(NWUserName, "Unknown")      ' Verkkok‰ytt‰j‰nimi (null-turvallinen)
            .Fields(1) = Nz(CurrentUser(), "Unknown")   ' Tietokannan k‰ytt‰j‰nimi (null-turvallinen)
            .Fields(2) = Nz(CName, "Unknown")           ' Tietokoneen nimi (null-turvallinen)
            .Fields(3) = Now             ' Aikaleima
            .Update
        End With  
            ws.CommitTrans ' Hyv‰ksyt‰‰n transaktio ó kaikki muutokset tallennetaan

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
  ws.Rollback ' Perutaan transaktio ó tietokannan eheys s‰ilyy virhetilanteessa
  ' Virhe kirjataan hiljaisesti ó tapahtuma per‰‰ntyy tietokannan eheyden suojaamiseksi
  Resume Cleanup
End Function

