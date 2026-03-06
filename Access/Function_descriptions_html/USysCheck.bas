Option Compare Database
Option Explicit

'================================================================================
' Moduuli: USysCheck
' Tarkoitus: Käyttäjäseuranta ja validointityökalut
' Päivitetty: 2025-11-13 — VBA7/64-bit tuki lisätty
'             2026-03-03 — ErrorHandler-lohkoon lisätty DB.Close ennen Nothing
'
' Kuvaus:
'   Kirjaa tietokantaan kirjautuvat käyttäjät tallentamalla verkkokäyttäjänimen,
'   tietokannan käyttäjänimen, tietokoneen nimen ja aikaleiman UsysUsers-tauluun.
'
' Riippuvuudet:
'   - Windows API (advapi32.dll, Kernel32)
'   - UsysUsers-taulu
'   - DAO.Recordset
'================================================================================

'---------------------------------------------
' Windows API -esittelyt — 64-bit-yhteensopiva
' Huom: nSize on Long (32-bit DWORD), ei LongPtr — vältetään Type Mismatch
'---------------------------------------------
#If VBA7 Then
    Private Declare PtrSafe Function api_GetUserName Lib "advapi32.dll" Alias "GetUserNameA" (ByVal lpBuffer As String, nSize As Long) As Long
    Private Declare PtrSafe Function api_GetComputerName Lib "Kernel32" Alias "GetComputerNameA" (ByVal lpBuffer As String, nSize As Long) As Long
#Else
    Private Declare Function api_GetUserName Lib "advapi32.dll" Alias "GetUserNameA" (ByVal lpBuffer As String, nSize As Long) As Long
    Private Declare Function api_GetComputerName Lib "Kernel32" Alias "GetComputerNameA" (ByVal lpBuffer As String, nSize As Long) As Long
#End If

'================================================================================
' Funktio: SniffUser
' Tarkoitus: Kirjaa nykyisen käyttäjän tiedot UsysUsers-tauluun
' Palauttaa: Ei paluuarvoa
'
' Kuvaus:
'   Hakee verkkokäyttäjänimen ja tietokoneen nimen Windows API:lta,
'   yhdistää ne Access CurrentUser() -funktioon, ja kirjaa tiedot
'   UsysUsers-tauluun aikaleiman kera.
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
    
    ' Haetaan Windows-verkkokäyttäjänimi
    If api_GetUserName(NBuffer, BuffSize) Then
      NWUserName = Left$(NBuffer, InStr(NBuffer, Chr(0)) - 1)
    Else
      NWUserName = "Unknown"
    End If
    
    ' Haetaan tietokoneen nimi
    BuffSize = 256
    NBuffer = Space$(BuffSize)
    If api_GetComputerName(NBuffer, BuffSize) Then
      CName = Left$(NBuffer, InStr(NBuffer, Chr(0)) - 1)
    Else
      CName = "Unknown"
    End If
       
    ' Kirjataan tiedot tietokantaan
    Set DB = CurrentDb
    Set Taulu = DB.OpenRecordset("UsysUsers", dbOpenTable)
    
    With Taulu
        .AddNew
        .Fields(0) = NWUserName     ' Verkkokäyttäjänimi
        .Fields(1) = CurrentUser()  ' Access-tietokannan käyttäjänimi
        .Fields(2) = CName          ' Tietokoneen nimi
        .Fields(3) = Now            ' Aikaleima
        .Update
    End With
    
    ' Suljetaan oikein — .Close ennen Set Nothing (DAO-sääntö)
    Taulu.Close
    Set Taulu = Nothing
    ' CurrentDb-viittausta EI suljeta .Close-kutsulla — vain Set Nothing
    Set DB = Nothing
    Exit Function
    
ErrorHandler:
    ' Hiljainen virheenkäsittely — ei keskeytetä sovelluksen toimintaa
    On Error Resume Next
    If Not Taulu Is Nothing Then Taulu.Close
    Set Taulu = Nothing
    ' CurrentDb-viittausta EI suljeta .Close-kutsulla — vain Set Nothing
    Set DB = Nothing
End Function
