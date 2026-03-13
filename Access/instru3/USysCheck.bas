Option Compare Database
Option Explicit
'================================================================================
' Moduuli: USysCheck
' Tarkoitus: Käyttäjän kirjautumisen seuranta ja lokitus
' Tekijä: VG Codes (2001)
' Päivitetty: 2025-11-11 - VBA7/64-bit-tuki lisätty
'             2026-03-03 - Kommentit suomeksi, nSize-tyyppikorjaus
'             2026-03-06 - Siirtyminen Unicode W -API-versioihin (GetUserNameW, GetComputerNameW)
'
' Kuvaus:
'   Kirjaa käyttäjän kirjautumistiedot UsysUsers-tauluun, mukaan lukien:
'   - Verkkokäyttäjänimi (Windows API — Unicode)
'   - Tietokonenimi (Windows API — Unicode)
'   - Access-tietokannan käyttäjänimi
'   - Kirjautumisaika
'
' Riippuvuudet:
'   - UsysUsers-taulu tietokannassa
'   - advapi32.dll (GetUserNameW API)
'   - kernel32.dll (GetComputerNameW API)
'================================================================================

' Unicode W -versiot tukevat skandinaavisia merkkejä (Ä, Ö) ilman merkistökorruptioriskiä
#If VBA7 Then
    Private Declare PtrSafe Function api_GetUserName _
                    Lib "advapi32.dll" _
                    Alias "GetUserNameW" _
                    (ByVal lpBuffer As String, ByRef nSize As Long) As Long
    Private Declare PtrSafe Function api_GetComputerName _
                    Lib "kernel32" _
                    Alias "GetComputerNameW" _
                    (ByVal lpBuffer As String, ByRef nSize As Long) As Long
#Else
    Private Declare PtrSafe Function api_GetUserName _
                    Lib "advapi32.dll" _
                    Alias "GetUserNameW" _
                    (ByVal lpBuffer As String, nSize As Long) As Long
    Private Declare PtrSafe Function api_GetComputerName _
                    Lib "kernel32" _
                    Alias "GetComputerNameW" _
                    (ByVal lpBuffer As String, nSize As Long) As Long
#End If

'--------------------------------------------------------------------------------
' Funktio: SniffUser
' Tarkoitus: Kirjaa nykyisen käyttäjän kirjautumistiedot seurantatauluun
'
' Palauttaa: Ei mitään (toimenpide suoritetaan hiljaisesti)
'
' Huomiot:
'   - Virheet vaiennetaan, jotta sovelluksen käynnistys ei keskeydy
'   - Kutsutaan tyypillisesti AutoExec-makrosta tai käynnistyslomakkeesta
'   - Vaatii UsysUsers-taulun kentäillä: NetworkUser, DBUser, ComputerName, LoginTime
'--------------------------------------------------------------------------------
Function SniffUser()
On Error GoTo ErrorHandler
    Dim DB As DAO.Database  ' Tietokantaviittaus
    Dim Taulu As DAO.Recordset  ' UsysUsers-taulun recordset
    Dim NWUserName As String  ' Verkkokäyttäjänimi Windowsista
    Dim CName As String  ' Tietokoneen nimi Windowsista
    Dim BuffSize As Long  ' Puskurin koko API-kutsuille
    Dim NBuffer As String  ' Merkkijonopuskuri API-kutsuille
    
    ' Haetaan verkkokäyttäjänimi Windows API:n avulla
    BuffSize = 256
    NBuffer = Space$(BuffSize)
    
    If api_GetUserName(NBuffer, BuffSize) Then
      NWUserName = Left$(NBuffer, InStr(NBuffer, Chr(0)) - 1)
    Else
      NWUserName = "Tuntematon"
    End If
    
    ' Haetaan tietokoneen nimi Windows API:n avulla
    BuffSize = 256
    NBuffer = Space$(BuffSize)
    If api_GetComputerName(NBuffer, BuffSize) Then
      CName = Left$(NBuffer, InStr(NBuffer, Chr(0)) - 1)
    Else
      CName = "Tuntematon"
    End If
       
    ' Kirjoitetaan kirjautumisrivi seurantatauluun
    Set DB = CurrentDb
    Set Taulu = DB.OpenRecordset("UsysUsers", dbOpenTable)
    With Taulu
        .AddNew
        .Fields(0) = NWUserName     'Käyttäjänimi verkossa
        .Fields(1) = CurrentUser()  'Käyttäjänimi tässä tietokannassa
        .Fields(2) = CName          'Tietokoneen nimi
        .Fields(3) = Now            'Kirjautumisaika
        .Update
    End With
    
    ' Siivotaan resurssit
    Taulu.Close
    Set Taulu = Nothing
    Set DB = Nothing
    Exit Function

ErrorHandler:
    ' Vaiennetaan virhe – ei saa keskeyyttää sovelluksen käynnistystä
    On Error Resume Next
    If Not Taulu Is Nothing Then Taulu.Close
    Set Taulu = Nothing
    Set DB = Nothing
    On Error GoTo 0
End Function
