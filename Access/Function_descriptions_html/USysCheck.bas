Option Compare Database
Option Explicit

'================================================================================
' Moduuli: USysCheck
' Tarkoitus: K‰ytt‰j‰seuranta ja validointityˆkalut
' P‰ivitetty: 2025-11-13 ó VBA7/64-bit tuki lis‰tty
'             2026-03-03 ó Poistettu virheellinen DB.Close CurrentDb-viittaukselta
'             2026-03-06 ó Kentt‰indeksit korvattu nimetyill‰ sarakkeilla (yll‰pidett‰vyys)
'             2026-11-10 - API-aliakset W -> A (ByVal String -puskuri on ANSI); dbAppendOnly; Left$-suojaus
'             2026-11-10 - Viimeistely: DAO-tyypit Object-muuttujien tilalle; virheen kirjaus Immediate-ikkunaan
'
' Kuvaus:
'   Kirjaa tietokantaan kirjautuvat k‰ytt‰j‰t tallentamalla verkkok‰ytt‰j‰nimen,
'   tietokannan k‰ytt‰j‰nimen, tietokoneen nimen ja aikaleiman UsysUsers-tauluun.
'
' Riippuvuudet:
'   - Windows API (advapi32.dll, Kernel32)
'   - UsysUsers-taulu
'   - DAO.Recordset
'================================================================================

'---------------------------------------------
' Windows API -esittelyt ó 64-bit-yhteensopiva
' Huom: nSize on Long (32-bit DWORD), ei LongPtr ó v‰ltet‰‰n Type Mismatch
'---------------------------------------------
#If VBA7 Then
    Private Declare PtrSafe Function api_GetUserName Lib "advapi32.dll" Alias "GetUserNameA" (ByVal lpBuffer As String, ByRef nSize As Long) As Long
    Private Declare PtrSafe Function api_GetComputerName Lib "Kernel32" Alias "GetComputerNameA" (ByVal lpBuffer As String, ByRef nSize As Long) As Long
#Else
  Private Declare Function api_GetUserName Lib "advapi32.dll" Alias "GetUserNameA" (ByVal lpBuffer As String, ByRef nSize As Long) As Long
  Private Declare Function api_GetComputerName Lib "Kernel32" Alias "GetComputerNameA" (ByVal lpBuffer As String, ByRef nSize As Long) As Long
#End If

'================================================================================
' Funktio: SniffUser
' Tarkoitus: Kirjaa nykyisen k‰ytt‰j‰n tiedot UsysUsers-tauluun
' Palauttaa: Ei paluuarvoa
'
' Kuvaus:
'   Hakee verkkok‰ytt‰j‰nimen ja tietokoneen nimen Windows API:lta,
'   yhdist‰‰ ne Access CurrentUser() -funktioon, ja kirjaa tiedot
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
    
    ' Haetaan Windows-verkkok‰ytt‰j‰nimi
    If api_GetUserName(NBuffer, BuffSize) Then
      NWUserName = Left$(NBuffer, InStr(NBuffer & vbNullChar, vbNullChar) - 1)
    Else
      NWUserName = "Unknown"
    End If
    
    ' Haetaan tietokoneen nimi
    BuffSize = 256
    NBuffer = Space$(BuffSize)
    If api_GetComputerName(NBuffer, BuffSize) Then
      CName = Left$(NBuffer, InStr(NBuffer & vbNullChar, vbNullChar) - 1)
    Else
      CName = "Unknown"
    End If
       
    ' Kirjataan tiedot tietokantaan
    Set DB = CurrentDb
    Set Taulu = DB.OpenRecordset("UsysUsers", dbOpenDynaset, dbAppendOnly)
    
    With Taulu
        .AddNew
        .Fields("Verkkotunnus") = NWUserName     ' Windows-verkkok‰ytt‰j‰nimi
        .Fields("AccessTunnus") = CurrentUser()  ' Access-tietokannan k‰ytt‰j‰nimi
        .Fields("KoneenNimi")   = CName          ' Tietokoneen nimi
        .Fields("Aikaleima")    = Now             ' Kirjaushetki
        .Update
    End With
    
    ' Suljetaan oikein ó .Close ennen Set Nothing (DAO-s‰‰ntˆ)
    Taulu.Close
    Set Taulu = Nothing
    ' CurrentDb-viittausta EI suljeta .Close-kutsulla ó vain Set Nothing
    Set DB = Nothing
    Exit Function
    
ErrorHandler:
    Debug.Print "SniffUser-virhe: " & Err.Description
    ' Hiljainen virheenk‰sittely - ei keskeytet‰ sovelluksen toimintaa
    On Error Resume Next
    If Not Taulu Is Nothing Then Taulu.Close
    Set Taulu = Nothing
    ' CurrentDb-viittausta EI suljeta .Close-kutsulla ó vain Set Nothing
    Set DB = Nothing
End Function
