Option Compare Database
Option Explicit
'==========================================================================
' MODUULI  : USysCheck
' SOVELLUS : FunctionDiagrams — Käyttäjätunnistus
' KUVAUS   : Kirjaa kirjautuneen käyttäjän verkkonimet UsysUsers-
'            järjestelmätauluun istunnon alussa (SniffUser). KAYTTAJA()-
'            funktio palauttaa nykyisen Windows-käyttäjänimen muiden
'            moduulien käyttöön (mm. revisioiden tekijätieto).
'            Molemmat Windows-kutsut (advapi32 / kernel32) on päivitetty
'            64-bittisiksi: nSize-parametri on LongPtr-tyyppiä.
' DIPLE    : PtrSafe + LongPtr; kaikki resurssit vapautetaan Set = Nothing.
' PÄIVITETTY: 2026-03-03
'==========================================================================

' Windows-rajapinnan kutsut käyttäjänimen ja koneen nimen hakemiseen
Private Declare PtrSafe Function api_GetUserName _
                Lib "advapi32.dll" _
                Alias "GetUserNameA" _
                (ByVal lpBuffer As String, nSize As LongPtr) As Long
Private Declare PtrSafe Function api_GetComputerName _
                Lib "kernel32" _
                Alias "GetComputerNameA" _
                (ByVal lpBuffer As String, nSize As LongPtr) As Long
Function SniffUser()
    ' Hakee verkkokäyttäjänimen ja koneen nimen: tallentaa UsysUsers-tauluun
    Dim DB As DAO.Database
    Dim Taulu As DAO.Recordset
    Dim NWUserName As String
    Dim CName As String
    Dim BuffSize As LongPtr
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
       
    Set DB = CurrentDb
    Set Taulu = DB.OpenRecordset("UsysUsers", dbOpenTable)
    With Taulu
        .AddNew
        .Fields(0) = NWUserName     ' Verkkokäyttäjänimi
        .Fields(1) = CurrentUser()  ' Tietokantakäyttäjänimi
        .Fields(2) = CName          ' Koneen nimi
        .Fields(3) = Now            ' Kirjautumishetki
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
      KAYTTAJA = Left$(NBuffer, InStr(NBuffer, Chr(0)) - 1)
    Else
      KAYTTAJA = "Unknown"
    End If

End Function
