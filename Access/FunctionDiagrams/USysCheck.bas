Option Compare Database
Option Explicit
'==========================================================================
' MODUULI  : USysCheck
' SOVELLUS : FunctionDiagrams ó K‰ytt‰j‰tunnistus
' KUVAUS   : Kirjaa kirjautuneen k‰ytt‰j‰n verkkonimet UsysUsers-
'            j‰rjestelm‰tauluun istunnon alussa (SniffUser). KAYTTAJA()-
'            funktio palauttaa nykyisen Windows-k‰ytt‰j‰nimen muiden
'            moduulien k‰yttˆˆn (mm. revisioiden tekij‰tieto).
'            Molemmat Windows-kutsut (advapi32 / kernel32) on p‰ivitetty
'            PtrSafe-muotoon; nSize on ByRef Long (LPDWORD), ei LongPtr.
' HUOM    : PtrSafe + LongPtr; kaikki resurssit vapautetaan Set = Nothing.
' PƒIVITETTY: 2026-03-03
'==========================================================================

' Windows-rajapinnan kutsut k‰ytt‰j‰nimen ja koneen nimen hakemiseen
' nSize on LPDWORD (osoitin 32-bittiseen DWORD:iin) ó ByRef Long, EI LongPtr
' ANSI-versiot (A), koska ByVal String v‰litet‰‰n API:lle ANSI-merkkijonona.
' W-versio kirjoittaisi UTF-16:ta ja nimest‰ j‰isi luettavaksi vain ensimm‰inen merkki.
#If VBA7 Then
Private Declare PtrSafe Function api_GetUserName _
                Lib "advapi32.dll" _
                Alias "GetUserNameA" _
                (ByVal lpBuffer As String, ByRef nSize As Long) As Long
Private Declare PtrSafe Function api_GetComputerName _
                Lib "kernel32" _
                Alias "GetComputerNameA" _
                (ByVal lpBuffer As String, ByRef nSize As Long) As Long
#Else
Private Declare Function api_GetUserName _
                Lib "advapi32.dll" _
                Alias "GetUserNameA" _
                (ByVal lpBuffer As String, ByRef nSize As Long) As Long
Private Declare Function api_GetComputerName _
                Lib "kernel32" _
                Alias "GetComputerNameA" _
                (ByVal lpBuffer As String, ByRef nSize As Long) As Long
#End If
' Poistaa API:n palauttaman puskurin p‰‰tt‰v‰n nollamerkin ja sen j‰lkeisen osan.
Private Function PoistaNollat(ByVal Puskuri As String) As String
    Dim Paikka As Long
    Paikka = InStr(Puskuri, vbNullChar)
    If Paikka > 0 Then
        PoistaNollat = Left$(Puskuri, Paikka - 1)
    Else
        PoistaNollat = Puskuri
    End If
End Function

' Palauttaa Windowsin k‰ytt‰j‰nimen tai "Unknown", jos haku ep‰onnistuu.
Private Function HaeKayttajanimi() As String
    Dim BuffSize As Long
    Dim NBuffer As String

    BuffSize = 256
    NBuffer = Space$(BuffSize)
    If api_GetUserName(NBuffer, BuffSize) <> 0 Then
        HaeKayttajanimi = PoistaNollat(NBuffer)
    Else
        HaeKayttajanimi = "Unknown"
    End If
End Function

' Palauttaa koneen nimen tai "Unknown", jos haku ep‰onnistuu.
Private Function HaeKoneenNimi() As String
    Dim BuffSize As Long
    Dim NBuffer As String

    BuffSize = 256
    NBuffer = Space$(BuffSize)
    If api_GetComputerName(NBuffer, BuffSize) <> 0 Then
        HaeKoneenNimi = PoistaNollat(NBuffer)
    Else
        HaeKoneenNimi = "Unknown"
    End If
End Function

Function SniffUser()
    ' Kirjataan k‰ytt‰j‰, kone ja kirjautumishetki UsysUsers-tauluun.
    ' Kirjauksen ep‰onnistuminen ei saa est‰‰ sovelluksen k‰ynnistymist‰.
    Dim DB As Object
    Dim Taulu As Object

    On Error GoTo Virhe

    Set DB = CurrentDb
    ' dbOpenDynaset (2): dbOpenTable (1) ei toimi linkitetyille tauluille
    Set Taulu = DB.OpenRecordset("UsysUsers", 2)
    With Taulu
        .AddNew
        .Fields(0) = HaeKayttajanimi()  ' Verkkok‰ytt‰j‰nimi
        .Fields(1) = CurrentUser()      ' Tietokantak‰ytt‰j‰nimi
        .Fields(2) = HaeKoneenNimi()    ' Koneen nimi
        .Fields(3) = Now                ' Kirjautumishetki
        .Update
    End With

Siivous:
    On Error Resume Next
    If Not Taulu Is Nothing Then Taulu.Close
    Set Taulu = Nothing
    Set DB = Nothing
    Exit Function

Virhe:
    Debug.Print "SniffUser: kirjaus ep‰onnistui: " & Err.Number & " - " & Err.Description
    Resume Siivous
End Function

Function KAYTTAJA() As String
    KAYTTAJA = HaeKayttajanimi()
End Function
