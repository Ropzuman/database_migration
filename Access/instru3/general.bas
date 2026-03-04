Option Compare Database
Option Explicit
'================================================================================
' Moduuli: general
' Tarkoitus: Yleiset apufunktiot ja tiedostodialogituki
' Päivitetty: 2025-11-11 - VBA7/64-bit-tuki lisätty
'             2026-03-03 - Kommentit suomeksi
'
' Kuvaus:
'   Tarjoaa apufunktioita:
'   - Lukumuotoilu (pilkku↔piste-muunnos)
'   - Revisioseuranta ja päivämääräjäsentäminen
'   - Loopin olemassaolon tarkistus
'   - Tiedoston avausdialogi (Windows Common Dialog)
'
' Riippuvuudet:
'   - comdlg32.dll (Common Dialog API)
'   - _Revisions-taulu (revisioseurantaan)
'   - qrysolvalve-kysely (looppitarkistukseen)
'================================================================================

' Moduulitason julkiset muuttujat sivunumeroinnille (käyttää Sivunumerointi.bas)
Public Sivunro As Integer  ' Nykyinen sivunumero
Public EdelArea As Integer  ' Edellinen aluekoodi
Public Sivuja As Integer  ' Sivulaskuri

'--------------------------------------------------------------------------------
' Windows Common Dialog API -määrittely
' Päivitetty 2025-11-11: VBA7/64-bit-tuki lisätty GetOpenFileName-APIlle
'--------------------------------------------------------------------------------
#If VBA7 Then
    Declare PtrSafe Function GetOpenFileName Lib "comdlg32.dll" Alias "GetOpenFileNameA" (pOpenfilename As OPENFILENAME) As Long
    Public Type OPENFILENAME
        lStructSize As Long
        hwndOwner As LongPtr  ' Updated for 64-bit (window handle)
        hInstance As LongPtr  ' Updated for 64-bit (instance handle)
        lpstrFilter As String
        lpstrCustomFilter As String
        nMaxCustFilter As Long
        nFilterIndex As Long
        lpstrFile As String
        nMaxFile As Long
        lpstrFileTitle As String
        nMaxFileTitle As Long
        lpstrInitialDir As String
        lpstrTitle As String
        flags As Long
        nFileOffset As Integer
        nFileExtension As Integer
        lpstrDefExt As String
        lCustData As LongPtr  ' Updated for 64-bit
        lpfnHook As LongPtr  ' Updated for 64-bit (callback pointer)
        lpTemplateName As String
    End Type
#Else
    Declare Function GetOpenFileName Lib "comdlg32.dll" Alias "GetOpenFileNameA" (pOpenfilename As OPENFILENAME) As Long
    Public Type OPENFILENAME
        lStructSize As Long
        hwndOwner As Long
        hInstance As Long
        lpstrFilter As String
        lpstrCustomFilter As String
        nMaxCustFilter As Long
        nFilterIndex As Long
        lpstrFile As String
        nMaxFile As Long
        lpstrFileTitle As String
        nMaxFileTitle As Long
        lpstrInitialDir As String
        lpstrTitle As String
        flags As Long
        nFileOffset As Integer
        nFileExtension As Integer
        lpstrDefExt As String
        lCustData As Long
        lpfnHook As Long
        lpTemplateName As String
    End Type
#End If

'--------------------------------------------------------------------------------
' Funktio: PilkkuPiste
' Tarkoitus: Muuntaa desimaalipilkun pisteeksi (suomalainen → kansainvälinen muoto)
'
' Parametrit:
'   Luku - Variantti, joka sisältää luvun pilkulla tai pisteellä erotettuna
'
' Palauttaa:
'   Merkkijono, jossa desimaalierotin on piste (esim. "3,14" → "3.14")
'
' Huomiot:
'   - Palauttaa tyhjän merkkijonon, jos syöte on null tai tyhjä
'   - Käytetty kansainväliseen lukuformaattimuunnokseen
'   - Tyypillisesti käytetty ennen CSV- tai ulkoista järjestelmävientii
'--------------------------------------------------------------------------------
Public Function PilkkuPiste(Luku As Variant) As String
On Error GoTo ErrorHandler
    Dim Osoitin As Long  ' Pilkun sijainti merkkijonossa
    
    ' Käsitellään null/tyhjä syöte
    If Nz(Luku) = "" Then
        PilkkuPiste = ""
        Exit Function
    End If

    ' Etsitään ja korvataan pilkku pisteellä
    Osoitin = InStr(Luku, ",")
    If Osoitin = 0 Then
        PilkkuPiste = Luku  ' Ei pilkkua, palautetaan sellaisenaan
    Else
        PilkkuPiste = Left$(Luku, Osoitin - 1) & "." & Mid$(Luku, Osoitin + 1)
    End If
    Exit Function

ErrorHandler:
    PilkkuPiste = ""
End Function

'--------------------------------------------------------------------------------
' Funktio: UdNoteToRev
' Tarkoitus: Poimii revisionumeron käyttäjämuistiinpanosta päivämäärän perusteella
'
' Parametrit:
'   UdNote - Variantti, joka sisältää merkkijonon muodossa "teksti:pvm|lisäteksti"
'
' Palauttaa:
'   Variantti – revisiokoodi _Revisions-taulusta tai Null, jos ei löydy
'
' Huomiot:
'   - Jäsentää päivämäärän UdNote-merkkijonosta (muoto "teksti:KK/PP/VVVV|...")
'   - Hakee vastaavan revision _Revisions-taulusta
'   - Palauttaa ensimmäisen revision, jossa BeforeDate > jäsennetty päivämäärä
'   - Käytetty historialliseen revisioseurantaan
'--------------------------------------------------------------------------------
Public Function UdNoteToRev(UdNote As Variant) As Variant
On Error GoTo ErrorHandler
    Dim Paiva As String  ' Päivämäärä poimittuna muistiinpanosta
    Dim Os As Long  ' Sijainti merkkijonon jäsentämisessä
    Dim VP As Date  ' Jäsennetty päivämääräarvo
    Dim RevTaul As DAO.Recordset  ' _Revisions-taulun recordset
    
    ' Käsitellään null-syöte
    If IsNull(UdNote) Then
        UdNoteToRev = Null
        Exit Function
    End If
    
    ' Jäsennetään päivämäärä muistiinpanomerkkijonosta (muoto "teksti:pvm|lisää")
    Os = InStr(UdNote, ":")
    If Os > 0 Then
        ' Extract date portion between : and |
        Paiva = Mid$(UdNote, Os + 1)
        Paiva = Left$(Paiva, InStr(Paiva, "|") - 1)
        VP = DateValue(Paiva)
        Paiva = Month(VP) & "/" & Day(VP) & "/" & Year(VP)   'Format: M/D/YYYY (e.g., 2/1/2007)
        
        ' Look up revision based on date
        Set RevTaul = CurrentDb.OpenRecordset("SELECT * FROM _Revisions WHERE (((BeforeDate) > #" & Paiva & "#)) ORDER BY BeforeDate ASC;")
        If RevTaul.RecordCount > 0 Then
            UdNoteToRev = RevTaul.Fields("Rev").Value
        Else
            UdNoteToRev = Null
        End If
        RevTaul.Close
        Set RevTaul = Nothing
    Else
        UdNoteToRev = Null
    End If
    Exit Function

ErrorHandler:
    On Error Resume Next
    If Not RevTaul Is Nothing Then RevTaul.Close
    Set RevTaul = Nothing
    On Error GoTo 0
    UdNoteToRev = Null
End Function

'--------------------------------------------------------------------------------
' Funktio: EtsiLoop
' Tarkoitus: Tarkistaa, onko looppi olemassa järjestelmässä
'
' Parametrit:
'   Alue   - Merkkijono, joka sisältää aluekoodin
'   Looppi - Merkkijono, joka sisältää loopinumeron
'
' Palauttaa:
'   Merkkijono – "1" jos looppi löytyy, "" (tyhjä) jos ei löydy
'
' Huomiot:
'   - Kyselee qrysolvalve-kyselystä AreaCode- ja LoopNo-kentillä
'   - Palauttaa yksinkertaisen olemassaolomerkin (ei boolean, yhteensopivuussyistä)
'   - Käytetty validointiin ennen uuden loopin luomista
'--------------------------------------------------------------------------------
Function EtsiLoop(Alue As String, Looppi As String) As String
On Error GoTo ErrorHandler
    Dim Taul As DAO.Recordset  ' Kyselytulokset
    
    ' Haetaan vastaava looppi
    Set Taul = CurrentDb.OpenRecordset("SELECT * From qrysolvalve WHERE AreaCode='" & Alue & "' AND LoopNo='" & Looppi & "'")
    If Taul.EOF Then
        EtsiLoop = ""  ' Ei löydy
    Else
        EtsiLoop = "1"  ' Löytyi
    End If
    Taul.Close
    Set Taul = Nothing
    Exit Function

ErrorHandler:
    On Error Resume Next
    If Not Taul Is Nothing Then Taul.Close
    Set Taul = Nothing
    On Error GoTo 0
    EtsiLoop = ""
End Function
