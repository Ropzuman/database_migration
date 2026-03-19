Option Explicit
'==========================================================================
' MODUULI  : GlobalVBAs
' SOVELLUS : DOCUMENTS — Yleiset apufunktiot
' KUVAUS   : Käynnistyskirjaus (SetStartup), merkkijono-operaatiot
'            (Yhdista, aReplace) sekä revisiotietojen parsintafunktiot
'            (HaeTekija, HaeRevisioija, HaeRevisio, jne.).
'            Windows-kutsut on päivitetty 64-bittisiksi: nSize on ByRef Long,
'            koska API kirjoittaa DWORD-arvoon (4 tavua), ei 8-tavuiseen LongPtr:iin.
' PÄIVITETTY: 2026-03-03
'==========================================================================

' Windows-rajapinnan kutsut käyttäjänimen ja koneen nimen hakemiseen
' nSize on LPDWORD (osoitin 32-bittiseen DWORD:iin) — ByRef Long, EI LongPtr
#If VBA7 Then
    Private Declare PtrSafe Function api_GetUserName Lib "advapi32.dll" Alias "GetUserNameA" (ByVal lpBuffer As String, ByRef nSize As Long) As Long
    Private Declare PtrSafe Function api_GetComputerName Lib "kernel32" Alias "GetComputerNameA" (ByVal lpBuffer As String, ByRef nSize As Long) As Long
#Else
    Private Declare Function api_GetUserName Lib "advapi32.dll" Alias "GetUserNameA" (ByVal lpBuffer As String, ByRef nSize As Long) As Long
    Private Declare Function api_GetComputerName Lib "kernel32" Alias "GetComputerNameA" (ByVal lpBuffer As String, ByRef nSize As Long) As Long
#End If

Private Const DB_OPEN_TABLE As Long = 1

Public Function SetStartup() As Boolean
  Dim DB As DAO.Database
  Dim taulu As DAO.Recordset
  Dim NWUserName As String
  Dim CName As String
  Dim BuffSize As Long    ' LPDWORD-yhteensopiva: 32-bittinen arvo
  Dim NBuffer As String

  On Error GoTo ErrorHandler

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

  ' AutoExec ei saa kaatua, vaikka lokitaulu puuttuisi tai olisi lukittu.
  If TauluOnOlemassa(DB, "UsysUsers") Then
    Set taulu = DB.OpenRecordset("UsysUsers", dbOpenDynaset)
    With taulu
      .AddNew
      .Fields(0) = NWUserName     ' Verkkokäyttäjänimi
      .Fields(1) = CurrentUser()   ' Tietokantakäyttäjänimi
      .Fields(2) = CName           ' Koneen nimi
      .Fields(3) = Now             ' Kirjautumishetki
      .Update
    End With
  End If

Cleanup:
  On Error Resume Next
  If Not taulu Is Nothing Then
    taulu.Close
    Set taulu = Nothing
  End If
  Set DB = Nothing
  On Error GoTo 0
  SetStartup = True
  Exit Function

ErrorHandler:
  ' Käynnistyksen pitää jatkua myös virhetilanteissa.
  Resume Cleanup
End Function

Private Function TauluOnOlemassa(ByVal DB As DAO.Database, ByVal TaulunNimi As String) As Boolean
  Dim T As DAO.TableDef

  On Error GoTo ErrorHandler
  For Each T In DB.TableDefs
    If StrComp(T.Name, TaulunNimi, vbTextCompare) = 0 Then
      TauluOnOlemassa = True
      Exit Function
    End If
  Next T

  TauluOnOlemassa = False
  Exit Function

ErrorHandler:
  TauluOnOlemassa = False
End Function
Public Function Yhdista(T1 As String, T2 As String, T3 As String) As String
' Yhdistää dokumentin nimikenttä yhdeksi sarakkeeksi.
' Tätä ei voi tehdä suoraan kyselyssä, koska Access tulkitsee
' CrLf-merkkiä kenttänimenee eikä rivinvaihtona.
Dim apu As String
apu = IIf(T1 = "", "0", "1") & IIf(T2 = "", "0", "1") & IIf(T3 = "", "0", "1")
Select Case apu
  Case "000"
    Yhdista = ""
  Case "001"
    Yhdista = T3
  Case "010"
    Yhdista = T2
  Case "011"
    Yhdista = T2 & vbCrLf & T3
  Case "100"
    Yhdista = T1
  Case "101"
    Yhdista = T1 & vbCrLf & T3
  Case "110"
    Yhdista = T1 & vbCrLf & T2
  Case "111"
    Yhdista = T1 & vbCrLf & T2 & vbCrLf & T3
End Select
End Function
'***************************************************************************
'* REMOVED: Custom Replace() function                                       *
'* Date: November 8, 2025                                                  *
'* Reason: Shadowed VBA's built-in Replace() function with identical       *
'*         functionality. VBA built-in is faster (compiled C vs VBA loop). *
'* Impact: No code changes needed - built-in has same signature.           *
'* Used in: Form_USysRevText.cls (2 locations)                             *
'***************************************************************************
Public Function aReplace(Source As String) As String
'***************************************************************************
'* Korvaa merkkijonosta kaikki tiedostonimeen sopimattomat erikoismerkit  *
'* väliviivalla ("-").                                                     *
'***************************************************************************
Dim Tmp As String
Dim Lahde As String
Dim Merkki As String
Dim i As Long
Lahde = Source
Tmp = ""
    For i = 1 To Len(Source)
      Merkki = Mid$(Lahde, i, 1)
      Select Case Merkki
        Case "/", "\", "?", "*", ":", ",", ";", "."
          Merkki = "-"
        Case Else
      End Select
      Tmp = Tmp & Merkki
    Next i
    aReplace = Tmp
End Function
'---------------------------------------------------------
' Revisiotekstin parsintafunktiot
' Kukin funktio ottaa syötteeksi revisionotaation ja palauttaa halutun osan.
'   HaeTekija      : palauttaa alkuperäisen tekijän (vanhin revisio)
'   HaeRevisioija  : palauttaa uusimman tarkistajan (jos vain yksi rivi, sama kuin tekijä)
'   HaeRevisio     : palauttaa uusimman revision tunnusmerkin (esim. "A", "B")
'   HaeViimPaiva   : palauttaa vanhimman revision päivämäärän
'   HaePaiva       : palauttaa uusimman revision päivämäärän
'
' - VG/22.3.2002
' - Päivitetty: 8.11.2025 — poistettu käyttämättömät muuttujat
'---------------------------------------------------------
Function HaeTekija(Revisio As Variant) As String
' Etsii alkuperäisen tekijän monirivisen revisiojonon vanhimmasta riviltä.
' Parsii merkkijonon takaperinjuurin löytääkseen ensimmäisen revision.
On Error GoTo ErrorHandler
Dim i As Long
  If IsNull(Revisio) Then
    HaeTekija = ""
  Else
    i = 2
    ' Etsitään ensimmäinen (vanhin) revisiorivi parsimalla merkkijonoa lopusta alkuun
    If InStr(Revisio, vbCrLf) Then
      Do
        i = i + 1
      Loop Until InStr(Right$(Revisio, i), vbCrLf) = 1 Or i = Len(Revisio)
      Revisio = Mid$(Revisio, Len(Revisio) - i + 3)
    End If
    Revisio = Mid$(Revisio, InStr(Revisio, "/") + 1)
    HaeTekija = Left$(Revisio, InStr(Revisio, "/") - 1)
  End If
  Exit Function

ErrorHandler:
  HaeTekija = ""
End Function
Function HaeRevisioija(Revisio As String) As String
On Error GoTo ErrorHandler
Dim Teksti As String
  Teksti = Revisio
  If InStr(Teksti, vbCrLf) Then ' Syöte sisältää rivinvaihdon: on useampi revisio
    Teksti = Mid$(Teksti, InStr(Teksti, "/") + 1)
    HaeRevisioija = Left$(Teksti, InStr(Teksti, "/") - 1)
  Else ' Vain yksi revisiorivi — tarkistajaa ei tarvita
    HaeRevisioija = ""
  End If
  Exit Function

ErrorHandler:
  HaeRevisioija = ""
End Function
Function HaeRevisioijaPvm(Revisio As String) As String
On Error GoTo ErrorHandler
Dim Teksti As String
Dim Tekija As String
Dim Pvm As String
  Teksti = Revisio
  If InStr(Teksti, vbCrLf) Then ' Syöte sisältää rivinvaihdon: on useampi revisio
    Pvm = Mid$(Teksti, InStr(Teksti, " ") + 1)
    Pvm = Left$(Pvm, InStr(Pvm, "/") - 1)
    Teksti = Mid$(Teksti, InStr(Teksti, "/") + 1)
    Tekija = Left$(Teksti, InStr(Teksti, "/") - 1)
    HaeRevisioijaPvm = Tekija & ": " & Pvm
  Else ' Vain yksi revisiorivi — tarkistajaa ei tarvita
    HaeRevisioijaPvm = ""
  End If
  Exit Function

ErrorHandler:
  HaeRevisioijaPvm = ""
End Function
Public Function EkaRevRivi(Revisio As String) As String
' Palauttaa ensimmäisen (uusimman) revisioner rivin monirivisen revisiojonosta.
On Error GoTo ErrorHandler
  If InStr(Revisio, vbCrLf) Then
    EkaRevRivi = Left$(Revisio, InStr(Revisio, vbCrLf) - 1)
  Else
    EkaRevRivi = Revisio
  End If
  
  Exit Function

ErrorHandler:
  EkaRevRivi = ""
End Function
Public Function HaeRevisio(Revisio As Variant) As String
' Palauttaa revision tunnusmerkin (esim. "A", "B", "0") revisiojonosta.
On Error GoTo ErrorHandler
  If IsNull(Revisio) Then
    HaeRevisio = ""
  Else
    HaeRevisio = Left$(Revisio, InStr(Revisio, " ") - 1)
  End If
  Exit Function

ErrorHandler:
  HaeRevisio = ""
End Function

Function HaeViimPaiva(Revisio As String) As String
' Palauttaa vanhimman revision päivämäärän parsimalla monirivinen revisijono lopusta alkuun.
On Error GoTo ErrorHandler
Dim i As Long
Dim Teksti As String
  Teksti = Revisio
  i = 2
  ' Etsitään ensimmäinen (vanhin) revisiorivi parsimalla takaperinjuurin
  If InStr(Teksti, vbCrLf) Then ' Syöte sisältää rivinvaihdon
    Do
      i = i + 1
    Loop Until InStr(Right$(Teksti, i), vbCrLf) = 1 Or i = Len(Teksti)
    Teksti = Mid$(Teksti, Len(Teksti) - i + 3)
  End If
  Teksti = Mid$(Teksti, InStr(Teksti, " ") + 1)
  HaeViimPaiva = Left$(Teksti, InStr(Teksti, "/") - 1)
  
  Exit Function

ErrorHandler:
  HaeViimPaiva = ""
End Function
Function HaePaiva(Revisio As String) As String
' Palauttaa uusimman (viimeisimmän) revision päivämäärän.
On Error GoTo ErrorHandler
Dim Teksti As String
  Teksti = Revisio
  Teksti = Mid$(Teksti, InStr(Teksti, " ") + 1)
  HaePaiva = Left$(Teksti, InStr(Teksti, "/") - 1)
  
  Exit Function

ErrorHandler:
  HaePaiva = ""
End Function
