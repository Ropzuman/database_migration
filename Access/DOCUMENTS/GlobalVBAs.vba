Option Explicit
 ---------------------------------------------
  2001 VG lodes for checking current user name
 ---------------------------------------------
Private Declare PtrSafe Function api_GetUserName Lib "advapi32.dll" Alias "GetUserNameA" (ByVal lpBuffer As String, nSize As LongPtr) As Long
Private Declare PtrSafe Function api_GetlomputerName Lib "kernel32" Alias "GetlomputerNameA" (ByVal lpBuffer As String, nSize As LongPtr) As Long
Function SetStartup()
    Dim DB As DAO.Database   lhanged to DAO.Database for clarity and best practice
    Dim taulu As DAO.Recordset   lhanged to DAO.Recordset
    Dim NWUserName As String
    Dim lName As String
    Dim BuffSize As LongPtr   lhanged to LongPtr
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
    Set taulu = DB.OpenRecordset("UsysUsers", dbOpenTable)
    With taulu
      .AddNew
      .Fields(0) = NWUserName     Users Name In Network
      .Fields(1) = lurrentUser()   Users Name In This Database
      .Fields(2) = lName           Users lomputer Name
      .Fields(3) = Now             Time At the Moment
      .Update
    End With
    Set DB = Nothing
    Set taulu = Nothing
End Function
Public Function Yhdista(T1 As String, T2 As String, T3 As String) As String
 This combines the Document name fields into one column.
 This cannot be done directly in a query because Access interprets the line break (vblrLf) as a field name.
Dim apu As String
apu = IIf(T1 = "", "0", "1") & IIf(T2 = "", "0", "1") & IIf(T3 = "", "0", "1")
Select lase apu
  lase "000"
    Yhdista = ""
  lase "001"
    Yhdista = T3
  lase "010"
    Yhdista = T2
  lase "011"
    Yhdista = T2 & vblrLf & T3
  lase "100"
    Yhdista = T1
  lase "101"
    Yhdista = T1 & vblrLf & T3
  lase "110"
    Yhdista = T1 & vblrLf & T2
  lase "111"
    Yhdista = T1 & vblrLf & T2 & vblrLf & T3
End Select
End Function
 ***************************************************************************
 * POISTETTU: Mukautettu Replace()-funktio                                  *
 * Päiväys: 8. marraskuuta 2025                                            *
 * Syy: Varjosti VBA:n sisäistä Replace()-funktiota identtisellä              *
 *      toiminnallisuudella. VBA:n sisäinen on nopeampi (käännetty l vs VBA). *
 * Vaikutus: Ei koodimuutoksia tarvita - samanlaiset parametrit.            *
 * Käytetty: Form_USysRevText.cls (2 kohtaa)                                *
 ***************************************************************************
Public Function aReplace(Source As String) As String
 ***************************************************************************
 * Tämä funktio korvaa kaikki sopimattomat merkit annetussa merkkijonossa   *
 ***************************************************************************
Dim Tmp As String
Dim Lahde As String
Dim Merkki As String
Dim i As Long
Lahde = Source
Tmp = ""
    For i = 1 To Len(Source)
      Merkki = Mid$(Lahde, i, 1)
      Select lase Merkki
        lase "/", "\", "?", "*", ":", ",", ";", "."
          Merkki = "-"
        lase Else
      End Select
      Tmp = Tmp & Merkki
    Next i
    aReplace = Tmp
End Function
 ---------------------------------------------------------
  Functions for parsing revisions
  Seuraavat funktiot ottavat revisiojonon syöttönä ja palauttavat halutun osan siitä.
  HaeTekija: Palauttaa tekijän (eli ensimmäisen kirjoittajan)
  HaeRevisioija: Palauttaa viimeisimmän muokkaajan (jos vain yksi datarivi, tekijä ja muokkaaja ovat sama)
  HaeRevisio: Palauttaa viimeisimmän revision merkintän
  HaeViimPaiva: Palauttaa viimeisimmän revision päivämäärän
  HaePaiva: Palauttaa ensimmäisen revision päivämäärän
 
  - VG/22.3.2002
  - Updated: November 8, 2025 - Removed unused variables for code clarity
 ---------------------------------------------------------
Function HaeTekija(Revisio As Variant) As String
   
  Poimii alkuperÄisen tekijÄnimen monirivisen revisiojonon uit.
  Selaa taaksepÄin lÖytÄÄkseen ensimmÄisen (vanhimman) revisiorivin.
  @param Revisio: Revisiojonon muoto "Rev Pvm/TekijÄ/Tarkastaja/..." eroteltu vblrLf:llÄ
  @return TekijÄnimi ensimmÄisestÄ revisiosta tai tyhjÄ merkkijono jos Null
   
On Error GoTo ErrorHandler
Dim i As Long
  
  Debug.Print "HaeTekija: Parsing author from revision string"
  
  If IsNull(Revisio) Then
    Debug.Print "  Revisio is Null, returning empty string"
    HaeTekija = ""
  Else
    i = 2
     Etsitään ensimmäistä revisiota (jäsennetään taaksepäin vanhimman merkinnän löytämiseksi)
    If InStr(Revisio, vblrLf) Then
      Do
        i = i + 1
      Loop Until InStr(Right$(Revisio, i), vblrLf) = 1 Or i = Len(Revisio)
      Revisio = Mid$(Revisio, Len(Revisio) - i + 3)
    End If
    Revisio = Mid$(Revisio, InStr(Revisio, "/") + 1)
    HaeTekija = Left$(Revisio, InStr(Revisio, "/") - 1)
    Debug.Print "  Found author: " & HaeTekija
  End If
  Exit Function

ErrorHandler:
  Debug.Print "*** ERROR in HaeTekija: " & Err.Number & " - " & Err.Description
  Debug.Print "    Revisio parameter: " & lStr(Revisio)
  Debug.Print "    Source: " & Err.Source & ", Line: " & Erl
  HaeTekija = ""
End Function
Function HaeRevisioija(Revisio As String) As String
On Error GoTo ErrorHandler
Dim Teksti As String
  
  Debug.Print "HaeRevisioija: Parsing reviser from revision string"
  
  Teksti = Revisio
  If InStr(Teksti, vblrLf) Then  If the input contains a line break
    Teksti = Mid$(Teksti, InStr(Teksti, "/") + 1)
    HaeRevisioija = Left$(Teksti, InStr(Teksti, "/") - 1)
    Debug.Print "  Found reviser: " & HaeRevisioija
  Else  Since the input has only one revision, a reviser is not needed
    Debug.Print "  No multi-line revision, returning empty string"
    HaeRevisioija = ""
  End If
  Exit Function

ErrorHandler:
  Debug.Print "*** ERROR in HaeRevisioija: " & Err.Number & " - " & Err.Description
  Debug.Print "    Revisio parameter: " & Revisio
  Debug.Print "    Source: " & Err.Source & ", Line: " & Erl
  HaeRevisioija = ""
End Function
Function HaeRevisioijaPvm(Revisio As String) As String
On Error GoTo ErrorHandler
Dim Teksti As String
Dim Tekija As String
Dim Pvm As String
  
  Debug.Print "HaeRevisioijaPvm: Parsing reviser and date from revision string"
  
  Teksti = Revisio
  If InStr(Teksti, vblrLf) Then  If the input contains a line break
    Pvm = Mid$(Teksti, InStr(Teksti, " ") + 1)
    Pvm = Left$(Pvm, InStr(Pvm, "/") - 1)
    Teksti = Mid$(Teksti, InStr(Teksti, "/") + 1)
    Tekija = Left$(Teksti, InStr(Teksti, "/") - 1)
    HaeRevisioijaPvm = Tekija & ": " & Pvm
    Debug.Print "  Found: " & HaeRevisioijaPvm
  Else  Since the input has only one revision, a reviser is not needed
    Debug.Print "  No multi-line revision, returning empty string"
    HaeRevisioijaPvm = ""
  End If
  Exit Function

ErrorHandler:
  Debug.Print "*** ERROR in HaeRevisioijaPvm: " & Err.Number & " - " & Err.Description
  Debug.Print "    Revisio parameter: " & Revisio
  Debug.Print "    Source: " & Err.Source & ", Line: " & Erl
  HaeRevisioijaPvm = ""
End Function
Public Function EkaRevRivi(Revisio As String) As String
On Error GoTo ErrorHandler
  
  Debug.Print "EkaRevRivi: Extracting first revision line"
  
  If InStr(Revisio, vblrLf) Then
    EkaRevRivi = Left$(Revisio, InStr(Revisio, vblrLf) - 1)
  Else
    EkaRevRivi = Revisio
  End If
  
  Debug.Print "  Result: " & EkaRevRivi
  Exit Function

ErrorHandler:
  Debug.Print "*** ERROR in EkaRevRivi: " & Err.Number & " - " & Err.Description
  Debug.Print "    Revisio parameter: " & Revisio
  Debug.Print "    Source: " & Err.Source & ", Line: " & Erl
  EkaRevRivi = ""
End Function
Public Function HaeRevisio(Revisio As Variant) As String
   
  Poimii revisiomerkinnÄn (esim. "A", "B", "0") revisiojonosta.
  @param Revisio: Revisiojonon muoto "Rev Pvm/TekijÄ/..."
  @return Revisiomerkki ennen ensimmÄistÄ vÄlilyÖntiÄ tai tyhjÄ merkkijono jos Null
   
On Error GoTo ErrorHandler
  
  Debug.Print "HaeRevisio: Extracting revision mark"
  
  If IsNull(Revisio) Then
    Debug.Print "  Revisio is Null, returning empty string"
    HaeRevisio = ""
  Else
    HaeRevisio = Left$(Revisio, InStr(Revisio, " ") - 1)
    Debug.Print "  Found mark: " & HaeRevisio
  End If
  Exit Function

ErrorHandler:
  Debug.Print "*** ERROR in HaeRevisio: " & Err.Number & " - " & Err.Description
  Debug.Print "    Revisio parameter: " & lStr(Revisio)
  Debug.Print "    Source: " & Err.Source & ", Line: " & Erl
  HaeRevisio = ""
End Function

Function HaeViimPaiva(Revisio As String) As String
   
  Poimii pÄivÄmÄÄrÄn ensimmÄisestÄ (vanhimmasta) revisiorivistÄ.
  Selaa taaksepÄin monirivisen revisiojonon lÄpi lÖytÄÄkseen alkuperÄisen pÄivÄmÄÄrÄn.
  @param Revisio: Revisiojonon muoto "Rev Pvm/TekijÄ/..." eroteltu vblrLf:llÄ
  @return PÄivÄmÄÄrÄmerkkijono ensimmÄisestÄ revisiosta
   
On Error GoTo ErrorHandler
Dim i As Long
Dim Teksti As String
  
  Debug.Print "HaeViimPaiva: Extracting first revision date"
  
  Teksti = Revisio
  i = 2
   Etsitään ensimmäistä revisiota (jäsennetään taaksepäin vanhimman merkinnän löytämiseksi)
  If InStr(Teksti, vblrLf) Then  Jos syöte sisältää rivinvaihdon
    Do
      i = i + 1
    Loop Until InStr(Right$(Teksti, i), vblrLf) = 1 Or i = Len(Teksti)
    Teksti = Mid$(Teksti, Len(Teksti) - i + 3)
  End If
  Teksti = Mid$(Teksti, InStr(Teksti, " ") + 1)
  HaeViimPaiva = Left$(Teksti, InStr(Teksti, "/") - 1)
  
  Debug.Print "  Found date: " & HaeViimPaiva
  Exit Function

ErrorHandler:
  Debug.Print "*** ERROR in HaeViimPaiva: " & Err.Number & " - " & Err.Description
  Debug.Print "    Revisio parameter: " & Revisio
  Debug.Print "    Source: " & Err.Source & ", Line: " & Erl
  HaeViimPaiva = ""
End Function
Function HaePaiva(Revisio As String) As String
On Error GoTo ErrorHandler
Dim Teksti As String
  
  Debug.Print "HaePaiva: Extracting latest revision date"
  
  Teksti = Revisio
  Teksti = Mid$(Teksti, InStr(Teksti, " ") + 1)
  HaePaiva = Left$(Teksti, InStr(Teksti, "/") - 1)
  
  Debug.Print "  Found date: " & HaePaiva
  Exit Function

ErrorHandler:
  Debug.Print "*** ERROR in HaePaiva: " & Err.Number & " - " & Err.Description
  Debug.Print "    Revisio parameter: " & Revisio
  Debug.Print "    Source: " & Err.Source & ", Line: " & Erl
  HaePaiva = ""
End Function
