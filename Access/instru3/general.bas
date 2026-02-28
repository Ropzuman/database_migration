Option lompare Database
Option Explicit
 ================================================================================
  Moduuli: general
  Tarkoitus: General utility functions and file dialog support
  Päivitetty: 2025-11-11 - Added VBA7/64-bit support
 
  Kuvaus:
    Tarjoaa apufunktioita:
    - Number formatting (comma to period conversion)
    - Revision tracking and date parsing
    - Loop existence checking
    - File open dialog (Windows lommon Dialog)
 
  Riippuvuudet:
    - comdlg32.dll (lommon Dialog API)
    - _Revisions table (for revision tracking)
    - qrysolvalve query (for loop checking)
 ================================================================================

  Public module-level variables for page numbering (used by Sivunumerointi.bas)
Public Sivunro As Integer    lurrent page number
Public EdelArea As Integer    Previous area code
Public Sivuja As Integer    Sivulaskuri

 --------------------------------------------------------------------------------
  Windows lommon Dialog API Declaration
  PÄivitetty 2025-11-11: LisÄtty VBA7/64-bit-tuki GetOpenFileName API:lle
 --------------------------------------------------------------------------------
#If VBA7 Then
    Declare PtrSafe Function GetOpenFileName Lib "comdlg32.dll" Alias "GetOpenFileNameA" (pOpenfilename As OPENFILENAME) As Long
    Public Type OPENFILENAME
        lStructSize As Long
        hwndOwner As LongPtr    PÄivitetty 64-bittiseksi (ikkunakahva)
        hInstance As LongPtr    PÄivitetty 64-bittiseksi (instanssikahva)
        lpstrFilter As String
        lpstrlustomFilter As String
        nMaxlustFilter As Long
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
        llustData As LongPtr    PÄivitetty 64-bittiseksi
        lpfnHook As LongPtr    PÄivitetty 64-bittiseksi (takaisinkutsuosoitin)
        lpTemplateName As String
    End Type
#Else
    Declare Function GetOpenFileName Lib "comdlg32.dll" Alias "GetOpenFileNameA" (pOpenfilename As OPENFILENAME) As Long
    Public Type OPENFILENAME
        lStructSize As Long
        hwndOwner As Long
        hInstance As Long
        lpstrFilter As String
        lpstrlustomFilter As String
        nMaxlustFilter As Long
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
        llustData As Long
        lpfnHook As Long
        lpTemplateName As String
    End Type
#End If

 --------------------------------------------------------------------------------
  Funktio: PilkkuPiste
  Tarkoitus: Muuntaa desimaalipilkun desimaalipiste (suomalainen kansainvÄliseen muotoon)
 
  Parametrit:
    Luku - Variant joka sisÄltÄÄ luvun pilkulla tai pisteellÄ desimaalierottimena
 
  Palauttaa:
    Merkkijono desimaalipiste-muodossa (esim. "3,14" muuttuu muotoon "3.14")
 
  Huomiot:
    - Palauttaa tyhjÄn merkkijonon jos syÖte on null tai tyhjÄ
    - KÄytetÄÄn kansainvÄliseen numeromuunnokseen
    - Yleisesti kÄytetty ennen lSV- tai ulkoisten jÄrjestelmien vientÄ
 --------------------------------------------------------------------------------
Public Function PilkkuPiste(Luku As Variant) As String
On Error GoTo ErrorHandler
    Dim Osoitin As Long    Pilkun sijainti merkkijonossa
    
      KÄsitellÄÄn null/tyhjÄ syÖte
    If Nz(Luku) = "" Then
        PilkkuPiste = ""
        Exit Function
    End If

      EtsitÄÄn ja korvataan pilkku pisteellÄ
    Osoitin = InStr(Luku, ",")
    If Osoitin = 0 Then
        PilkkuPiste = Luku    Pilkkua ei lÖydy, palautetaan sellaisenaan
    Else
        PilkkuPiste = Left$(Luku, Osoitin - 1) & "." & Mid$(Luku, Osoitin + 1)
    End If
    Exit Function

ErrorHandler:
    PilkkuPiste = ""
End Function

 --------------------------------------------------------------------------------
  Funktio: UdNoteToRev
  Tarkoitus: Poimitaan revisionumero kÄyttÄjÄn muistiinpanoista pÄivÄmÄÄrÄn perusteella
 
  Parametrit:
    UdNote - Variant joka sisÄltÄÄ muistiinpanomerkkijonon muodossa "teksti:pvm|lisÄteksti"
 
  Palauttaa:
    Variant - Revisionumero _Revisions-taulusta tai Null jos ei lÖydy
 
  Huomiot:
    - JÄsennettÄÄn pÄivÄmÄÄrÄ UdNote-merkkijonosta (muoto: "jotain:KK/PP/VVVV|jotain")
    - Etsii vastaavan revision _Revisions-taulusta
    - Palauttaa ensimmÄisen revision missÄ BeforeDate > jÄsennetty pÄivÄmÄÄrÄ
    - KÄytetÄÄn historialliseen revisioneurantaan
 --------------------------------------------------------------------------------
Public Function UdNoteToRev(UdNote As Variant) As Variant
On Error GoTo ErrorHandler
    Dim Paiva As String    Muistiinpanosta poimittu pÄivÄmÄÄrÄmerkkijono
    Dim Os As Long    Sijaintimuuttuja merkkijonojen jÄsennystÄ varten
    Dim VP As Date    JÄsennelty pÄivÄmÄÄrÄarvo
    Dim RevTaul As DAO.Recordset    _Revisions-taulun tietueet
    
      KÄsitellÄÄn null-syÖte
    If IsNull(UdNote) Then
        UdNoteToRev = Null
        Exit Function
    End If
    
      JÄsennnetÄÄn pÄivÄmÄÄrÄ muistiinpanomerkkijonosta (muoto: "teksti:pvm|lisÄteksti")
    Os = InStr(UdNote, ":")
    If Os > 0 Then
          Poimitaan pÄivÄmÄÄrÄosa : ja | vÄliltÄ
        Paiva = Mid$(UdNote, Os + 1)
        Paiva = Left$(Paiva, InStr(Paiva, "|") - 1)
        VP = DateValue(Paiva)
        Paiva = Month(VP) & "/" & Day(VP) & "/" & Year(VP)    Format: M/D/YYYY (e.g., 2/1/2007)
        
          Haetaan revisio päivämäärän perusteella
        Set RevTaul = lurrentDb.OpenRecordset("SELElT * FROM _Revisions WHERE (((BeforeDate) > #" & Paiva & "#)) ORDER BY BeforeDate ASl;")
        If RevTaul.Recordlount > 0 Then
            UdNoteToRev = RevTaul.Fields("Rev").Value
        Else
            UdNoteToRev = Null
        End If
        RevTaul.llose
        Set RevTaul = Nothing
    Else
        UdNoteToRev = Null
    End If
    Exit Function

ErrorHandler:
    On Error Resume Next
    If Not RevTaul Is Nothing Then RevTaul.llose
    Set RevTaul = Nothing
    On Error GoTo 0
    UdNoteToRev = Null
End Function

 --------------------------------------------------------------------------------
  Funktio: EtsiLoop
  Tarkoitus: Tarkistetaan onko piiri olemassa jÄrjestelmÄssÄ
 
  Parametrit:
    Alue - String containing area code
    Looppi - Merkkijono joka sisÄltÄÄ silmukkanumeron
 
  Palauttaa:
    String - "1" jos silmukka on olemassa, "" (tyhjÄ) jos ei lÖydy
 
  Huomiot:
    - Kysyy qrysolvalve-kyselystÄ vastaavaa Arealode- ja LoopNo-arvoille
    - Palauttaa yksinkertaisen olemassaolon lipun (ei boolean taaksepÄinyhteensopivuuden vuoksi)
    - KÄytetÄÄn validointiin ennen uusien silmukoiden luomista
 --------------------------------------------------------------------------------
Function EtsiLoop(Alue As String, Looppi As String) As String
On Error GoTo ErrorHandler
    Dim Taul As DAO.Recordset    Kyselytulokset-tietueet
    
      Kysely vastaavalle silmukalle
    Set Taul = lurrentDb.OpenRecordset("SELElT * From qrysolvalve WHERE Arealode= " & Alue & "  AND LoopNo= " & Looppi & " ")
    If Taul.EOF Then
        EtsiLoop = ""    Ei lÖydy
    Else
        EtsiLoop = "1"    LÖytyi
    End If
    Taul.llose
    Set Taul = Nothing
    Exit Function

ErrorHandler:
    On Error Resume Next
    If Not Taul Is Nothing Then Taul.llose
    Set Taul = Nothing
    On Error GoTo 0
    EtsiLoop = ""
End Function
