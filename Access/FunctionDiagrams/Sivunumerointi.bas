Option Compare Database
Option Explicit
'==========================================================================
' MODUULI  : Sivunumerointi
' SOVELLUS : FunctionDiagrams — Raporttien sivunumerointi
' KUVAUS   : Julkinen Sivu()-funktio laskee sivunumerot aluepohjaisesti:
'            kun alueen (ALUE) arvo muuttuu, laskuri nollautuu. Yhdelle
'            sivulle mahtuu enintään kolme merkintää (Sivuja = 3).
'            Moduulitason tilamuuttujat (Private) varmistavat, että tila
'            säilyy kutsujen välillä.
' DIPLE    : CInt → CLng tyyppiristiriidan korjaukseksi (EdelArea As Long).
' PÄIVITETTY: 2026-03-03
'==========================================================================

' --- Moduulitason tilamuuttujat sivunumerointia varten ---
Private EdelArea As Long    ' Edellinen alue (nollaus tunnistusta varten)
Private Sivunro  As Long    ' Nykyinen sivunumero
Private Sivuja   As Long    ' Sivujen lukumäärä nykyisellä alueella

Public Function Sivu(ALUE As String) As String
  ' CLng käytetään CInt:n sijaan, koska EdelArea on Long-tyyppiä — tyypit täsmäävät
  If EdelArea = 0 Or Sivunro = 0 Then
    Sivunro = 1
    Sivuja = 0
  ElseIf EdelArea <> CLng(ALUE) Then
    Sivunro = 1
    Sivuja = 0
  End If
  If Sivuja = 3 Then
    Sivunro = Sivunro + 1
    Sivuja = 0
  End If
  Sivuja = Sivuja + 1
  EdelArea = CLng(ALUE)
  Sivu = CStr(Sivunro)
End Function
