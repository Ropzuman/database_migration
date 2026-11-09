Option Compare Database
Option Explicit
'==========================================================================
' MODUULI  : Sivunumerointi
' SOVELLUS : FunctionDiagrams - Raporttien sivunumerointi
' KUVAUS   : Julkinen Sivu()-funktio laskee sivunumerot aluepohjaisesti:
'            kun alueen (ALUE) arvo muuttuu, laskuri nollautuu. Yhdelle
'            sivulle mahtuu enint‰‰n kolme merkint‰‰ (Sivuja = 3).
'            Moduulitason tilamuuttujat (Private) varmistavat, ett‰ tila
'            s‰ilyy kutsujen v‰lill‰.
' HUOM    : CInt korvattu CLng-muunnoksella tyyppiristiriidan korjaamiseksi (EdelArea As Long).
' PƒIVITETTY: 2026-03-03
'==========================================================================

' --- Moduulitason tilamuuttujat sivunumerointia varten ---
Private EdelArea As Long    ' Edellinen alue (nollaus tunnistusta varten)
Private Sivunro  As Long    ' Nykyinen sivunumero
Private Sivuja   As Long    ' Sivujen lukum‰‰r‰ nykyisell‰ alueella

Public Function Sivu(ALUE As String) As String
  Dim AlueNro As Long
  ' Val ei kaadu ei-numeeriseen arvoon, toisin kuin CLng
  AlueNro = CLng(Val(ALUE))
  ' CLng k‰ytet‰‰n CInt:n sijaan, koska EdelArea on Long-tyyppi‰ ó tyypit t‰sm‰‰v‰t
  If EdelArea = 0 Or Sivunro = 0 Then
    Sivunro = 1
    Sivuja = 0
  ElseIf EdelArea <> AlueNro Then
    Sivunro = 1
    Sivuja = 0
  End If
  If Sivuja = 3 Then
    Sivunro = Sivunro + 1
    Sivuja = 0
  End If
  Sivuja = Sivuja + 1
  EdelArea = AlueNro
  Sivu = CStr(Sivunro)
End Function
