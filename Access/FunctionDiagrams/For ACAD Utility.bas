Option Compare Database
Option Explicit
'==========================================================================
' MODUULI  : For ACAD Utility
' SOVELLUS : FunctionDiagrams — AutoCAD-apurakenteet
' KUVAUS   : Määrittelee moduulitason tietotyypit ja Windows-rajapinnan
'            API-kutsut AutoCAD-yhteyden koordinaattilaskentaa varten.
'            iPoint-rakenne pitää sisällään XYZ-koordinaatit blokkien
'            insertointipaikkojen hallintaan. GetCursorPos-kutsu mahdollistaa
'            hiiren sijainnin hakemisen Windows GDI -rajapinnasta.
' DIPLE    : 64-bittinen siirto (PtrSafe-julistukset), M365-yhteydensp.
' PÄIVITETTY: 2026-03-03
'==========================================================================

' Insertointipaikkarakenne AutoCAD-koordinaateille (X, Y, Z)
Public Type iPoint
  Pisteet(2) As Double
End Type
Public Paikat() As iPoint

' Hiiren kursorin sijainnin määrittäminen Windows-rajapinnasta
Private Type POINTAPI
    X As Long
    Y As Long
End Type
' 64-bittinen PtrSafe-julistus hiiren sijainnille
Declare PtrSafe Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long

