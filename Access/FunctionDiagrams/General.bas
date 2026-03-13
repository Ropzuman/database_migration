Option Compare Database
Option Explicit
'==========================================================================
' MODUULI  : General
' SOVELLUS : FunctionDiagrams — Yleiset apufunktiot
' KUVAUS   : Sisältää Windows API -julistukset käyttäjänimen hakemiseen
'            (advapi32) sekä tiedostonvalintaikkunan avaamiseen (comdlg32).
'            OPENFILENAME-rakenne on päivitetty 64-bittiseksi: kaikki
'            osoitin- ja kahvakentät ovat LongPtr-tyyppiä. Moduulitasoiset
'            muuttujat last_criteria ja last_used tallentavat viimeksi
'            käytetyn hakuarvon lomakkeiden välillä.
' DIPLE    : Global → Public As Variant; PtrSafe + LongPtr läpi koko tyypin.
' PÄIVITETTY: 2026-03-03
'==========================================================================

' Kirjautuneen käyttäjän poiminta Windows-rajapinnasta
Declare PtrSafe Function wu_GetUserName Lib "advapi32" Alias "GetUserNameA" (ByVal lpBuffer As String, ByRef nSize As Long) As Long

' --------- [ VALITSE TIEDOSTO -ikkuna ] -----------------
Declare PtrSafe Function GetOpenFileName Lib "comdlg32.dll" Alias "GetOpenFileNameA" (pOpenfilename As OPENFILENAME) As Long
Public Type OPENFILENAME
    lStructSize As Long
    hwndOwner As LongPtr
    hInstance As LongPtr
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
    lCustData As LongPtr
    lpfnHook As LongPtr
    lpTemplateName As String
End Type

' Edellinen hakuehto ja arvo taltioidaan moduulitasolla
Public last_criteria As Variant
Public last_used As Variant

Function Set_last(Values As Variant, criterias As Variant) As Variant
    ' Talletetaan viimeksi käytetty hakuarvo ja ehto muistiin
    last_criteria = criterias
    last_used = Values
    Set_last = last_used
End Function

Function Show_last(criterias As Variant) As Variant
    ' Palauttaa viimeksi käytetyn hakuarvon
    Show_last = last_used
End Function

Function Show_last_criteria(criterias As Variant) As Variant
    ' Palauttaa viimeksi käytetyn hakuehdon
    Show_last_criteria = last_criteria
End Function

