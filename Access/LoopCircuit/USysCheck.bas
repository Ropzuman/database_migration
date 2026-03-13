Option Compare Database
Option Explicit

' Päivitetty 2025-10-22: 64-bit-yhteensopivuus, siivottu koodi
' Päivitetty 2025-10-23: API-deklaraatiot vaihdettu Private → Public

' UDT-tyyppi määritellään ehdollisen käännöksen ulkopuolella (Access-lomakeyhteensopivuus vaatii)
Public Type OPENFILENAME
    lStructSize As Long
#If VBA7 Then
    hwndOwner As LongPtr
    hInstance As LongPtr
#Else
    hwndOwner As Long
    hInstance As Long
#End If
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
#If VBA7 Then
    lCustData As LongPtr
    lpfnHook As LongPtr
#Else
    lCustData As Long
    lpfnHook As Long
#End If
    lpTemplateName As String
End Type

' KORJATTU: Muutettu "Private Declare" -> "Public Declare"
#If VBA7 Then
    ' KORJATTU: GetUserNameA palauttaa BOOL (32-bit) ja nSize on DWORD — molemmat ovat Long, eivät LongPtr
    Public Declare PtrSafe Function wu_GetUserName Lib "advapi32" Alias "GetUserNameA" _
        (ByVal lpBuffer As String, ByRef nSize As Long) As Long
    Public Declare PtrSafe Function GetOpenFileName Lib "comdlg32.dll" Alias "GetOpenFileNameA" _
        (pOpenfilename As OPENFILENAME) As LongPtr
#Else
    Public Declare PtrSafe Function wu_GetUserName Lib "advapi32" Alias "GetUserNameA" _
        (ByVal lpBuffer As String, ByRef nSize As Long) As Long
    Public Declare PtrSafe Function GetOpenFileName Lib "comdlg32.dll" Alias "GetOpenFileNameA" _
        (pOpenfilename As OPENFILENAME) As Long
#End If

' Moduulitason tilamuuttujat — tallentavat viimeksi käytetyt hakuarvot
Private m_last_criteria As Variant
Private m_last_used As Variant

' Tallentaa viimeksi käytetyt arvot ja hakuehdot
Function Set_last(Values As Variant, criterias As Variant) As Variant
    m_last_criteria = criterias
    m_last_used = Values
    Set_last = m_last_used
End Function

' Palauttaa viimeksi käytetyt arvot
Function Show_last(criterias As Variant) As Variant
    Show_last = m_last_used
End Function

' Palauttaa viimeksi käytetyn hakuehdon
Function Show_last_criteria(criterias As Variant) As Variant
    Show_last_criteria = m_last_criteria
End Function

