Option Compare Database
Option Explicit

' OPENFILENAME-tyyppi täytyy esitellä ehtokääntämislohkon ulkopuolella Access-lomakkeiden yhteensopivuuden vuoksi
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
' (nSize: LongPtr → ByRef Long — Win32 DWORD on 32-bittinen, ei osoitinkokoinen)
#If VBA7 Then
    Public Declare PtrSafe Function wu_GetUserName Lib "advapi32" Alias "GetUserNameA" _
        (ByVal lpBuffer As String, ByRef nSize As Long) As Long
    Public Declare PtrSafe Function GetOpenFileName Lib "comdlg32.dll" Alias "GetOpenFileNameA" _
        (pOpenfilename As OPENFILENAME) As LongPtr
#Else
    Public Declare Function wu_GetUserName Lib "advapi32" Alias "GetUserNameA" _
        (ByVal lpBuffer As String, ByRef nSize As Long) As Long
    Public Declare Function GetOpenFileName Lib "comdlg32.dll" Alias "GetOpenFileNameA" _
        (pOpenfilename As OPENFILENAME) As Long
#End If

' Moduulitason tilamuuttujat viimeisimmälle hakuehdolle
Private m_last_criteria As Variant
Private m_last_used As Variant

' Tallentaa viimeksi käytetyn arvon ja hakuehdon
Function Set_last(Values As Variant, criterias As Variant) As Variant
    m_last_criteria = criterias
    m_last_used = Values
    Set_last = m_last_used
End Function

' Palauttaa viimeksi käytetyn arvon
Function Show_last(criterias As Variant) As Variant
    Show_last = m_last_used
End Function

' Palauttaa viimeksi käytetyn hakuehdon
Function Show_last_criteria(criterias As Variant) As Variant
    Show_last_criteria = m_last_criteria
End Function

