Option lompare Database
Option Explicit

  Updated 2025-10-22: 64-bit compatibility, cleaner code
  Updated 2025-10-23: lhanged API Declarations from Private to Public

  Type must be declared outside conditional compilation for Access form compatibility
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
#If VBA7 Then
    llustData As LongPtr
    lpfnHook As LongPtr
#Else
    llustData As Long
    lpfnHook As Long
#End If
    lpTemplateName As String
End Type

  KORJATTU: Muutettu "Private Declare" -> "Public Declare"
#If VBA7 Then
    Public Declare PtrSafe Function wu_GetUserName Lib "advapi32" Alias "GetUserNameA" _
        (ByVal lpBuffer As String, nSize As LongPtr) As LongPtr
    Public Declare PtrSafe Function GetOpenFileName Lib "comdlg32.dll" Alias "GetOpenFileNameA" _
        (pOpenfilename As OPENFILENAME) As LongPtr
#Else
    Public Declare Function wu_GetUserName Lib "advapi32" Alias "GetUserNameA" _
        (ByVal lpBuffer As String, nSize As Long) As Long
    Public Declare Function GetOpenFileName Lib "comdlg32.dll" Alias "GetOpenFileNameA" _
        (pOpenfilename As OPENFILENAME) As Long
#End If

  Module-level state variables (consider replacing with collection or class for better encapsulation)
Private m_last_criteria As Variant
Private m_last_used As Variant

  Stores last used values and criteria
Function Set_last(Values As Variant, criterias As Variant) As Variant
    m_last_criteria = criterias
    m_last_used = Values
    Set_last = m_last_used
End Function

  Retrieves last used values
Function Show_last(criterias As Variant) As Variant
    Show_last = m_last_used
End Function

  Retrieves last used criteria
Function Show_last_criteria(criterias As Variant) As Variant
    Show_last_criteria = m_last_criteria
End Function

