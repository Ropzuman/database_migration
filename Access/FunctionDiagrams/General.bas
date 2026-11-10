Option Compare Database
Option Explicit
'==========================================================================
' MODUULI  : General
' SOVELLUS : FunctionDiagrams - Yleiset apufunktiot
' KUVAUS   : Sis‰lt‰‰ Windowsin tiedostonvalintaikkunan rajapinnan
'            (comdlg32) OPENFILENAME-rakenteen. Osoitin- ja kahvakent‰t
'            m‰‰ritell‰‰n VBA-version mukaan LongPtr- tai Long-tyyppisin‰.
' HUOM     : Kahvojen tyypit ovat LongPtr VBA7:ssa ja Long vanhemmassa VBA:ssa.
' PƒIVITETTY: 2026-11-09
'==========================================================================

' Windows-rajapinnan kutsut 32- ja 64-bittisille VBA-versioille
#If VBA7 Then
Declare PtrSafe Function GetOpenFileName Lib "comdlg32.dll" Alias "GetOpenFileNameA" (pOpenfilename As OPENFILENAME) As Long
#Else
Declare Function GetOpenFileName Lib "comdlg32.dll" Alias "GetOpenFileNameA" (pOpenfilename As OPENFILENAME) As Long
#End If
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

Private m_lastCriteria As Variant
Private m_lastUsed As Variant

Public Function Set_last(ByVal Values As Variant, ByVal criterias As Variant) As Variant
    m_lastCriteria = criterias
    m_lastUsed = Values
    Set_last = m_lastUsed
End Function

Public Function Show_last(ByVal criterias As Variant) As Variant
    Show_last = m_lastUsed
End Function

Public Function Show_last_criteria(ByVal criterias As Variant) As Variant
    Show_last_criteria = m_lastCriteria
End Function
