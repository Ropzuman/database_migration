Attribute VB_Name = "USysCheck"
Option Compare Database
Option Explicit

' Updated 2025-10-22: 64-bit compatibility, cleaner code

#If VBA7 Then
    Private Declare PtrSafe Function wu_GetUserName Lib "advapi32" Alias "GetUserNameA" _
        (ByVal lpBuffer As String, nSize As LongPtr) As LongPtr
    Private Declare PtrSafe Function GetOpenFileName Lib "comdlg32.dll" Alias "GetOpenFileNameA" _
        (pOpenfilename As OPENFILENAME) As LongPtr
    
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
#Else
    Private Declare Function wu_GetUserName Lib "advapi32" Alias "GetUserNameA" _
        (ByVal lpBuffer As String, nSize As Long) As Long
    Private Declare Function GetOpenFileName Lib "comdlg32.dll" Alias "GetOpenFileNameA" _
        (pOpenfilename As OPENFILENAME) As Long
    
    Public Type OPENFILENAME
        lStructSize As Long
        hwndOwner As Long
        hInstance As Long
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
        lCustData As Long
        lpfnHook As Long
        lpTemplateName As String
    End Type
#End If

' Module-level state variables (consider replacing with collection or class for better encapsulation)
Private m_last_criteria As Variant
Private m_last_used As Variant

' Stores last used values and criteria
Function Set_last(Values As Variant, criterias As Variant) As Variant
    m_last_criteria = criterias
    m_last_used = Values
    Set_last = m_last_used
End Function

' Retrieves last used values
Function Show_last(criterias As Variant) As Variant
    Show_last = m_last_used
End Function

' Retrieves last used criteria
Function Show_last_criteria(criterias As Variant) As Variant
    Show_last_criteria = m_last_criteria
End Function
