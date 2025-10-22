Attribute VB_Name = "USysCheck"
Option Compare Database
Option Explicit

 #If VBA7 Then
    Private Declare PtrSafe Function wu_GetUserName Lib "advapi32" Alias "GetUserNameA" (ByVal lpBuffer As String, ByVal nSize As LongPtr) As LongPtr
    ' --------- [ CHOOSE FILE ] -----------------
    Private Declare PtrSafe Function GetOpenFileName Lib "comdlg32.dll" Alias "GetOpenFileNameA" (pOpenfilename As OPENFILENAME) As LongPtr
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
    Declare Function wu_GetUserName Lib "advapi32" Alias "GetUserNameA" (ByVal lpBuffer As String, nSize As Long) As Long
    ' --------- [ CHOOSE FILE ] -----------------
    Declare Function GetOpenFileName Lib "comdlg32.dll" Alias "GetOpenFileNameA" (pOpenfilename As OPENFILENAME) As Long
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
'Etsii temp hakemiston
'Public Declare Function GetTempPath Lib "kernel32" Alias "GetTempPathA" (ByVal nBufferLength As Long, ByVal lpBuffer As String) As Long
Global last_criteria As Variant, last_used As Variant
 
Function Set_last(Values As Variant, criterias As Variant) As Variant
    ' Store last used values and criteria (keeps state in module-level globals)
    last_criteria = criterias
    last_used = Values
    Set_last = last_used
End Function

Function Show_last(criterias As Variant) As Variant
    Show_last = last_used
End Function

Function Show_last_criteria(criterias As Variant) As Variant
    Show_last_criteria = last_criteria
End Function
