Attribute VB_Name = "For ACAD Utility"
Option Compare Database
Option Explicit

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
Public Type iPoint
  Pisteet(2) As Double
End Type
Public Paikat() As iPoint
'Etsii temp hakemiston
'Public Declare Function GetTempPath Lib "kernel32" Alias "GetTempPathA" (ByVal nBufferLength As Long, ByVal lpBuffer As String) As Long
'Hiiren kursorin sijainnin m‰‰ritt‰minen
Type POINTAPI ' Declare types
    X As Long
    Y As Long
End Type
Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long ' Declare API

