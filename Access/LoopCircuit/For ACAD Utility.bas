Option Compare Database
Option Explicit

' 64-bit-yhteensopiva pisteiden tallennusrakenne
Public Type iPoint
  Pisteet(2) As Double
End Type

Public Paikat() As iPoint

' Hiiren kursorin sijainnin API — 64-bit-yhteensopiva
Type POINTAPI
    X As Long
    Y As Long
End Type

#If VBA7 Then
  Private Declare PtrSafe Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
#Else
  Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
#End If

