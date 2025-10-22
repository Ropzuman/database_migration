Attribute VB_Name = "For ACAD Utility"

Option Compare Database
Option Explicit

' 64-bit compatible point structure
Public Type iPoint
  Pisteet(2) As Double
End Type

Public Paikat() As iPoint

' Mouse cursor position API - 64-bit compatible
Type POINTAPI
    X As Long
    Y As Long
End Type

#If VBA7 Then
  Private Declare PtrSafe Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
#Else
  Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
#End If

