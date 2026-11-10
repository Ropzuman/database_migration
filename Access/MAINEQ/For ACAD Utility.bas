Option Compare Database
Option Explicit

' Piste (X, Y, Z) AutoCAD-koordinaateille
Public Type iPoint
  Pisteet(2) As Double
End Type

Public Paikat() As iPoint

' Hiiren kohdistimen sijainnin API (64-bittinen, PtrSafe)
Type POINTAPI
    X As Long
    Y As Long
End Type

Private Declare PtrSafe Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long

