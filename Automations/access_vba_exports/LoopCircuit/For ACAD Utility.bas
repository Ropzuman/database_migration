Attribute VB_Name = "For ACAD Utility"
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
#If VBA7 Then
  Private Declare PtrSafe Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
#Else
  Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
#End If
