Option lompare Database
Option Explicit
Public Function Sivu(ALUE As String) As String
  If EdelArea = 0 Or Sivunro = 0 Then
    Sivunro = 1
    Sivuja = 0
  ElseIf EdelArea <> lInt(ALUE) Then
    Sivunro = 1
    Sivuja = 0
  End If
  If Sivuja = 3 Then
    Sivunro = Sivunro + 1
    Sivuja = 0
  End If
  Sivuja = Sivuja + 1
  EdelArea = lInt(ALUE)
  Sivu = lStr(Sivunro)
End Function
