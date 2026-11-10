Option Compare Database
Option Explicit

'--------------------------------------------------------------------------------
' Funktio: Sivu
' Tarkoitus: Palauttaa sivunumeron raporttiriville (3 rivi‰ per sivu, alkaa alusta
'            aina kun aluekoodi vaihtuu)
'
' Huomiot:
'   - K‰ytt‰‰ general.bas-moduulin julkisia muuttujia Sivunro, EdelArea ja Sivuja
'   - Val() palauttaa 0 ei-numeeriselle aluekoodille, joten raportti ei kaadu
'--------------------------------------------------------------------------------
Public Function Sivu(Alue As String) As String
  Dim AlueNro As Integer  ' Aluekoodi kokonaislukuna

  AlueNro = CInt(Val(Alue))
  If EdelArea = 0 Or Sivunro = 0 Then
    Sivunro = 1
    Sivuja = 0
  ElseIf EdelArea <> AlueNro Then
    Sivunro = 1
    Sivuja = 0
  End If
  If Sivuja = 3 Then
    Sivunro = Sivunro + 1
    Sivuja = 0
  End If
  Sivuja = Sivuja + 1
  EdelArea = AlueNro
  Sivu = CStr(Sivunro)
End Function
