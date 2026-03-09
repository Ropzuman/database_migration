Option Compare Database
Option Explicit

Private Sub Form_Current()
  VaihdaTiedot
End Sub

Public Sub VaihdaTiedot()
  Me.Parent.TyhjennaKentat
  Me.Parent.Text_4.VALUE = VALUE
  Me.Parent.T_OUT4.VALUE = VALUE1
End Sub
