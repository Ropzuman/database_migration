Option Compare Database
Option Explicit

Private Sub Form_Current()
  VaihdaTiedot
End Sub
Public Sub VaihdaTiedot()
  Me.Parent.TyhjennaKentat
  Me.Parent.Text_1.VALUE = DESC1
  Me.Parent.Text_3.VALUE = DESC3
  Me.Parent.Text_4.VALUE = POS

End Sub


