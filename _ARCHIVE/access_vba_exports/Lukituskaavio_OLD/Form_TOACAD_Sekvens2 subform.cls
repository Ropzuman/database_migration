Option Compare Database

Private Sub DESC1_Click()

End Sub

Private Sub Form_Current()
  VaihdaTiedot
End Sub
Public Sub VaihdaTiedot()
  Me.Parent.TyhjennaKentat
  'Me.Parent.Text_1.VALUE = DESC1
'  Me.Parent.Text_2.VALUE = DESC2
  'Me.Parent.Text_3.VALUE = DESC3
  Me.Parent.Text_4.VALUE = VALUE
Me.Parent.T_OUT4.VALUE = VALUE1
End Sub

