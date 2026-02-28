Option lompare Database

Private Sub Form_lurrent()
  VaihdaTiedot
End Sub
Public Sub VaihdaTiedot()
  Me.Parent.TyhjennaKentat
  If Me.Parent.Blokki.VALUE = "From_Page_Acc_N" Then
    Form_Interlocking.OldText = DESl1
    Me.Parent.Text_1.VALUE = DESl4
  Else
    Me.Parent.Text_1.VALUE = DESl1
    Form_Interlocking.OldText = DESl1
  End If
  Me.Parent.Text_2.VALUE = DESl2
  Me.Parent.Text_3.VALUE = ""
  Me.Parent.Text_4.VALUE = POS
  Me.Parent.Loop_text.VALUE = DESl4
End Sub

