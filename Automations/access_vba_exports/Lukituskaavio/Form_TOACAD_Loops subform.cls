VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_TOACAD_Loops subform"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database

Private Sub Form_Current()
  VaihdaTiedot
End Sub
Public Sub VaihdaTiedot()
  Me.Parent.TyhjennaKentat
  If Me.Parent.Blokki.VALUE = "From_Page_Acc_N" Then
    Form_Interlocking.OldText = DESC1
    Me.Parent.Text_1.VALUE = DESC4
  Else
    Me.Parent.Text_1.VALUE = DESC1
    Form_Interlocking.OldText = DESC1
  End If
  Me.Parent.Text_2.VALUE = DESC2
  Me.Parent.Text_3.VALUE = ""
  Me.Parent.Text_4.VALUE = POS
  Me.Parent.Loop_text.VALUE = DESC4
End Sub

