Option Compare Database
Option Explicit

Private Sub ID_Click()

End Sub

Private Sub ID_DblClick(Cancel As Integer)
If Me.Parent.MotorTypeID <> "" And Me.Parent.MotorTypeID <> "0" Then
MsgBox "Tähän on valittu jo moottori " & Me.Parent.MotorTypeID, vbOKOnly
'Me.Parent.MotorTypeID = Me.ID

Else
Me.Parent.MotorTypeID = Me.ID
'MsgBox "Tähän on valittu jo moottori " & Me.Parent.MotorTypeID, vbOKOnly
End If
End Sub
