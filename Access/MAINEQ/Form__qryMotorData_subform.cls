Option Compare Database



Private Sub ID_Click()

End Sub

Private Sub ID_DblClick(Cancel As Integer)
If Me.Parent.MotorTypeID <> "" And Me.Parent.MotorTypeID <> "0" Then
MsgBox "Tähän on valittu jo moottori " & Me.Parent.MotorTypeID, vbOKOnly
'Me.Parent.MotorTypeID = ID

Else
Me.Parent.MotorTypeID = ID
'MsgBox "Tähän on valittu jo moottori " & Me.Parent.MotorTypeID, vbOKOnly
End If

'Debug.Print Me.Parent.MotorTypeID
End Sub
