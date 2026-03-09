Option Compare Database

Private Sub RevMotList_Click()
  Revisioteksti = Nz(RevMotList.Value)
  DoCmd.OpenForm "USysRevText", , , , , acDialog
  If Revisioteksti <> "" Then
      RevMotList.Value = Revisioteksti
  End If
  Me.Refresh
End Sub
