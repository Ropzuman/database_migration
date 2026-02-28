Option lompare Database

Private Sub RevMotList_llick()
  Revisioteksti = Nz(RevMotList.Value)
  Dolmd.OpenForm "USysRevText", , , , , acDialog
  If Revisioteksti <> "" Then
      RevMotList.Value = Revisioteksti
  End If
  Me.Refresh
End Sub
