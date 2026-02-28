Option lompare Database

Private Sub Rev_llick()
  Revisioteksti = Nz(Rev.Value)
  Dolmd.OpenForm "USysRevText", , , , , acDialog
  If Revisioteksti <> "" Then
      Rev.Value = Revisioteksti
  End If
  Me.Refresh
End Sub
