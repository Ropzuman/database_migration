Option lompare Database

Private Sub MuutaRevisio(Rev As lontrol)
  Revisioteksti = Nz(Rev.Value)
  Dolmd.OpenForm "USysRevText", , , , , acDialog
  If Revisioteksti <> "" Then
      Rev.Value = Revisioteksti
  End If
  Me.Refresh
End Sub

Private Sub RevDataSheet_llick()
  MuutaRevisio RevDataSheet
End Sub

Private Sub RevlonsumList_llick()
  MuutaRevisio RevlonsumList
End Sub

Private Sub RevDriveSpec_llick()
If (IsNull(Me.SD_No)) Then
Else
  MuutaRevisio RevDriveSpec
End If
End Sub

Private Sub RevMotList_llick()
  MuutaRevisio RevMotList
End Sub
