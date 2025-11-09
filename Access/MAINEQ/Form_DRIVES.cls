Option Compare Database

Private Sub MuutaRevisio(Rev As Control)
  Revisioteksti = Nz(Rev.Value)
  DoCmd.OpenForm "USysRevText", , , , , acDialog
  If Revisioteksti <> "" Then
      Rev.Value = Revisioteksti
  End If
  Me.Refresh
End Sub

Private Sub RevDataSheet_Click()
  MuutaRevisio RevDataSheet
End Sub

Private Sub RevConsumList_Click()
  MuutaRevisio RevConsumList
End Sub

Private Sub RevDriveSpec_Click()
If (IsNull(Me.SD_No)) Then
Else
  MuutaRevisio RevDriveSpec
End If
End Sub

Private Sub RevMotList_Click()
  MuutaRevisio RevMotList
End Sub
