Option Compare Database
Option Explicit

Private Sub RevMotList_Click()
  Revisioteksti = Nz(Me.RevMotList.Value)
  DoCmd.OpenForm "USysRevText", , , , , acDialog
  If Revisioteksti <> "" Then
      Me.RevMotList.Value = Revisioteksti
  End If
  Me.Refresh
End Sub
