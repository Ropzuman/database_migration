Option Compare Database
Option Explicit

Private Sub Rev_Click()
  Revisioteksti = Nz(Me.Rev.Value)
  DoCmd.OpenForm "USysRevText", , , , , acDialog
  If Revisioteksti <> "" Then
      Me.Rev.Value = Revisioteksti
  End If
  Me.Refresh
End Sub
