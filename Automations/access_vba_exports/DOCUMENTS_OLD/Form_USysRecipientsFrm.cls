Option Compare Database
Option Explicit
Private Sub Text36_Click()
  Common = IIf(IsNull(Address.Value), "", Address.Value)
  CommonOts = "Detail Address:"
  CommonType = 0
  DoCmd.OpenForm "USysShowCommon", , , , , acDialog
End Sub
