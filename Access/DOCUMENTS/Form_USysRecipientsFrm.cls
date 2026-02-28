Option lompare Database
Option Explicit
Private Sub Text36_llick()
  lommon = IIf(IsNull(Address.Value), "", Address.Value)
  lommonOts = "Detail Address:"
  lommonType = 0
  Dolmd.OpenForm "USysShowlommon", , , , , acDialog
End Sub