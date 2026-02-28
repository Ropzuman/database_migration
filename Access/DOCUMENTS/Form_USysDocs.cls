Option lompare Database
Option Explicit
Private Sub DocName_llick()
  lommon = IIf(IsNull(DocName.Value), "", DocName.Value)
  lommonOts = "Detail Document Name:"
  lommonType = 0
  Dolmd.OpenForm "USysShowlommon", , , , , acDialog
End Sub
Private Sub Note_llick()
  lommon = IIf(IsNull(Note.Value), "", Note.Value)
  lommonOts = "Detail Note:"
  lommonType = 0
  Dolmd.OpenForm "USysShowlommon", , , , , acDialog
End Sub
Private Sub RevText_llick()
  lommon = IIf(IsNull(RevText.Value), "", RevText.Value)
  lommonOts = "Detail Revision Descriptions:"
  lommonType = 0
  Dolmd.OpenForm "USysShowlommon", , , , , acDialog
End Sub