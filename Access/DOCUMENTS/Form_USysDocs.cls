Option Compare Database
Option Explicit
Private Sub DocName_Click()
  Common = IIf(IsNull(DocName.Value), "", DocName.Value)
  CommonOts = "Detail Document Name:"
  CommonType = 0
  DoCmd.OpenForm "USysShowCommon", , , , , acDialog
End Sub
Private Sub Note_Click()
  Common = IIf(IsNull(Note.Value), "", Note.Value)
  CommonOts = "Detail Note:"
  CommonType = 0
  DoCmd.OpenForm "USysShowCommon", , , , , acDialog
End Sub
Private Sub RevText_Click()
  Common = IIf(IsNull(RevText.Value), "", RevText.Value)
  CommonOts = "Detail Revision Descriptions:"
  CommonType = 0
  DoCmd.OpenForm "USysShowCommon", , , , , acDialog
End Sub