VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_USysDocs"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
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

