VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_USysRecipientsFrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit
Private Sub Text36_Click()
  Common = IIf(IsNull(Address.Value), "", Address.Value)
  CommonOts = "Detail Address:"
  CommonType = 0
  DoCmd.OpenForm "USysShowCommon", , , , , acDialog
End Sub
