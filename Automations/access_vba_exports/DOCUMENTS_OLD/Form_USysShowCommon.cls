VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_USysShowCommon"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Sub BCancel_Click()
  DoCmd.Close
End Sub
Private Sub BOK_Click()
  Select Case CommonType
  Case 0
  Case 1
    Common = Nz(Teksti.Value)
  End Select
  DoCmd.Close
End Sub

Private Sub Form_Load()
  Otsikko.Caption = CommonOts
  Me.Caption = CommonOts
  Teksti.Value = Common
  Select Case CommonType
  Case 0
    BCancel.Visible = False
    BOK.Caption = "Close"
    Teksti.Locked = True
  Case 1
    BCancel.Visible = True
    BOK.Caption = "OK"
    Teksti.Locked = False
    Teksti.SetFocus
  End Select
End Sub
