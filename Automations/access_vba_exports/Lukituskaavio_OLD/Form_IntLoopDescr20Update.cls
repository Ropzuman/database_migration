VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_IntLoopDescr20Update"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database

Private Sub Command0_Click()
Dim DB As Database

DoCmd.Hourglass True

DoCmd.OpenQuery ("IntLoopDescr20Upd")

DoCmd.Hourglass False

End Sub
