Option Compare Database

Private Sub Command0_Click()
Dim DB As Database

DoCmd.Hourglass True

DoCmd.OpenQuery ("IntLoopDescr20Upd")

DoCmd.Hourglass False

End Sub
