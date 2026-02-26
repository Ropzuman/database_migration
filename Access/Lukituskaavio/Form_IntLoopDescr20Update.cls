Option Compare Database

Private Sub Command0_Click()
Dim DB As DAO.Database

DoCmd.Hourglass True

DoCmd.OpenQuery ("IntLoopDescr20Upd")

DoCmd.Hourglass False

End Sub
