Option Compare Database
Option Explicit

Private Sub Command0_Click()
    ' Ajetaan IntLoopDescr20Upd-kyselyä – päivittää loop-kuvausten 20-merkkiset kentät
    On Error GoTo ErrorHandler
    
    DoCmd.Hourglass True
    DoCmd.OpenQuery "IntLoopDescr20Upd"
    DoCmd.Hourglass False
    Exit Sub

ErrorHandler:
    DoCmd.Hourglass False
    MsgBox "Virhe kyselyn ajossa: " & Err.Description, vbCritical, "IntLoopDescr20Update"
End Sub
