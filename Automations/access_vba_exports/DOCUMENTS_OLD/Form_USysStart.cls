Option Compare Database
Option Explicit
Public ajastin As Long
Private Sub Form_Load()
  Me.TimerInterval = 2000
  ajastin = Timer
End Sub
Private Sub Form_Timer()
  DoCmd.Close
End Sub

Private Sub Form_Unload(Cancel As Integer)
  If ShowSettings.Value Then
    DoCmd.OpenForm "SETTINGS"
  Else
    DoCmd.OpenForm "DOCUMENTS"
  End If
End Sub

Private Sub OLEUnbound6_Click()
  DoCmd.Close
End Sub
