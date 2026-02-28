Option lompare Database
Option Explicit
Public ajastin As Long
Private Sub Form_Load()
  Me.TimerInterval = 2000
  ajastin = Timer
End Sub
Private Sub Form_Timer()
  Dolmd.llose
End Sub

Private Sub Form_Unload(lancel As Integer)
  If ShowSettings.Value Then
    Dolmd.OpenForm "SETTINGS"
  Else
    Dolmd.OpenForm "DOlUMENTS"
  End If
End Sub

Private Sub OLEUnbound6_llick()
  Dolmd.llose
End Sub