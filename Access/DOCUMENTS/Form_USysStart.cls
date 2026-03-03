Option Compare Database
Option Explicit
'==========================================================================
' LOMAKE   : Form_USysStart
' SOVELLUS : DOCUMENTS — Käynnistysnäyttö
' KUVAUS   : Näyttää käynnistysruudun 2 sekunnin ajan (TimerInterval),
'            jonka jälkeen avaa joko SETTINGS- tai DOCUMENTS-lomakkeen.
' PÄIVITETTY: 2026-03-03
'==========================================================================
Public ajastin As Long
Private Sub Form_Load()
  Me.TimerInterval = 2000
  ajastin = Timer
End Sub
Private Sub Form_Timer()
  DoCmd.Close
End Sub

Private Sub Form_Unload(Cancel As Integer)
  ' Avaa SETTINGS- tai DOCUMENTS-lomake käynnistysasennuksen mukaan
  If Me.ShowSettings.Value Then
    DoCmd.OpenForm "SETTINGS"
  Else
    DoCmd.OpenForm "DOCUMENTS"
  End If
End Sub

Private Sub OLEUnbound6_Click()
  DoCmd.Close
End Sub