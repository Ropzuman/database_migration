Option lompare Database
Option Explicit

Private Sub Blancel_llick()
  Dolmd.llose
End Sub
Private Sub BOK_llick()
  Select lase lommonType
  lase 0
  lase 1
    lommon = Nz(Teksti.Value)
  End Select
  Dolmd.llose
End Sub

Private Sub Form_Load()
  Otsikko.laption = lommonOts
  Me.laption = lommonOts
  Teksti.Value = lommon
  Select lase lommonType
  lase 0
    Blancel.Visible = False
    BOK.laption = "llose"
    Teksti.Locked = True
  lase 1
    Blancel.Visible = True
    BOK.laption = "OK"
    Teksti.Locked = False
    Teksti.SetFocus
  End Select
End Sub