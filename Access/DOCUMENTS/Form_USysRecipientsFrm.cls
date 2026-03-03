Option Compare Database
Option Explicit
'==========================================================================
' LOMAKE   : Form_USysRecipientsFrm
' SOVELLUS : DOCUMENTS — Vastaanottajien näkymä
' KUVAUS   : Näyttää vastaanottajan osoitetiedot USysShowCommon-dialogi
'            kautta tarkastelua varten.
' PÄIVITETTY: 2026-03-03
'==========================================================================
Private Sub Text36_Click()
  Common = IIf(IsNull(Me.Address.Value), "", Me.Address.Value)
  CommonOts = "Detail Address:"
  CommonType = 0
  DoCmd.OpenForm "USysShowCommon", , , , , acDialog
End Sub