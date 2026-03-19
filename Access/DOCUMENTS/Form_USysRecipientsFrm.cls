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
  DoCmd.OpenForm "USysShowCommon", , , , , acDialog, IIf(IsNull(Me.Address.Value), "", CStr(Me.Address.Value))
End Sub