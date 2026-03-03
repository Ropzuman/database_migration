Option Compare Database
Option Explicit
'==========================================================================
' LOMAKE   : Form_USysShowCommon
' SOVELLUS : DOCUMENTS — Yleiskäyttöinen tekstin näyttö/muokkaus
' KUVAUS   : Yleiskäyttöinen teksti-ikkuna lomakkeiden väliseen tiedonsiirtoon.
'            CommonType 0 = näytä ainoastaan (Locked), 1 = muokkaa.
'            SetFocus-kutsu on kosmeettinen — kaatumattomuuden vuoksi
'            suojattu On Error Resume Next -lohkolla.
' PÄIVITETTY: 2026-03-03
'==========================================================================

Private Sub BCancel_Click()
  DoCmd.Close
End Sub
Private Sub BOK_Click()
  Select Case CommonType
  Case 0
  Case 1
    Common = Nz(Me.Teksti.Value)
  End Select
  DoCmd.Close
End Sub

Private Sub Form_Load()
  Me.Otsikko.Caption = CommonOts
  Me.Caption = CommonOts
  Me.Teksti.Value = Common
  Select Case CommonType
  Case 0
    Me.BCancel.Visible = False
    Me.BOK.Caption = "Close"
    Me.Teksti.Locked = True
  Case 1
    Me.BCancel.Visible = True
    Me.BOK.Caption = "OK"
    Me.Teksti.Locked = False
    ' SetFocus on kosmeettinen toiminto — suojataan kaatumiselta
    On Error Resume Next
    Me.Teksti.SetFocus
    On Error GoTo 0
  End Select
End Sub