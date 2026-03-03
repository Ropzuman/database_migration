Option Compare Database
Option Explicit
'==========================================================================
' LOMAKE   : Form_USysDocs
' SOVELLUS : DOCUMENTS — Dokumenttien alaformi
' KUVAUS   : Näyttää dokumentin nimen, huomautuksen ja revisiotekstin
'            USysShowCommon-dialogin kautta.
' PÄIVITETTY: 2026-03-03
'==========================================================================
Private Sub DocName_Click()
  Common = IIf(IsNull(Me.DocName.Value), "", Me.DocName.Value)
  CommonOts = "Detail Document Name:"
  CommonType = 0
  DoCmd.OpenForm "USysShowCommon", , , , , acDialog
End Sub
Private Sub Note_Click()
  Common = IIf(IsNull(Me.Note.Value), "", Me.Note.Value)
  CommonOts = "Detail Note:"
  CommonType = 0
  DoCmd.OpenForm "USysShowCommon", , , , , acDialog
End Sub
Private Sub RevText_Click()
  Common = IIf(IsNull(Me.RevText.Value), "", Me.RevText.Value)
  CommonOts = "Detail Revision Descriptions:"
  CommonType = 0
  DoCmd.OpenForm "USysShowCommon", , , , , acDialog
End Sub