Option Compare Database
Option Explicit
'==========================================================================
' LOMAKE   : Form_USysDocs
' SOVELLUS : DOCUMENTS — Dokumenttien alaformi
' KUVAUS   : Näyttää dokumentin nimen, huomautuksen ja revisiotekstin
'            USysShowCommon-dialogin kautta.
' PÄIVITETTY: 2026-03-03
'==========================================================================
Private Const TV_DOC_COMMON_TEXT As String = "DOC_Common_Text"
Private Const TV_DOC_COMMON_TITLE As String = "DOC_Common_Title"
Private Const TV_DOC_COMMON_TYPE As String = "DOC_Common_Type"

Private Sub DocName_Click()
  AsetaTempVar TV_DOC_COMMON_TEXT, IIf(IsNull(Me.DocName.Value), "", Me.DocName.Value)
  AsetaTempVar TV_DOC_COMMON_TITLE, "Detail Document Name:"
  AsetaTempVar TV_DOC_COMMON_TYPE, "0"
  DoCmd.OpenForm "USysShowCommon", , , , , acDialog
End Sub
Private Sub Note_Click()
  AsetaTempVar TV_DOC_COMMON_TEXT, IIf(IsNull(Me.Note.Value), "", Me.Note.Value)
  AsetaTempVar TV_DOC_COMMON_TITLE, "Detail Note:"
  AsetaTempVar TV_DOC_COMMON_TYPE, "0"
  DoCmd.OpenForm "USysShowCommon", , , , , acDialog
End Sub
Private Sub RevText_Click()
  AsetaTempVar TV_DOC_COMMON_TEXT, IIf(IsNull(Me.RevText.Value), "", Me.RevText.Value)
  AsetaTempVar TV_DOC_COMMON_TITLE, "Detail Revision Descriptions:"
  AsetaTempVar TV_DOC_COMMON_TYPE, "0"
  DoCmd.OpenForm "USysShowCommon", , , , , acDialog
End Sub

Private Sub AsetaTempVar(ByVal Nimi As String, ByVal Arvo As String)
  On Error Resume Next
  TempVars.Remove Nimi
  On Error GoTo 0
  TempVars.Add Nimi, Arvo
End Sub