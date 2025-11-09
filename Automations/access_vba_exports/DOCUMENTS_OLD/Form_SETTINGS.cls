VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_SETTINGS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Sub Sulje_Click()
  DoCmd.Close
End Sub
Private Sub Form_Unload(Cancel As Integer)
  CurrentDb.Properties("AppTitle").Value = Nz(Nimi.Value)
  Application.RefreshTitleBar
  DoCmd.OpenForm "DOCUMENTS"
End Sub
Private Sub PoimiPolku(Kentta As Control)
Dim Hakemisto As String
  If IsNull(Kentta.Value) Then
    Hakemisto = ValitseHakem(Me.hWnd)
  Else
    Hakemisto = ValitseHakem(Me.hWnd, Kentta.Value)
  End If
  If Hakemisto <> "" Then Kentta.Value = Hakemisto
End Sub

Private Sub TPath_DblClick(Cancel As Integer)
  PoimiPolku TPath
End Sub
