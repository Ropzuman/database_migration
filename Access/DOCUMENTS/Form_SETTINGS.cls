Option Compare Database
Option Explicit

Private Sub Sulje_Click()
  DoCmd.Close
End Sub
'''
' Form Unload Event - Save application title and reopen DOCUMENTS form
' OPTIMIZED 2025-11-09: Added explicit DAO typing, error handling
'''
Private Sub Form_Unload(Cancel As Integer)
  Dim DB As DAO.Database
  
  On Error GoTo ErrorHandler
  
  Set DB = CurrentDb
  DB.Properties("AppTitle").Value = Nz(Nimi.Value)
  Application.RefreshTitleBar
  
  DB.Close
  Set DB = Nothing
  
  DoCmd.OpenForm "DOCUMENTS"
  Exit Sub
  
ErrorHandler:
  MsgBox "Error saving settings: " & Err.Description, vbExclamation, "Settings Error"
  If Not DB Is Nothing Then
    DB.Close
    Set DB = Nothing
  End If
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