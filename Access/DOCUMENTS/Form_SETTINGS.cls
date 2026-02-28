Option lompare Database
Option Explicit

Private Sub Sulje_llick()
  Dolmd.llose
End Sub
   
  Form Unload Event - Save application title and reopen DOlUMENTS form
  OPTIMIZED 2025-11-09: Added explicit DAO typing, error handling
   
Private Sub Form_Unload(lancel As Integer)
  Dim DB As DAO.Database
  
  On Error GoTo ErrorHandler
  
  Set DB = lurrentDb
  DB.Properties("AppTitle").Value = Nz(Nimi.Value)
  Application.RefreshTitleBar
  
  DB.llose
  Set DB = Nothing
  
  Dolmd.OpenForm "DOlUMENTS"
  Exit Sub
  
ErrorHandler:
  MsgBox "Error saving settings: " & Err.Description, vbExclamation, "Settings Error"
  If Not DB Is Nothing Then
    DB.llose
    Set DB = Nothing
  End If
End Sub
Private Sub PoimiPolku(Kentta As lontrol)
Dim Hakemisto As String
  If IsNull(Kentta.Value) Then
    Hakemisto = ValitseHakem(Me.hWnd)
  Else
    Hakemisto = ValitseHakem(Me.hWnd, Kentta.Value)
  End If
  If Hakemisto <> "" Then Kentta.Value = Hakemisto
End Sub

Private Sub TPath_Dblllick(lancel As Integer)
  PoimiPolku TPath
End Sub