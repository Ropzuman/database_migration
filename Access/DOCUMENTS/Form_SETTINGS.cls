Option Compare Database
Option Explicit
'==========================================================================
' LOMAKE   : Form_SETTINGS
' SOVELLUS : DOCUMENTS — Sovellusasetukset
' KUVAUS   : Tallennetaan sovelluksen nimi (AppTitle) tietokantaan
'            sulkemisen yhteydessä ja avataan DOCUMENTS-päälomake.
'            Hakemiston valintadialogia käytetään polun valintaan.
' PÄIVITETTY: 2026-03-03
'==========================================================================

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
  ' Varmistetaan, ettei tallenneta tyhjää sovelluksen otsikkoa
  If Len(Nz(Me.Nimi.Value, "")) > 0 Then
    DB.Properties("AppTitle").Value = Me.Nimi.Value
    Application.RefreshTitleBar
  End If
  
  ' EI SAA SULKEA: CurrentDb-viitettä ei tule sulkea koodista!
  Set DB = Nothing
  
  DoCmd.OpenForm "DOCUMENTS"
  Exit Sub
  
ErrorHandler:
  MsgBox "Error saving settings: " & Err.Description, vbExclamation, "Settings Error"
  If Not DB Is Nothing Then
    ' DB.Close <- Poistettu; CurrentDb-viitettä ei tule sulkea
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
  PoimiPolku Me.TPath
End Sub