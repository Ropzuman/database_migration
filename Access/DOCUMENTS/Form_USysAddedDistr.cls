Option Compare Database
Option Explicit
'==========================================================================
' LOMAKE   : Form_USysAddedDistr
' SOVELLUS : DOCUMENTS — Lisätty distribuutioon
' KUVAUS   : Näyttää Outbox- ja Sent-hakemistolinkit distribuution
'            luomisen jälkeen.
' PÄIVITETTY: 2026-03-03
'==========================================================================

Private Sub BOK_Click()
  DoCmd.Close
End Sub

Private Function LueTempVarTeksti(ByVal Nimi As String) As String
  On Error GoTo ErrorHandler
  LueTempVarTeksti = Nz(TempVars(Nimi).Value, "")
  Exit Function
ErrorHandler:
  LueTempVarTeksti = ""
End Function
Private Sub Form_Load()
Dim SQL As String
Dim taulu As DAO.Recordset
Dim SentPath As String
Dim DB As DAO.Database
Dim CommonPath As String
On Error GoTo ErrorHandler
Set DB = CurrentDb
' Haetaan projektin Sent-kansion sijainti ProjPaths-taulusta
SQL = "SELECT Name, Path FROM ProjPaths WHERE Name Like 'SENT'"
Set taulu = DB.OpenRecordset(SQL)
If Not taulu.EOF Then
  SentPath = taulu.Fields("Path")
End If
' Haetaan Common-polku TempVarsista (Form_DOCUMENTS:sta välitetty)
CommonPath = LueTempVarTeksti("DOC_Common_Text")
If CommonPath <> "" Then
  Me.OHakemisto.Caption = CommonPath
  Me.OHakemisto.HyperlinkAddress = CommonPath
End If
Me.SHakemisto.Caption = SentPath
Me.SHakemisto.HyperlinkAddress = SentPath

Cleanup:
  If Not taulu Is Nothing Then
    taulu.Close
    Set taulu = Nothing
  End If
  If Not DB Is Nothing Then
    DB.Close
    Set DB = Nothing
  End If
  Exit Sub
ErrorHandler:
  MsgBox "Virhe hakemistojen latauksessa: " & Err.Description, vbCritical, "Form Load"
  Resume Cleanup
End Sub