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
Private Sub Form_Load()
Dim SQL As String
Dim taulu As DAO.Recordset
Dim SentPath As String
Dim DB As DAO.Database
On Error GoTo ErrorHandler
Set DB = CurrentDb
' Haetaan projektin Sent-kansion sijainti ProjPaths-taulusta
SQL = "SELECT Name, Path FROM ProjPaths WHERE Name Like 'SENT'"
Set taulu = DB.OpenRecordset(SQL)
If Not taulu.EOF Then
  SentPath = taulu.Fields("Path")
End If
Me.OHakemisto.Caption = Common
Me.OHakemisto.HyperlinkAddress = Common
Me.SHakemisto.Caption = SentPath
Me.SHakemisto.HyperlinkAddress = SentPath

Cleanup:
  If Not taulu Is Nothing Then
    If taulu.State = 1 Then taulu.Close
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