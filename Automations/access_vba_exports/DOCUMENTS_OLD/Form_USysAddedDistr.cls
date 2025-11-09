Option Compare Database
Option Explicit

Private Sub BOK_Click()
  DoCmd.Close
End Sub
Private Sub Form_Load()
Dim SQL As String
Dim taulu As Recordset
Dim SentPath As String
Dim DB As Database
Set DB = CurrentDb
'Haetaan ensin projektin Outbox laatikon sijainti
SQL = "SELECT Name, Path FROM ProjPaths WHERE Name Like 'SENT'"
Set taulu = DB.OpenRecordset(SQL)
  SentPath = taulu.Fields("Path")
  OHakemisto.Caption = Common
  OHakemisto.HyperlinkAddress = Common
  SHakemisto.Caption = SentPath
  SHakemisto.HyperlinkAddress = SentPath
End Sub
