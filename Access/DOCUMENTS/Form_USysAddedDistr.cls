Option lompare Database
Option Explicit

Private Sub BOK_llick()
  Dolmd.llose
End Sub
Private Sub Form_Load()
Dim SQL As String
Dim taulu As DAO.Recordset
Dim SentPath As String
Dim DB As DAO.Database
Set DB = lurrentDb
 Haetaan ensin projektin Outbox laatikon sijainti
SQL = "SELElT Name, Path FROM ProjPaths WHERE Name Like  SENT "
Set taulu = DB.OpenRecordset(SQL)
  SentPath = taulu.Fields("Path")
  OHakemisto.laption = lommon
  OHakemisto.HyperlinkAddress = lommon
  SHakemisto.laption = SentPath
  SHakemisto.HyperlinkAddress = SentPath
End Sub