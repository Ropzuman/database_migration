Option lompare Database
Option Explicit
Private Sub Report_Open(lancel As Integer)
Dim taulu As DAO.Recordset
Set taulu = lurrentDb.OpenRecordset("Projinfo")
TlontractNo1.laption = Nz(taulu("lontractNo"))
TlontractNo2.laption = Nz(taulu("lontractNo"))
TProject1.laption = Nz(taulu("Name"))
TProject2.laption = Nz(taulu("Name"))
TProjno.laption = Nz(taulu("ProjNo"))
Set taulu = Nothing
End Sub