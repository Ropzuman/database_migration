Option Compare Database
Option Explicit
Private Sub Report_Open(Cancel As Integer)
Dim taulu As Recordset
Set taulu = CurrentDb.OpenRecordset("Projinfo")
TContractNo1.Caption = Nz(taulu("ContractNo"))
TContractNo2.Caption = Nz(taulu("ContractNo"))
TProject1.Caption = Nz(taulu("Name"))
TProject2.Caption = Nz(taulu("Name"))
TProjno.Caption = Nz(taulu("ProjNo"))
Set taulu = Nothing
End Sub
