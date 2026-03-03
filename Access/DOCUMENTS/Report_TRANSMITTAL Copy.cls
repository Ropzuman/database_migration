Option Compare Database
Option Explicit
'==========================================================================
' RAPORTTI : Report_TRANSMITTAL Copy
' SOVELLUS : DOCUMENTS — Lähetyslistan kopio
' KUVAUS   : Täyttää raportin otsikot ProjInfo-taulusta avauksen yhteydessä.
' PÄIVITETTY: 2026-03-03
'==========================================================================
Private Sub Report_Open(Cancel As Integer)
Dim taulu As DAO.Recordset
Set taulu = CurrentDb.OpenRecordset("Projinfo")
Me.TContractNo1.Caption = Nz(taulu("ContractNo"))
Me.TContractNo2.Caption = Nz(taulu("ContractNo"))
Me.TProject1.Caption = Nz(taulu("Name"))
Me.TProject2.Caption = Nz(taulu("Name"))
Me.TProjno.Caption = Nz(taulu("ProjNo"))
Set taulu = Nothing
End Sub