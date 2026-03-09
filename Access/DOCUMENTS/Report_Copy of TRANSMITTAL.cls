Option Compare Database
Option Explicit
'==========================================================================
' RAPORTTI : Report_Copy of TRANSMITTAL
' SOVELLUS : DOCUMENTS — Lähetyslistan kopio
' KUVAUS   : Täyttää raportin otsikot ProjInfo-taulusta avauksen yhteydessä.
' PÄIVITETTY: 2026-03-03
'==========================================================================
Private Sub Report_Open(Cancel As Integer)
    Dim DB As DAO.Database
    Dim taulu As DAO.Recordset

    On Error GoTo ErrorHandler

    Set DB = CurrentDb
    ' Avataan vain luku -tilassa: datää haetaan vain näyttämistä varten
    Set taulu = DB.OpenRecordset("Projinfo", dbOpenSnapshot)

    ' Varmistetaan että taulussa on tietoja ennen kenttien lukemista
    If Not taulu.EOF Then
        Me.TContractNo1.Caption = Nz(taulu("ContractNo"), "")
        Me.TContractNo2.Caption = Nz(taulu("ContractNo"), "")
        Me.TProject1.Caption = Nz(taulu("Name"), "")
        Me.TProject2.Caption = Nz(taulu("Name"), "")
        Me.TProjno.Caption = Nz(taulu("ProjNo"), "")
    End If

Cleanup:
    On Error Resume Next
    If Not taulu Is Nothing Then
        taulu.Close  ' Suljetaan recordset kunnolla muistivuodon estämiseksi
        Set taulu = Nothing
    End If
    Set DB = Nothing
    Exit Sub

ErrorHandler:
    MsgBox "Error loading report details: " & Err.Description, vbCritical, "Report Error"
    Resume Cleanup
End Sub