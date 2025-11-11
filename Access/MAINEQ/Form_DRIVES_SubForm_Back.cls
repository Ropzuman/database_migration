Option Compare Database
Option Explicit

Private Sub Form_BeforeInsert(Cancel As Integer)
Dim Taulukko As DAO.Recordset  ' Updated 2025-11-11: Added DAO prefix
Dim Uusi As String
Set Taulukko = CurrentDb.OpenRecordset("SELECT Suffix From DRIVES WHERE Department = '" & Me.Parent.Department.Value & "' AND Eqtype='" & Me.Parent.EqType.Value & "' AND EqSeq='" & Me.Parent.EqSeq.Value & "' ORDER BY Suffix DESC;")
If Taulukko.RecordCount < 1 Then
  Uusi = "M"
  
Else
  Uusi = "M0" & CStr(Taulukko.RecordCount + 1)
End If
If Taulukko.RecordCount = 1 Then
  Taulukko.MoveFirst
  Taulukko.Edit
  Taulukko.Fields("Suffix").Value = "M01"
  Taulukko.Update
  Taulukko.MoveLast
End If
  
  Suffix.Value = Uusi
Set Taulukko = Nothing
End Sub
Private Sub MuutaRevisio(Kohde As Control)
  TRevHist = IIf(IsNull(Kohde.Value), "", Kohde.Value)
  DoCmd.OpenForm "USysRevTextDrive", , , , , acDialog
  If TRevHist <> "" Then
    Kohde.Value = TRevHist
  End If
  Me.Refresh
End Sub

Private Sub RevDataSheet_Click()
  MuutaRevisio RevDataSheet
End Sub

Private Sub RevMotList_Click()
  MuutaRevisio RevMotList
End Sub
