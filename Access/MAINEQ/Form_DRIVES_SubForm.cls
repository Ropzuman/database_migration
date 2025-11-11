Option Compare Database
Option Explicit

Private Sub Form_BeforeInsert(Cancel As Integer)
Dim Taulukko As DAO.Recordset  ' Updated 2025-11-11: Added DAO prefix for early binding
Dim Uusi As String
Set Taulukko = CurrentDb.OpenRecordset("SELECT Suffix From DRIVES WHERE Department = '" & Me.Parent.Department.Value & "' AND Eqtype='" & Me.Parent.EqType.Value & "' AND EqSeq='" & Me.Parent.EqSeq.Value & "' ORDER BY Suffix DESC;")
If Taulukko.RecordCount < 9 Then
  Uusi = "0" & CStr(Taulukko.RecordCount + 1)
Else
  Uusi = CStr(Taulukko.RecordCount + 1)
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

Private Sub Rpm_Exit(Cancel As Integer)
If Me.Transmission Like "*GM*" Then
Exit Sub
Else
If Not IsNull(Me.Rpm) And Me.Rpm <> "0" Then
    If Me.Rpm <> "3000" And Me.Rpm <> "1500" And Me.Rpm <> "1000" And Me.Rpm <> "750" Then
        MsgBox "Huom!@Arvo pit�� olla synkroninen nopeus! @Esim. 3000, 1500, 1000 tai 750. @Paitsi vaihdemoottoreilla (GM)", vbOKOnly
        Me.Rpm.SetFocus
    End If
End If
End If
End Sub



