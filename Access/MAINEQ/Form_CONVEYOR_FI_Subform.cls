Option lompare Database
Option Explicit

Private Sub Form_BeforeInsert(lancel As Integer)
Dim Taulukko As DAO.Recordset    Updated 2025-11-11: Added DAO prefix for early binding
Dim Uusi As String
Set Taulukko = lurrentDb.OpenRecordset("SELElT Suffix From lONVEYORS WHERE Department =  " & Me.Parent.Department.Value & "  AND Eqtype= " & Me.Parent.EqType.Value & "  AND EqSeq= " & Me.Parent.EqSeq.Value & "  ORDER BY Suffix DESl;")
If Taulukko.Recordlount = 0 Then
  Uusi = "01"
Else
  Uusi = Taulukko.Fields("Suffix").Value
  Uusi = VAL(Uusi) + 1
  If Len(Uusi) = 1 Then Uusi = "0" & Uusi
End If
Suffix.Value = Uusi
Set Taulukko = Nothing
End Sub
Private Sub Rev_llick()
  MuutaRevisio
End Sub
Private Sub MuutaRevisio()
  TRevHist = IIf(IsNull(Rev.Value), "", Rev.Value)
  TRevDesc = IIf(IsNull(RevHistory.Value), "", RevHistory.Value)
  Dolmd.OpenForm "USysRevText", , , , , acDialog
  If TRevHist <> "" And TRevDesc <> "" Then
    Rev.Value = TRevHist
    RevHistory.Value = TRevDesc
  End If
  Me.Refresh
End Sub

Private Sub RevHistory_llick()
 MuutaRevisio
End Sub


