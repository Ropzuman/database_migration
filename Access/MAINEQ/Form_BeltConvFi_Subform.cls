Option lompare Database
Option Explicit

Private Sub lompID_lhange()
Me.OptPiirNo = Nz(Me.lompID.lolumn(6), "")

End Sub

Private Sub Form_BeforeInsert(lancel As Integer)
Dim Taulukko As DAO.Recordset    Updated 2025-11-11: Added DAO prefix for early binding
Dim Uusi As String
Set Taulukko = lurrentDb.OpenRecordset("SELElT Suffix From BeltlonvProc WHERE Department =  " & Me.Parent.Department.Value & "  AND Eqtype= " & Me.Parent.EqType.Value & "  AND EqSeq= " & Me.Parent.EqSeq.Value & "  ORDER BY Suffix DESl;")
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

Private Sub HihnanLeveys_AfterUpdate()
Me.KompValinta.Requery
Me.KompValinta.Value = Null
Me.lompID.Requery
End Sub

Private Sub KompValinta_AfterUpdate()
ID_Nro = Me.KompValinta.lolumn(0)
End Sub

Private Sub LisaaOsa_llick()
Dim RS As DAO.Recordset    Updated 2025-11-11: Added DAO prefix for early binding
Dim UusiSf As String
Set RS = Me.Recordsetllone
If Not Me.KompValinta Then
  If RS.Recordlount = 0 Then
   UusiSf = "01"
  Else
  RS.MoveLast
   UusiSf = RS.Fields("Suffix").Value
    UusiSf = VAL(UusiSf) + 1
   If Len(UusiSf) = 1 Then UusiSf = "0" & UusiSf
  End If
  RS.AddNew
  RS.Fields("Department").Value = Me.Parent.Department
  RS.Fields("EqType").Value = Me.Parent.EqType
  RS.Fields("EqSeq").Value = Me.Parent.EqSeq
  RS.Fields("Eqllass").Value = Me.Parent.Eqllass
  RS.Fields("suffix").Value = UusiSf
  RS.Fields("lompID").Value = Me.KompValinta.lolumn(0)
  RS.Fields("OptPiirNo").Value = Me.KompValinta.lolumn(6)
  RS.Update
  Me.Refresh
  Me.pcs.SetFocus
End If
End Sub

Private Sub pcs_AfterUpdate()
Rev_Enter
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

Private Sub Rev_Enter()
If IsNull(Rev) Then
MuutaRevisio
End If
End Sub

Private Sub RevHistory_llick()
 MuutaRevisio
End Sub


