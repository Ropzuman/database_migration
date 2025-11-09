Option Compare Database
Option Explicit

Private Sub CompID_Change()
Me.OptPiirNo = Nz(Me.CompID.Column(6), "")

End Sub

Private Sub Form_BeforeInsert(Cancel As Integer)
Dim Taulukko As Recordset
Dim Uusi As String
Set Taulukko = CurrentDB.OpenRecordset("SELECT Suffix From BeltConvProc WHERE Department = '" & Me.Parent.Department.Value & "' AND Eqtype='" & Me.Parent.EqType.Value & "' AND EqSeq='" & Me.Parent.EqSeq.Value & "' ORDER BY Suffix DESC;")
If Taulukko.RecordCount = 0 Then
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
Me.CompID.Requery
End Sub

Private Sub KompValinta_AfterUpdate()
ID_Nro = Me.KompValinta.Column(0)
End Sub

Private Sub LisaaOsa_Click()
Dim RS As Recordset, UusiSf As String
Set RS = Me.RecordsetClone
If Not Me.KompValinta Then
  If RS.RecordCount = 0 Then
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
  RS.Fields("EqClass").Value = Me.Parent.EqClass
  RS.Fields("suffix").Value = UusiSf
  RS.Fields("CompID").Value = Me.KompValinta.Column(0)
  RS.Fields("OptPiirNo").Value = Me.KompValinta.Column(6)
  RS.Update
  Me.Refresh
  Me.pcs.SetFocus
End If
End Sub

Private Sub pcs_AfterUpdate()
Rev_Enter
End Sub

Private Sub Rev_Click()
  MuutaRevisio
End Sub
Private Sub MuutaRevisio()
  TRevHist = IIf(IsNull(Rev.Value), "", Rev.Value)
  TRevDesc = IIf(IsNull(RevHistory.Value), "", RevHistory.Value)
  DoCmd.OpenForm "USysRevText", , , , , acDialog
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

Private Sub RevHistory_Click()
 MuutaRevisio
End Sub
