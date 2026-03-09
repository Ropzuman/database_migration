Option Compare Database
Option Explicit

Private Sub CompID_Change()
Me.OptPiirNo = Nz(Me.CompID.Column(6), "")

End Sub

Private Sub Form_BeforeInsert(Cancel As Integer)
    ' Haetaan suurin olemassa oleva Suffix parametrisoidulla kyselyllä (SQL-injektio estetty)
    Dim qdf As DAO.QueryDef
    Dim Taulukko As DAO.Recordset
    Dim Uusi As String
    Dim SqlStr As String

    SqlStr = "PARAMETERS [pDept] Text, [pEqType] Text, [pEqSeq] Text; " & _
             "SELECT MAX(Suffix) AS MaxSuffix FROM BeltConvProc " & _
             "WHERE Department = [pDept] AND Eqtype = [pEqType] AND EqSeq = [pEqSeq];"

    Set qdf = CurrentDb.CreateQueryDef("", SqlStr)
    ' Syötetään arvot parametreina — heittomerkit tai erikoismerkit eivät aiheuta SQL-injektiota
    qdf.Parameters("[pDept]") = IIf(IsNull(Me.Parent.Department.Value), "", Me.Parent.Department.Value)
    qdf.Parameters("[pEqType]") = IIf(IsNull(Me.Parent.EqType.Value), "", Me.Parent.EqType.Value)
    qdf.Parameters("[pEqSeq]") = IIf(IsNull(Me.Parent.EqSeq.Value), "", Me.Parent.EqSeq.Value)

    Set Taulukko = qdf.OpenRecordset(dbOpenSnapshot)

    If Taulukko.EOF Or IsNull(Taulukko!MaxSuffix) Then
        Uusi = "01"  ' Ei aiempia tietueita tälle kombinaatiolle
    Else
        ' Lisätään 1 suurimpaan arvoon — turvallinen myös poistettujen rivien jälkeen
        Uusi = Format$(CStr(Val(Taulukko!MaxSuffix) + 1), "00")
    End If

    Me.Suffix.Value = Uusi

    Taulukko.Close
    Set Taulukko = Nothing
    Set qdf = Nothing
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
Dim RS As DAO.Recordset  ' Updated 2025-11-11: Added DAO prefix for early binding
Dim UusiSf As String
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


