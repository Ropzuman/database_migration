Option Compare Database
Option Explicit

Private Sub Form_BeforeInsert(Cancel As Integer)
    ' Haetaan suurin olemassa oleva Suffix parametrisoidulla kyselyllä (SQL-injektio estetty)
    Dim qdf As DAO.QueryDef
    Dim Taulukko As DAO.Recordset
    Dim Uusi As String
    Dim SqlStr As String

    SqlStr = "PARAMETERS [pDept] Text, [pEqType] Text, [pEqSeq] Text; " & _
             "SELECT MAX(Suffix) AS MaxSuffix FROM CONVEYORS " & _
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

Private Sub RevHistory_Click()
 MuutaRevisio
End Sub


