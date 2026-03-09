Option Compare Database
Option Explicit

Private Sub Form_BeforeInsert(Cancel As Integer)
    ' Haetaan SUURIN olemassa oleva Suffix MAX-funktiolla (turvallinen, ei rip RecordCount-bug)
    ' Parametrisoitu kysely estää myös SQL-injektion
    Dim qdf As DAO.QueryDef
    Dim Taulukko As DAO.Recordset
    Dim Uusi As String
    Dim SqlStr As String

    SqlStr = "PARAMETERS [pDept] Text, [pEqType] Text, [pEqSeq] Text; " & _
             "SELECT MAX(Suffix) AS MaxSuffix FROM DRIVES " & _
             "WHERE Department = [pDept] AND Eqtype = [pEqType] AND EqSeq = [pEqSeq];"

    Set qdf = CurrentDb.CreateQueryDef("", SqlStr)
    qdf.Parameters("[pDept]") = IIf(IsNull(Me.Parent.Department.Value), "", Me.Parent.Department.Value)
    qdf.Parameters("[pEqType]") = IIf(IsNull(Me.Parent.EqType.Value), "", Me.Parent.EqType.Value)
    qdf.Parameters("[pEqSeq]") = IIf(IsNull(Me.Parent.EqSeq.Value), "", Me.Parent.EqSeq.Value)

    Set Taulukko = qdf.OpenRecordset(dbOpenSnapshot)

    If Taulukko.EOF Or IsNull(Taulukko!MaxSuffix) Then
        Uusi = "01"  ' Ei aiempia tietueita
    Else
        ' MAX(Suffix) + 1 — toimii oikein myös poistojen jälkeen
        Uusi = Format$(CStr(Val(Taulukko!MaxSuffix) + 1), "00")
    End If

    Me.Suffix.Value = Uusi

    Taulukko.Close
    Set Taulukko = Nothing
    Set qdf = Nothing
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
        MsgBox "Huom!@Arvo pitää olla synkroninen nopeus! @Esim. 3000, 1500, 1000 tai 750. @Paitsi vaihdemoottoreilla (GM)", vbOKOnly
        Me.Rpm.SetFocus
    End If
End If
End If
End Sub



