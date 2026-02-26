Option Compare Database
Private Sub NEWVALUE_DblClick(Cancel As Integer)
Dim DB As DAO.Database
Dim Taulu As DAO.Recordset
Dim ALUE As String
Dim LooppiSymb As String
Dim LooppiNo As String
Dim Suffiksi As String
Dim Blokki As String
Dim tagi As String
Dim taginarvo As String
Dim tagiid As String

Set DB = CurrentDb
Set Taulu = DB.OpenRecordset("FunBlockAttr", dbOpenTable)
Blokki = Me!BLOCK
tagi = Me!Atag
taginarvo = Me!NEWVALUE
ALUE = Forms!funktiokaavio.FUNC.Form!AreaCode
LooppiNo = Forms!funktiokaavio.FUNC.Form!LoopNo
LooppiSymb = Forms!funktiokaavio.FUNC.Form!LoopSymb
Suffiksi = Forms!funktiokaavio.FUNC.Form!Suffix
tagiid = Forms!funktiokaavio.FUNC.Form!Text68
Taulu.AddNew
Taulu.Fields("Areacode") = ALUE
Taulu.Fields("loopno") = LooppiNo
Taulu.Fields("loopsymb") = LooppiSymb
Taulu.Fields("suffix") = Suffiksi
Taulu.Fields("fblock") = Blokki
Taulu.Fields("atag") = tagi
Taulu.Fields("value") = taginarvo
Taulu.Fields("TagID") = tagiid
Taulu.Update

MsgBox "Uusi arvo lisätty"

End Sub

Private Sub NEWVALUE_Enter()
  If UCase$(Left$(Atag, 5)) = "FTAGI" Then
      NEWVALUE.RowSource = "SELECT POS FROM POS_List;"
  ElseIf UCase$(Left$(Atag, 7)) = "REFFROM" Or UCase$(Left$(Atag, 5)) = "REFTO" Then
      NEWVALUE.RowSource = "SELECT FunBlockAttr.Value FROM FunBlockAttr GROUP BY FunBlockAttr.Value, FunBlockAttr.Atag HAVING (((FunBlockAttr.Atag) Like 'REFTO*'));"
  Else
      NEWVALUE.RowSource = "SELECT FunBlockAttr.Value FROM FunBlockAttr GROUP BY FunBlockAttr.Value, FunBlockAttr.Atag HAVING (((FunBlockAttr.Atag) = '" & Atag & "'));"
  End If
  NEWVALUE.Requery
End Sub

Private Sub SelectValue_AfterUpdate()
  NEWVALUE.Value = SelectValue.Value
End Sub

Private Sub SelectValue_Enter()
Dim ALUE As String
Dim looppi As String
Dim tagi As String
Dim Sfx As String

  SelectValue.RowSource = "SELECT Number, POS FROM POS_List;"
  If UCase$(Atag) = "RANGE" Then
      ALUE = Me.Parent.Parent.AreaCode
      looppi = Me.Parent.Parent.LoopNo
      tagi = "T" 'Right$(Me.Parent.LoopSymb, 1)
      Sfx = Me.Parent.Suffix
      
      SelectValue.BoundColumn = 1
      SelectValue.ColumnCount = 1
      SelectValue.RowSource = "SELECT LRange & '-' & URange & EngUnits as RANGE From Devtable WHERE AreaCOde = '" & ALUE & "' AND TagID = '" & tagi & "' AND LoopNo = '" & looppi & "' AND  Suffix = '" & Sfx & "';"
      SelectValue.ColumnWidths = ""
    
  ElseIf UCase$(Left$(Atag, 5)) = "FTAGI" Then
      SelectValue.BoundColumn = 2
      SelectValue.ColumnCount = 2
      SelectValue.RowSource = "SELECT Number, POS FROM POS_List;"
      SelectValue.ColumnWidths = "567;1701"
  ElseIf UCase$(Left$(Atag, 7)) = "REFFROM" Or UCase$(Left$(Atag, 5)) = "REFTO" Then
      SelectValue.BoundColumn = 1
      SelectValue.ColumnCount = 1
      SelectValue.RowSource = "SELECT FunBlockAttr.Value FROM FunBlockAttr GROUP BY FunBlockAttr.Value, FunBlockAttr.Atag HAVING (((FunBlockAttr.Atag) Like 'REFTO*'));"
      SelectValue.ColumnWidths = ""
  Else
      SelectValue.BoundColumn = 1
      SelectValue.ColumnCount = 1
      SelectValue.RowSource = "SELECT FunBlockAttr.Value FROM FunBlockAttr GROUP BY FunBlockAttr.Value, FunBlockAttr.Atag HAVING (((FunBlockAttr.Atag) = '" & Atag & "'));"
      SelectValue.ColumnWidths = ""
  End If
  SelectValue.Requery
  Me.Refresh

End Sub
