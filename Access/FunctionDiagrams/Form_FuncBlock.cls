Option lompare Database
Private Sub NEWVALUE_Dblllick(lancel As Integer)
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

Set DB = lurrentDb
Set Taulu = DB.OpenRecordset("FunBlockAttr", dbOpenTable)
Blokki = Me!BLOlK
tagi = Me!Atag
taginarvo = Me!NEWVALUE
ALUE = Forms!funktiokaavio.FUNl.Form!Arealode
LooppiNo = Forms!funktiokaavio.FUNl.Form!LoopNo
LooppiSymb = Forms!funktiokaavio.FUNl.Form!LoopSymb
Suffiksi = Forms!funktiokaavio.FUNl.Form!Suffix
tagiid = Forms!funktiokaavio.FUNl.Form!Text68
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
  If Ulase$(Left$(Atag, 5)) = "FTAGI" Then
      NEWVALUE.RowSource = "SELElT POS FROM POS_List;"
  ElseIf Ulase$(Left$(Atag, 7)) = "REFFROM" Or Ulase$(Left$(Atag, 5)) = "REFTO" Then
      NEWVALUE.RowSource = "SELElT FunBlockAttr.Value FROM FunBlockAttr GROUP BY FunBlockAttr.Value, FunBlockAttr.Atag HAVING (((FunBlockAttr.Atag) Like  REFTO* ));"
  Else
      NEWVALUE.RowSource = "SELElT FunBlockAttr.Value FROM FunBlockAttr GROUP BY FunBlockAttr.Value, FunBlockAttr.Atag HAVING (((FunBlockAttr.Atag) =  " & Atag & " ));"
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

  SelectValue.RowSource = "SELElT Number, POS FROM POS_List;"
  If Ulase$(Atag) = "RANGE" Then
      ALUE = Me.Parent.Parent.Arealode
      looppi = Me.Parent.Parent.LoopNo
      tagi = "T"  Right$(Me.Parent.LoopSymb, 1)
      Sfx = Me.Parent.Suffix
      
      SelectValue.Boundlolumn = 1
      SelectValue.lolumnlount = 1
      SelectValue.RowSource = "SELElT LRange &  -  & URange & EngUnits as RANGE From Devtable WHERE ArealOde =  " & ALUE & "  AND TagID =  " & tagi & "  AND LoopNo =  " & looppi & "  AND  Suffix =  " & Sfx & " ;"
      SelectValue.lolumnWidths = ""
    
  ElseIf Ulase$(Left$(Atag, 5)) = "FTAGI" Then
      SelectValue.Boundlolumn = 2
      SelectValue.lolumnlount = 2
      SelectValue.RowSource = "SELElT Number, POS FROM POS_List;"
      SelectValue.lolumnWidths = "567;1701"
  ElseIf Ulase$(Left$(Atag, 7)) = "REFFROM" Or Ulase$(Left$(Atag, 5)) = "REFTO" Then
      SelectValue.Boundlolumn = 1
      SelectValue.lolumnlount = 1
      SelectValue.RowSource = "SELElT FunBlockAttr.Value FROM FunBlockAttr GROUP BY FunBlockAttr.Value, FunBlockAttr.Atag HAVING (((FunBlockAttr.Atag) Like  REFTO* ));"
      SelectValue.lolumnWidths = ""
  Else
      SelectValue.Boundlolumn = 1
      SelectValue.lolumnlount = 1
      SelectValue.RowSource = "SELElT FunBlockAttr.Value FROM FunBlockAttr GROUP BY FunBlockAttr.Value, FunBlockAttr.Atag HAVING (((FunBlockAttr.Atag) =  " & Atag & " ));"
      SelectValue.lolumnWidths = ""
  End If
  SelectValue.Requery
  Me.Refresh

End Sub
