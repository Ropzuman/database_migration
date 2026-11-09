Option Compare Database
Option Explicit
'==========================================================================
' LOMAKE   : Form_FuncBlock
' SOVELLUS : FunctionDiagrams — Funktiolohkon attribuuttien muokkaus
' KUVAUS   : Alilomake funktiopiirien yksittäisten lohkoattribuuttien
'            käsittelyyn. NEWVALUE_DblClick tallentaa käyttäjän syöttämän
'            uuden arvon FunBlockAttr-tauluun. NEWVALUE_Enter ja
'            SelectValue_Enter päivittävät listalaatikon lähteen attribuutin
'            tyypin mukaan (FTAGI, REFFROM, REFTO tai yleinen arvo).
' HUOM     : DAO-oliot vapautetaan sekä onnistumis- että virhepolulla.
' PÄIVITETTY: 2026-11-09
'==========================================================================

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

On Error GoTo Virhe
If IsNull(Me!NEWVALUE) Then
  MsgBox "Anna ensin uusi arvo.", vbExclamation
  Exit Sub
End If

Set DB = CurrentDb
' dbOpenDynaset: dbOpenTable ei toimi linkitetylle taululle
Set Taulu = DB.OpenRecordset("FunBlockAttr", dbOpenDynaset)
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

' Päivitetään näkyvä attribuuttilista heti tallennuksen jälkeen, jotta uusi rivi näkyy lomakkeella
Me.Requery

' Vapautetaan DAO-resurssit muistivuodon ehkäisemiseksi
Taulu.Close
Set Taulu = Nothing
Set DB = Nothing

MsgBox "Uusi arvo lisätty"

Exit Sub

Virhe:
  Debug.Print "*** VIRHE NEWVALUE_DblClick: " & Err.Number & " - " & Err.Description
  MsgBox "Virhe uuden arvon tallennuksessa: " & Err.Description, vbCritical
  On Error Resume Next
  If Not Taulu Is Nothing Then Taulu.Close
  Set Taulu = Nothing
  Set DB = Nothing
End Sub

Private Sub NEWVALUE_Enter()
  Dim sAtag As String
  sAtag = Replace(Nz(Atag, ""), "'", "''")  ' Heittomerkit kaksinkertaistetaan dynaamista SQL-ehtoa varten
  If UCase$(Left$(sAtag, 5)) = "FTAGI" Then
      NEWVALUE.RowSource = "SELECT POS FROM POS_List;"
  ElseIf UCase$(Left$(sAtag, 7)) = "REFFROM" Or UCase$(Left$(sAtag, 5)) = "REFTO" Then
      NEWVALUE.RowSource = "SELECT FunBlockAttr.Value FROM FunBlockAttr GROUP BY FunBlockAttr.Value, FunBlockAttr.Atag HAVING (((FunBlockAttr.Atag) Like 'REFTO*'));"
  Else
      NEWVALUE.RowSource = "SELECT FunBlockAttr.Value FROM FunBlockAttr GROUP BY FunBlockAttr.Value, FunBlockAttr.Atag HAVING (((FunBlockAttr.Atag) = '" & sAtag & "'));"
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
Dim sAtag As String

  sAtag = Replace(Nz(Atag, ""), "'", "''")  ' Heittomerkit kaksinkertaistetaan dynaamista SQL-ehtoa varten
  SelectValue.RowSource = "SELECT Number, POS FROM POS_List;"
  If UCase$(Nz(Atag, "")) = "RANGE" Then
      ALUE   = Replace(Me.Parent.Parent.AreaCode, "'", "''")
      looppi = Replace(Me.Parent.Parent.LoopNo,   "'", "''")
      tagi = "T" ' Käytetään kiinteää T-tunnistetyyppiä alueeseen sidotuissa hauissa
      Sfx    = Replace(Me.Parent.Suffix,           "'", "''")
      
      SelectValue.BoundColumn = 1
      SelectValue.ColumnCount = 1
      SelectValue.RowSource = "SELECT LRange & '-' & URange & EngUnits as RANGE From Devtable WHERE AreaCOde = '" & ALUE & "' AND TagID = '" & tagi & "' AND LoopNo = '" & looppi & "' AND  Suffix = '" & Sfx & "';"
      SelectValue.ColumnWidths = ""
    
  ElseIf UCase$(Left$(Nz(Atag, ""), 5)) = "FTAGI" Then
      SelectValue.BoundColumn = 2
      SelectValue.ColumnCount = 2
      SelectValue.RowSource = "SELECT Number, POS FROM POS_List;"
      SelectValue.ColumnWidths = "567;1701"
  ElseIf UCase$(Left$(Nz(Atag, ""), 7)) = "REFFROM" Or UCase$(Left$(Nz(Atag, ""), 5)) = "REFTO" Then
      SelectValue.BoundColumn = 1
      SelectValue.ColumnCount = 1
      SelectValue.RowSource = "SELECT FunBlockAttr.Value FROM FunBlockAttr GROUP BY FunBlockAttr.Value, FunBlockAttr.Atag HAVING (((FunBlockAttr.Atag) Like 'REFTO*'));"
      SelectValue.ColumnWidths = ""
  Else
      SelectValue.BoundColumn = 1
      SelectValue.ColumnCount = 1
      SelectValue.RowSource = "SELECT FunBlockAttr.Value FROM FunBlockAttr GROUP BY FunBlockAttr.Value, FunBlockAttr.Atag HAVING (((FunBlockAttr.Atag) = '" & sAtag & "'));"
      SelectValue.ColumnWidths = ""
  End If
  SelectValue.Requery
  Me.Refresh

End Sub
