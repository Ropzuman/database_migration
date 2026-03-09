Option Compare Database
Option Explicit

Private Sub Command1_Click()
On Error GoTo Err_Command1_Click


    Screen.PreviousControl.SetFocus
    DoCmd.FindNext

Exit_Command1_Click:
    Exit Sub

Err_Command1_Click:
    MsgBox Err.Description
    Resume Exit_Command1_Click
    
End Sub


Private Sub Command10_Click()
DoCmd.OpenQuery "Function_zdetails15_loopsprintselect", acNormal, acEdit
End Sub

Private Sub Command11_Click()
Dim DB As DAO.Database
Dim Taulu As Recordset
Dim i As Integer
Dim j As Integer
Set DB = CurrentDb
Set Taulu = DB.OpenRecordset(KaikkiTaulukot.Value)
Do While Not Taulu.EOF
  For i = 0 To Taulu.Fields.Count - 1
    If Taulu.Fields(i) = "" Then
      Taulu.Edit
      Taulu.Fields(i) = Null
      Taulu.Update
    End If
  Next i
  Taulu.MoveNext
Loop
MsgBox "Valmis"
Taulu.Close
DB.Close
Set Taulu = Nothing
Set DB = Nothing
End Sub

Private Sub Command3_Click()
On Error GoTo Err_Command3_Click
    
    Dim tbl As Recordset
    Dim q As String
    Dim q1 As String
    Dim qq As String
    Dim qq1 As String
    Dim i As String
    
'    DoCmd.OpenQuery "FunctionDetails01", acNormal, acEdit
    Set tbl = CurrentDb.OpenRecordset("SELECT [AreaCode] & [Loopno] & [pageno] AS Loop, zDetails.Blockplace, zDetails.AllDevTag FROM zDetails ORDER BY [AreaCode] & [Loopno] & [pageno], zDetails.AllDevTag;")
  
    tbl.MoveFirst
  Do
    i = 14
    Do
    q = tbl!Loop
    tbl.Edit
    tbl!blockplace = i
    tbl.Update
    i = i + 1
    tbl.MoveNext
    If tbl.EOF Then Exit Do
    qq = tbl!Loop
    If q <> qq Then Exit Do
    Loop
    
  Loop Until tbl.EOF
  
  'DoCmd.OpenQuery "Tag_1-7_ZDetailsKuviin", acNormal, acEdit
  
  MsgBox "Homma hoidettu"
Exit_Command3_Click:
    Exit Sub

Err_Command3_Click:
    MsgBox Err.Description
    Resume Exit_Command3_Click
    
End Sub

Private Sub Command4_Click()
DoCmd.OpenQuery "Function_zDetails03_lis�tt�v�t", acNormal, acEdit
DoCmd.OpenQuery "Function_zDetails04_FileID", acNormal, acEdit
DoCmd.OpenQuery "Function_zDetails09_update", acNormal, acEdit
DoCmd.OpenQuery "Function_zDetails14_Fileid2", acNormal, acEdit
End Sub


Private Sub Command5_Click()
'tuhoaa zdetails taulukosta tiedot joita ei ole en��n instrumentindexsiss�
Dim DB As Database
Dim tbl As Recordset
Dim tbl2 As Recordset
Dim qry As Recordset
Dim sSQL As String
Dim ssSQL As String
Dim acode As String
Dim LoopNo As String
Dim LoopSymb As String
Dim Suffix As String
Dim L As String

L = Chr(34)

Set DB = CurrentDb

sSQL = ""
sSQL = sSQL & "SELECT zDetails.AreaCode, zDetails.LoopNo, zDetails.LoopSymb, zDetails.Suffix"
sSQL = sSQL & " FROM zDetails LEFT JOIN Function_zdetails01 ON (zDetails.Suffix = Function_zdetails01.Suffix) AND (zDetails.LoopSymb = Function_zdetails01.LoopSymb) AND (zDetails.LoopNo = Function_zdetails01.LoopNo) AND (zDetails.AreaCode = Function_zdetails01.AreaCode)"
sSQL = sSQL & " WHERE (((Function_zdetails01.AreaCode) Is Null) AND ((Function_zdetails01.LoopNo) Is Null) AND ((Function_zdetails01.LoopSymb) Is Null) AND ((Function_zdetails01.Suffix) Is Null));"


Set tbl = DB.OpenRecordset(sSQL)


If Not tbl.EOF Then
tbl.MoveFirst
Do Until tbl.EOF
    DoCmd.Hourglass True

    acode = tbl!AreaCode
    LoopNo = tbl!LoopNo
    Suffix = tbl!Suffix
    LoopSymb = tbl!LoopSymb

    ssSQL = ""
    ssSQL = ssSQL & "SELECT zDetails.AreaCode, zDetails.LoopNo, zDetails.LoopSymb, zDetails.Suffix, zDetails.CirDWG, zDetails.Pageno, zDetails.AllDevTag, zDetails.Name, zDetails.Manufacturer, zDetails.Type, zDetails.ProcConn, zDetails.Faixatotal, zDetails.Faixa, zDetails.Blockname, zDetails.Blockplace, zDetails.[Select]"
    ssSQL = ssSQL & " FROM zDetails"
    ssSQL = ssSQL & " WHERE (((zDetails.AreaCode)=" & L & acode & L & ") AND ((zDetails.LoopNo)=" & L & LoopNo & L & ") AND ((zDetails.LoopSymb)=" & L & LoopSymb & L & ") AND ((zDetails.Suffix)=" & L & Suffix & L & "));"
    Set tbl2 = DB.OpenRecordset(ssSQL)

       tbl2.Delete
       tbl2.MovePrevious
 '
 
       tbl.MoveNext
Loop
DoCmd.Hourglass False

tbl.Close
Else
End If
End Sub

Private Sub Command6_Click()
DoCmd.OpenQuery "Function_zDetails06_clearprint", acNormal, acEdit
End Sub

Private Sub Command7_Click()
DoCmd.OpenQuery "Function_zDetails05_printselect", acNormal, acEdit
End Sub

Private Sub Command8_Click()
DoCmd.OpenQuery "Function_Zdetails08_Clearzdetailskuviin", acNormal, acEdit
End Sub

Private Sub Command9_Click()
DoCmd.OpenQuery "Tag_1-7_ZDetailsKuviin", acNormal, acEdit
DoCmd.OpenQuery "Tag1_CR0-1_FC0-1_F0-2_PWC0_CCM0-1_JB0120", acNormal, acEdit
DoCmd.OpenQuery "Tag1_DCS", acNormal, acEdit
DoCmd.OpenQuery "Tag1_F10-11", acNormal, acEdit
DoCmd.OpenQuery "Tag2_CCM0-1_PWC0", acNormal, acEdit
DoCmd.OpenQuery "Tag1_JB10-12", acNormal, acEdit
DoCmd.OpenQuery "TAG1_FC10-11", acNormal, acEdit
'DoCmd.OpenQuery "Tag_1-7_ZDetailsKuviin", acNormal, acEdit
End Sub

Private Sub Form_Load()
  Dim DB As DAO.Database
  Dim i As Long
  Dim Taulukko As String
  Dim TJoukko As String
  Set DB = CurrentDb
  For i = 0 To DB.TableDefs.Count - 1
    Taulukko = DB.TableDefs(i).Name
    If Left(Taulukko, 4) <> "USys" And Left(Taulukko, 4) <> "MSys" And Left(Taulukko, 3) <> "Dev" And InStr(Taulukko, "old") = 0 Then
      TJoukko = TJoukko & Taulukko & ";"
    End If
  Next i
  KaikkiTaulukot.RowSource = TJoukko
  KaikkiTaulukot.Value = Left(TJoukko, InStr(TJoukko, ";") - 1)
End Sub
