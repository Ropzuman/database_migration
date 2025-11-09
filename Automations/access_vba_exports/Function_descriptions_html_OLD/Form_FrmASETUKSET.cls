Option Compare Database

Private Sub Command11_Click()
DoCmd.OpenForm "frmMUOKKAUS"
DoCmd.Close , "FrmASETUKSET"
End Sub

Private Sub Command2_Click()
'Linkkaa taulukot
   Linkkaa "Alarms", TINSTRU
   Linkkaa "Loops", TINSTRU
   Linkkaa "DevTable", TINSTRU
   Linkkaa "DevSpec", TINSTRU
   Linkkaa "TAGID_exception", TINSTRU
   Linkkaa "MAINEQ", TMAINEQ
   Linkkaa "DRIVES", TMAINEQ
'   Linkkaa "CONTROL", TFUNCDIAG
   MsgBox "Taulukot linkattu", vbOKOnly, "Linkkaa taulukot"
 
End Sub
Private Sub Linkkaa(Taulukko As String, KANTA As String)
Dim tdf As TableDef
Dim i As Integer
    For i = 0 To CurrentDb.TableDefs.Count - 1
      If LCase(CurrentDb.TableDefs(i).Name) = LCase(Taulukko) Then
        CurrentDb.Execute "DROP TABLE " & Taulukko
        Exit For
      End If
    Next i
    Set tdf = CurrentDb.CreateTableDef(Taulukko)
    With tdf
        .Connect = ";DATABASE=" & KANTA
        .SourceTableName = Taulukko
    End With
    CurrentDb.TableDefs.Append tdf
    CurrentDb.TableDefs.Refresh
End Sub
Private Sub Form_Load()
Dim POLKU As String
  POLKU = CurrentDb.Name
  POLKU = Left(POLKU, InStrRev(POLKU, "\"))
  TINSTRU.Value = POLKU & "instru3.accdb"
  TMAINEQ.Value = POLKU & "MAINEQ.accdb"
  TFUNCDIAG.Value = POLKU & "FunctionDiagrams.accdb"
  Call SniffUser
End Sub
