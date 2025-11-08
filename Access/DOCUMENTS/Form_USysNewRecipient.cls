Option Compare Database   'Use database order for string comparisons
Option Explicit
Private Sub BExit_Click()
  DoCmd.Close
End Sub
Private Sub Lista_DblClick(Cancel As Integer)
  Dim DB As Database
  Dim taulu As Recordset
  Dim SQL As String
  
  Set DB = CurrentDb
  SQL = "SELECT * FROM USysRecipients WHERE ID=" & Lista.ItemData(Lista.ItemsSelected(0))
  Set taulu = DB.OpenRecordset(SQL)
  With taulu
    TName = .Fields("name")
    TUnit = .Fields("unit")
    TDescription = .Fields("description")
    TCompany = .Fields("company")
    TCopies = .Fields("copies")
    TAddress = .Fields("address")
    Temail = .Fields("email")
    .Close
  End With
  Set DB = Nothing
  Set taulu = Nothing
End Sub
Private Sub Valmis_Click()
'T�m� lis�� uuden distribuution USysDISTRIBUTION taulukkoon ja
'Vastaanottajat USysRecipByDistr taulukkoon
If IsNull(TName) Then
 MsgBox "You have to give name for the recipient!", vbCritical, "Error!"
ElseIf IsNull(TAddress) And IsNull(Temail) Then
 MsgBox "You have to give some address to the recipient!", vbCritical, "Error!"
Else
  Dim DB As Database
  Dim taulu As Recordset
  Set DB = CurrentDb
  Set taulu = DB.OpenRecordset("USysRecipients")
  With taulu
    .AddNew
    If IsNull(TCopies) = False Then
      .Fields("To Distribution") = True
    End If
    .Fields("Name") = TName
    .Fields("Unit") = TUnit
    .Fields("Description") = TDescription
    .Fields("Company") = TCompany
    .Fields("Copies") = TCopies
    .Fields("Address") = TAddress
    .Fields("email") = Temail
    .Update
    .Close
  End With
  Set taulu = Nothing
  Set DB = Nothing
  DoCmd.Close
End If
End Sub