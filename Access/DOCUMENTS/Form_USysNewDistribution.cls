Option Compare Database   'Use database order for string comparisons
Option Explicit
Private Sub BExit_Click()
  DoCmd.Close
End Sub

'''
' Form Load Event - Initialize distribution form
' OPTIMIZED 2025-11-09: Added explicit DAO typing, error handling
'''
Private Sub Form_Load()
  Dim DB As DAO.Database
  
  On Error GoTo ErrorHandler
  
  Set DB = CurrentDb
  DB.Execute "UPDATE USysRecipients SET USysRecipients.[To Distribution] = No;"
  TDate.Value = Date
  
  DB.Close
  Set DB = Nothing
  Exit Sub
  
ErrorHandler:
  MsgBox "Error initializing form: " & Err.Description, vbExclamation, "Form Load Error"
  If Not DB Is Nothing Then
    DB.Close
    Set DB = Nothing
  End If
End Sub

Private Sub Lisaa_Click()
'Lisää uusi vastaanottaja
   DoCmd.OpenForm "USysNewRecipient", acNormal, , , , acDialog
   Lista.Requery
End Sub

'''
' Lista Double Click - Toggle recipient selection for distribution
' OPTIMIZED 2025-11-09: Added error handling, explicit DAO typing, null-checking
'''
Private Sub Lista_DblClick(Cancel As Integer)
  Dim DB As DAO.Database
  Dim taulu As DAO.Recordset
  Dim SQL As String
  Dim Oletus As String
  Dim Kopioita As String
  
  On Error GoTo ErrorHandler
  
  Set DB = CurrentDb
  SQL = "SELECT * FROM USysRecipients WHERE ID=" & Lista.ItemData(Lista.ItemsSelected(0))
  Set taulu = DB.OpenRecordset(SQL, dbOpenDynaset)
  
  If Not taulu.EOF Then
    taulu.Edit
    If taulu.Fields("To Distribution") Then
       taulu.Fields("To Distribution") = False
    Else
      Oletus = IIf(IsNull(taulu.Fields("Copies")), "", taulu.Fields("Copies"))
      Kopioita = InputBox("How many copies are been sent? E.g. 4 paper copies and e-mail are marked as ""4+e"".", "Copies", Oletus)
      If Kopioita <> "" Then
        taulu.Fields("To Distribution") = True
        taulu.Fields("Copies") = Kopioita
      End If
    End If
    taulu.Update
  End If
  
  Lista.Requery
  
Cleanup:
  If Not taulu Is Nothing Then
    If taulu.State = 1 Then taulu.Close
    Set taulu = Nothing
  End If
  If Not DB Is Nothing Then
    DB.Close
    Set DB = Nothing
  End If
  Exit Sub
  
ErrorHandler:
  MsgBox "Error updating recipient: " & Err.Description, vbExclamation, "Update Error"
  Resume Cleanup
End Sub

Private Sub Valmis_Click()
' Tämä lisää uuden distribuution USysDISTRIBUTION taulukkoon ja
' Vastaanottajat USysRecipByDistr taulukkoon
Dim TRecipients As String
'Dim FSO As New FileSystemObject
Dim i As Integer
Dim SQL As String
Dim DB As Database
Dim taulu As Recordset
Dim Taulu2 As Recordset
Dim Hakemisto As String
Dim OutboxPath As String
Dim DistrID As Long
Set DB = CurrentDb
'Haetaan ensin projektin Outbox laatikon sijainti
'SQL = "SELECT Name, Path FROM ProjPaths WHERE Name Like 'OUTBOX'"
'Set Taulu = DB.OpenRecordset(SQL)
'OutboxPath = Taulu.Fields("Path")

SQL = "SELECT USysRecipients.*, USysRecipients.[To Distribution] FROM USysRecipients WHERE (((USysRecipients.[To Distribution])=Yes));"
Set taulu = DB.OpenRecordset(SQL)
If taulu.EOF Then
 MsgBox "You have to select recipients for Distribution!", vbCritical, "Error!"
ElseIf IsNull(DistrNo) Then
 MsgBox "You have to give number for distribution!", vbCritical, "Error!"
ElseIf IsNull(TDate) Then
 MsgBox "You have to select date!", vbCritical, "Error!"
ElseIf IsNull(Desc) Then
 MsgBox "You have to put description!", vbCritical, "Error!"
ElseIf IsNull(TReason) Then
 MsgBox "Distribution must have reason! Select reason and try again", vbCritical, "Error!"
Else
  SQL = "SELECT * FROM UsysDISTRIBUTION WHERE No='" & DistrNo & "'"
  Set taulu = DB.OpenRecordset(SQL)
  If taulu.EOF Then 'Ei vielä yhtään saman ID:n omaavaa Distribuutiota
    Hakemisto = aReplace(DistrNo) & "\"
'    If FSO.FolderExists(OutboxPath & Hakemisto) = False Then
'      FSO.CreateFolder (OutboxPath & Hakemisto)
'    End If
    With taulu
      .AddNew
      .Fields("No") = DistrNo
      .Fields("Description") = Desc
      .Fields("Date") = TDate
      .Fields("Sender") = TSender
      .Fields("Reason") = TReason
      .Fields("Note") = Note
      DistrID = .Fields("ID")
      .Update
      .Close
    End With
    SQL = "SELECT USysRecipients.*, USysRecipients.[To Distribution] FROM USysRecipients WHERE (((USysRecipients.[To Distribution])=Yes));"
    Set taulu = DB.OpenRecordset(SQL)
    Set Taulu2 = DB.OpenRecordset("USysRecipByDistr")
    Do While Not taulu.EOF
      With Taulu2
        .AddNew
        .Fields("DistrID") = DistrID
        .Fields("Copies") = taulu.Fields("Copies")
        .Fields("Unit") = taulu.Fields("Unit")
        .Fields("Company") = taulu.Fields("Company")
        .Fields("Name") = taulu.Fields("Name")
        .Fields("Description") = taulu.Fields("Description")
        .Fields("Address") = taulu.Fields("Address")
        .Fields("email") = taulu.Fields("email")
        .Update
      End With
      taulu.MoveNext
    Loop
    Common = OutboxPath & Hakemisto
'    DoCmd.OpenForm "UsysAddedDistr", acNormal, , , , acDialog
  Else
    MsgBox "Distirubution number """ & DistrNo & """ were already selected! Please select another number.", vbCritical, "Error!"
  End If
  taulu.Close
  DoCmd.Close
End If
Set taulu = Nothing
Set DB = Nothing
End Sub