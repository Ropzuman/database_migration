Option Compare Database
Option Explicit

#If VBA7 Then
Private Declare PtrSafe Function CoCreateGuid Lib "ole32.dll" (ByRef Guid As Byte) As Long
#Else
Private Declare Function CoCreateGuid Lib "ole32.dll" (ByRef Guid As Byte) As Long
#End If

Private mSessionID As String
Private mTallennettu As Boolean

Private Sub BExit_Click()
  DoCmd.Close
End Sub

Private Sub Form_Load()
  Dim DB As DAO.Database
  Dim taulu As DAO.Recordset

  On Error GoTo ErrorHandler
  mSessionID = LuoIstuntotunniste()
  Set DB = CurrentDb
  Set taulu = DB.OpenRecordset("SELECT SessionID FROM USysDistrSelection WHERE 1=0", dbOpenSnapshot)
  taulu.Close
  Set taulu = Nothing
  Set DB = Nothing
  Me.TDate.Value = Date
  PaivitaVastaanottajalista
  Exit Sub

ErrorHandler:
  MsgBox "Error initializing distribution: " & Err.Description, vbCritical, "Form Load Error"
  On Error Resume Next
  If Not taulu Is Nothing Then taulu.Close
  Set taulu = Nothing
  Set DB = Nothing
  mSessionID = ""
  Me.Valmis.Enabled = False
  Me.Lista.Enabled = False
  On Error GoTo 0
End Sub

Private Function LuoIstuntotunniste() As String
  Dim Tavut(0 To 15) As Byte
  Dim i As Long
  Dim Tunniste As String
  Dim Tulos As Long

  Tulos = CoCreateGuid(Tavut(0))
  If Tulos <> 0 Then
    Err.Raise vbObjectError + 1200, "LuoIstuntotunniste", "A unique session identifier could not be created. Error: " & Tulos
  End If
  For i = 0 To 15
    Tunniste = Tunniste & Right$("0" & Hex$(Tavut(i)), 2)
  Next i
  LuoIstuntotunniste = Tunniste
End Function

Private Sub PaivitaVastaanottajalista()
  Me.Lista.RowSource = _
    "SELECT R.ID, R.Unit, IIf(S.RecipientID Is Null, False, True) AS [To Distribution], " & _
    "IIf(S.RecipientID Is Null, R.Copies, S.Copies) AS Copies, R.Company, R.Name, " & _
    "R.Description, R.email, R.Address FROM USysRecipients AS R " & _
    "LEFT JOIN (SELECT RecipientID, Copies FROM USysDistrSelection WHERE SessionID='" & mSessionID & "') AS S " & _
    "ON R.ID=S.RecipientID ORDER BY R.ID DESC;"
End Sub

Private Sub Lisaa_Click()
  DoCmd.OpenForm "USysNewRecipient", acNormal, , , , acDialog
  Me.Lista.Requery
End Sub

Private Sub Lista_DblClick(Cancel As Integer)
  Dim DB As DAO.Database
  Dim taulu As DAO.Recordset
  Dim VastaanottajaID As Long
  Dim SQL As String
  Dim Oletus As String
  Dim Kopioita As String

  On Error GoTo ErrorHandler
  If mSessionID = "" Then Err.Raise vbObjectError + 1201, "Lista_DblClick", "The distribution session is not initialized."
  If Me.Lista.ItemsSelected.Count = 0 Then Exit Sub
  VastaanottajaID = CLng(Me.Lista.ItemData(Me.Lista.ItemsSelected(0)))
  Set DB = CurrentDb
  SQL = "SELECT RecipientID FROM USysDistrSelection WHERE SessionID='" & mSessionID & "' AND RecipientID=" & VastaanottajaID
  Set taulu = DB.OpenRecordset(SQL, dbOpenSnapshot)

  If Not taulu.EOF Then
    taulu.Close
    Set taulu = Nothing
    DB.Execute "DELETE FROM USysDistrSelection WHERE SessionID='" & mSessionID & "' AND RecipientID=" & VastaanottajaID, dbFailOnError
  Else
    taulu.Close
    Set taulu = Nothing
    Set taulu = DB.OpenRecordset("SELECT Copies FROM USysRecipients WHERE ID=" & VastaanottajaID, dbOpenSnapshot)
    If taulu.EOF Then Err.Raise vbObjectError + 1202, "Lista_DblClick", "The selected recipient no longer exists."
    Oletus = Nz(taulu.Fields("Copies").Value, "")
    taulu.Close
    Set taulu = Nothing
    Kopioita = InputBox("How many copies are being sent? E.g. 4 paper copies and e-mail are marked as ""4+e"".", "Copies", Oletus)
    If Kopioita <> "" Then
      SQL = "INSERT INTO USysDistrSelection (SessionID, RecipientID, Copies, CreatedAt) VALUES ('" & _
            mSessionID & "', " & VastaanottajaID & ", '" & Replace(Kopioita, "'", "''") & "', Now());"
      DB.Execute SQL, dbFailOnError
    End If
  End If

  Set DB = Nothing
  Me.Lista.Requery
  Exit Sub

ErrorHandler:
  MsgBox "Error updating recipient selection: " & Err.Description, vbExclamation, "Update Error"
  On Error Resume Next
  If Not taulu Is Nothing Then taulu.Close
  Set taulu = Nothing
  Set DB = Nothing
  On Error GoTo 0
End Sub

Private Sub Valmis_Click()
  Dim SQL As String
  Dim DB As DAO.Database
  Dim taulu As DAO.Recordset
  Dim DistrID As Long
  Dim Valintoja As Long
  Dim TransaktioKaynnissa As Boolean
  Dim Virhekuvaus As String

  On Error GoTo ErrorHandler
  If mTallennettu Then
    MsgBox "This distribution has already been saved. Close the form to avoid a duplicate.", vbExclamation, "Distribution saved"
    Exit Sub
  End If
  If mSessionID = "" Then Err.Raise vbObjectError + 1203, "Valmis_Click", "The distribution session is not initialized."
  If IsNull(Me.DistrNo) Then
    MsgBox "You have to give number for distribution!", vbCritical, "Error!"
    Exit Sub
  ElseIf IsNull(Me.TDate) Then
    MsgBox "You have to select date!", vbCritical, "Error!"
    Exit Sub
  ElseIf IsNull(Me.Desc) Then
    MsgBox "You have to put description!", vbCritical, "Error!"
    Exit Sub
  ElseIf IsNull(Me.TReason) Then
    MsgBox "Distribution must have reason! Select reason and try again", vbCritical, "Error!"
    Exit Sub
  End If

  Set DB = CurrentDb
  SQL = "SELECT ID FROM UsysDISTRIBUTION WHERE No='" & Replace(CStr(Me.DistrNo.Value), "'", "''") & "'"
  Set taulu = DB.OpenRecordset(SQL, dbOpenSnapshot)
  If Not taulu.EOF Then
    MsgBox "Distribution number """ & Me.DistrNo & """ is already in use.", vbCritical, "Error!"
    taulu.Close
    Set taulu = Nothing
    Set DB = Nothing
    Exit Sub
  End If
  taulu.Close
  Set taulu = Nothing

  DBEngine.BeginTrans
  TransaktioKaynnissa = True
  Set taulu = DB.OpenRecordset("SELECT Count(*) AS SelectionCount FROM USysDistrSelection WHERE SessionID='" & mSessionID & "'", dbOpenSnapshot)
  Valintoja = CLng(taulu.Fields("SelectionCount").Value)
  taulu.Close
  Set taulu = Nothing
  If Valintoja = 0 Then Err.Raise vbObjectError + 1204, "Valmis_Click", "Select at least one recipient for the distribution."

  Set taulu = DB.OpenRecordset("UsysDISTRIBUTION", dbOpenDynaset)
  With taulu
    .AddNew
    .Fields("No") = Me.DistrNo
    .Fields("Description") = Me.Desc
    .Fields("Date") = Me.TDate
    .Fields("Sender") = Me.TSender
    .Fields("Reason") = Me.TReason
    .Fields("Note") = Me.Note
    .Update
    .Bookmark = .LastModified
    DistrID = CLng(.Fields("ID").Value)
    .Close
  End With
  Set taulu = Nothing

  SQL = "INSERT INTO USysRecipByDistr (DistrID, Copies, Unit, Company, Name, Description, Address, email) " & _
        "SELECT " & DistrID & ", S.Copies, R.Unit, R.Company, R.Name, R.Description, R.Address, R.email " & _
        "FROM USysRecipients AS R INNER JOIN USysDistrSelection AS S ON R.ID=S.RecipientID " & _
        "WHERE S.SessionID='" & mSessionID & "';"
  DB.Execute SQL, dbFailOnError
  If DB.RecordsAffected <> Valintoja Then
    Err.Raise vbObjectError + 1205, "Valmis_Click", "The recipient list changed; check the selected recipients."
  End If
  DB.Execute "DELETE FROM USysDistrSelection WHERE SessionID='" & mSessionID & "'", dbFailOnError
  DBEngine.CommitTrans
  TransaktioKaynnissa = False
  mTallennettu = True
  Set DB = Nothing
  DoCmd.Close
  Exit Sub

ErrorHandler:
  Virhekuvaus = Err.Description
  On Error Resume Next
  If TransaktioKaynnissa Then
    Err.Clear
    DBEngine.Rollback
    If Err.Number <> 0 Then Virhekuvaus = Virhekuvaus & vbCrLf & "Rollback failed: " & Err.Description
  End If
  If Not taulu Is Nothing Then
    Err.Clear
    taulu.Close
    If Err.Number <> 0 Then Virhekuvaus = Virhekuvaus & vbCrLf & "Closing the recordset failed: " & Err.Description
    Set taulu = Nothing
  End If
  Set DB = Nothing
  On Error GoTo 0
  If mTallennettu Then
    MsgBox "The distribution was saved, but the form could not be closed. Do not save it again." & vbCrLf & Virhekuvaus, vbExclamation, "Distribution saved"
  Else
    MsgBox "Distribution save could not be confirmed: " & Virhekuvaus, vbCritical, "Database Error"
  End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
  Dim DB As DAO.Database

  If mSessionID = "" Or mTallennettu Then Exit Sub
  On Error GoTo ErrorHandler
  Set DB = CurrentDb
  DB.Execute "DELETE FROM USysDistrSelection WHERE SessionID='" & mSessionID & "'", dbFailOnError
  Set DB = Nothing
  Exit Sub

ErrorHandler:
  MsgBox "Could not remove this session's temporary recipient selections: " & Err.Description & vbCrLf & _
         "Session ID: " & mSessionID, vbExclamation, "Selection cleanup"
  Set DB = Nothing
End Sub