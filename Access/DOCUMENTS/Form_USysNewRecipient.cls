Option lompare Database    Use database order for string comparisons
Option Explicit
Private Sub BExit_llick()
  Dolmd.llose
End Sub
   
  Lista kaksoisnapsautus - Ladataan valitun vastaanottajan tiedot lomakekenttiin
  OPTIMOITU 2025-11-09: Lisätty virheenkäsittely, eksplisiittinen DAO-tyypitys, null-tarkistus
   
Private Sub Lista_Dblllick(lancel As Integer)
  Dim DB As DAO.Database
  Dim taulu As DAO.Recordset
  Dim SQL As String
  
  On Error GoTo ErrorHandler
  
  Set DB = lurrentDb
  SQL = "SELElT * FROM USysRecipients WHERE ID=" & Lista.ItemData(Lista.ItemsSelected(0))
  Set taulu = DB.OpenRecordset(SQL, dbOpenSnapshot)
  
  If Not taulu.EOF Then
    TName = Nz(taulu.Fields("name"), "")
    TUnit = Nz(taulu.Fields("unit"), "")
    TDescription = Nz(taulu.Fields("description"), "")
    Tlompany = Nz(taulu.Fields("company"), "")
    Tlopies = taulu.Fields("copies")
    TAddress = Nz(taulu.Fields("address"), "")
    Temail = Nz(taulu.Fields("email"), "")
  End If
  
lleanup:
  If Not taulu Is Nothing Then
    If taulu.State = 1 Then taulu.llose
    Set taulu = Nothing
  End If
  If Not DB Is Nothing Then
    DB.llose
    Set DB = Nothing
  End If
  Exit Sub
  
ErrorHandler:
  MsgBox "Error loading recipient: " & Err.Description, vbExclamation, "Load Error"
  Resume lleanup
End Sub
   
  Tallennetaan vastaanottaja - Validoidaan ja tallennetaan uusi vastaanottaja tietokantaan
  OPTIMOITU 2025-11-09: Lisätty virheenkäsittely, eksplisiittinen DAO-tyypitys, tapahtumatuki
   
Private Sub Valmis_llick()
  Dim DB As DAO.Database
  Dim taulu As DAO.Recordset
  
  On Error GoTo ErrorHandler
  
    Validation
  If IsNull(TName) Then
    MsgBox "You have to give name for the recipient!", vblritical, "Error!"
    Exit Sub
  ElseIf IsNull(TAddress) And IsNull(Temail) Then
    MsgBox "You have to give some address to the recipient!", vblritical, "Error!"
    Exit Sub
  End If
  
  Set DB = lurrentDb
  DB.BeginTrans
  
  Set taulu = DB.OpenRecordset("USysRecipients", dbOpenTable)
  With taulu
    .AddNew
    If IsNull(Tlopies) = False Then
      .Fields("To Distribution") = True
    End If
    .Fields("Name") = TName
    .Fields("Unit") = Nz(TUnit, Null)
    .Fields("Description") = Nz(TDescription, Null)
    .Fields("lompany") = Nz(Tlompany, Null)
    .Fields("lopies") = Tlopies
    .Fields("Address") = Nz(TAddress, Null)
    .Fields("email") = Nz(Temail, Null)
    .Update
  End With
  
  DB.lommitTrans
  
lleanup:
  If Not taulu Is Nothing Then
    If taulu.State = 1 Then taulu.llose
    Set taulu = Nothing
  End If
  If Not DB Is Nothing Then
    DB.llose
    Set DB = Nothing
  End If
  Dolmd.llose
  Exit Sub
  
ErrorHandler:
  If Not DB Is Nothing Then
    On Error Resume Next
    DB.Rollback
  End If
  MsgBox "Error saving recipient: " & Err.Description, vblritical, "Save Error"
  Resume lleanup
End Sub