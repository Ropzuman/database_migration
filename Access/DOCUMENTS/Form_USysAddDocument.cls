Option Compare Database
Option Explicit
Public SelDept As String
Public Varaaja As String
Private Sub ChoosePath_Click()
Dim Hakemisto As String
  Hakemisto = ValitseHakem(Me.hWnd, IIf(IsNull(TPath.Value), "", TPath.Value))
  If Hakemisto <> "" Then
    TPath.Value = Hakemisto
  End If
End Sub
'''
' Form Load Event - Initialize form with default values from previous document
' OPTIMIZED 2025-11-09: Added error handling, explicit DAO typing, null-checking
'''
Private Sub Form_Load()
  Dim DB As DAO.Database
  Dim taulu As DAO.Recordset
  
  On Error GoTo ErrorHandler
  
  Varaaja = NetworkUserName
  
  ' Set default values from previous document entry
  Set DB = CurrentDb
  Set taulu = DB.OpenRecordset("SELECT Path FROM ProjInfo", dbOpenSnapshot)
  
  TDocName1.Value = IIf(DefDName1 <> "", DefDName1, Null)
  TDocName2.Value = IIf(DefDName2 <> "", DefDName2, Null)
  TDocName3.Value = IIf(DefDName3 <> "", DefDName3, Null)
  TPages.Value = IIf(DefPages <> "", DefPages, 1)
  TPath.Value = IIf(DefPath <> "", DefPath, Nz(taulu("Path"), ""))
  TRev.Value = IIf(DefRev <> "", DefRev, Null)
  
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
  MsgBox "Error loading form defaults: " & Err.Description, vbExclamation, "Form Load Error"
  Resume Cleanup
End Sub
Private Sub Peruuta_Click()
  DoCmd.Close
End Sub
'''
' Save Document - Validate and save new document to database
' OPTIMIZED 2025-11-09: Added transaction support, error handling, DAO typing, null-checking
'''
Private Sub TalletaNappi_Click()
    Dim DB As DAO.Database
    Dim taulu As DAO.Recordset
    Dim SQL As String
    
    On Error GoTo ErrorHandler
    
    ' Validation
    If IsNull(TDocName1.Value) And IsNull(TDocName2.Value) Then
      MsgBox "You can't leave Document Name empty!" & vbCrLf _
           & "If you don't know document name now, put something like:" & vbCrLf _
           & """Reserved for subcontractor"", and write name down afterwards.", vbCritical
      Exit Sub
    ElseIf IsNull(TClientDocNo.Value) And IsNull(TMetsoDocNo.Value) And IsNull(TExtDocID.Value) Then
      MsgBox "Document has to have some kind of ID!" & vbCrLf _
           & "If document is not going to client put at least Extra Document ID", vbCritical
      Exit Sub
    End If
    
    Set DB = CurrentDb
    DB.BeginTrans
    
    ' Check and reserve client document number if provided
    If IsNull(TClientDocNo.Value) = False And TClientDocNo.Value <> "" Then
        SQL = "SELECT * FROM UsysClientDocNos WHERE Number='AV-2090-210-DD-" & TNumber.Value & "'"
        Set taulu = DB.OpenRecordset(SQL, dbOpenDynaset, dbDenyWrite)
        
        If taulu.EOF = False Then
          If IsNull(taulu.Fields("Reserved")) Then
            taulu.Edit
            taulu.Fields("Number") = TClientDocNo.Value
            taulu.Fields("Reserved") = Varaaja & "@" & Now
            taulu.Update
            taulu.Close
            Set taulu = Nothing
          Else
            ' Number already reserved
            taulu.Close
            Set taulu = Nothing
            DB.Rollback
            DB.Close
            Set DB = Nothing
            MsgBox "Someone reserved number " & TClientDocNo.Value & " at the same time!" & vbCrLf & "Please choose another number.", vbCritical, "Error!"
            Exit Sub
          End If
        Else
          taulu.Close
          Set taulu = Nothing
        End If
    End If
    
    ' Add new document
    Set taulu = DB.OpenRecordset("DOCUMENTS", dbOpenTable)
    With taulu
        .AddNew
        .Fields("DocName1") = TDocName1.Value
        .Fields("DocName2") = TDocName2.Value
        .Fields("DocName3") = TDocName3.Value
        .Fields("Area") = TArea.Value
        .Fields("Pick2") = TPick2.Value
        .Fields("Rev") = TRev.Value
        .Fields("ClientNo") = TClientDocNo.Value
        .Fields("MetsoDocNo") = TMetsoDocNo.Value
        .Fields("ExtDocID") = TExtDocID.Value
        .Fields("Pages") = TPages.Value
        .Fields("WorkPath") = TPath.Value
        .Fields("Note") = TNote.Value
        CurRecord = .Fields("Counter")
        .Update
    End With
    
    DB.CommitTrans
    
    ' Save defaults for next document
    DefDName1 = Nz(TDocName1.Value, "")
    DefDName2 = Nz(TDocName2.Value, "")
    DefDName3 = Nz(TDocName3.Value, "")
    DefArea = Nz(TArea.Value, "")
    DefSort = Nz(TPick2.Value, "")
    DefPages = Nz(TPages.Value, "")
    DefPath = Nz(TPath.Value, "")
    DefRev = Nz(TRev.Value, "")
    
Cleanup:
    If Not taulu Is Nothing Then
      If taulu.State = 1 Then taulu.Close
      Set taulu = Nothing
    End If
    If Not DB Is Nothing Then
      DB.Close
      Set DB = Nothing
    End If
    DoCmd.Close
    Exit Sub
    
ErrorHandler:
    If Not DB Is Nothing Then
        On Error Resume Next
        DB.Rollback
    End If
    MsgBox "Error saving document: " & Err.Description, vbCritical, "Save Error"
    Resume Cleanup
End Sub

Private Sub TDiscipline_Change()
  MuutaNumero
End Sub

Private Sub TNumber_Change()
  MuutaNumero
End Sub
Private Sub TPaperSize_Change()
  MuutaNumero
End Sub
Private Sub TRev_Click()
  AvaaRevisio
End Sub
Private Sub MuutaNumero()
  If IsNull(TPaperSize.Value) Or IsNull(TDiscipline.Value) Or IsNull(TNumber.Value) Then
  Else
    TClientDocNo.Value = TPaperSize.Value & "-2090-210-" & TDiscipline.Value & "-" & TNumber.Value
  End If
End Sub
Private Sub AvaaRevisio()
  Common = "ENSIMMAINEN"
  Revisioteksti = IIf(IsNull(TRev.Value), "", TRev.Value)
  DoCmd.OpenForm "USysRevText", , , , , acDialog
  TRev.Value = Revisioteksti
End Sub