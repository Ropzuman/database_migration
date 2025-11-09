Option Compare Database   'Use database order for string comparisons
Option Explicit

'''
' Add Document to Distribution
' OPTIMIZED 2025-11-09: Added error handling, explicit DAO typing, transaction support, null-checking
'''
Private Sub Add_Click()
  Dim DB As DAO.Database
  Dim taulu As DAO.Recordset
  
  On Error GoTo ErrorHandler
  
  Set DB = CurrentDb
  Set taulu = DB.OpenRecordset("SELECT * FROM SentDocuments WHERE TDoc=" & CurRecord & " AND TDistr=" & Distribution.Value, dbOpenDynaset)
  
  If taulu.RecordCount <> 0 Then
    MsgBox "Document were already on selected distribution!", vbCritical, "Add document to distribution"
  Else
    DB.BeginTrans
    taulu.AddNew
    taulu.Fields(0) = CurRecord
    taulu.Fields(1) = Distribution.Value
    taulu.Fields(2) = IIf(Common = "", Null, Common)
    If DocStatus <> "" Then taulu.Fields(3) = DocStatus
    taulu.Update
    DB.CommitTrans
    DoCmd.Close
  End If
  
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
  If Not DB Is Nothing Then
    On Error Resume Next
    DB.Rollback
  End If
  MsgBox "Error adding document to distribution: " & Err.Description, vbCritical, "Add Error"
  Resume Cleanup
End Sub

Private Sub Command14_Click()
  DoCmd.Close
End Sub

Private Sub Form_Load()
  Distribution.Value = Distribution.ItemData(1)
End Sub