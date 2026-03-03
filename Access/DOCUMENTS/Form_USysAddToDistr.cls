Option Compare Database
Option Explicit
'==========================================================================
' LOMAKE   : Form_USysAddToDistr
' SOVELLUS : DOCUMENTS — Lisää dokumentti distribuutioon
' KUVAUS   : Tarkistaa onko dokumentti jo lisätty valittuun distribuutioon
'            ja lisää sen, jos ei. Käyttää transaktioita atomaarisuuden
'            takaamiseksi.
' PÄIVITETTY: 2026-03-03
'==========================================================================

'''
' Add Document to Distribution
' OPTIMIZED 2025-11-09: Added error handling, explicit DAO typing, transaction support, null-checking
'''
Private Sub Add_Click()
  Dim DB As DAO.Database
  Dim taulu As DAO.Recordset
  
  On Error GoTo ErrorHandler
  
  Set DB = CurrentDb
Set taulu = DB.OpenRecordset("SELECT * FROM SentDocuments WHERE TDoc=" & CurRecord & " AND TDistr=" & Me.Distribution.Value, dbOpenDynaset)

  If taulu.RecordCount <> 0 Then
    MsgBox "Document were already on selected distribution!", vbCritical, "Add document to distribution"
  Else
    DBEngine.BeginTrans
    taulu.AddNew
    taulu.Fields(0) = CurRecord
    taulu.Fields(1) = Me.Distribution.Value
    taulu.Fields(2) = IIf(Common = "", Null, Common)
    If DocStatus <> "" Then taulu.Fields(3) = DocStatus
    taulu.Update
    DBEngine.CommitTrans
    DoCmd.Close
  End If
  
Cleanup:
  If Not taulu Is Nothing Then
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
    DBEngine.Rollback
  End If
  MsgBox "Error adding document to distribution: " & Err.Description, vbCritical, "Add Error"
  Resume Cleanup
End Sub

Private Sub Command14_Click()
  DoCmd.Close
End Sub

Private Sub Form_Load()
  ' Valitaan oletuksena ensimmäinen distribuutiovaihtoehto
  Me.Distribution.Value = Me.Distribution.ItemData(1)
End Sub