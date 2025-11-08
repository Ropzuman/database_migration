Option Compare Database   'Use database order for string comparisons
Option Explicit

Private Sub Add_Click()
Dim DB As Database
Dim taulu As Recordset
   Set DB = CurrentDb
   Set taulu = DB.OpenRecordset("SELECT * FROM SentDocuments WHERE TDoc=" & CurRecord & " AND TDistr=" & Distribution.Value)
   If taulu.RecordCount <> 0 Then 'Dokumentti on jo lis�tty
     MsgBox "Document were already on selected distribution!", vbCritical, "Add document to distribution"
     taulu.Close
     Set taulu = Nothing
     DB.Close
     Set DB = Nothing
   Else
     taulu.AddNew
     taulu.Fields(0) = CurRecord
     taulu.Fields(1) = Distribution.Value
     taulu.Fields(2) = IIf(Common = "", Null, Common)
    ' Taulu.Fields(3) = IIf(DefRev = "", Null, DefRev)
     If DocStatus <> "" Then taulu.Fields(3) = DocStatus
     taulu.Update
     taulu.Close
     Set taulu = Nothing
     DB.Close
     Set DB = Nothing
     DoCmd.Close
   End If
End Sub

Private Sub Command14_Click()
  DoCmd.Close
End Sub

Private Sub Form_Load()
  Distribution.Value = Distribution.ItemData(1)
End Sub