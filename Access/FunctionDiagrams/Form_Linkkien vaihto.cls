Option Compare Database
Option Explicit

Private Sub Command0_Click()
On Error GoTo ErrorHandler
Debug.Print "Command0_Click: Starting database link update operation"
Dim i As Integer, j As Integer
Dim Polku As String
Dim Polku2 As String
Dim Nimi As String
Dim Taulut() As String
Dim s As Integer
Dim Taul As DAO.Recordset
Dim Taulu As String


Polku = Left$(CurrentDb.Name, Len(CurrentDb.Name) - Len(Dir(CurrentDb.Name)))
Debug.Print "  Current database path: " & Polku
Set Taul = CurrentDb.OpenRecordset("SELECT [DATABASE], ForeignName, [Name] FROM MSysObjects WHERE (([Database] Is Not Null) AND ([Name] Not Like 'MotorData'))")
Debug.Print "  Processing linked tables..."
Do While Not Taul.EOF
  If Taul.Fields("Name") <> Taulu Then
      Debug.Print "  Checking table: " & Taul.Fields("Name")
      For i = Len(Taul.Fields("Database")) To 1 Step -1
        If Mid$(Taul.Fields("Database"), i, 1) = "\" Then
          Polku2 = LCase$(Left$(Taul.Fields("Database"), i))
          Nimi = Mid$(Taul.Fields("Database"), i + 1)
          Exit For
        End If
      Next i
    If Polku2 <> LCase$(Polku) Then
      Debug.Print "  Relinking table: " & Taul.Fields("Name") & " from: " & Taul.Fields("Database")
      Taulu = Taul.Fields("Name")
      CurrentDb.Execute "DROP TABLE [" & Taul.Fields("Name") & "]"
      On Error Resume Next
      DoCmd.TransferDatabase acLink, "Microsoft Access", Polku & Nimi, acTable, Taulu, Taulu, False
      Debug.Print "  Relinked to: " & Polku & Nimi
      Err.Clear
      On Error GoTo ErrorHandler
    End If
  End If
  Taul.MoveNext
Loop
Debug.Print "Command0_Click: COMPLETED successfully - All table links updated"
MsgBox "Valmis"
Exit Sub

ErrorHandler:
    Debug.Print "*** ERROR in Command0_Click: " & Err.Number & " - " & Err.Description
    Debug.Print "    Current Path: " & Polku
    Debug.Print "    Table: " & Taulu
    Debug.Print "    Source: " & Err.Source & ", Line: " & Erl
    MsgBox "Error updating links: " & Err.Description, vbCritical
End Sub
