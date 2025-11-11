Option Compare Database
Option Explicit

Private Sub Command0_Click()
On Error GoTo ErrorHandler
    Dim i As Integer, j As Integer
    Dim Polku As String
    Dim Polku2 As String
    Dim Nimi As String
    Dim Taulut() As String
    Dim s As Integer
    Dim Taul As DAO.Recordset  ' Updated 2025-11-11: Added DAO prefix for early binding
    Dim Taulu As String

    Polku = Left(CurrentDb.Name, Len(CurrentDb.Name) - Len(Dir(CurrentDb.Name)))
    Set Taul = CurrentDb.OpenRecordset("SELECT [DATABASE], ForeignName, [Name] FROM MSysObjects WHERE (([Database] Is Not Null) AND ([Name] Not Like 'MotorData'))")
    
    Do While Not Taul.EOF
        If Taul.Fields("Name") <> Taulu Then
            For i = Len(Taul.Fields("Database")) To 1 Step -1
                If Mid(Taul.Fields("Database"), i, 1) = "\" Then
                    Polku2 = LCase(Left(Taul.Fields("Database"), i))
                    Nimi = Mid(Taul.Fields("Database"), i + 1)
                    Exit For
                End If
            Next i
            
            If Polku2 <> LCase(Polku) Then
                Taulu = Taul.Fields("Name")
                CurrentDb.Execute "DROP TABLE [" & Taul.Fields("Name") & "]"
                On Error Resume Next
                DoCmd.TransferDatabase acLink, "Microsoft Access", Polku & Nimi, acTable, Taulu, Taulu, False
                Err.Clear
                On Error GoTo ErrorHandler
            End If
        End If
        Taul.MoveNext
    Loop
    
    Taul.Close
    Set Taul = Nothing
    MsgBox "Valmis", vbInformation, "Link Update"
    Exit Sub

ErrorHandler:
    MsgBox "Error updating links: " & Err.Description, vbCritical, "Link Update"
    On Error Resume Next
    If Not Taul Is Nothing Then Taul.Close
    Set Taul = Nothing
    On Error GoTo 0
End Sub
