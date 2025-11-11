Option Compare Database
Option Explicit

Private Sub Command8_Click()
On Error GoTo ErrorHandler
    Dim DB As DAO.Database  ' Updated 2025-11-11: Added DAO prefix for early binding
    Dim Taulu As DAO.Recordset  ' Updated 2025-11-11: Added DAO prefix for early binding
    Dim FSO As New FileSystemObject
    Dim Tiedosto As Object
    Dim Path As String
    Dim i As Integer
    Dim Osoitin As Long
    Dim Arvo As String
    
    Set DB = CurrentDb
    Path = Hakem.Value
    If Right(Path, 1) <> "\" Then
        Path = Path & "\"
    End If
    
    If Not FSO.FolderExists(Path) Then
        MsgBox "Hakemistoa: " & Path & " ei ollut", vbCritical, "Export Error"
        Exit Sub
    End If
    
    Set Taulu = DB.OpenRecordset("SELECT * FROM " & Taulukko.Value)
    Set Tiedosto = FSO.CreateTextFile(Path & FileName.Value & ".csv")
    
    Do While Not Taulu.EOF
        Tiedosto.Write """" & Taulu.Fields(0).Value & """"
        For i = 1 To Taulu.Fields.Count - 1
            If IsNull(Taulu.Fields(i).Value) Then
                Arvo = ""
            Else
                Arvo = Taulu.Fields(i).Value
                Osoitin = InStr(Arvo, ",")
                If Osoitin <> 0 Then
                    Arvo = Left(Arvo, Osoitin - 1) & "." & Mid(Arvo, Osoitin + 1)
                End If
            End If
            Tiedosto.Write ";""" & Arvo & """"
        Next i
        Tiedosto.Write vbCrLf
        Taulu.MoveNext
    Loop
    
    ' Cleanup
    Taulu.Close
    Set Taulu = Nothing
    Set DB = Nothing
    Set Tiedosto = Nothing
    Set FSO = Nothing
    
    MsgBox "Valmis", vbInformation, "Export Complete"
    Call Shell("EXPLORER " & Path, vbNormalFocus)
    Exit Sub

ErrorHandler:
    MsgBox "Error exporting data: " & Err.Description, vbCritical, "Export Error"
    On Error Resume Next
    If Not Taulu Is Nothing Then Taulu.Close
    Set Taulu = Nothing
    Set DB = Nothing
    Set Tiedosto = Nothing
    Set FSO = Nothing
    On Error GoTo 0
End Sub

Private Sub Form_Load()
  Taulukko.Value = "qryValveSizingOut"
  'Hakem.VALUE = "d:\tilap\"
End Sub
