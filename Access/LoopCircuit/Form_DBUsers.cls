Option Compare Database
Option Explicit

' Updated 2025-10-22: Explicit DAO, better error handling, cleaned up logic

Private Type UserRec
   bMach(1 To 32) As Byte  ' Machine name bytes
   bUser(1 To 32) As Byte  ' User name bytes
End Type

Private Sub Form_Open(Cancel As Integer)
   On Error GoTo ErrHandler
   
   Dim sRowSource As String
   sRowSource = WhosOn()
   
   If Len(sRowSource) > 0 Then
      Me.LoggedOn.RowSource = sRowSource
   Else
      Me.LoggedOn.RowSource = ""
   End If
   
   Exit Sub
   
ErrHandler:
   ' Don't cancel form open, just show error
   MsgBox "Error opening form: " & Err.Number & " - " & Err.Description, vbExclamation, "Form Error"
   Me.LoggedOn.RowSource = ""
End Sub

Private Sub OKBtn_Click()
   DoCmd.Close
End Sub

Private Sub UpdateBtn_Click()
    On Error GoTo ErrHandler
    
    Dim sRowSource As String
    sRowSource = WhosOn()
    
    If Len(sRowSource) > 0 Then
       Me.LoggedOn.RowSource = sRowSource
    Else
       Me.LoggedOn.RowSource = ""
    End If
    
    Me.Requery
    Exit Sub
    
ErrHandler:
    MsgBox "Error updating: " & Err.Number & " - " & Err.Description, vbExclamation, "Update Error"
End Sub

' Reads the .LACCDB lock file to determine who's currently logged on
' Lock file has 64-byte records: station name (bytes 1-32) and login name (bytes 33-64), null-terminated
Private Function WhosOn() As String
   On Error GoTo Err_WhosOn

   Dim iLDBFile As Integer
   Dim SPath As String
   Dim X As String
   Dim sLogStr As String
   Dim sLogins As String
   Dim sMach As String
   Dim sUser As String
   Dim rUser As UserRec
   Dim i As Long

   ' Get current database path and construct lock file path
   SPath = CurrentDb.Name
   
   ' For Access 2007+ the lock file extension is .laccdb (not .ldb)
   SPath = Left$(SPath, InStr(1, SPath, ".")) & "laccdb"

   ' Test for valid file
   X = Dir(SPath)
   If X = "" Then
      ' Try legacy .ldb extension for older Access versions
      SPath = Left$(SPath, InStr(1, SPath, ".")) & "ldb"
      X = Dir(SPath)
      If X = "" Then
         Err.Raise 68 ' File not found
      End If
   End If
   
   iLDBFile = FreeFile
   Open SPath For Binary Access Read Shared As iLDBFile

   ' Read 64-byte records until EOF
   Do While Not EOF(iLDBFile)
      Get iLDBFile, , rUser
      
      With rUser
         ' Extract machine name (null-terminated)
         i = 1
         sMach = ""
         Do While i <= 32 And .bMach(i) <> 0
            sMach = sMach & Chr(.bMach(i))
            i = i + 1
         Loop
         
         ' Extract user name (null-terminated)
         i = 1
         sUser = ""
         Do While i <= 32 And .bUser(i) <> 0
            sUser = sUser & Chr(.bUser(i))
            i = i + 1
         Loop
      End With
      
      ' Build unique login string
      sLogStr = sMach & " -- " & sUser
      If InStr(sLogins, sLogStr) = 0 Then
         sLogins = sLogins & sLogStr & ";"
      End If
   Loop
   
   Close iLDBFile
   
   ' Remove trailing semicolon if present
   If Len(sLogins) > 0 And Right$(sLogins, 1) = ";" Then
      sLogins = Left$(sLogins, Len(sLogins) - 1)
   End If
   
   WhosOn = sLogins

Exit_WhosOn:
   Exit Function

Err_WhosOn:
   On Error Resume Next
   Close iLDBFile
   On Error GoTo 0
   
   WhosOn = ""
   
   If Err.Number = 68 Then
      MsgBox "Couldn't populate the list - lock file not found.", vbExclamation, "No Lock File"
   Else
      MsgBox "Error: " & Err.Number & vbCrLf & Err.Description, vbCritical, "WhosOn Error"
   End If
   Resume Exit_WhosOn
End Function

