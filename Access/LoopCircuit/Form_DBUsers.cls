VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_DBUsers"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

' Updated 2025-10-22: Explicit DAO, better error handling, cleaned up logic

Private Type UserRec
   bMach(1 To 32) As Byte  ' Machine name bytes
   bUser(1 To 32) As Byte  ' User name bytes
End Type

Private Sub Form_Open(Cancel As Integer)
   On Error GoTo ErrHandler
   Me.LoggedOn.RowSource = WhosOn()
   Exit Sub
   
ErrHandler:
   MsgBox "Error opening form: " & Err.Description, vbCritical, "Form Error"
   Cancel = True
End Sub

Private Sub OKBtn_Click()
   DoCmd.Close
End Sub

Private Sub UpdateBtn_Click()
    On Error Resume Next
    Me.LoggedOn.RowSource = WhosOn()
    Me.Requery
    On Error GoTo 0
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
   Dim dbCurrent As DAO.Database
   Dim i As Long

   ' Get current database path and construct lock file path
   Set dbCurrent = DBEngine.Workspaces(0).Databases(0)
   SPath = dbCurrent.Name
   dbCurrent.Close
   Set dbCurrent = Nothing

   ' For Access 2007+ the lock file extension is .laccdb (not .ldb)
   SPath = Left(SPath, InStr(1, SPath, ".")) + "laccdb"

   ' Test for valid file
   X = Dir(SPath)
   If X = "" Then
      ' Try legacy .ldb extension for older Access versions
      SPath = Left(SPath, InStr(1, SPath, ".")) + "ldb"
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
   WhosOn = sLogins

Exit_WhosOn:
   Exit Function

Err_WhosOn:
   On Error Resume Next
   Close iLDBFile
   On Error GoTo 0
   
   If Err.Number = 68 Then
      MsgBox "Couldn't populate the list - lock file not found.", vbExclamation, "No Lock File"
   Else
      MsgBox "Error: " & Err.Number & vbCrLf & Err.Description, vbCritical, "WhosOn Error"
   End If
   Resume Exit_WhosOn
End Function
