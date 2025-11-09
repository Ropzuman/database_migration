Option Compare Database

' Declare a record type to break down the user info

Private Type UserRec
   bMach(1 To 32) As Byte  ' 1st 32 bytes hold machine name
   bUser(1 To 32) As Byte  ' 2nd 32 bytes hold user name
End Type
Private Sub Command27_Click()
Dim Viesti As String
  Viesti = InputBox("Send message to user " & NetworkName.Value & vbCrLf & vbCrLf & vbCrLf & vbCrLf & vbCrLf & vbCrLf & vbCrLf & "Give message:", "Send message")
  If Viesti <> "" Then
    Call Shell("net send " & NetworkName.Value & " """ & Viesti & """")
  End If
End Sub

Private Sub Form_Open(Cancel As Integer)
   
   Me.LoggedOn.RowSource = WhosOn()

End Sub

Private Sub OKBtn_Click()
   
   DoCmd.Close

End Sub

Private Sub UpdateBtn_Click()
    Me.LoggedOn.RowSource = WhosOn()
    Me.Requery
End Sub

'-------------------------------------------------------------------------------------
'   Subject : WhosOn()
'   Purpose : Will read *.LDB file and read who's currently
'             logged on and their station name.
'
'             The LDB file has a 64 byte record.
'
'             The station name starts at byte 1 and is null
'             terminated.
'
'             Log-in names start at the 33rd byte and are
'             also null terminated.
'
'-------------------------------------------------------------------------------------
'''
' WhosOn - Read LDB file to show who's currently logged on
' OPTIMIZED 2025-11-09: Added explicit DAO typing, removed redundant CurrentDb call
' @return String - Semicolon-delimited list of "Machine -- User" entries
'''
Private Function WhosOn() As String

On Error GoTo Err_WhosOn

   Dim iLDBFile As Integer, iStart As Integer
   Dim iLOF As Integer, i As Integer
   Dim sPath As String, X As String
   Dim sLogStr As String, sLogins As String
   Dim sMach As String, sUser As String
   Dim rUser As UserRec    ' Defined in General
   Dim dbCurrent As DAO.Database

' Get Path of current database.  Should substitute this code
' for an attached table path in a multi-user environment.

   Set dbCurrent = CurrentDb
   sPath = dbCurrent.Name

' Iterate thru dbCurrent.LDB file for login names.

   sPath = Left(sPath, InStr(1, sPath, ".")) + "laccdb"

' Test for valid file, else Error

   X = Dir(sPath)
   iStart = 1
   iLDBFile = FreeFile

   Open sPath For Binary Access Read Shared As iLDBFile
   iLOF = LOF(iLDBFile)
   Do While Not EOF(iLDBFile)
      Get iLDBFile, , rUser
      With rUser
         i = 1
         sMach = ""
         While .bMach(i) <> 0
            sMach = sMach & Chr(.bMach(i))
            i = i + 1
         Wend
         i = 1
         sUser = ""
         While .bUser(i) <> 0
            sUser = sUser & Chr(.bUser(i))
            i = i + 1
         Wend
      End With
      sLogStr = sMach & " -- " & sUser
      If InStr(sLogins, sLogStr) = 0 Then
         sLogins = sLogins & sLogStr & ";"
      End If
      iStart = iStart + 64 'increment to next record offset
   Loop
   Close iLDBFile
   WhosOn = sLogins

Exit_WhosOn:
   If Not dbCurrent Is Nothing Then
      dbCurrent.Close
      Set dbCurrent = Nothing
   End If
   Exit Function

Err_WhosOn:
   If Err = 68 Then
      MsgBox "Couldn't populate the list", 48, "No LDB File"
   Else
      MsgBox "Error: " & Err.Number & vbCrLf & Err.Description
      Close iLDBFile
   End If
   Resume Exit_WhosOn

End Function