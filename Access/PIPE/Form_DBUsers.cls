Option Compare Database
Option Explicit

'================================================================================
' Form: DBUsers
' Purpose: Display currently logged-in database users
' Updated: 2025-11-12 - Added VBA7/64-bit support
'
' Description:
'   Shows active database users by reading .LACCDB lock file:
'   - Displays machine name and username for each connection
'   - Refresh button to update list
'   - Optional send message functionality (net send command)
'
' Dependencies:
'   - .LACCDB file (Access 2013+ lock file)
'   - WhosOn() function to parse lock file
'
' Notes:
'   - Similar to instru3 and MAINEQ versions
'   - Uses binary file reading of 64-byte records
'   - No Windows APIs needed (pure VBA)
'================================================================================

' Declare a record type to break down the user info
Private Type UserRec
   bMach(1 To 32) As Byte  ' 1st 32 bytes hold machine name
   bUser(1 To 32) As Byte  ' 2nd 32 bytes hold user name
End Type
'--------------------------------------------------------------------------------
' Event: Command27_Click
' Purpose: Send network message to selected user (legacy feature)
'
' Notes:
'   - Uses legacy "net send" command (may not work on modern Windows)
'   - Requires NetBIOS messaging service enabled
'--------------------------------------------------------------------------------
Private Sub Command27_Click()
On Error GoTo ErrorHandler
Dim Viesti As String
  Viesti = InputBox("Send message to user " & NetworkName.Value & vbCrLf & vbCrLf & vbCrLf & vbCrLf & vbCrLf & vbCrLf & vbCrLf & "Give message:", "Send message")
  If Viesti <> "" Then
    Call Shell("net send " & NetworkName.Value & " """ & Viesti & """")
  End If
  Exit Sub

ErrorHandler:
    MsgBox "Error sending message: " & Err.Description, vbCritical, "Send Message"
End Sub

'--------------------------------------------------------------------------------
' Event: Form_Open
' Purpose: Populate logged-on users list when form opens
'--------------------------------------------------------------------------------
Private Sub Form_Open(Cancel As Integer)
On Error GoTo ErrorHandler
   Me.LoggedOn.RowSource = WhosOn()
   Exit Sub

ErrorHandler:
    MsgBox "Error opening form: " & Err.Description, vbCritical, "DBUsers"
    Cancel = True
End Sub

'--------------------------------------------------------------------------------
' Event: OKBtn_Click
' Purpose: Close form
'--------------------------------------------------------------------------------
Private Sub OKBtn_Click()
   DoCmd.Close
End Sub

'--------------------------------------------------------------------------------
' Event: UpdateBtn_Click
' Purpose: Refresh the logged-on users list
'--------------------------------------------------------------------------------
Private Sub UpdateBtn_Click()
On Error GoTo ErrorHandler
    Me.LoggedOn.RowSource = WhosOn()
    Me.Requery
    Exit Sub

ErrorHandler:
    MsgBox "Error updating list: " & Err.Description, vbCritical, "Update"
End Sub

'--------------------------------------------------------------------------------
' Function: WhosOn
' Purpose: Read .LACCDB lock file to determine currently logged-in users
'
' Returns:
'   String - Semicolon-delimited list of "MachineName -- Username" pairs
'
' Process:
'   1. Get current database path
'   2. Replace .accdb extension with .LACCDB
'   3. Open lock file in binary mode
'   4. Read 64-byte records (32 bytes machine name + 32 bytes user name)
'   5. Parse null-terminated strings
'   6. Build semicolon-delimited list
'
' Notes:
'   - LACCDB file format: 64-byte records
'   - Machine name: bytes 1-32 (null-terminated)
'   - User name: bytes 33-64 (null-terminated)
'   - Opens file with shared read access
'   - Handles missing file gracefully
'--------------------------------------------------------------------------------
Private Function WhosOn() As String
On Error GoTo Err_WhosOn

   Dim iLDBFile As Integer, iStart As Integer
   Dim iLOF As Integer, i As Integer
   Dim sPath As String, X As String
   Dim sLogStr As String, sLogins As String
   Dim sMach As String, sUser As String
   Dim rUser As UserRec    ' User record structure
   Dim dbCurrent As DAO.Database

' Get path of current database
   Set dbCurrent = DBEngine.Workspaces(0).Databases(0)
   sPath = dbCurrent.Name
   dbCurrent.Close
   Set dbCurrent = Nothing

' Build path to .LACCDB lock file
   sPath = Left$(sPath, InStr(1, sPath, ".")) + "LACCDB"

' Test for valid file
   X = Dir$(sPath)
   If X = "" Then
      WhosOn = ""
      GoTo Exit_WhosOn
   End If
   
   iStart = 1
   iLDBFile = FreeFile

   Open sPath For Binary Access Read Shared As iLDBFile
   iLOF = LOF(iLDBFile)
   
   Do While Not EOF(iLDBFile)
      Get iLDBFile, , rUser
      With rUser
         ' Extract machine name (null-terminated)
         i = 1
         sMach = ""
         While .bMach(i) <> 0 And i <= 32
            sMach = sMach & Chr$(.bMach(i))
            i = i + 1
         Wend
         
         ' Extract user name (null-terminated)
         i = 1
         sUser = ""
         While .bUser(i) <> 0 And i <= 32
            sUser = sUser & Chr$(.bUser(i))
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
   Exit Function

Err_WhosOn:
   On Error Resume Next
   Close iLDBFile
   On Error GoTo 0
   
   If Err = 68 Then
      MsgBox "Couldn't populate the list", 48, "No LACCDB File"
   Else
      MsgBox "Error: " & Err.Number & vbCrLf & Err.Description, vbCritical, "WhosOn"
   End If
   WhosOn = ""
   Resume Exit_WhosOn

End Function

