Option Compare Database
Option Explicit

'================================================================================
' Form: DBUsers
' Purpose: Display currently logged-in database users
' Updated: 2025-11-13 - Added VBA7/64-bit support and optimization
'
' Description:
'   Reads the .LACCDB lock file to show who is currently using the database.
'   Displays station name and login name for each active connection.
'
' Dependencies:
'   - .LACCDB lock file (Access 2007+)
'   - LoggedOn control (list/combo box)
'
' Lock File Format:
'   - 64-byte records
'   - Bytes 1-32: Machine/station name (null-terminated)
'   - Bytes 33-64: User login name (null-terminated)
'================================================================================

' User record structure for reading lock file
Private Type UserRec
    bMach(1 To 32) As Byte  ' Machine name (1st 32 bytes)
    bUser(1 To 32) As Byte  ' User name (2nd 32 bytes)
End Type

'================================================================================
' Event: Form_Open
' Purpose: Initialize form with current user list
'================================================================================
Private Sub Form_Open(Cancel As Integer)
    On Error GoTo ErrorHandler
    Me.LoggedOn.RowSource = WhosOn()
    Exit Sub
ErrorHandler:
    MsgBox "Error loading user list: " & Err.Description, vbExclamation
End Sub

'================================================================================
' Event: OKBtn_Click
' Purpose: Close the form
'================================================================================
Private Sub OKBtn_Click()
    On Error GoTo ErrorHandler
    DoCmd.Close
    Exit Sub
ErrorHandler:
    ' Silent error - form likely already closing
End Sub

'================================================================================
' Event: UpdateBtn_Click
' Purpose: Refresh the user list
'================================================================================
Private Sub UpdateBtn_Click()
    On Error GoTo ErrorHandler
    Me.LoggedOn.RowSource = WhosOn()
    Me.Requery
    Exit Sub
ErrorHandler:
    MsgBox "Error updating user list: " & Err.Description, vbExclamation
End Sub

'================================================================================
' Function: WhosOn
' Purpose: Read lock file and build list of logged-in users
' Returns: Semicolon-separated list of "Machine;User" pairs for RowSource
'
' Description:
'   Opens the .LACCDB lock file and reads all 64-byte user records.
'   Extracts machine name and user name from each record, building
'   a value list suitable for a list box RowSource property.
'
' Lock File Location:
'   Same directory as database, same name but .LACCDB extension
'================================================================================
Private Function WhosOn() As String
    Dim iLDBFile As Integer
    Dim iLOF As Integer
    Dim i As Integer
    Dim SPath As String
    Dim X As String
    Dim sLogins As String
    Dim sMach As String
    Dim sUser As String
    Dim rUser As UserRec
    Dim dbCurrent As DAO.Database

    On Error GoTo Err_WhosOn

    ' Get current database path
    Set dbCurrent = DBEngine.Workspaces(0).Databases(0)
    SPath = dbCurrent.Name
    dbCurrent.Close
    Set dbCurrent = Nothing

    ' Build lock file path (.LACCDB extension)
    SPath = Left$(SPath, InStr(1, SPath, ".")) & "LACCDB"

    ' Verify lock file exists
    X = Dir$(SPath)
    If X = "" Then
        WhosOn = "No users currently logged in"
        Exit Function
    End If

    ' Open lock file for binary reading
    iLDBFile = FreeFile
    Open SPath For Binary Access Read Shared As iLDBFile
    iLOF = LOF(iLDBFile)

    ' Read all user records
    Do While Not EOF(iLDBFile)
        Get iLDBFile, , rUser
        
        ' Extract machine name (null-terminated string)
        With rUser
            i = 1
            sMach = ""
            While .bMach(i) <> 0 And i <= 32
                sMach = sMach & Chr$(.bMach(i))
                i = i + 1
            Wend

            ' Extract user name (null-terminated string)
            i = 1
            sUser = ""
            While .bUser(i) <> 0 And i <= 32
                sUser = sUser & Chr$(.bUser(i))
                i = i + 1
            Wend
        End With

        ' Add to results (semicolon-separated for RowSource)
        If sLogins <> "" Then sLogins = sLogins & ";"
        sLogins = sLogins & sMach & ";" & sUser
    Loop

    Close iLDBFile
    WhosOn = sLogins
    Exit Function

Err_WhosOn:
    On Error Resume Next
    Close iLDBFile
    WhosOn = "Error reading user list: " & Err.Description
End Function
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

