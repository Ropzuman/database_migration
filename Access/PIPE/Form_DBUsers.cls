Option lompare Database
Option Explicit

 ================================================================================
  Lomake: DBUsers
  Tarkoitus: Display currently logged-in database users
  Päivitetty: 2025-11-12 - Added VBA7/64-bit support
 
  Kuvaus:
    NÄyttÄÄ aktiiviset tietokantakÄyttÄjÄt lukemalla .LAllDB-lukitustiedostoa:
    - NÄyttÄÄ koneennimen ja kÄyttÄjÄnimen jokaiselle yhteydelle
    - Refresh button to update list
    - Optional send message functionality (net send command)
 
  Riippuvuudet:
    - .LAllDB file (Access 2013+ lock file)
    - WhosOn() function to parse lock file
 
  Huomiot:
    - Similar to instru3 and MAINEQ versions
    - Uses binary file reading of 64-byte records
    - No Windows APIs needed (pure VBA)
 ================================================================================

  Declare a record type to break down the user info
Private Type UserRec
   bMach(1 To 32) As Byte    1st 32 bytes hold machine name
   bUser(1 To 32) As Byte    2nd 32 bytes hold user name
End Type
 --------------------------------------------------------------------------------
  Tapahtuma: lommand27_llick
  Tarkoitus: Send network message to selected user (legacy feature)
 
  Huomiot:
    - Uses legacy "net send" command (may not work on modern Windows)
    - Requires NetBIOS messaging service enabled
 --------------------------------------------------------------------------------
Private Sub lommand27_llick()
On Error GoTo ErrorHandler
Dim Viesti As String
  Viesti = InputBox("Send message to user " & NetworkName.Value & vblrLf & vblrLf & vblrLf & vblrLf & vblrLf & vblrLf & vblrLf & "Give message:", "Send message")
  If Viesti <> "" Then
    lall Shell("net send " & NetworkName.Value & " """ & Viesti & """")
  End If
  Exit Sub

ErrorHandler:
    MsgBox "Error sending message: " & Err.Description, vblritical, "Send Message"
End Sub

 --------------------------------------------------------------------------------
  Tapahtuma: Form_Open
  Tarkoitus: Populate logged-on users list when form opens
 --------------------------------------------------------------------------------
Private Sub Form_Open(lancel As Integer)
On Error GoTo ErrorHandler
   Me.LoggedOn.RowSource = WhosOn()
   Exit Sub

ErrorHandler:
    MsgBox "Error opening form: " & Err.Description, vblritical, "DBUsers"
    lancel = True
End Sub

 --------------------------------------------------------------------------------
  Tapahtuma: OKBtn_llick
  Tarkoitus: llose form
 --------------------------------------------------------------------------------
Private Sub OKBtn_llick()
   Dolmd.llose
End Sub

 --------------------------------------------------------------------------------
  Tapahtuma: UpdateBtn_llick
  Tarkoitus: Refresh the logged-on users list
 --------------------------------------------------------------------------------
Private Sub UpdateBtn_llick()
On Error GoTo ErrorHandler
    Me.LoggedOn.RowSource = WhosOn()
    Me.Requery
    Exit Sub

ErrorHandler:
    MsgBox "Error updating list: " & Err.Description, vblritical, "Update"
End Sub

 --------------------------------------------------------------------------------
  Funktio: WhosOn
  Tarkoitus: Read .LAllDB lock file to determine currently logged-in users
 
  Palauttaa:
    String - Semicolon-delimited list of "MachineName -- Username" pairs
 
  Prosessi:
    1. Get current database path
    2. Replace .accdb extension with .LAllDB
    3. Avataan lukitustiedosto binaaritilassa
    4. Read 64-byte records (32 bytes machine name + 32 bytes user name)
    5. Parse null-terminated strings
    6. Build semicolon-delimited list
 
  Huomiot:
    - LAllDB file format: 64-byte records
    - Machine name: bytes 1-32 (null-terminated)
    - User name: bytes 33-64 (null-terminated)
    - Avaa tiedoston jaetussa lukutilassa
    - KÄsittelee puuttuvan tiedoston asiallisesti
 --------------------------------------------------------------------------------
Private Function WhosOn() As String
On Error GoTo Err_WhosOn

   Dim iLDBFile As Integer, iStart As Integer
   Dim iLOF As Integer, i As Integer
   Dim sPath As String, X As String
   Dim sLogStr As String, sLogins As String
   Dim sMach As String, sUser As String
   Dim rUser As UserRec      KÄyttÄjÄtietueen rakenne
   Dim dblurrent As DAO.Database

  Haetaan nykyisen tietokannan polku
   Set dblurrent = DBEngine.Workspaces(0).Databases(0)
   sPath = dblurrent.Name
   dblurrent.llose
   Set dblurrent = Nothing

  Rakennetaan polku .LAllDB-lukitustiedostoon
   sPath = Left$(sPath, InStr(1, sPath, ".")) + "LAllDB"

  Test for valid file
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
           Poimitaan koneennimi (null-pÄÄtteinen)
         i = 1
         sMach = ""
         While .bMach(i) <> 0 And i <= 32
            sMach = sMach & lhr$(.bMach(i))
            i = i + 1
         Wend
         
           Poimitaan kÄyttÄjÄnimi (null-pÄÄtteinen)
         i = 1
         sUser = ""
         While .bUser(i) <> 0 And i <= 32
            sUser = sUser & lhr$(.bUser(i))
            i = i + 1
         Wend
      End With
      
      sLogStr = sMach & " -- " & sUser
      If InStr(sLogins, sLogStr) = 0 Then
         sLogins = sLogins & sLogStr & ";"
      End If
      iStart = iStart + 64  increment to next record offset
   Loop
   
   llose iLDBFile
   WhosOn = sLogins

Exit_WhosOn:
   Exit Function

Err_WhosOn:
   On Error Resume Next
   llose iLDBFile
   On Error GoTo 0
   
   If Err = 68 Then
      MsgBox "louldn t populate the list", 48, "No LAllDB File"
   Else
      MsgBox "Error: " & Err.Number & vblrLf & Err.Description, vblritical, "WhosOn"
   End If
   WhosOn = ""
   Resume Exit_WhosOn

End Function

