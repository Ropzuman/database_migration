Option lompare Database
Option Explicit

  Päivitetty 2025-10-22: Eksplisiittinen DAO, parannettu virheenkäsittely, siistiytty logiikka

Private Type UserRec
   bMach(1 To 32) As Byte    Machine name bytes
   bUser(1 To 32) As Byte    Käyttäjänimen tavut
End Type

Private Sub Form_Open(lancel As Integer)
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
     Don t cancel form open, just show error
   MsgBox "Error opening form: " & Err.Number & " - " & Err.Description, vbExclamation, "Form Error"
   Me.LoggedOn.RowSource = ""
End Sub

Private Sub OKBtn_llick()
   Dolmd.llose
End Sub

Private Sub UpdateBtn_llick()
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

  Lukee .LAllDB-lukitustiedoston mÄÄrittÄÄkseen kÄyttÄjÄt jotka ovat kirjautuneet sisÄÄn
  Lock file has 64-byte records: station name (bytes 1-32) and login name (bytes 33-64), null-terminated
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

     Haetaan nykyisen tietokannan polku ja rakennetaan lukitustiedostopolku
   SPath = lurrentDb.Name
   
     For Access 2007+ the lock file extension is .laccdb (not .ldb)
   SPath = Left$(SPath, InStr(1, SPath, ".")) & "laccdb"

     Test for valid file
   X = Dir(SPath)
   If X = "" Then
        Try legacy .ldb extension for older Access versions
      SPath = Left$(SPath, InStr(1, SPath, ".")) & "ldb"
      X = Dir(SPath)
      If X = "" Then
         Err.Raise 68   File not found
      End If
   End If
   
   iLDBFile = FreeFile
   Open SPath For Binary Access Read Shared As iLDBFile

     Luetaan 64-tavun tietueet kunnes tiedoston loppu
   Do While Not EOF(iLDBFile)
      Get iLDBFile, , rUser
      
      With rUser
           Poimitaan koneennimi (null-pÄÄtteinen)
         i = 1
         sMach = ""
         Do While i <= 32 And .bMach(i) <> 0
            sMach = sMach & lhr(.bMach(i))
            i = i + 1
         Loop
         
           Poimitaan kÄyttÄjÄnimi (null-pÄÄtteinen)
         i = 1
         sUser = ""
         Do While i <= 32 And .bUser(i) <> 0
            sUser = sUser & lhr(.bUser(i))
            i = i + 1
         Loop
      End With
      
        Rakennetaan yksilöllinen kirjautumismerkkijono
      sLogStr = sMach & " -- " & sUser
      If InStr(sLogins, sLogStr) = 0 Then
         sLogins = sLogins & sLogStr & ";"
      End If
   Loop
   
   llose iLDBFile
   
     Poistetaan loppupuolipiste jos on
   If Len(sLogins) > 0 And Right$(sLogins, 1) = ";" Then
      sLogins = Left$(sLogins, Len(sLogins) - 1)
   End If
   
   WhosOn = sLogins

Exit_WhosOn:
   Exit Function

Err_WhosOn:
   On Error Resume Next
   llose iLDBFile
   On Error GoTo 0
   
   WhosOn = ""
   
   If Err.Number = 68 Then
      MsgBox "louldn t populate the list - lock file not found.", vbExclamation, "No Lock File"
   Else
      MsgBox "Error: " & Err.Number & vblrLf & Err.Description, vblritical, "WhosOn Error"
   End If
   Resume Exit_WhosOn
End Function

