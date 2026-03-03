Option Compare Database
Option Explicit

' Päivitetty 2025-10-22: Eksplisiittinen DAO, parannettu virheenkäsittely, siivottu logiikka

Private Type UserRec
   bMach(1 To 32) As Byte  ' Tietokoneen nimen tavut
   bUser(1 To 32) As Byte  ' Käyttäjänimen tavut
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
   ' Lomakkeen avaaminen ei peruunnu — virhe näytetään käyttäjälle
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

   ' Haetaan tietokannan polku lukitustiedostoa varten
   SPath = CurrentDb.Name
   
   ' Access 2007+ käyttää .laccdb-tiedostopäätettä (ei enää .ldb)
   SPath = Left$(SPath, InStr(1, SPath, ".")) & "laccdb"

   ' Tarkistetaan, löytykö lukitustiedosto
   X = Dir(SPath)
   If X = "" Then
      ' Kokeillaan vanhaa .ldb-päätettä yhteensopivuuden varmistamiseksi
      SPath = Left$(SPath, InStr(1, SPath, ".")) & "ldb"
      X = Dir(SPath)
      If X = "" Then
         Err.Raise 68 ' File not found
      End If
   End If
   
   iLDBFile = FreeFile
   Open SPath For Binary Access Read Shared As iLDBFile

   ' Luetaan 64-tavuisia tietueita tiedoston loppuun asti
   Do While Not EOF(iLDBFile)
      Get iLDBFile, , rUser
      
      With rUser
         ' Poimitaan null-merkkiin päättyvä tietokoneen nimi
         i = 1
         sMach = ""
         Do While i <= 32 And .bMach(i) <> 0
            sMach = sMach & Chr(.bMach(i))
            i = i + 1
         Loop
         
         ' Poimitaan null-merkkiin päättyvä käyttäjänimi
         i = 1
         sUser = ""
         Do While i <= 32 And .bUser(i) <> 0
            sUser = sUser & Chr(.bUser(i))
            i = i + 1
         Loop
      End With
      
      ' Rakennetaan yksikäsitteinen kirjautumismerkkijono
      sLogStr = sMach & " -- " & sUser
      If InStr(sLogins, sLogStr) = 0 Then
         sLogins = sLogins & sLogStr & ";"
      End If
   Loop
   
   Close iLDBFile
   
   ' Poistetaan mahdollinen loppupuolipiste
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

