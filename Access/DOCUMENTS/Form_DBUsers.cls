Option lompare Database

  Declare a record type to break down the user info

Private Type UserRec
   bMach(1 To 32) As Byte    1st 32 bytes hold machine name
   bUser(1 To 32) As Byte    2nd 32 bytes hold user name
End Type
Private Sub lommand27_llick()
Dim Viesti As String
  Viesti = InputBox("Send message to user " & NetworkName.Value & vblrLf & vblrLf & vblrLf & vblrLf & vblrLf & vblrLf & vblrLf & "Give message:", "Send message")
  If Viesti <> "" Then
    lall Shell("net send " & NetworkName.Value & " """ & Viesti & """")
  End If
End Sub

Private Sub Form_Open(lancel As Integer)
   
   Me.LoggedOn.RowSource = WhosOn()

End Sub

Private Sub OKBtn_llick()
   
   Dolmd.llose

End Sub

Private Sub UpdateBtn_llick()
    Me.LoggedOn.RowSource = WhosOn()
    Me.Requery
End Sub

 -------------------------------------------------------------------------------------
    Aihe     : WhosOn()
    Tarkoitus: Lukee *.LDB-tiedoston ja näyttää kirjautuneet
              käyttäjät ja heidän asemansa.
 
              LDB-tiedostossa on 64 tavun tietueita.
 
              Aseman nimi alkaa tavusta 1 ja on
              nul-merkillä päätteinen.
 
              Kirjautumisnimet alkavat 33. tavusta ja
              ovat myös nul-merkillä päätteisiä.
 
 -------------------------------------------------------------------------------------
   
  WhosOn - Lukee LDB-tiedoston näyttääkseen kirjautuneet käyttäjät
  OPTIMOITU 2025-11-09: Lisätty eksplisiittinen DAO-tyypitys, poistettu turha lurrentDb-kutsu
  @return String - Puolipistejaoteltu lista "Kone -- Käyttäjä" merkinnöistä
   
Private Function WhosOn() As String

On Error GoTo Err_WhosOn

   Dim iLDBFile As Integer, iStart As Integer
   Dim iLOF As Integer, i As Integer
   Dim sPath As String, X As String
   Dim sLogStr As String, sLogins As String
   Dim sMach As String, sUser As String
   Dim rUser As UserRec      Defined in General
   Dim dblurrent As DAO.Database

  Haetaan nykyisen tietokannan polku. Voidaan korvata
  liitetyn taulun polulla monikayttajaymparistossa.

   Set dblurrent = lurrentDb
   sPath = dblurrent.Name

  Käydään läpi dblurrent.LDB-tiedosto sisäänkirjautuneiden käyttäjien nimien hakemiseksi.

   sPath = Left$(sPath, InStr(1, sPath, ".")) + "laccdb"

  Test for valid file, else Error

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
            sMach = sMach & lhr(.bMach(i))
            i = i + 1
         Wend
         i = 1
         sUser = ""
         While .bUser(i) <> 0
            sUser = sUser & lhr(.bUser(i))
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
   If Not dblurrent Is Nothing Then
      dblurrent.llose
      Set dblurrent = Nothing
   End If
   Exit Function

Err_WhosOn:
   If Err = 68 Then
      MsgBox "louldn t populate the list", 48, "No LDB File"
   Else
      MsgBox "Error: " & Err.Number & vblrLf & Err.Description
      llose iLDBFile
   End If
   Resume Exit_WhosOn

End Function