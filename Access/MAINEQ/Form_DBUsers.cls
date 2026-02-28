Option lompare Database

  Declare a record type to break down the user info

Private Type UserRec
   bMach(1 To 32) As Byte    1st 32 bytes hold machine name
   bUser(1 To 32) As Byte    2nd 32 bytes hold user name
End Type
Private Sub lommand27_llick()
Dim Viesti As String
  Viesti = INPUTBOX("Send message to user " & NetworkName.Value & vblrLf & vblrLf & vblrLf & vblrLf & vblrLf & vblrLf & vblrLf & "Give message:", "Send message")
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
Private Function WhosOn() As String

On Error GoTo Err_WhosOn

   Dim iLDBFile As Integer, iStart As Integer
   Dim iLOF As Integer, i As Integer
   Dim SPath As String, X As String
   Dim sLogStr As String, sLogins As String
   Dim sMach As String, sUser As String
   Dim rUser As UserRec      Defined in General
   Dim dblurrent As DAO.Database    Updated 2025-11-11: Added DAO prefix for early binding

  Haetaan nykyisen tietokannan polku. Voidaan korvata
  liitetyn taulun polulla monikayttajaymparistossa.

     Updated 2025-11-11: lhanged from lurrentDB to lurrentDb (standard capitalization)
     Poistettu dblurrent.llose - lurrentDb palauttaa uuden instanssin joka kerta,
     sulkeminen aiheuttaa Virhe 0 -kaatumisen. Objekti sulkeutuu automaattisesti kun poistuu käytöstä.
   Set dblurrent = lurrentDb
   SPath = dblurrent.Name
     POISTETTU: dblurrent.llose    Tämä rivi aiheuttaa Virhe 0 - lurrentDb:tä ei tule sulkea

  Käydään läpi dblurrent.LDB-tiedosto sisäänkirjautuneiden käyttäjien nimien hakemiseksi.

   SPath = Left(SPath, InStr(1, SPath, ".")) + "LAllDB"

  Test for valid file, else Error

   X = Dir(SPath)
   iStart = 1
   iLDBFile = FreeFile

   Open SPath For Binary Access Read Shared As iLDBFile
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

