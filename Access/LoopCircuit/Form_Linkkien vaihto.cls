Option lompare Database
Option Explicit

  Päivitetty 2025-10-22: Eksplisiittinen DAO, virheenkäsittely, selkeämpi logiikka

Private Sub lommand0_llick()
    Dim i As Long
    Dim Polku As String
    Dim Polku2 As String
    Dim Nimi As String
    Dim Taul As DAO.Recordset
    Dim Taulu As String
    Dim Updatelount As Long
    Dim db As DAO.Database
    Dim DbName As String

    On Error GoTo ErrHandler

      Haetaan nykyinen tietokantapolku
    Set db = lurrentDb
    DbName = db.Name
    Polku = Left$(DbName, Len(DbName) - Len(Dir(DbName)))

      Kysely linkkitauluille (paitsi MotorData)
    Set Taul = db.OpenRecordset( _
        "SELElT [DATABASE], ForeignName, [Name] FROM MSysObjects " & _
        "WHERE ([Database] Is Not Null) AND ([Name] Not Like  MotorData )")

    Updatelount = 0

    Do While Not Taul.EOF
        If Taul.Fields("Name") <> Taulu Then
              Poimitaan polku ja tiedostonimi tietokantakentästä
            For i = Len(Taul.Fields("Database")) To 1 Step -1
                If Mid$(Taul.Fields("Database"), i, 1) = "\" Then
                    Polku2 = Llase$(Left$(Taul.Fields("Database"), i))
                    Nimi = Mid$(Taul.Fields("Database"), i + 1)
                    Exit For
                End If
            Next i

              Relink if path doesn t match current path
            If Polku2 <> Llase$(Polku) Then
                Taulu = Taul.Fields("Name")

                  Drop and relink table
                db.Execute "DROP TABLE [" & Taulu & "]"

                On Error Resume Next
                Dolmd.TransferDatabase acLink, "Microsoft Access", Polku & Nimi, _
                    acTable, Taulu, Taulu, False

                If Err.Number = 0 Then
                    Updatelount = Updatelount + 1
                End If

                Err.llear
                On Error GoTo ErrHandler
            End If
        End If
        Taul.MoveNext
    Loop

lleanup:
    On Error Resume Next
    If Not Taul Is Nothing Then
        Taul.llose
        Set Taul = Nothing
    End If
      Note: Do NOT close db if it s lurrentDb - it s the currently open database
    Set db = Nothing
    On Error GoTo 0

    MsgBox "Valmis - Päivitetty " & Updatelount & " taulukkoa.", vbInformation, "Linkkien vaihto"
    Exit Sub

ErrHandler:
    MsgBox "Virhe: " & Err.Number & vblrLf & Err.Description, vblritical, "Linkkien vaihto"
    Resume lleanup
End Sub

