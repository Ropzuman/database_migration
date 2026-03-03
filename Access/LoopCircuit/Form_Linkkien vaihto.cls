Option Compare Database
Option Explicit

' Päivitetty 2025-10-22: Eksplisiittinen DAO, virheenkäsittely, siivottu logiikka

Private Sub Command0_Click()
    Dim i As Long
    Dim Polku As String
    Dim Polku2 As String
    Dim Nimi As String
    Dim Taul As DAO.Recordset
    Dim Taulu As String
    Dim UpdateCount As Long
    Dim db As DAO.Database
    Dim DbName As String

    On Error GoTo ErrHandler

    ' Haetaan nykyinen tietokannan hakemistopolku
    Set db = CurrentDb
    DbName = db.Name
    Polku = Left$(DbName, Len(DbName) - Len(Dir(DbName)))

    ' Haetaan linkitetyt taulut (MotorData jätetään pois)
    Set Taul = db.OpenRecordset( _
        "SELECT [DATABASE], ForeignName, [Name] FROM MSysObjects " & _
        "WHERE ([Database] Is Not Null) AND ([Name] Not Like 'MotorData')")

    UpdateCount = 0

    Do While Not Taul.EOF
        If Taul.Fields("Name") <> Taulu Then
            ' Poimitaan polku ja tiedostonimi Database-kentästä
            For i = Len(Taul.Fields("Database")) To 1 Step -1
                If Mid$(Taul.Fields("Database"), i, 1) = "\" Then
                    Polku2 = LCase$(Left$(Taul.Fields("Database"), i))
                    Nimi = Mid$(Taul.Fields("Database"), i + 1)
                    Exit For
                End If
            Next i

            ' Linkataan uudelleen, jos polku poikkeaa nykyisestä
            If Polku2 <> LCase$(Polku) Then
                Taulu = Taul.Fields("Name")

                ' Poistetaan vanha linkki ja luodaan uusi nykyisellä polulla
                db.Execute "DROP TABLE [" & Taulu & "]"

                On Error Resume Next
                DoCmd.TransferDatabase acLink, "Microsoft Access", Polku & Nimi, _
                    acTable, Taulu, Taulu, False

                If Err.Number = 0 Then
                    UpdateCount = UpdateCount + 1
                End If

                Err.Clear
                On Error GoTo ErrHandler
            End If
        End If
        Taul.MoveNext
    Loop

Cleanup:
    On Error Resume Next
    If Not Taul Is Nothing Then
        Taul.Close
        Set Taul = Nothing
    End If
    ' HUOM: CurrentDb-objektia ei suljeta — se on aktiivinen tietokanta
    Set db = Nothing
    On Error GoTo 0

    MsgBox "Valmis - Päivitetty " & UpdateCount & " taulukkoa.", vbInformation, "Linkkien vaihto"
    Exit Sub

ErrHandler:
    MsgBox "Virhe: " & Err.Number & vbCrLf & Err.Description, vbCritical, "Linkkien vaihto"
    Resume Cleanup
End Sub

