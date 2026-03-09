Option Compare Database
Option Explicit

Private Sub Command0_Click()
    Dim i As Integer, j As Integer
    Dim Polku As String
    Dim Polku2 As String
    Dim Nimi As String
    Dim Taulut() As String
    Dim s As Integer
    Dim Taul As DAO.Recordset
    Dim Taulu As String
    Dim UusiPolku As String

    Polku = Left(CurrentDb.Name, Len(CurrentDb.Name) - Len(Dir(CurrentDb.Name)))
    Set Taul = CurrentDb.OpenRecordset("SELECT [DATABASE], ForeignName, [Name] FROM MSysObjects WHERE (([Database] Is Not Null) AND ([Name] Not Like 'MotorData'))")
    Do While Not Taul.EOF
      If Taul.Fields("Name") <> Taulu Then
          For i = Len(Taul.Fields("Database")) To 1 Step -1
            If Mid(Taul.Fields("Database"), i, 1) = "\" Then
              Polku2 = LCase(Left(Taul.Fields("Database"), i))
              Nimi = Mid(Taul.Fields("Database"), i + 1)
              Exit For
            End If
          Next i
        If Polku2 <> LCase(Polku) Then
          UusiPolku = Polku & Nimi
          ' Tarkistetaan uuden tietokantatiedoston olemassaolo ENNEN vanhan poistoa
          ' — näin vältetään taulujen pysyvä häviäminen, jos verkkolevyä ei löydy
          If Dir(UusiPolku) = "" Then
            MsgBox "Tietokantatiedostoa ei löydy:" & vbCrLf & UusiPolku & vbCrLf & vbCrLf & _
                   "Taulun '" & Taul.Fields("Name") & "' linkitystä ei muutettu.", _
                   vbCritical, "Linkitysvirhe"
            Taul.MoveNext
            GoTo JatkaSilmukka
          End If
          Taulu = Taul.Fields("Name")
          CurrentDb.Execute "DROP TABLE [" & Taul.Fields("Name") & "]"
          On Error Resume Next
          DoCmd.TransferDatabase acLink, "Microsoft Access", UusiPolku, acTable, Taulu, Taulu, False
          Err.Clear
          On Error GoTo 0
        End If
      End If
JatkaSilmukka:
      Taul.MoveNext
    Loop
    Taul.Close
    Set Taul = Nothing
    MsgBox "Valmis"
End Sub

