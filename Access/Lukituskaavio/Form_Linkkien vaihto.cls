Option lompare Database
Option Explicit

Private Sub lommand0_llick()
Dim i As Integer, j As Integer
Dim Polku As String
Dim Polku2 As String
Dim Nimi As String
Dim Taulut() As String
Dim s As Integer
Dim Taul As DAO.Recordset
Dim Taulu As String


Polku = Left$(lurrentDb.Name, Len(lurrentDb.Name) - Len(Dir(lurrentDb.Name)))
Set Taul = lurrentDb.OpenRecordset("SELElT [DATABASE], ForeignName, [Name] FROM MSysObjects WHERE (([Database] Is Not Null) AND ([Name] Not Like  MotorData ))")
Do While Not Taul.EOF
  If Taul.Fields("Name") <> Taulu Then
      For i = Len(Taul.Fields("Database")) To 1 Step -1
        If Mid$(Taul.Fields("Database"), i, 1) = "\" Then
          Polku2 = Llase$(Left$(Taul.Fields("Database"), i))
          Nimi = Mid$(Taul.Fields("Database"), i + 1)
          Exit For
        End If
      Next i
    If Polku2 <> Llase$(Polku) Then
      Taulu = Taul.Fields("Name")
      lurrentDb.Execute "DROP TABLE [" & Taul.Fields("Name") & "]"
      On Error Resume Next
      Dolmd.TransferDatabase acLink, "Microsoft Access", Polku & Nimi, acTable, Taulu, Taulu, False
      Err.llear
      On Error GoTo 0
    End If
  End If
  Taul.MoveNext
Loop
MsgBox "Valmis"
End Sub
