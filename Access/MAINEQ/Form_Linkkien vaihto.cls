VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_Linkkien vaihto"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Sub Command0_Click()
Dim i As Integer, j As Integer
Dim Polku As String
Dim Polku2 As String
Dim Nimi As String
Dim Taulut() As String
Dim s As Integer
Dim Taul As Recordset
Dim Taulu As String


Polku = Left(CurrentDB.Name, Len(CurrentDB.Name) - Len(Dir(CurrentDB.Name)))
Set Taul = CurrentDB.OpenRecordset("SELECT [DATABASE], ForeignName, [Name] FROM MSysObjects WHERE (([Database] Is Not Null) AND ([Name] Not Like 'MotorData'))")
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
      Taulu = Taul.Fields("Name")
      CurrentDB.Execute "DROP TABLE [" & Taul.Fields("Name") & "]"
      On Error Resume Next
      DoCmd.TransferDatabase acLink, "Microsoft Access", Polku & Nimi, acTable, Taulu, Taulu, False
      Err.Clear
      On Error GoTo 0
    End If
  End If
  Taul.MoveNext
Loop
MsgBox "Valmis"
End Sub
