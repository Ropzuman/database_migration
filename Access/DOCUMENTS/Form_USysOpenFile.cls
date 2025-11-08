Option Compare Database
Option Explicit
Dim Hakem As String
Dim fso As New FileSystemObject
Private Sub Command4_Click()
  DoCmd.Close
End Sub
Private Sub Command52_Click()
  If fso.FolderExists(Hakem) Then
    Call Shell("C:\WINNT\EXPLORER.EXE /e, " & Hakem, vbNormalFocus)
  Else 'Hakemistoa ei ole
    MsgBox "Invalid path! " & Hakem, vbCritical, "Open Path"
  End If
End Sub

Private Sub Form_Load()
  Dim DB As Database
  Dim taulu As Recordset
  Dim Tiedostot As String
  Dim Rivit As String
  Dim Osoitin As Long
  
  Set DB = CurrentDb
  Set taulu = DB.OpenRecordset("SELECT File, WorkPath FROM DOCUMENTS WHERE Counter=" & Common)
  Tiedostot = taulu.Fields(0).Value & vbCrLf
  Hakem = taulu.Fields(1).Value
  MyPath.Caption = "Path: " & Hakem
  Osoitin = InStr(Tiedostot, vbCrLf)
  Rivit = Left(Tiedostot, Osoitin - 1)
  Tiedostot = Mid(Tiedostot, Osoitin + 2)
  Do
    Osoitin = InStr(Tiedostot, vbCrLf)
    If Osoitin = 0 Then Exit Do
    Rivit = Rivit & ";" & Left(Tiedostot, Osoitin - 1)
    Tiedostot = Mid(Tiedostot, Osoitin + 2)
  Loop
  Tied.RowSource = Rivit
  Set taulu = Nothing
  Set DB = Nothing
End Sub
Private Sub Form_Unload(Cancel As Integer)
  Set fso = Nothing
End Sub
Private Sub Tied_DblClick(Cancel As Integer)
  If fso.FileExists(Hakem & Tied.Value) Then
    Application.FollowHyperlink Hakem & Tied.Value, , True, False
  Else
      MsgBox "File not found: " & Hakem & Tied.Value, vbCritical, "Open File"
  End If
End Sub