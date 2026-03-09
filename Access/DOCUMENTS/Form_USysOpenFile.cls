Option Compare Database
Option Explicit
'==========================================================================
' LOMAKE   : Form_USysOpenFile
' SOVELLUS : DOCUMENTS — Tiedoston avaus
' KUVAUS   : Näyttää dokumenttiin liitetyt tiedostot listassa.
'            Kaksoisnapsautuksella avataan tiedosto FollowHyperlink-kutsulla.
'            Hakemiston voi avata Explorerissa Shell-komennolla.
' PÄIVITETTY: 2026-03-03
'==========================================================================
Dim Hakem As String
Dim fso As New FileSystemObject
Private Sub Command4_Click()
  DoCmd.Close
End Sub
Private Sub Command52_Click()
  ' Avataan hakemisto turvallisesti käyttöjärjestelmän omalla menettelyllä
  ' FollowHyperlink välttää Shell-komentoinjektioriskin kokonaan
  If fso.FolderExists(Hakem) Then
    On Error Resume Next
    Application.FollowHyperlink Hakem
    If Err.Number <> 0 Then
      MsgBox "Failed to open folder.", vbCritical, "Error"
    End If
    On Error GoTo 0
  Else
    MsgBox "Invalid path! " & Hakem, vbCritical, "Open Path"
  End If
End Sub

'''
' Form Load Event - Load document files from database
' OPTIMIZED 2025-11-09: Added error handling, explicit DAO typing, null-checking
'''
Private Sub Form_Load()
  Dim DB As DAO.Database
  Dim taulu As DAO.Recordset
  Dim Tiedostot As String
  Dim Rivit As String
  Dim Osoitin As Long
  
  On Error GoTo ErrorHandler
  
  Set DB = CurrentDb
  Set taulu = DB.OpenRecordset("SELECT File, WorkPath FROM DOCUMENTS WHERE Counter=" & Common, dbOpenSnapshot)
  
  If Not taulu.EOF Then
    Tiedostot = Nz(taulu.Fields(0).Value, "") & vbCrLf
    Hakem = Nz(taulu.Fields(1).Value, "")
    Me.MyPath.Caption = "Path: " & Hakem
    
    Osoitin = InStr(Tiedostot, vbCrLf)
    If Osoitin > 0 Then
      Rivit = Left$(Tiedostot, Osoitin - 1)
      Tiedostot = Mid$(Tiedostot, Osoitin + 2)
      Do
        Osoitin = InStr(Tiedostot, vbCrLf)
        If Osoitin = 0 Then Exit Do
        Rivit = Rivit & ";" & Left$(Tiedostot, Osoitin - 1)
        Tiedostot = Mid$(Tiedostot, Osoitin + 2)
      Loop
      Me.Tied.RowSource = Rivit
    End If
  End If
  
Cleanup:
  If Not taulu Is Nothing Then
    taulu.Close
    Set taulu = Nothing
  End If
  If Not DB Is Nothing Then
    DB.Close
    Set DB = Nothing
  End If
  Exit Sub
  
ErrorHandler:
  MsgBox "Error loading files: " & Err.Description, vbExclamation, "Form Load Error"
  Resume Cleanup
End Sub
Private Sub Form_Unload(Cancel As Integer)
  Set fso = Nothing
End Sub
Private Sub Tied_DblClick(Cancel As Integer)
  If fso.FileExists(Hakem & Me.Tied.Value) Then
    Application.FollowHyperlink Hakem & Me.Tied.Value, , True, False
  Else
      MsgBox "File not found: " & Hakem & Me.Tied.Value, vbCritical, "Open File"
  End If
End Sub