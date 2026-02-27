VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} StartForm 
   Caption         =   "Interlocking"
   ClientHeight    =   1695
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4680
   OleObjectBlob   =   "StartForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "StartForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub CommandButton1_Click()
Dim TAULU As New ADODB.Recordset
  If Dir(Tietokannat.Value) <> "" Then
    DB.Open "Driver=Microsoft Access Driver (*.mdb);DBQ=" & Tietokannat.Value & ";"
    Tietokanta = Tietokannat.Value
    On Error GoTo Virhe
    Set TAULU = DB.Execute("Select * From TOACAD")
    Set TAULU = DB.Execute("Select * From SETTINGS")
    Set TAULU = DB.Execute("Select * From TOACADLoopEq")
    Set TAULU = Nothing
    StartOK = True
    Unload Me
    Exit Sub
Virhe:
    Tietokannat.BackColor = RGB(255, 150, 150)
    MsgBox "Database: " & vbCrLf & Tietokannat.Value & vbCrLf & _
           "did not contain necessary tables (TOACAD, TOACADLoopEq AND SETTINGS)!", vbCritical, "Error!"
    Tietokannat.SetFocus
    Err.Clear
    Set TAULU = Nothing
    Exit Sub
  Else 'Tietokantaa ei löytynyt
    Tietokannat.BackColor = RGB(255, 150, 150)
    MsgBox "Database: " & vbCrLf & Tietokannat.Value & vbCrLf & "did not found!", vbCritical, "Error!"
    Tietokannat.SetFocus
  End If
End Sub
Private Sub CommandButton2_Click()
  Set DB = Nothing
  Unload Me
End Sub
Private Sub CommandButton3_Click()
'Tämä valitsee tiedoston hakemistosta
    Dim OpenFile As OPENFILENAME
    Dim lReturn As Long
    Dim Filtteri As String
    Dim AHakem As String
    Dim Otsikko As String
    Dim WHandle As Long
    
    WHandle = FindWindow(0&, "Interlocking")
    AHakem = "L:\Projdata\"
    
    Filtteri = "MS Access Database (*.mdb)" & Chr(0) & "*.MDB" & Chr(0)
    Otsikko = "Choose Database"
    With OpenFile
      .lStructSize = Len(OpenFile)
      .hwndOwner = WHandle
      .hInstance = 0
      .lpstrFilter = Filtteri
      .nFilterIndex = 1
      .lpstrFile = String(257, 0)
      .nMaxFile = Len(.lpstrFile) - 1
      .lpstrFileTitle = .lpstrFile
      .nMaxFileTitle = .nMaxFile
      .lpstrInitialDir = AHakem
      .lpstrTitle = Otsikko
      .flags = 0
    End With
    lReturn = GetOpenFileName(OpenFile)
    If lReturn = 0 Then
        'Painettiin Cancel painiketta. Ei tehdä mitään
    Else 'Otetaan ylös Tiedostonimi ja Hakemisto
      Tietokannat.AddItem Mid(OpenFile.lpstrFile, 1, InStr(OpenFile.lpstrFile, Chr(0)) - 1)
      Tietokannat.ListIndex = Tietokannat.ListCount - 1
    End If
End Sub
Private Sub UserForm_Initialize()
  Tietokannat.AddItem "K:\Projects\N_whldata\Santa Fe 220018\DATABASE\Instru.mdb"
  Tietokannat.ListIndex = 0
End Sub
