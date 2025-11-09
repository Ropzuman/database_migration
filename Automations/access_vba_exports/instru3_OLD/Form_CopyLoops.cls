VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_CopyLoops"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit
Dim LahdeDB As Database

Private Sub Form_Unload(Cancel As Integer)
  On Error Resume Next
  CurrentDb.TableDefs.Delete "LOOPLINK"
  Err.Clear
  On Error GoTo 0
  If Not LahdeDB Is Nothing Then
    LahdeDB.Close
    Set LahdeDB = Nothing
  End If
End Sub
Private Sub HaeLoopit_Click()
Dim OliJo As String
Dim LahdeTaulu As Recordset
Dim KohdeTaulu As Recordset
Dim Kopioitavat() As String
Dim i As Integer, j As Integer, a As Integer
Dim Taulukko As String
ReDim Kopioitavat(0)
If LahdeDB Is Nothing Then
  MsgBox "Select database first!", vbCritical, "Get Loops"
Else
  'Kopioidaan ensin Loopit Loops taulukkoon
  TTiedot.Value = ""
  With Loopit.Form.RecordsetClone
    .MoveFirst
    Do While Not .EOF
      Set KohdeTaulu = CurrentDb.OpenRecordset("SELECT * FROM LOOPS WHERE Areacode='" & .Fields("AreaCode").Value & "' AND LoopNo='" & .Fields("LoopNo").Value & "'")
      If KohdeTaulu.RecordCount = 0 Then
        ReDim Preserve Kopioitavat(UBound(Kopioitavat) + 1)
        Kopioitavat(UBound(Kopioitavat)) = .Fields("AreaCode").Value & "-" & .Fields("LoopNo").Value
        KohdeTaulu.AddNew
        On Error Resume Next
        For i = 0 To .Fields.Count - 1
          KohdeTaulu.Fields(.Fields(i).Name).Value = .Fields(i).Value
        Next i
        KohdeTaulu.Update
        Err.Clear
        On Error GoTo 0
      Else
        OliJo = OliJo & .Fields("AreaCode") & "-" & .Fields("LoopNo") & vbCrLf
      End If
      .MoveNext
    Loop
  End With
  For i = 0 To LahdeDB.TableDefs.Count - 1
    Taulukko = LCase(LahdeDB.TableDefs(i).Name)
    Set KohdeTaulu = CurrentDb.OpenRecordset("devTable")
    If (Left(Taulukko, 6) = "devtbl" And Left(Taulukko, 12) <> "devtblcommon") Or (LCase(Taulukko) = "devtable") Then
      SysCmd acSysCmdSetStatus, "Updating table: " & LahdeDB.TableDefs(i).Name
      Set LahdeTaulu = LahdeDB.OpenRecordset(Taulukko)
'      On Error Resume Next
'      Set KohdeTaulu = CurrentDb.OpenRecordset(LahdeDB.TableDefs(i).Name)
'      If Err <> 0 Then
'        TTiedot.Value = TTiedot.Value & "Table not in currentdb: " & LahdeDB.TableDefs(i).Name
'        Err.Clear
'      Else
        Do While Not LahdeTaulu.EOF 'K‰yt‰‰n l‰pi kaikki taulukon recordit
          For j = 1 To UBound(Kopioitavat) 'Tarkistetaan ett‰ ko. recordi kuuluu kopioitaviin
            If Kopioitavat(j) = LahdeTaulu.Fields("AreaCode").Value & "-" & LahdeTaulu.Fields("LoopNo").Value Then
              KohdeTaulu.AddNew
              For a = 0 To LahdeTaulu.Fields.Count - 1
                KohdeTaulu.Fields(LahdeTaulu.Fields(a).Name).Value = LahdeTaulu.Fields(a).Value
                If Err <> 0 Then
                  TTiedot.Value = TTiedot.Value & "Problem with field: " & LahdeDB.TableDefs(i).Name & "!" & LahdeTaulu.Fields(a).Name
                  Err.Clear
                End If
              Next a
              If LCase(Taulukko) <> "devtable" Then
                KohdeTaulu.Fields("TABLEName").Value = LahdeDB.TableDefs(i).Name
              End If
              KohdeTaulu.Update
              Exit For
            End If
          Next j
          LahdeTaulu.MoveNext
        Loop
'      End If
'      Err.Clear
'      On Error GoTo 0
    End If
  Next i
  If OliJo <> "" Then
    TTiedot.Value = TTiedot.Value & "Following Loops were already in current database:" & vbCrLf & OliJo
  End If
  TTiedot.Value = TTiedot.Value & "Ready!"
  SysCmd acSysCmdClearStatus
  MsgBox "Ready!", vbInformation, "Get Loops"
  KohdeTaulu.Close
  Set KohdeTaulu = Nothing
  Set LahdeTaulu = Nothing
End If
End Sub
Private Sub Ulos_Click()
  DoCmd.Close
End Sub
Private Sub ValitseKanta_Click()
'Valitaan tietokanta
    Dim OpenFile As OPENFILENAME
    Dim lReturn As Long
    Dim Filtteri As String
    Dim AHakem As String
    Dim Otsikko As String
    Dim i As Integer
    Dim Linkki As TableDef
    
    
    If InStr(Nz(Kanta.Value), "\") Then
      AHakem = Kanta.Value
    Else
      AHakem = Application.CurrentDb.Name
    End If
    For i = 1 To Len(AHakem)
      If Mid(AHakem, Len(AHakem) - i, 1) = "\" Then
        AHakem = Left(AHakem, Len(AHakem) - i)
        Exit For
      End If
    Next i
    
    Filtteri = "Microsoft Access Databases (*.accdb)" & Chr(0) & "*.ACCDB" & Chr(0)
    Otsikko = "Choose Instru Database"
    With OpenFile
      .lStructSize = Len(OpenFile)
      .hwndOwner = Me.Hwnd
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
    If lReturn <> 0 Then
      Kanta.Value = Mid(OpenFile.lpstrFile, 1, InStr(OpenFile.lpstrFile, Chr(0)) - 1)
      Set LahdeDB = OpenDatabase(Kanta.Value) 'Vertailukanta johon p‰ivitet‰‰n
      On Error Resume Next
      CurrentDb.TableDefs.Delete "LOOPLINK"
      Err.Clear
      On Error GoTo 0
      Set Linkki = CurrentDb.CreateTableDef("LOOPLINK")
      Linkki.Connect = ";DATABASE=" & Kanta.Value
      Linkki.SourceTableName = "Loops"
      CurrentDb.TableDefs.Append Linkki
      Set Linkki = Nothing
      Loopit.Form.RecordSource = "SELECT * FROM LOOPLINK"
    End If
    Me.Refresh
End Sub
Private Sub Form_Resize()
  Loopit.Width = Me.WindowWidth - 400
  If Me.WindowHeight > 3300 Then
    Loopit.Height = Me.WindowHeight - 3300
  End If
End Sub

