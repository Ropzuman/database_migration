VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Formi 
   Caption         =   "Execute Utility"
   ClientHeight    =   5400
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8145
   OleObjectBlob   =   "Formi.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "Formi"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim PFile As Integer
Dim Ajastin As Long
Dim SDocMode As Boolean
Dim FSO As FileSystemObject

Private Sub Tiedostosta_Click()
'Tämä valitsee tiedoston hakemistosta
    Dim OpenFile As OPENFILENAME
    Dim lReturn As Long
    Dim Filtteri As String
    Dim AHakem As String
    Dim Otsikko As String
    Dim WHandle As LongPtr
    Dim Tiedosto As String
    Dim Tiedot As TextStream
    
    WHandle = FindWindow(0&, "Execute Utility")
    If InStr(FPath.Value, "\") Then
      AHakem = FPath.Text
    Else
      AHakem = "P:\PROJEKTI\"
    End If
    
    Filtteri = "Script Files (*.scr; *.txt)" & Chr(0) & "*.SCR;*.TXT" & Chr(0) _
             & "All Files (*.*)" & Chr(0) & "*.*" & Chr(0)
    Otsikko = "Choose File for Script"
    With OpenFile
      .lStructSize = LenB(OpenFile)
      .hwndOwner = WHandle
      .hInstance = 0
      .lpstrFilter = Filtteri
      .nFilterIndex = 1
      .lpstrFile = String(257, 0)
      .nMaxFile = LenB(.lpstrFile) - 1
      .lpstrFileTitle = .lpstrFile
      .nMaxFileTitle = .nMaxFile
      .lpstrInitialDir = AHakem
      .lpstrTitle = Otsikko
      .flags = 0
    End With
    lReturn = GetOpenFileName(OpenFile)
    If lReturn <> 0 Then
      Tiedosto = Left(OpenFile.lpstrFile, InStr(OpenFile.lpstrFile, Chr(0)) - 1)
      Set Tiedot = FSO.OpenTextFile(Tiedosto)
      Komennot.Value = Tiedot.ReadAll
      Set Tiedot = Nothing
    End If
End Sub

Private Sub Ulos_Click()
  Unload Formi
End Sub
Private Sub UserForm_Initialize()
'Avaus
Dim i As Integer
Dim Tulostimet As Variant
Dim Tyylit As Variant
Dim Koot As Variant
Dim oTiedosto As Scripting.TextStream
Dim Tiedot As String
  'Laitetaan pysty ja vaaka vaihtoehdot
  Set FSO = New FileSystemObject
  If FSO.FileExists(Application.Preferences.Files.TempFilePath & "ACADExec.TXT ") Then
    Set oTiedosto = FSO.OpenTextFile(Application.Preferences.Files.TempFilePath & "ACADExec.TXT")
    Tiedot = oTiedosto.ReadLine
    oTiedosto.Close
    FPath.Value = Left(Tiedot, InStr(Tiedot, vbTab) - 1)
    FName.Value = Mid(Tiedot, InStr(Tiedot, vbTab) + 1)
  End If
  Set oTiedosto = Nothing
End Sub
Private Sub KirjoitaTiedot()
Dim oTiedosto As Scripting.TextStream
  Set oTiedosto = FSO.CreateTextFile(Application.Preferences.Files.TempFilePath & "ACADExec.TXT", True)
  oTiedosto.Write FPath.Value & vbTab
  oTiedosto.Write FName.Value & vbCrLf
  oTiedosto.Close
  Set oTiedosto = Nothing
End Sub
Private Sub Suorita_Click()
  Kay_Lapi_Tiedostot True    'Suorittaa toiminnon kaikille dokumenteille
End Sub
Private Sub Testi_Click()
  Kay_Lapi_Tiedostot False 'Suorittaa toiminnon yhdelle listan dokumenteista
End Sub
Private Sub TestaaNykyinen_Click()
  SuoritaToiminto
End Sub
Private Sub Kay_Lapi_Tiedostot(Kaikki As Boolean)
Dim Paate As String
Dim Tiedosto As String
Dim i As Integer
    Virhe = False
    Ajastin = Timer
    If FPath.Value = "" Or FName.Value = "" Then
      MsgBox "Please choose path and file(s) first.", vbCritical, "Plot"
      Exit Sub
    End If
    If MsgBox("This will close alla open documents and start plotting. Do you wish to do this?", vbOKCancel, "Plot Files") = vbCancel Then
      Exit Sub
    End If
    'Varmistetaan että ollaan monen dokumentin tilassa
   ' SDocMode = Application.Preferences.System.SingleDocumentMode
   ' Application.Preferences.System.SingleDocumentMode = False
    'Tiedoston nimi on määritelty
    Formi.MousePointer = fmMousePointerHourGlass
    PFile = 1
    Ajastin = Timer
    
    If Right(FPath.Value, 1) <> "\" Then FPath.Value = FPath.Value & "\"
    Tiedosto = Dir(FPath.Value & FName.Value)
    Do While Tiedosto <> ""
      Paate = UCase(Right(Tiedosto, 4))
      If Paate = ".DWG" Or Paate = ".DXF" Then
        If AvaaTiedosto(FPath.Value & Tiedosto) = True Then 'Tiedoston avaus onnistui
          Tila Tiedosto
          SuoritaToiminto
          If Virhe Then Exit Do
          PFile = PFile + 1
          If Kaikki = False Then  'Lopetetaan tähän jos on kysymys testitulostuksesta
            Exit Do
          End If
        End If
      ElseIf Paate = ".LST" Then
        LueTiedosto FPath.Value & Tiedosto, Kaikki
        If Kaikki = False Then  'Lopetetaan tähän jos on kysymys testitulostuksesta
          Exit Do
        End If
      End If
      Tiedosto = Dir
    Loop
    'Suljetaan viimeinen dokumentti, jotta se ei jää vahingossa auki
    For i = 0 To Application.Documents.Count - 1
      Application.Documents(0).Close False
    Next i
 '   Application.Documents.Add 'Tehdään uusi dokumentti jotta voidaan siirtyä single tilaan. Single tilaan ei voi siirtyä, jollei yhtään dokumenttia ole auki
 '   Application.Preferences.System.SingleDocumentMode = SDocMode 'Siirrytään tilaan, joka oli ennen ohjelman käynnistystä
    TilaTieto.Caption = ""
    MsgBox "Script Executed for Drawing(s)", , "Ready"
    Formi.MousePointer = fmMousePointerDefault
    
    'Formi.Hide
    'KirjoitaTiedot
    ActiveDocument.SetVariable "FILEDIA", 1 'MRU 20111017
    
End Sub
Private Sub LueTiedosto(Tiedosto As String, Kaikki As Boolean)
Dim oTiedosto As Scripting.TextStream
Dim Rivi As String
Dim i As Integer
Dim Alku As String
    Set oTiedosto = FSO.OpenTextFile(Tiedosto, 1)
    Do While Not oTiedosto.AtEndOfStream
      DoEvents
      Rivi = oTiedosto.ReadLine
      Alku = Mid(Rivi, 1, 1)
      If Alku = " " Or Alku = ";" Or Rivi = "" Then
        'Ei mitään
      ElseIf Alku = "@" Then
        LueTiedosto Mid(Rivi, 2), Kaikki
      Else
        If Len(Rivi) > 41 Then
          Tiedosto = FPath.Value & Mid(Rivi, 40)
        Else
          Tiedosto = FPath.Value & Rivi
        End If
        If Right(UCase(Tiedosto), 4) <> ".DWG" Then Tiedosto = Tiedosto & ".DWG"
        If AvaaTiedosto(Tiedosto) Then
            Tila Tiedosto
            SuoritaToiminto
            PFile = PFile + 1
            If Kaikki = False Then  'Lopetetaan tähän jos on kysymys testitulostuksesta
              Exit Do
            End If
        End If
      End If
    Loop
    Set oTiedosto = Nothing
End Sub
Private Sub SuoritaToiminto()
  ActiveDocument.SendCommand Replace(Komennot.Value, vbCrLf, vbCr)
  Do Until Application.GetAcadState.IsQuiescent = True
  Loop
  
End Sub
Private Function AvaaTiedosto(Avattava As String) As Boolean
'Avaa tiedoston AutoCADissä.
Dim i As Integer
    If FSO.FileExists(Avattava) Then
      For i = 0 To Application.Documents.Count - 1
        Application.Documents(0).Close False   'Suljetaan documentit (Multi Document tilassa AutoCAD 2002 ei herjaa vaikkei documentteja olisikaan ja vaikka documentteja ei olisi talletettu)
      Do Until Application.GetAcadState.IsQuiescent = True
      Loop
      Next i
      Application.Documents.Open (Avattava)  'Avataan uusi dokumentti AutoCadiin (ACAD 2002 osaa avata myös DXF documentit ilman kyselyjä)
      AvaaTiedosto = True
      Do Until Application.GetAcadState.IsQuiescent = True
      Loop
    Else
      AvaaTiedosto = False
      MsgBox "File not found: " & Avattava, vbCritical, "Error!"
    End If
End Function
Private Sub UserForm_Terminate()
  KirjoitaTiedot
  Set FSO = Nothing
End Sub
Private Sub Valitse_Click()
'Tämä valitsee tiedoston hakemistosta
    Dim OpenFile As OPENFILENAME
    Dim lReturn As Long
    Dim Filtteri As String
    Dim AHakem As String
    Dim Otsikko As String
    Dim WHandle As LongPtr
    
    WHandle = FindWindow(0&, "Execute Utility")
    If InStr(FPath.Value, "\") Then
      AHakem = FPath.Text
    Else
      AHakem = "P:\PROJEKTI\"
    End If
    
    Filtteri = "AutoCAD Drawing (*.dwg)" & Chr(0) & "*.DWG" & Chr(0) _
             & "AutoCAD DXF (*.dxf)" & Chr(0) & "*.DXF" & Chr(0) _
             & "List Of Drawings (*.lst)" & Chr(0) & "*.LST" & Chr(0) _
             & "All Types (*.dwg;*.dxf;*.lst)" & Chr(0) & "*.DWG;*.LST;*.DXF" & Chr(0)
    Otsikko = "Choose AutoCAD Drawing or List"
    With OpenFile
      .lStructSize = LenB(OpenFile)
      .hwndOwner = WHandle
      .hInstance = 0
      .lpstrFilter = Filtteri
      .nFilterIndex = 1
      .lpstrFile = String(257, 0)
      .nMaxFile = LenB(.lpstrFile) - 1
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
       FName.Value = Mid(OpenFile.lpstrFileTitle, 1, InStr(OpenFile.lpstrFileTitle, Chr(0)) - 1)
       FPath.Value = Mid(OpenFile.lpstrFile, 1, InStr(OpenFile.lpstrFile, Chr(0)) - 1 - Len(FName.Value))
    End If
End Sub
Private Sub Tila(Tiedosto As String)
  sec = CInt(Timer - Ajastin)
  Min = sec \ 60
  sec = sec Mod 60
  TilaTieto.Caption = "Executing script. File: " & PFile & " (" & Tiedosto & ")  Time: " & Min & " min  " & sec & " sec"
  Formi.Repaint
End Sub
