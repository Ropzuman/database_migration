VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Formi 
   Caption         =   "Kuvien Selaus"
   ClientHeight    =   7590
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7680
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

' --- 64-bitti korjattu: PtrSafe lisatty, Long -> LongPtr handleille ---
Private Type BrowseInfo
    hWndOwner As LongPtr
    pidlRoot As LongPtr
    sDisplayName As String
    sTitle As String
    ulFlags As Long
    lpfn As LongPtr
    lParam As LongPtr
    iImage As Long
End Type

Private Declare PtrSafe Function SHBrowseForFolder Lib "Shell32.dll" (bBrowse As BrowseInfo) As LongPtr
Private Declare PtrSafe Function SHGetPathFromIDList Lib "Shell32.dll" (ByVal lItem As LongPtr, ByVal sDir As String) As Long

Dim fso As FileSystemObject
Private Sub CommandButton1_Click()
    Formi.Hide
    KirjoitaTiedot
    
    ActiveDocument.SetVariable "FILEDIA", 1 'MRU 20111017
    
End Sub

Private Sub Luo_lst_Click()
Dim Paate As String
Dim tiedosto As String
Dim intI As Integer
Dim CheckPaate As String

    If Right(FPath.Value, 1) <> "\" Then FPath.Value = FPath.Value & "\"

    CheckPaate = UCase(Right(FName.Value, 4))
    If CheckPaate = ".LST" Then
      MsgBox "Tulostuslistaa ei luoda lst-tiedostoista. Kirjoita suodatustieto kohtaan [Tiedosto tai jokerimerkki].", vbOKOnly, "Tulostuslistan luonti"
    FName.SetFocus
    With FName
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
      Exit Sub
End If

tiedosto = Dir(FPath.Value & FName.Value)

Open FPath.Value & "tulostus.lst" For Output As #1
    Do While tiedosto <> ""
      Paate = UCase(Right(tiedosto, 4))
      If Paate = ".DWG" Or Paate = ".DXF" Then
'      MsgBox Tiedosto, vbOKOnly, Tiedostolistaus
    Print #1, tiedosto
'        If AvaaTiedosto(FPath.Value & Tiedosto) = True Then 'Tiedoston avaus onnistui
'          Tila Tiedosto
'          TulostaDOC
'          PFile = PFile + 1
'          If Kaikki = False Then  'Lopetetaan t�h�n jos on kysymys testitulostuksesta
'            Exit Do
'          End If
'        End If
      ElseIf Paate = ".LST" Then
'        LueTiedosto FPath.Value & Tiedosto, Kaikki
'        If Kaikki = False Then  'Lopetetaan t�h�n jos on kysymys testitulostuksesta
'          Exit Do
'        End If
      End If
      tiedosto = Dir
      intI = intI + 1
    Loop
Close #1
Dim Response
Response = MsgBox("Homma hoidettu. Tulostustiedosto luotu " & intI & " kuvan tulostusta varten." & Chr(13) & "Avataanko luotu tiedosto?", vbYesNo)
If Response = vbYes Then    ' User chose Yes.
    'Avaa_lst_tiedosto (FPath.Value & "tulostus.lst")
    MyFinalPathFile = FPath.Value & "tulostus.lst"
    Shell "Notepad.exe" & " " & MyFinalPathFile, vbMaximizedFocus
    'MsgBox "Homma ehk� toimii!", vbOKOnly    ' Perform some action.
'Else    ' User chose No.
'    MsgBox "Homma EI toimi!", vbOKOnly    ' Perform some action.
End If

'MsgBox "Homma hoidettu. Tulostustiedosto luotu " & intI & " kuvan tulostusta varten." & Chr(13) & "Avataanko luotu tiedosto?", vbYesNo
'MsgBox "Tulostustiedosto luotu as vbYesNo, Homma hoidettu"
FName.Value = "tulostus.lst"
Nayta_lista.Enabled = True
End Sub
Private Sub Nayta_lista_Click()
    MyFinalPathFile = FPath.Value & "tulostus.lst"
    Shell "Notepad.exe" & " " & MyFinalPathFile, vbMaximizedFocus
End Sub

Private Sub PltTied_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
If PltTiedCheck.Value = True Then
    PltTied.Value = BrowseForDirectory
End If
End Sub
Private Sub PltTiedCheck_Click()
If PltTiedCheck.Value = False Then
    PltTied.Enabled = False
    PltTied.BackColor = &H80000013
 Else
'    PltTied.Value = BrowseForDirectory
    PltTied.BackColor = &H80000005
    PltTied.Enabled = True
 End If
End Sub
Private Sub select_dwg_Click()
FName.Value = "*.dwg"
Nayta_lista.Enabled = False
Luo_lst.Enabled = True
End Sub

Private Sub UserForm_Initialize()
'Avaus
Dim i As Integer
Dim oTiedosto As Scripting.TextStream
Dim Tiedot As String
Set fso = New FileSystemObject
  If fso.FileExists(Application.Preferences.Files.TempFilePath & "ACADSel.TXT") Then
    Set oTiedosto = fso.OpenTextFile(Application.Preferences.Files.TempFilePath & "ACADSel.TXT")
    Tiedot = oTiedosto.ReadLine
    oTiedosto.Close
    FPath.Value = Left(Tiedot, InStr(Tiedot, vbTab) - 1)
    Tiedot = Mid(Tiedot, InStr(Tiedot, vbTab) + 1)
    FName.Value = Mid(Tiedot, 1)
    'Tiedot = Mid(Tiedot, InStr(Tiedot, vbTab) + 1)
  End If
  Set oTiedosto = Nothing
End Sub
Private Sub KirjoitaTiedot()
Dim oTiedosto As Scripting.TextStream
  Set oTiedosto = fso.CreateTextFile(Application.Preferences.Files.TempFilePath & "ACADSel.TXT", True)
  oTiedosto.Write FPath.Value & vbTab
  oTiedosto.Write FName.Value & vbTab
  oTiedosto.Close
  Set oTiedosto = Nothing
End Sub

Private Sub Tulosta_Click()
Set oBatTiedosto = Nothing
    Kay_Lapi_Tiedostot True 'Tulostaa kaikki dokumentit
'    YhdistaPSjamuunnaPDF
End Sub

Private Sub Kay_Lapi_Tiedostot(Kaikki As Boolean)
Dim Paate As String
Dim TiedostoNimi As String
Dim tiedosto As String
'Dim LueTiedosto As String
Dim i As Integer
  
    Ajastin = Timer
    If FPath.Value = "" Or FPath.Value = "Hakemistoa ei ole valittu" Or FName.Value = "" Then
      MsgBox "Ole hyv� ja valitse polku ja tiedosto ensin.", vbCritical, "Plot"
      Exit Sub
    End If
    If MsgBox("T�m� sulkee kaikki avoimet dokumentit. Oletko varma ett� n�in voidaan tehd�?", vbOKCancel, "Plot Files") = vbCancel Then
      Exit Sub
    End If
   ' If Size.Value <> "A4" Then
   ' If MsgBox("Paperikoko ei ole A4 vaan " & Size.Value & ". Haluatko varmasti jatkaa tulostusta?", vbOKCancel, "Plot Files") = vbCancel Then
   '   Exit Sub
    '  End If
   ' End If
'Application.Visible = False
    'Varmistetaan ett� ollaan monen dokumentin tilassa
    SDocMode = Application.Preferences.System.SingleDocumentMode
    Application.Preferences.System.SingleDocumentMode = False
    'Tiedoston nimi on m��ritelty
    Formi.MousePointer = fmMousePointerHourGlass
    PFile = 1
    Ajastin = Timer
    
    
Set oBatTiedosto = fso.CreateTextFile(Application.Preferences.Files.TempFilePath & "tiedostot.bat", True)
Set oBatTiedosto = Nothing
    
    
    
    If Right(FPath.Value, 1) <> "\" Then FPath.Value = FPath.Value & "\"
    tiedosto = Dir(FPath.Value & FName.Value)
    
    
Dim YhdKansio As String
   If PltTied.Value = "" Then
   YhdKansio = "C:\temp\"
   Else
   If Right(PltTied.Value, 1) <> "\" Then PltTied.Value = PltTied.Value & "\"
   YhdKansio = PltTied.Value
   End If
   
   If PltTied.Value = "" And PltTiedCheck.Value = True Then
    PDFHakemisto = Application.Preferences.Files.TempFilePath
   ElseIf PltTied.Value <> "" And PltTiedCheck.Value = True Then
    PDFHakemisto = PltTied.Value
   End If

    
    
    Do While tiedosto <> ""
      Paate = UCase(Right(tiedosto, 4))
      If Plotter.Value = "Generic PostScript Printer" Then
       TiedostoNimi = (Left(tiedosto, Len(tiedosto) - Len(Paate))) & ".ps"
      Else
       TiedostoNimi = (Left(tiedosto, Len(tiedosto) - Len(Paate))) & ".plt"
      End If
      
      
      
      
      If Paate = ".DWG" Or Paate = ".DXF" Then
        If AvaaTiedosto(FPath.Value & tiedosto) = True Then 'Tiedoston avaus onnistui
          Tila tiedosto
          
If Plotter.Value = "Generic PostScript Printer" Then
'       If Palokortit.Value = False Then
'       TiedostoNimi = (Left(TiedostoNimi, Len(TiedostoNimi) - 4)) & ".ps"
Set oBatTiedosto = fso.CreateTextFile(Application.Preferences.Files.TempFilePath & "tiedostot.bat", True)
       oBatTiedosto.WriteLine "TYPE " & Chr(34) & Application.Preferences.Files.TempFilePath & TiedostoNimi & Chr(34) & Chr(32) & Chr(62) & Chr(62) & Chr(32) & Chr(34) & Application.Preferences.Files.TempFilePath & "Valmis.ps" & Chr(34)
       oBatTiedosto.WriteLine "DEL " & Chr(34) & Application.Preferences.Files.TempFilePath & TiedostoNimi & Chr(34)
      Set oBatTiedosto = Nothing
End If
          
          TulostaDOC TiedostoNimi, True
          PFile = PFile + 1
          If Kaikki = False Then  'Lopetetaan t�h�n jos on kysymys testitulostuksesta
            Exit Do
          End If
        End If
      
      ElseIf Paate = ".LST" Then

'        LueTiedosto FPath.Value & Tiedosto, Kaikki
        TiedostoNimi = LueTiedosto(FPath.Value & tiedosto, Kaikki)
        If Kaikki = False Then  'Lopetetaan t�h�n jos on kysymys testitulostuksesta
          Exit Do
        End If
      End If
      
      tiedosto = Dir
    Loop
    
    
    'Suljetaan viimeinen dokumentti, jotta se ei j�� vahingossa auki
    For i = 0 To Application.Documents.Count - 1
      Application.Documents(0).Close False
    Next i
    Application.Documents.Add 'Tehd��n uusi dokumentti jotta voidaan siirty� single tilaan. Single tilaan ei voi siirty�, jollei yht��n dokumenttia ole auki
    Application.Preferences.System.SingleDocumentMode = SDocMode 'Siirryt��n tilaan, joka oli ennen ohjelman k�ynnistyst�
    
    TilaTieto.Caption = ""
    MsgBox "Drawing(s) Plotted", , "Ready"
    Formi.MousePointer = fmMousePointerDefault
'    Application.Visible = True

End Sub
Public Function LueTiedosto(LstTiedosto As String, Kaikki As Boolean) As String
Dim oTiedosto As Scripting.TextStream
Dim Rivi As String
Dim i As Integer
Dim Paate As String
Dim Alku As String
Dim tiedosto As String
Dim TiedostoArray As String
Dim nro As Integer
Dim TiedostoNimi As String

Set oBatTiedosto = fso.OpenTextFile(Application.Preferences.Files.TempFilePath & "tiedostot.bat", ForWriting)

    Set oTiedosto = fso.OpenTextFile(LstTiedosto, 1)

      Rivi = oTiedosto.ReadLine
      tiedosto = FPath.Value & Rivi
        End If
        nro = 0
    If Plotter.Value = "Generic PostScript Printer" Then
'       If Palokortit.Value = False Then
       TiedostoNimi = (Left(TiedostoNimi, Len(TiedostoNimi) - 4)) & ".ps"
       oBatTiedosto.WriteLine "TYPE " & Chr(34) & Application.Preferences.Files.TempFilePath & TiedostoNimi & Chr(34) & Chr(32) & Chr(62) & Chr(62) & Chr(32) & Chr(34) & Application.Preferences.Files.TempFilePath & "Valmis.ps" & Chr(34)
       oBatTiedosto.WriteLine "DEL " & Chr(34) & Application.Preferences.Files.TempFilePath & TiedostoNimi & Chr(34)
      
      Else
       TiedostoNimi = (Left(TiedostoNimi, Len(TiedostoNimi) - 4)) & ".plt"

        If Right(UCase(tiedosto), 4) <> ".DWG" Then tiedosto = tiedosto & ".DWG"
      End If
      
      
        If AvaaTiedosto(tiedosto) Then
            Tila tiedosto

            'TulostaDOC TiedostoNimi, True
            
            PFile = PFile + 1
            If Kaikki = False Then  'Lopetetaan t�h�n jos on kysymys testitulostuksesta
              Exit Do
            End If
        End If
      End If
  
LueTiedosto = TiedostoNimi

End Function

Private Function AvaaTiedosto(Avattava As String) As Boolean
'Avaa tiedoston AutoCADiss�.
Dim i As Integer
    If fso.FileExists(Avattava) Then
      For i = 0 To Application.Documents.Count - 1
        Application.Documents(0).Close False   'Suljetaan documentit (Multi Document tilassa AutoCAD 2002 ei herjaa vaikkei documentteja olisikaan ja vaikka documentteja ei olisi talletettu)
      Next i
      Application.Documents.Open (Avattava)  'Avataan uusi dokumentti AutoCadiin (ACAD 2002 osaa avata my�s DXF documentit ilman kyselyj�)
      AvaaTiedosto = True
    Else
      AvaaTiedosto = False
      MsgBox "File not found: " & Avattava, vbCritical, "Error!"
    End If
End Function
Private Sub UserForm_Terminate()
  KirjoitaTiedot
  Set fso = Nothing
End Sub
Private Sub Valitse_Click()
'T�m� valitsee tiedoston hakemistosta
    Dim OpenFile As OPENFILENAME
    Dim lReturn As Long
    Dim Filtteri As String
    Dim AHakem As String
    Dim Otsikko As String
    Dim WHandle As LongPtr
    
    WHandle = FindWindow(0&, "Plot Utility")
    If InStr(FPath.Value, "\") Then
      AHakem = FPath.Text
    Else
      AHakem = "I:\"
    End If
    
    Filtteri = "AutoCAD Drawing (*.dwg)" & Chr(0) & "*.DWG" & Chr(0) _
             & "AutoCAD DXF (*.dxf)" & Chr(0) & "*.DXF" & Chr(0) _
             & "List Of Drawings (*.lst)" & Chr(0) & "*.LST" & Chr(0) _
             & "All Types (*.dwg;*.dxf;*.lst)" & Chr(0) & "*.DWG;*.LST;*.DXF" & Chr(0)
    
    'Tiedostop��tteen mukaan s��det��n filtteri oikeanlaiseksi
    If FName.Value = "*.dwg" Then
    FiltteriIndeksi = 1
    ElseIf FName.Value = "*.dxf" Then
    FiltteriIndeksi = 2
    ElseIf FName.Value = "*.lst" Then
    FiltteriIndeksi = 3
    Else
    FiltteriIndeksi = 4
    End If
    
    'Otsikko = "Choose AutoCAD Drawing or List"
    Otsikko = "Valitse AutoCAD piirustus tai lista (*.dwg, *.dxf tai *.lst)"
    With OpenFile
      .lStructSize = Len(OpenFile)
      .hWndOwner = WHandle
      .hInstance = 0
      .lpstrFilter = Filtteri
      .nFilterIndex = FiltteriIndeksi
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
        'Painettiin Cancel painiketta. Ei tehd� mit��n
    Else 'Otetaan yl�s Tiedostonimi ja Hakemisto
       FName.Value = Mid(OpenFile.lpstrFileTitle, 1, InStr(OpenFile.lpstrFileTitle, Chr(0)) - 1)
       FPath.Value = Mid(OpenFile.lpstrFile, 1, InStr(OpenFile.lpstrFile, Chr(0)) - 1 - Len(FName.Value))
       Nayta_lista.Enabled = True
    End If
End Sub
Private Sub Valitse_H_Click()
'    MsgBox BrowseForDirectory
    FPath.Value = BrowseForDirectory
End Sub

' Let the user browse for a directory. Return the
' selected directory. Return an empty string if
' the user cancels.
Private Function BrowseForDirectory() As String
Dim browse_info As BrowseInfo
Dim item As LongPtr
Dim dir_name As String

'OletusHakemisto
    'If InStr(SHakemisto.Value, "\") Then
    ' SHakem = SHakemisto.Value
    'Else
    '  SHakem = "I:\"
    'End If


   'modified for MS Access/VBA
   With browse_info
'       .hWndOwner = Application.hWndAccessApp
'       .hWndOwner = Me.Handle
       .pidlRoot = 0
       .sDisplayName = Space$(260)
       .sTitle = "Valitse hakemisto"
       .ulFlags = 1 ' Return directory name.
       .lpfn = 0
       .lParam = 0
       .iImage = 0
   End With

   item = SHBrowseForFolder(browse_info)
   If item Then
       dir_name = Space$(260)
       If SHGetPathFromIDList(item, dir_name) Then
           BrowseForDirectory = Left(dir_name, _
                                InStr(dir_name, Chr$(0)) - 1)
       Else
           BrowseForDirectory = "Hakemistoa ei ole valittu"
       End If
    Else
       BrowseForDirectory = "Hakemistoa ei ole valittu"
   End If
End Function
Private Sub Tila(tiedosto As String)
  sec = CInt(Timer - Ajastin)
  Min = sec \ 60
  sec = sec Mod 60
  TilaTieto.Caption = "Printing File: " & PFile & " (" & tiedosto & ")  " & blockno & "  Time: " & Min & " min  " & sec & " sec"
  Formi.Repaint
End Sub

