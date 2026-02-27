VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Formi 
   Caption         =   "Plot Utility"
   ClientHeight    =   6120
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
Dim Virhe As Boolean
Dim OliAuki As Boolean
Dim TempPolku As String
Dim TulostusNimi As String
Dim oDOC As AcadDocument





Private Sub BLst_Click()
Dim P As Double
Dim FS As New FileSystemObject
Dim TS As TextStream
Dim Tied As String
If FPath.Value = "" Then
  MsgBox "Choose path first", vbCritical, "Make lst file"
Else
  If Right(FPath.Value, 1) <> "\" Then FPath.Value = FPath.Value & "\"
  Tied = Dir(FPath & FName)
  If Dir(FPath & FName) = "" Then
    MsgBox "No files found with filter: " & FName.Value, vbCritical, "Make lst file"
  Else
    Set TS = FS.CreateTextFile(FPath.Value & "tied.lst", True)
    Do While Tied <> ""
      TS.WriteLine FPath.Value & Tied
      Tied = Dir
    Loop
    TS.Close
    If MsgBox("Do you want to put listname to filename field?", vbYesNo, "Make lst file") = vbYes Then
      FName.Value = "tied.lst"
    End If
    
    P = Shell("NOTEPAD """ & FPath.Value & "tied.lst""", vbNormalFocus)
  End If
End If
End Sub
Private Sub TeePDF_Click()
Dim Vastaus As VbMsgBoxResult
  Vastaus = MsgBox("Do you want to join files to one PDF?", vbYesNoCancel, "Plot PDF")
    If PPath.Value = "" Then
        PPath.Value = FPath.Value
    End If
  If Vastaus = vbYes Then 'Yhdistetään dokumentit yhdeksi PDF tiedostoksi
    TulostaPDF True
  ElseIf Vastaus = vbNo Then 'Tulostetaan kaikki erikseen
    TulostaPDF False
  End If
End Sub


Private Sub Ulos_Click()
  Unload Formi
End Sub
Private Sub UserForm_Initialize()
'Avaus
Dim i As Integer
Dim Tulostimet As Variant
Dim Skalet(1 To 17, 1 To 2) As Variant
Dim Tyylit As Variant
Dim Koot As Variant
Dim oTiedosto As Scripting.TextStream
Dim Tiedot As Variant
  'Oletusarvot
  TempPolku = FSO.GetSpecialFolder(TemporaryFolder).Path & "\" '  "L:\tilapain\pdf\temp\"
  SPaikka = "BL"
  SSuunta = "Vertical"
  SKork = 5
  SYlosalais = False
  SXOffset = 2.54
  SYOffset = 2.54
  SFName = True
  SDate = True
   
  ThisDrawing.SetVariable "BACKGROUNDPLOT", 0

  'Laitetaan pysty ja vaaka vaihtoehdot
  Orientation.AddItem "Portrait (Pysty)"
  Orientation.AddItem "Landscape (Vaaka)"
  Orientation.AddItem "Portrait & Upside-down"
  Orientation.AddItem "Landscape & Upside-down"
  Orientation.ListIndex = 1
  'Laitetaan Skaalausvaihtoehdot
  Skalet(1, 2) = "Scale to Fit": Skalet(1, 1) = acScaleToFit
  Skalet(2, 2) = "1:1": Skalet(2, 1) = ac1_1
  Skalet(3, 2) = "1:2": Skalet(3, 1) = ac1_2
  Skalet(4, 2) = "1:4": Skalet(4, 1) = ac1_4
  Skalet(5, 2) = "1:8": Skalet(5, 1) = ac1_8
  Skalet(6, 2) = "1:10": Skalet(6, 1) = ac1_10
  Skalet(7, 2) = "1:16": Skalet(7, 1) = ac1_16
  Skalet(8, 2) = "1:20": Skalet(8, 1) = ac1_20
  Skalet(9, 2) = "1:30": Skalet(9, 1) = ac1_30
  Skalet(10, 2) = "1:40": Skalet(10, 1) = ac1_40
  Skalet(11, 2) = "1:50": Skalet(11, 1) = ac1_50
  Skalet(12, 2) = "1:100": Skalet(12, 1) = ac1_100
  Skalet(13, 2) = "2:1": Skalet(13, 1) = ac2_1
  Skalet(14, 2) = "4:1": Skalet(14, 1) = ac4_1
  Skalet(15, 2) = "8:1": Skalet(15, 1) = ac8_1
  Skalet(16, 2) = "10:1": Skalet(16, 1) = ac10_1
  Skalet(17, 2) = "100:1": Skalet(17, 1) = ac100_1
  Skaalaus.List = Skalet
  Skaalaus.ListIndex = 0
  
  'Laitetaan kopioita vaihtoehdot
  For i = 1 To 30
    Copies.AddItem i
  Next i
  Copies.ListIndex = 0
  Tulostimet = ActiveDocument.Layouts(0).GetPlotDeviceNames()
  For i = LBound(Tulostimet) To UBound(Tulostimet)
    Plotter.AddItem Tulostimet(i)
  Next i
  Plotter.ListIndex = 1
  Tyylit = ActiveDocument.Layouts(0).GetPlotStyleTableNames()
  For i = LBound(Tyylit) To UBound(Tyylit)
    If InStr(Tyylit(i), ".stb") = 0 Then
      Pens.AddItem Tyylit(i)
    End If
  Next i
  Pens.ListIndex = 0
  On Error Resume Next
  If FSO.FileExists(Application.Preferences.Files.TempFilePath & "ACADPlot.TXT") Then
    Set oTiedosto = FSO.OpenTextFile(Application.Preferences.Files.TempFilePath & "ACADPlot.TXT")
    Tiedot = Split(oTiedosto.ReadLine, vbTab)
    oTiedosto.Close
    FPath.Value = Tiedot(0)
    FName.Value = Tiedot(1)
    Plotter.Value = Tiedot(2)
    Orientation.Value = Tiedot(3)
    Copies.Value = Tiedot(4)
    Pens.Value = Tiedot(5)
    Size.Value = Tiedot(6)
    If UBound(Tiedot) > 6 Then
      PPath.Value = Tiedot(7)
      If UBound(Tiedot) > 7 Then
        If Tiedot(8) = 1 Then PTFile.Value = True
          If UBound(Tiedot) > 8 Then
            SPaikka = Tiedot(9)
            SSuunta = Tiedot(10)
            If Tiedot(11) = 1 Then SYlosalais = True
            SXOffset = CDbl(Tiedot(12))
            SYOffset = CDbl(Tiedot(13))
            SKork = CDbl(Tiedot(14))
            If Tiedot(15) = 1 Then PPSPACE.Value = True
            If Tiedot(15) = 2 Then PLAYOUTS.Value = True
            If Tiedot(16) = 1 Then SFName = True
            If Tiedot(17) = 1 Then SDate = True
          End If
      End If
    End If
  End If
  Err.Clear
  On Error GoTo 0
  Set oTiedosto = Nothing
End Sub
Private Sub KirjoitaTiedot()
Dim oTiedosto As Scripting.TextStream
  Set oTiedosto = FSO.CreateTextFile(Application.Preferences.Files.TempFilePath & "ACADPlot.TXT", True)
  oTiedosto.Write FPath.Value & vbTab
  oTiedosto.Write FName.Value & vbTab
  oTiedosto.Write Plotter.Value & vbTab
  oTiedosto.Write Orientation.Value & vbTab
  oTiedosto.Write Copies.Value & vbTab
  oTiedosto.Write Pens.Value & vbTab
  oTiedosto.Write Size.Value & vbTab
  oTiedosto.Write PPath.Value & vbTab
  oTiedosto.Write IIf(PTFile.Value = True, 1, 0) & vbTab
  oTiedosto.Write SPaikka & vbTab
  oTiedosto.Write SSuunta & vbTab
  oTiedosto.Write IIf(SYlosalais = True, 1, 0) & vbTab
  oTiedosto.Write SXOffset & vbTab
  oTiedosto.Write SYOffset & vbTab
  oTiedosto.Write SKork & vbTab
  oTiedosto.Write IIf(PPSPACE.Value = True, 1, IIf(PLAYOUTS.Value = True, 2, 0)) & vbTab
  oTiedosto.Write IIf(SFName = True, 1, 0) & vbTab
  oTiedosto.Write IIf(SDate = True, 1, 0) & vbCrLf
  oTiedosto.Close
  Set oTiedosto = Nothing
End Sub
Private Sub Plotter_Change()
Dim Koot As Variant
  On Error Resume Next
  ActiveDocument.Layouts(0).ConfigName = Plotter.Value
  Koot = ActiveDocument.Layouts(0).GetCanonicalMediaNames()
  For i = 0 To Size.ListCount - 1
    Size.RemoveItem 0
  Next i
  For i = LBound(Koot) To UBound(Koot)
    Size.AddItem Koot(i)
    Size.List(i, 1) = ActiveDocument.Layouts(0).GetLocaleMediaName(Koot(i))
  Next i
  If UBound(Koot) > 0 Then Size.ListIndex = 0
  Err.Clear
  On Error GoTo 0
End Sub
Private Sub Tulosta_Click()
  If KeraaTiedostot(True) Then
    TulostaDocs
  End If
End Sub
Private Sub TulostaTest_Click()
  If KeraaTiedostot(False) Then
    TulostaDocs
  End If
End Sub
Private Sub TulostaNykyinen_Click()
  Set oDOC = ActiveDocument
  TulostaDOC
End Sub
Private Sub TulostaDocs(Optional Temppiin As Boolean)
Dim i As Integer
    'Varmistetaan että ollaan monen dokumentin tilassa
    Ajastin = Timer
    Formi.MousePointer = fmMousePointerHourGlass
    For i = 1 To UBound(Tiedostot)
      PFile = i
      Tila Mid(Tiedostot(i), InStrRev(Tiedostot(i), "\") + 1)
      If Not oDOC Is Nothing Then
        If OliAuki = False Then oDOC.Close False
      End If
      Set oDOC = AvaaTiedosto(Tiedostot(i))
      TulostusNimi = Right("00" & i, 3)
      TulostaDOC Temppiin
    Next i
    If Not oDOC Is Nothing Then
      If OliAuki = False Then oDOC.Close False
    End If
    Set oDOC = Nothing
    TilaTieto.Caption = "Print completed"
    Formi.MousePointer = fmMousePointerDefault
End Sub
Private Sub TulostaDOC(Optional Temppiin As Boolean)
Dim Tulostettava As Object
Dim Origo(0 To 1) As Double
Dim Layout() As String
Dim Komento As String
Dim i As Integer, j As Integer
Dim Paate As String
'   If TeeMerkki Then
'     Komento = "-plotstamp on "
'     Komento = Komento & "f " & IIf(SFName, "y ", "n ") & "n " & IIf(SDate, "y ", "n ") & "n n n n "
'     Komento = Komento & "log n " & vbCr
'     Komento = Komento & "loc " & SPaikka & " " & SSuunta & " " & IIf(SYlosalais, "y ", "n ")
'     Komento = Komento & Replace(CStr(SXOffset), ",", ".") & "," & Replace(CStr(SYOffset), ",", ".") & " a "
'     Komento = Komento & "t Arial" & vbCr & Replace(CStr(SKork), ",", ".") & " y "
'     Komento = Komento & "un m " & vbCr
'   Else
'     Komento = "-plotstamp off " & vbCr
'   End If
'
'   oDOC.SendCommand Komento
   Origo(0) = 0
   Origo(1) = 0
   
   If PMSPACE.Value Then 'Model space
     Set Tulostettava = oDOC.ModelSpace
     ReDim Layout(1)
     Layout(0) = Tulostettava.Layout.Name
   ElseIf PPSPACE.Value Then 'Eka layout
      ReDim Layout(0)
      For i = 0 To oDOC.Layouts.Count - 1
         If InStr(oDOC.Layouts(i).Name, 1) Then
           Set Tulostettava = oDOC.Layouts(i)
           Layout(0) = oDOC.Layouts(i).Name
         End If
       Next i
   Else 'Kaikki layoutit
     Set Tulostettava = oDOC.Layouts(1)
     ReDim Layout(oDOC.Layouts.Count - 2)
     j = 0
     For i = 0 To oDOC.Layouts.Count - 1
       If LCase(oDOC.Layouts(i).Name) <> "model" Then
         Layout(j) = oDOC.Layouts(i).Name
         With oDOC.Layouts(i)
            .ConfigName = Plotter.Value
            .PlotType = acExtents
            .PlotOrigin = Origo
            .PlotRotation = Orientation.ListIndex  '0=ac0degrees (pysty), 1=ac90degrees (vaaka), 2=ac180degrees (pysty) upside-down, 3=ac270degrees (vaaka) upside-down
            .StandardScale = Skaalaus.Value  ' acScaleToFit
            .PlotWithPlotStyles = True
            .StyleSheet = Pens.Value
            .CanonicalMediaName = Size.Value
            .PaperUnits = acMillimeters
         End With
         j = j + 1
       End If
     Next i
   End If
    
   On Error Resume Next
   If PMSPACE.Value Then 'Tulostetaan model space
     With Tulostettava.Layout
       .ConfigName = Plotter.Value
       .PlotType = acExtents
       .PlotOrigin = Origo
       .StandardScale = Skaalaus.Value  ' acScaleToFit
       .PlotWithPlotStyles = True
       .StyleSheet = Pens.Value
       .CanonicalMediaName = Size.Value
       .PlotRotation = Orientation.ListIndex  '0=ac0degrees (pysty), 1=ac90degrees (vaaka), 2=ac180degrees (pysty) upside-down, 3=ac270degrees (vaaka) upside-down
       .PaperUnits = acMillimeters
     End With
   ElseIf PPSPACE.Value Then ' Tulostetaan eka layout
     With Tulostettava.Layout
       .ConfigName = Plotter.Value
       .PlotType = acExtents
       .PlotOrigin = Origo
  '     .CenterPlot = True
       .StandardScale = Skaalaus.Value  ' acScaleToFit
       .PlotWithPlotStyles = True
       .StyleSheet = Pens.Value
       .CanonicalMediaName = Size.Value
       .PlotRotation = Orientation.ListIndex  '0=ac0degrees (pysty), 1=ac90degrees (vaaka), 2=ac180degrees (pysty) upside-down, 3=ac270degrees (vaaka) upside-down
       .PaperUnits = acMillimeters
     End With
   Else 'tulostetaan kaikki Layoutit
   
   End If
   
   Err.Clear
   On Error GoTo 0
   oDOC.Regen acActiveViewport
   
   With oDOC.Plot
     .QuietErrorMode = True
     .SetLayoutsToPlot Layout
     .NumberOfCopies = Copies.Value
     If PTFile.Value = True Then
        If InStr(UCase(Plotter.Value), "PDF") Then
            Paate = ".pdf"
        Else
            Paate = ".plt"
        End If

        If Temppiin Then 'Jos välikansioon tulostus, tulostetetaan sinne ensisijaisesti
          If UBound(Layout) > 1 Then
            .PlotToFile TempPolku & "TULOSTUS\"
          Else
            .PlotToFile TempPolku & "TULOSTUS\" & TulostusNimi & Paate
          End If
        ElseIf PTSourcePath.Value = True Then
           .PlotToFile ActiveDocument.Path & "\" & Left(ActiveDocument.Name, Len(ActiveDocument.Name) - 4) & Paate
        ElseIf FSO.FolderExists(PPath.Value) Then
          If Right(PPath.Value, 1) <> "\" Then
              PPath.Value = PPath.Value & "\"
          End If
          .PlotToFile PPath.Value & Left(ActiveDocument.Name, Len(ActiveDocument.Name) - 4) & Paate
        Else
          Virhe = True
          MsgBox "Path for plot file must be chosen!", vbCritical, "Plot"
        End If
        
     Else
       .PlotToDevice
     End If
   End With
End Sub
Private Function AvaaTiedosto(Avattava As String) As AcadDocument
'Avaa tiedoston AutoCADissä.
Dim i As Integer
OliAuki = False
Nimi = UCase(Mid(Avattava, InStrRev(Avattava, "\") + 1))
    If FSO.FileExists(Avattava) Then
      For i = 0 To Application.Documents.Count - 1
        If UCase(Application.Documents(i).Name) = Nimi Then
          OliAuki = True
          Application.Documents(i).Activate
          Set AvaaTiedosto = Application.Documents(i)
          Exit Function
        End If
      Next i
      Set AvaaTiedosto = Application.Documents.Open(Avattava, True)  'Avataan uusi dokumentti AutoCadiin (ACAD 2002 osaa avata myös DXF documentit ilman kyselyjä)
    Else
      Set AvaaTiedosto = Nothing
      MsgBox "File not found: " & Avattava, vbCritical, "Error!"
    End If
End Function
Private Sub UserForm_Terminate()
'Dim Asema As Scripting.Drive
Dim Nimi As String
  KirjoitaTiedot
'  Exit Sub
'  For Each Asema In FSO.Drives
'    If Asema.IsReady Then
'      If Asema.ShareName <> "" Then
'        If LCase(Left(VBE.activeVBproject.Filename, Len(Asema.ShareName))) = LCase(Asema.ShareName) Then
'          Nimi = Asema.Path & Mid(VBE.activeVBproject.Filename, Len(Asema.ShareName) + 1)
'          Exit For
'        End If
'      End If
'    End If
'  Next
  Application.Visible = True
'  Set Drive = Nothing
  Set FSO = Nothing
  UnloadDVB VBE.activeVBproject.Filename 'Nimi
End Sub
Private Sub Valitse_Click()
'Tämä valitsee tiedoston hakemistosta
    Dim OpenFile As OPENFILENAME
    Dim lReturn As Long
    Dim Filtteri As String
    Dim AHakem As String
    Dim Otsikko As String
    Dim WHandle As Long
    
    WHandle = FindWindow(0&, "Plot Utility")
    If InStr(FPath.Value, "\") Then
      AHakem = FPath.Text
    Else
      AHakem = "K:\PROJECTS\"
    End If
    
    Filtteri = "AutoCAD Drawing (*.dwg)" & Chr(0) & "*.DWG" & Chr(0) _
             & "AutoCAD DXF (*.dxf)" & Chr(0) & "*.DXF" & Chr(0) _
             & "List Of Drawings (*.lst)" & Chr(0) & "*.LST" & Chr(0) _
             & "All Types (*.dwg;*.dxf;*.lst)" & Chr(0) & "*.DWG;*.LST;*.DXF" & Chr(0)
    Otsikko = "Choose AutoCAD Drawing or List"
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
       FName.Value = Mid(OpenFile.lpstrFileTitle, 1, InStr(OpenFile.lpstrFileTitle, Chr(0)) - 1)
       FPath.Value = Mid(OpenFile.lpstrFile, 1, InStr(OpenFile.lpstrFile, Chr(0)) - 1 - Len(FName.Value))
    End If
End Sub
Private Sub Tila(Tiedosto As String)
  sec = CInt(Timer - Ajastin)
  Min = sec \ 60
  sec = sec Mod 60
  TilaTieto.Caption = "Printing File: " & PFile & " (" & FSO.GetFileName(Tiedosto) & ")  Time: " & Min & " min  " & sec & " sec"
  Formi.Repaint
End Sub
Private Sub ValitseHakemisto_Click()
Dim Hakemisto As String
  Hakemisto = ValitseHakem(0, PPath.Value)
  If Hakemisto <> "" Then
    PPath.Value = Hakemisto
  End If
End Sub
Private Sub TulostaPDF(Yhdista As Boolean)
Dim i As Integer
Dim Loytyi As Boolean
Dim Valmis As String
Dim Muunna As String
Dim oMuunnos As TextStream
Dim OK As Long
Dim Tied1 As String
Dim Tied2 As String
Dim Polku As String
Dim Plotteri As String
Dim oFile As File
For i = 0 To Plotter.ListCount - 1
  If InStr(UCase(Plotter.List(i)), "DWG TO PDF") Then
    Plotteri = Plotter.List(i)
    Loytyi = True
    Exit For
  End If
Next i
If Loytyi Then
  If UCase(Plotter.Value) <> UCase(Plotteri) Then
    Plotter.Value = Plotteri
    Size.Value = "ISO_full_bleed_A4_(210.00_x_297.00_MM)"
  End If
  PTSourcePath.Value = False
  PTFile.Value = True
  If KeraaTiedostot(True) Then
    If FSO.FolderExists(TempPolku & "TULOSTUS\") Then
        FSO.DeleteFolder TempPolku & "TULOSTUS"
    End If
    
    FSO.CreateFolder TempPolku & "TULOSTUS\"
    
    TulostaDocs Yhdista
    If Yhdista Then 'Yhdistetään tulostetut kuvat yhdeksi tiedostoksi
      Valmis = Mid(Tiedostot(1), InStrRev(Tiedostot(1), "\") + 1)
      Valmis = Left(Valmis, Len(Valmis) - 4) & ".pdf"
      
      Muunna = "Yhdista.bat"
      Set oMuunnos = FSO.CreateTextFile(TempPolku & "TULOSTUS\" & Muunna, True, True)
      oMuunnos.WriteLine "@echo off"
      oMuunnos.WriteLine "echo Joining files..."
      oMuunnos.WriteLine "chcp 1252"
      oMuunnos.WriteLine "set PATH=W:\System\PDFTK\bin" 'PDF työkalun sijainti
      oMuunnos.WriteLine Left(TempPolku, 2)
      oMuunnos.WriteLine "CD " & TempPolku & "TULOSTUS"
      oMuunnos.WriteLine "pdftk *.pdf cat output vali.pdf"
      oMuunnos.WriteLine "MOVE vali.pdf """ & PPath.Value & Valmis & """"
      oMuunnos.WriteLine "CD .."
      oMuunnos.WriteLine "RMDIR TULOSTUS /S /Q"
      oMuunnos.Close
      Set oMuunnos = Nothing
      'Aloitetaan muunnos
      Shell TempPolku & "TULOSTUS\" & Muunna, vbNormalFocus
      'Kerrotaan käyttäjälle että muunnos on aloitettu
      MsgBox "Files are being joined." & vbCrLf & "When finished, completed file can be found from:" & vbCrLf & PPath.Value & Valmis, vbOKOnly, "Plot PDF"
      Call Shell("EXPLORER " & PPath.Value, vbNormalFocus)
    End If
  Else

  End If
Else
  MsgBox "Printer DWG to PDF not found. You must install printer before continue.", vbCritical, "Plot PDF"
End If
End Sub

