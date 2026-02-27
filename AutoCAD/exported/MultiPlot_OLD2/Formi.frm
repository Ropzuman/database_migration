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
Dim FSO As FileSystemObject
Dim Virhe As Boolean
Dim OliAuki As Boolean
Dim Tiedostot() As String

Private Sub Settings_Click()
  Asetus.Show
End Sub

Private Sub TeePDF_Click()
Dim Vastaus As VbMsgBoxResult
  Vastaus = MsgBox("Do you want to join files to one PDF?", vbYesNoCancel, "Plot PDF")
  If Vastaus = vbYes Then 'Yhdistetään dokumentit yhdeksi PDF tiedostoksi
    PPath.Value = "L:\tilapain\pdf\temp\"
    TulostaPDF True
  ElseIf Vastaus = vbNo Then 'Tulostetaan kaikki erikseen
    PPath.Value = "L:\tilapain\pdf\in\"
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
  SPaikka = "BL"
  SSuunta = "Vertical"
  SKork = 5
  SYlosalais = False
  SXOffset = 2.54
  SYOffset = 2.54
  SFName = True
  SDate = True
  
  'Laitetaan pysty ja vaaka vaihtoehdot
  Set FSO = New FileSystemObject
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
    Pens.AddItem Tyylit(i)
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
            If Tiedot(15) = 1 Then TeeMerkki.Value = True
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
  oTiedosto.Write IIf(TeeMerkki.Value = True, 1, 0) & vbTab
  oTiedosto.Write IIf(SFName = True, 1, 0) & vbTab
  oTiedosto.Write IIf(SDate = True, 1, 0) & vbCrLf
  oTiedosto.Close
  Set oTiedosto = Nothing
End Sub
Private Sub Plotter_Change()
Dim Koot As Variant
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
End Sub
Private Sub Tulosta_Click()
    Kay_Lapi_Tiedostot True 'Tulostaa kaikki dokumentit
End Sub
Private Sub TulostaTest_Click()
    Kay_Lapi_Tiedostot False 'Tulostaa yhden testiksi
End Sub
Private Sub TulostaNykyinen_Click()
    TulostaDOC
End Sub
Private Sub Kay_Lapi_Tiedostot(Kaikki As Boolean)
Dim Paate As String
Dim Tiedosto As String
Dim i As Integer
Dim oDOC As AcadDocument
    ReDim Tiedostot(0)
    Virhe = False
    Ajastin = Timer
    If FPath.Value = "" Or FName.Value = "" Then
      MsgBox "Please choose path and file(s) first.", vbCritical, "Plot"
      Exit Sub
    End If
    'Varmistetaan että ollaan monen dokumentin tilassa
    SDocMode = Application.Preferences.System.SingleDocumentMode
    Application.Preferences.System.SingleDocumentMode = False
    'Tiedoston nimi on määritelty
    Formi.MousePointer = fmMousePointerHourGlass
    PFile = 1
    Ajastin = Timer
    If Right(FPath.Value, 1) <> "\" Then FPath.Value = FPath.Value & "\"
    Tiedosto = Dir(FPath.Value & FName.Value)
    Do While Tiedosto <> ""
      Paate = UCase(Right(Tiedosto, 4))
      If Paate = ".DWG" Or Paate = ".DXF" Then
        Set oDOC = AvaaTiedosto(FPath.Value & Tiedosto)
        If Not oDOC Is Nothing Then 'Tiedoston avaus onnistui
          Tila Tiedosto
          TulostaDOC
          If Virhe Then Exit Do
          PFile = PFile + 1
          If OliAuki = False Then oDOC.Close False
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
'    For i = 0 To Application.Documents.Count - 1
'      Application.Documents(0).Close False
'    Next i
'    Application.Documents.Add 'Tehdään uusi dokumentti jotta voidaan siirtyä single tilaan. Single tilaan ei voi siirtyä, jollei yhtään dokumenttia ole auki
    Application.Preferences.System.SingleDocumentMode = SDocMode 'Siirrytään tilaan, joka oli ennen ohjelman käynnistystä
    TilaTieto.Caption = "Print completed"
'    MsgBox "Drawing(s) Plotted", , "Ready"
    Formi.MousePointer = fmMousePointerDefault
End Sub
Private Sub LueTiedosto(Tiedosto As String, Kaikki As Boolean)
Dim oTiedosto As Scripting.TextStream
Dim Rivi As String
Dim i As Integer
Dim Alku As String
Dim oDOC As AcadDocument
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
        If InStr(Rivi, "\") Then
          Tiedosto = Rivi
        ElseIf Len(Rivi) > 38 Then
          Tiedosto = FPath.Value & Mid(Rivi, 38)
        Else
          Tiedosto = FPath.Value & Rivi
        End If
        If Right(UCase(Tiedosto), 4) <> ".DWG" Then Tiedosto = Tiedosto & ".DWG"
        Set oDOC = AvaaTiedosto(Tiedosto)
        If Not oDOC Is Nothing Then
            Tila Tiedosto
            TulostaDOC
            PFile = PFile + 1
            If OliAuki = False Then oDOC.Close False
            If Kaikki = False Then  'Lopetetaan tähän jos on kysymys testitulostuksesta
              Exit Do
            End If
        End If
      End If
    Loop
    Set oTiedosto = Nothing
End Sub
Private Sub TulostaDOC()
Dim Tulostettava As Object
Dim Origo(0 To 1) As Double
Dim Layout(1) As String
Dim Komento As String
   ReDim Preserve Tiedostot(UBound(Tiedostot) + 1)
   Tiedostot(UBound(Tiedostot)) = Left(ActiveDocument.Name, Len(ActiveDocument.Name) - 4)
   If TeeMerkki Then
     Komento = "-plotstamp on "
     Komento = Komento & "f " & IIf(SFName, "y ", "n ") & "n " & IIf(SDate, "y ", "n ") & "n n n n "
     Komento = Komento & "log n " & vbCr
     Komento = Komento & "loc " & SPaikka & " " & SSuunta & " " & IIf(SYlosalais, "y ", "n ")
     Komento = Komento & Replace(CStr(SXOffset), ",", ".") & "," & Replace(CStr(SYOffset), ",", ".") & " a "
     Komento = Komento & "t Arial" & vbCr & Replace(CStr(SKork), ",", ".") & " y "
     Komento = Komento & "un m " & vbCr
   Else
     Komento = "-plotstamp off " & vbCr
   End If
   ActiveDocument.SendCommand Komento
   Origo(0) = 0
   Origo(1) = 0
   If PMSPACE.Value Then
     Set Tulostettava = Application.ActiveDocument.ModelSpace
   Else
     Set Tulostettava = Application.ActiveDocument.PaperSpace
   End If

   On Error Resume Next
   With Tulostettava.Layout
     Layout(0) = .Name
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
   Err.Clear
   On Error GoTo 0
   ActiveDocument.Regen acActiveViewport
   With ActiveDocument.Plot
     .QuietErrorMode = True
     .SetLayoutsToPlot Layout
     .NumberOfCopies = Copies.Value
     If PTFile.Value = True Then
       If PTSourcePath.Value = True Then
         .PlotToFile ActiveDocument.Path & "\" & Left(ActiveDocument.Name, Len(ActiveDocument.Name) - 4) & ".plt"
       ElseIf FSO.FolderExists(PPath.Value) Then
         If Right(PPath.Value, 1) <> "\" Then PPath.Value = PPath.Value & "\"
         .PlotToFile PPath.Value & Left(ActiveDocument.Name, Len(ActiveDocument.Name) - 4) & ".plt"
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
          Set AvaaTiedosto = Application.Documents(i)
          Exit Function
        End If
      Next i
      Set AvaaTiedosto = Application.Documents.Open(Avattava, True)  'Avataan uusi dokumentti AutoCadiin (ACAD 2002 osaa avata myös DXF documentit ilman kyselyjä)
    Else
      AvaaTiedosto = Nothing
      MsgBox "File not found: " & Avattava, vbCritical, "Error!"
    End If
End Function
Private Sub UserForm_Terminate()
Dim Asema As Scripting.Drive
Dim Nimi As String
  KirjoitaTiedot
  For Each Asema In FSO.Drives
    If Asema.IsReady Then
      If Asema.ShareName <> "" Then
        If Left(VBE.activeVBproject.Filename, Len(Asema.ShareName)) = Asema.ShareName Then
          Nimi = Asema.Path & Mid(VBE.activeVBproject.Filename, Len(Asema.ShareName) + 1)
          Exit For
        End If
      End If
    End If
  Next
  Set Drive = Nothing
  Set FSO = Nothing
  UnloadDVB Nimi
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
      AHakem = "P:\PROJEKTI\"
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
Private Sub Tila2(Tiedosto As String)
  sec = CInt(Timer - Ajastin)
  Min = sec \ 60
  sec = sec Mod 60
  TilaTieto.Caption = "Joining file: " & PFile & "/" & UBound(Tiedostot) & " (" & FSO.GetFileName(Tiedosto) & ")  Time: " & Min & " min  " & sec & " sec"
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
For i = 0 To Plotter.ListCount - 1
  If UCase(Plotter.List(i)) = "\\PORSCPRINT\PORPCPS1" Then
    Loytyi = True
    Exit For
  End If
Next i
If Loytyi Then
  If UCase(Plotter.Value) <> "\\PORSCPRINT\PORPCPS1" Then
    Plotter.Value = "\\PORSCPRINT\PORPCPS1"
    Size.Value = "User233"
  End If
  PTSourcePath = False
  PTFile.Value = True
  Kay_Lapi_Tiedostot True
  If Yhdista Then 'Yhdistetään tulostetut kuvat yhdeksi tiedostoksi
    Valmis = Replace(Replace(Replace(FPath, ":", ""), "\", "_"), " ", "")
    Muunna = Valmis & "PDFYhdista.bat"
    Valmis = Valmis & "ACadPDF"
    Set oMuunnos = FSO.CreateTextFile("L:\TILAPAIN\PDF\Temp\" & Muunna, True)
    For i = 1 To UBound(Tiedostot)
      PFile = i
      Tila2 Tiedostot(i)
      oMuunnos.WriteLine "TYPE ""L:\TILAPAIN\PDF\TEMP\" & Tiedostot(i) & ".plt"" >> ""L:\TILAPAIN\PDF\TEMP\" & Valmis & ".ps"""
      oMuunnos.WriteLine "DEL ""L:\TILAPAIN\PDF\TEMP\" & Tiedostot(i) & ".plt"""
    Next i
    oMuunnos.WriteLine "MOVE ""L:\TILAPAIN\PDF\TEMP\" & Valmis & ".ps"" ""L:\TILAPAIN\PDF\IN\" & Valmis & ".ps"""
    oMuunnos.WriteLine "DEL ""L:\TILAPAIN\PDF\TEMP\" & Muunna
    oMuunnos.Close
    Set oMuunnos = Nothing
    Do While Dir("L:\Tilapain\PDF\Temp\" & Tiedostot(UBound(Tiedostot)) & ".plt") = ""
      'Kierrätetään tätä kunnes tiedosto on ilmestynyt kansioon
    Loop
    Shell "L:\Tilapain\PDF\Temp\" & Muunna, vbHide
    MsgBox "Files joined." & vbCrLf & "You can find it from:" & vbCrLf & "L:\TILAPAIN\PDF\Out\" & Valmis & ".PDF", vbOKOnly, "Plot PDF"
  End If
Else
  MsgBox "Printer PORPCPS1 not found. You must install printer before continuing.", vbCritical, "Plot PDF"
End If
End Sub

