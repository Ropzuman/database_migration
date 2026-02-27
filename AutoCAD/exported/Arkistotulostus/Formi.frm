VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Formi 
   Caption         =   "Arkistojonotulostus"
   ClientHeight    =   3780
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
Dim FSO As FileSystemObject
Dim Virhe As Boolean
Private Sub Ulos_Click()
  Unload Formi
End Sub
Private Sub UserForm_Initialize()
'Avaus
Dim i As Integer
Dim Tulostimet As Variant
Dim Tyylit As Variant
Dim oTiedosto As Scripting.TextStream
Dim Loytyi As Boolean
Dim Tiedot As Variant
  
  Set FSO = New FileSystemObject
  
  Tulostimet = ActiveDocument.Layouts(0).GetPlotDeviceNames()
  For i = LBound(Tulostimet) To UBound(Tulostimet)
    If LCase(Tulostimet(i)) = "arkiston hpgl2.pc3" Then
      Loytyi = True
      Exit For
    End If
  Next i
  If Loytyi = False Then
    MsgBox "Tavittavaa tulostinajuria ei löytynyt (Arkiston hpgl2.pc3)", vbCritical, "Arkistotulostus"
    Unload Me
    Exit Sub
  End If
  Loytyi = False
  
  Tyylit = ActiveDocument.Layouts(0).GetPlotStyleTableNames()
  For i = LBound(Tyylit) To UBound(Tyylit)
    If LCase(Tyylit(i)) = "genius.ctb" Then
      Loytyi = True
      Exit For
    End If
  Next i
  If Loytyi = False Then
    MsgBox "Tavittavaa kynäasetuksia ei löytynyt (Genius.ctb)", vbCritical, "Arkistotulostus"
    Unload Me
    Exit Sub
  End If
  
  On Error Resume Next
  If FSO.FileExists(Application.Preferences.Files.TempFilePath & "ArkistoPlot.TXT") Then
    Set oTiedosto = FSO.OpenTextFile(Application.Preferences.Files.TempFilePath & "ArkistoPlot.TXT")
    Tiedot = Split(oTiedosto.ReadLine, vbTab)
    oTiedosto.Close
    FPath.Value = Tiedot(0)
    FName.Value = Tiedot(1)
  End If
  Err.Clear
  On Error GoTo 0
  Set oTiedosto = Nothing
End Sub
Private Sub KirjoitaTiedot()
Dim oTiedosto As Scripting.TextStream
  Set oTiedosto = FSO.CreateTextFile(Application.Preferences.Files.TempFilePath & "ArkistoPlot.TXT", True)
  oTiedosto.Write FPath.Value & vbTab
  oTiedosto.Write FName.Value & vbCrLf
  oTiedosto.Close
  Set oTiedosto = Nothing
End Sub
Private Sub Tulosta_Click()
  Kay_Lapi_Tiedostot True 'Tulostaa kaikki dokumentit
End Sub
Private Sub Kay_Lapi_Tiedostot(Kaikki As Boolean)
Dim Paate As String
Dim Tiedosto As String
Dim i As Integer
    Virhe = False
    Ajastin = Timer
    If FPath.Value = "" Or FName.Value = "" Then
      MsgBox "Valitse polku ja tiedostot ensin.", vbCritical, "Tulosta"
      Exit Sub
    End If
    If MsgBox("Ohjelma sulkee kaikki avoimet kuvat ja aloittaa tulostuksen." & vbCrLf & "Haluatko sulkea kaikki kuvat?", vbOKCancel, "Tulosta") = vbCancel Then
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
        If AvaaTiedosto(FPath.Value & Tiedosto) = True Then 'Tiedoston avaus onnistui
          Tila Tiedosto
          TulostaDOC
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
    Application.Documents.Add 'Tehdään uusi dokumentti jotta voidaan siirtyä single tilaan. Single tilaan ei voi siirtyä, jollei yhtään dokumenttia ole auki
    Application.Preferences.System.SingleDocumentMode = SDocMode 'Siirrytään tilaan, joka oli ennen ohjelman käynnistystä
    Tilatieto.Caption = ""
    MsgBox "Drawing(s) Plotted", , "Ready"
    Formi.MousePointer = fmMousePointerDefault
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
        If InStr(Rivi, "\") Then
          Tiedosto = Rivi
        ElseIf Len(Rivi) > 38 Then
          Tiedosto = FPath.Value & Mid(Rivi, 38)
        Else
          Tiedosto = FPath.Value & Rivi
        End If
        If Right(UCase(Tiedosto), 4) <> ".DWG" Then Tiedosto = Tiedosto & ".DWG"
        If AvaaTiedosto(Tiedosto) Then
            Tila Tiedosto
            TulostaDOC
            PFile = PFile + 1
            If Kaikki = False Then  'Lopetetaan tähän jos on kysymys testitulostuksesta
              Exit Do
            End If
        End If
      End If
    Loop
    Set oTiedosto = Nothing
End Sub
Private Sub TulostaDOC()
Dim Origo(0 To 1) As Double
Dim Layout(0) As String
'Dim Skale As AcPlotScale
Dim TSkale As String
Dim Koko As String
Dim Nimi As String
Dim NimiA As String
Dim Rev As String
Dim Lehti As String
Dim PlotRotation As AcPlotRotation
Dim SkaleA As Double
Dim SkaleB As Double
   
'  Haetaan ensin kuvasta tiedot jolla kuva tulostetaan
Dim Joukko As AcadSelectionSet  'Joukko, jolla valitaan kaikki halutut blokit
Dim oBlock As AcadBlockReference
Dim BlockArray As Variant       'Array muuttuja Blokkia varten
Dim FilterType(0) As Integer
Dim FilterData(0) As Variant
  
  FilterType(0) = 2       'ObjecName
  FilterData(0) = "TITLEBLOCK_METSOP_2002"
  
  Set Joukko = ActiveDocument.ActiveSelectionSet
  Joukko.Select acSelectionSetAll, , , FilterType, FilterData
  If Joukko.Count > 0 Then
    Set oBlock = Joukko(0)
    BlockArray = oBlock.GetAttributes
    For i = LBound(BlockArray) To UBound(BlockArray)
      Select Case UCase(BlockArray(i).TagString)
        Case "ITEM_DOC_NO"
          NimiA = BlockArray(i).TextString
        Case "REV"
          If BlockArray(i).TextString = "-" Or BlockArray(i).TextString = "" Then
            Rev = "0"
          Else
            If Len(BlockArray(i).TextString) > 2 Then
               Rev = UCase(Left(Replace(BlockArray(i).TextString, ".", ""), 2))
            Else
               Rev = UCase(Replace(BlockArray(i).TextString, ".", ""))
            End If
          End If
        Case "SHEET"
          Lehti = Val(BlockArray(i).TextString)
        Case "GEN-TITLE-SCA"
          TSkale = BlockArray(i).TextString
      End Select
    Next i
  End If
  Set Joukko = Nothing
  If NimiA = "" Or Lehti = "" Or TSkale = "" Then
    Virheet.Text = Virheet.Text & "Kuvassa " & ActiveDocument.Name & " ei ollut tarpeellista otsikkotaulua." & vbCrLf
    Exit Sub
  End If
  Select Case Mid(NimiA, 4, 1)
    Case "0"
      Koko = "ISO_A0_(841.00_x_1189.00_MM)"
      PlotRotation = ac270degrees
    Case "1"
      'Koko = "ISO_A1_(841.00_x_594.00_MM)"
      Koko = "ISO_A1_(594.00_x_841.00_MM)"
      PlotRotation = ac270degrees
    Case "2"
      'Koko = "ISO_A2_(594.00_x_420.00_MM)"
      Koko = "ISO_A2_(420.00_x_594.00_MM)"
      PlotRotation = ac270degrees
    Case "3"
      'Koko = "ISO_A3_(420.00_x_297.00_MM)"
      Koko = "ISO_A3_(297.00_x_420.00_MM)"
      PlotRotation = ac270degrees
    Case "4"
      Koko = "ISO_A4_(297.00_x_210.00_MM)"
      'Koko = "ISO_A4_(210.00_x_297.00_MM)"
      PlotRotation = ac0degrees
    Case Else
      Virheet.Text = Virheet.Text & "Kuvasta " & ActiveDocument.Name & " ei voitu määrittää piirustuskokoa." & vbCrLf
      Exit Sub
  End Select
  SkaleA = Val(Left(TSkale, InStr(TSkale, ":") - 1))
  SkaleB = Val(Mid(TSkale, InStr(TSkale, ":") + 1))
  If SkaleA = 0 Or SkaleB = 0 Then
    Virheet.Text = Virheet.Text & "Makro ei pystynyt määrittämään mittasuhdetta kuvasta " & ActiveDocument.Name & ", merkintä: " & TSkale & vbCrLf
    Exit Sub
  End If
  Nimi = NimiA & "."
  
  If Len(Rev) = 0 Then
    Nimi = Nimi & "00"
  ElseIf Len(Rev) = 1 Then
    If IsNumeric(Rev) Then
      Nimi = Nimi & "0" & Rev
    Else
      Nimi = Nimi & Rev
    End If
  Else
    Nimi = Nimi & Rev
  End If
  If Lehti <> "" Then
    Nimi = Nimi & "-" & Lehti & ".plt"
  Else
    Nimi = Nimi & "-1.plt"
  End If
   
   Origo(0) = 0
   Origo(1) = 0
'   Set Tulostettava = Application.ActiveDocument.ModelSpace

  On Error Resume Next
  With Application.ActiveDocument.ModelSpace.Layout
    Layout(0) = .Name
    .ConfigName = "Arkiston hpgl2.pc3"
    .PlotType = acLimits
    .PlotOrigin = Origo
    .CenterPlot = True
    .SetCustomScale SkaleA, SkaleB
'     .StandardScale = Skale
    .PlotWithPlotStyles = True
    .StyleSheet = "Genius.ctb"
    .CanonicalMediaName = Koko
    .PlotRotation = PlotRotation
    .PaperUnits = acMillimeters
  End With
  If Err <> 0 Then
    Virheet.Text = Virheet.Text & "Virhe: " & Err & " " & Err.Description & " kuvassa (" & ActiveDocument.Name & ")" & vbCrLf
  End If
  Err.Clear
  On Error GoTo 0
   
  ActiveDocument.Regen acActiveViewport
  With ActiveDocument.Plot
    .QuietErrorMode = True
    .SetLayoutsToPlot Layout
    .NumberOfCopies = 1
    .PlotToFile ActiveDocument.Path & "\" & Nimi
  End With
End Sub
Private Function AvaaTiedosto(Avattava As String) As Boolean
'Avaa tiedoston AutoCADissä.
Dim i As Integer
    If FSO.FileExists(Avattava) Then
      For i = 0 To Application.Documents.Count - 1
        Application.Documents(0).Close False   'Suljetaan documentit (Multi Document tilassa AutoCAD 2002 ei herjaa vaikkei documentteja olisikaan ja vaikka documentteja ei olisi talletettu)
      Next i
      Application.Documents.Open (Avattava)  'Avataan uusi dokumentti AutoCadiin (ACAD 2002 osaa avata myös DXF documentit ilman kyselyjä)
      AvaaTiedosto = True
    Else
      AvaaTiedosto = False
      MsgBox "Tiedostoa ei löytynyt: " & Avattava, vbCritical, "Virhe!"
      Virheet.Text = Virheet.Text & "Tiedostoa " & Avattava & " ei löytynyt" & vbCrLf
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
  Tilatieto.Caption = "Tulostaa tiedostoa: " & PFile & " (" & FSO.GetFileName(Tiedosto) & ")  " & blockno & "  Time: " & Min & " min  " & sec & " s"
  Formi.Repaint
End Sub
