VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Formi 
   Caption         =   "Plot Utility"
   ClientHeight    =   6210
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
Dim FSO As New FileSystemObject
Dim Tulostin As String
Dim Kansio As String
Dim WLowerLeft As Variant
Dim WUpperRight As Variant
Private Sub DWGNo_Change()
  KokoaNimi
End Sub
Private Sub DWGNo_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
  If KeyAscii > 96 And KeyAscii < 123 Then 'Pikkuaakkonen
    KeyAscii = KeyAscii - 32
  ElseIf KeyAscii > 64 And KeyAscii < 91 Then 'Suuraakkonen
  ElseIf KeyAscii > 47 And KeyAscii < 58 Then 'numero
  ElseIf KeyAscii = 45 Or KeyAscii = 95 Or KeyAscii = 32 Then '"-" "_" tai " "
  Else
    KeyAscii = 0
  End If
End Sub
Private Sub DWGSheet_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
  If KeyAscii < 48 Or KeyAscii > 57 Then KeyAscii = 0
End Sub
Private Sub Esikatselu_Click()
  Me.Hide
  If AsetaTulostus Then
    ActiveDocument.Plot.DisplayPlotPreview acFullPreview
  End If
  Me.Show
End Sub
Private Sub OsaEsikatselu_Click()
  Me.Hide
  If AsetaTulostus Then
    ActiveDocument.Plot.DisplayPlotPreview acPartialPreview
  End If
  Me.Show
End Sub
Private Sub Ikkuna_Click()
On Error GoTo Loppu
  Me.Hide
  ActiveDocument.Utility.Prompt "Specify window for printing"
  WLowerLeft = ActiveDocument.Utility.GetPoint(, "Specify first corner:")
  WUpperRight = ActiveDocument.Utility.GetPoint(WLowerLeft, "Specify opposite corner:")
  ReDim Preserve WLowerLeft(0 To 1)
  ReDim Preserve WUpperRight(0 To 1)
  ActiveDocument.ActiveLayout.SetWindowToPlot WLowerLeft, WUpperRight
  OWindow.Value = True
Loppu:
  Me.Show
  Err.Clear
End Sub
Private Sub OWindow_Click()
  If IsEmpty(WLowerLeft) Then
    OExtents.Value = True
  End If
End Sub

Private Sub Rev_Change()
  KokoaNimi
End Sub
Private Sub DWGSheet_Change()
  KokoaNimi
End Sub
Private Sub KokoaNimi()
  Filename.Caption = DWGNo.Value & "."
  If Len(Rev.Value) = 0 Then
    Filename.Caption = Filename.Caption & "00"
  ElseIf Len(Rev.Value) = 1 Then
    If IsNumeric(Rev.Value) Then
      Filename.Caption = Filename.Caption & "0" & Rev.Value
    Else
      Filename.Caption = Filename.Caption & Rev.Value
    End If
  Else
    Filename.Caption = Filename.Caption & Rev.Value
  End If
  If DWGSheet.Value <> "" Then
    Filename.Caption = Filename.Caption & "-" & DWGSheet.Value & ".plt"
  Else
    Filename.Caption = Filename.Caption & "-1.plt"
  End If
End Sub
Private Sub OLandscape_Change()
  PaperinSuunta
End Sub
Private Sub OPortrait_Change()
  PaperinSuunta
End Sub
Private Sub Rev_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
  If KeyAscii > 96 And KeyAscii < 123 Then 'Pikkuaakkonen
    KeyAscii = KeyAscii - 32
  ElseIf KeyAscii > 64 And KeyAscii < 91 Then 'Suuraakkonen
  ElseIf KeyAscii > 47 And KeyAscii < 58 Then 'numero
  Else
    KeyAscii = 0
  End If
End Sub

Private Sub Scale1_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
  If KeyAscii < 48 Or KeyAscii > 57 Then KeyAscii = 0
End Sub
Private Sub Scale2_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
  If KeyAscii < 48 Or KeyAscii > 57 Then KeyAscii = 0
End Sub
Private Sub Skaalaus_Change()
  If Skaalaus.Value = 99 Then
  ElseIf Skaalaus.Value = 0 Then
    Scale1.Value = ""
    Scale2.Value = ""
  Else
    Scale1.Value = Left(Skaalaus.Text, InStr(Skaalaus.Text, ":") - 1)
    Scale2.Value = Mid(Skaalaus.Text, InStr(Skaalaus.Text, ":") + 1)
  End If
End Sub
Private Sub Upsidedown_Change()
  PaperinSuunta
End Sub
Private Sub PaperinSuunta()
  IPysty.Visible = False
  IVaaka.Visible = False
  IPystyR.Visible = False
  IVaakaR.Visible = False
  If OPortrait.Value Then
    If Upsidedown.Value Then
      IPystyR.Visible = True
    Else
      IPysty.Visible = True
    End If
  Else
    If Upsidedown.Value Then
      IVaakaR.Visible = True
    Else
      IVaaka.Visible = True
    End If
  End If
End Sub
Private Sub Ulos_Click()
  Unload Formi
End Sub
Private Sub UserForm_Initialize()
'Avaus
Dim i As Integer
Dim Tulostimet As Variant
Dim Loytyi As Boolean
Dim Joukko As AcadSelectionSet  'Joukko, jolla valitaan kaikki halutut blokit
Dim oBlock As AcadBlockReference
Dim BlockArray As Variant       'Array muuttuja Blokkia varten
Dim FilterType(0 To 1) As Integer
Dim FilterData(0 To 1) As Variant
Dim Skalet(1 To 18, 1 To 2) As Variant
Dim Tyylit As Variant
Dim Koot As Variant
Dim oTiedosto As Scripting.TextStream
Dim Tiedot As Variant
  
  FilterType(1) = 66
  FilterData(1) = 1
  FilterType(0) = 2       'ObjecName
  FilterData(0) = "TITLEBLOCK_METSOP_2001,TITLEBLOCK_METSOP_2002,TITLEBLOCK_METSOP_2003,3LINEBOM,3LINEBOM_VALMET,F1A0,F1A1,F1A2,F1A3,F1A4"
  FilterData(0) = FilterData(0) & ",POR_OTS_TAU,POROTSI,POROTS_TAU"


  Kansio = "K:\PLTFILES\_HPGL2\"
  Tulostin = "Arkiston hpgl2.pc3"
  
  Tulostimet = ActiveDocument.Layouts(0).GetPlotDeviceNames()
  For i = LBound(Tulostimet) To UBound(Tulostimet)
    If Tulostimet(i) = Tulostin Then
      Loytyi = True
      Exit For
    End If
  Next i
  If Loytyi = False Then
    MsgBox "Can't find plotter!" & vbCrLf & "You have to have plotter: """ & Tulostin & """ installed.", vbCritical, "LTVF Plot"
    Unload Formi
  End If

  For i = 0 To ActiveDocument.SelectionSets.Count - 1
    If ActiveDocument.SelectionSets(i).Name = "LTVFHAKU" Then
      ActiveDocument.SelectionSets(i).Delete
      Exit For
    End If
  Next i
  Set Joukko = ActiveDocument.SelectionSets.Add("LTVFHAKU")
  Joukko.Select acSelectionSetAll, , , FilterType, FilterData
  If Joukko.Count > 0 Then
    Set oBlock = Joukko(0)
    BlockArray = oBlock.GetAttributes
    For i = LBound(BlockArray) To UBound(BlockArray)
      Select Case UCase(BlockArray(i).TagString)
        Case "ITEM_DOC_NO", "HEADER4", "HEADER1", "NUMBER"
          DWGNo.Value = BlockArray(i).TextString
        Case "REV", "HEADER5", "HEADER2", "R"
          If BlockArray(i).TextString = "-" Then
            Rev.Value = "0"
          Else
            If Len(BlockArray(i).TextString) > 2 Then
               Rev.Value = UCase(Left(Replace(BlockArray(i).TextString, ".", ""), 2))
            Else
               Rev.Value = UCase(Replace(BlockArray(i).TextString, ".", ""))
            End If
          End If
        Case "SHEET", "HEADER11", "SHE"
          DWGSheet.Value = Val(BlockArray(i).TextString)
      End Select
    Next i
  End If
  Joukko.Delete
  Set Joukko = Nothing
  If DWGNo.Value = "" Then
    DWGNo.Value = Left(ActiveDocument.Name, Len(ActiveDocument.Name) - 4)
  End If
  
'Laitetaan Skaalausvaihtoehdot
  Skalet(1, 2) = "Custom": Skalet(1, 1) = 99
  Skalet(2, 2) = "Scale to Fit": Skalet(2, 1) = acScaleToFit
  Skalet(3, 2) = "1:1": Skalet(3, 1) = ac1_1
  Skalet(4, 2) = "1:2": Skalet(4, 1) = ac1_2
  Skalet(5, 2) = "1:4": Skalet(5, 1) = ac1_4
  Skalet(6, 2) = "1:8": Skalet(6, 1) = ac1_8
  Skalet(7, 2) = "1:10": Skalet(7, 1) = ac1_10
  Skalet(8, 2) = "1:16": Skalet(8, 1) = ac1_16
  Skalet(9, 2) = "1:20": Skalet(9, 1) = ac1_20
  Skalet(10, 2) = "1:30": Skalet(10, 1) = ac1_30
  Skalet(11, 2) = "1:40": Skalet(11, 1) = ac1_40
  Skalet(12, 2) = "1:50": Skalet(12, 1) = ac1_50
  Skalet(13, 2) = "1:100": Skalet(13, 1) = ac1_100
  Skalet(14, 2) = "2:1": Skalet(14, 1) = ac2_1
  Skalet(15, 2) = "4:1": Skalet(15, 1) = ac4_1
  Skalet(16, 2) = "8:1": Skalet(16, 1) = ac8_1
  Skalet(17, 2) = "10:1": Skalet(17, 1) = ac10_1
  Skalet(18, 2) = "100:1": Skalet(18, 1) = ac100_1
  Skaalaus.List = Skalet
  Skaalaus.ListIndex = 0


'Kynäasetusten valinnat
  Tyylit = ActiveDocument.ActiveLayout.GetPlotStyleTableNames()
  For i = LBound(Tyylit) To UBound(Tyylit)
    Pens.AddItem Tyylit(i)
  Next i
  Pens.ListIndex = 0
  
'Kokojen valinnat
  ActiveDocument.ActiveLayout.ConfigName = Tulostin
  Koot = ActiveDocument.ActiveLayout.GetCanonicalMediaNames()
  For i = LBound(Koot) To UBound(Koot)
    Size.AddItem Koot(i)
    Size.List(i, 1) = ActiveDocument.ActiveLayout.GetLocaleMediaName(Koot(i))
  Next i
  Size.ListIndex = 0

'Näkymien valinnat
  If ActiveDocument.Views.Count > 0 Then
    For i = 0 To ActiveDocument.Views.Count - 1
      Nakymat.AddItem ActiveDocument.Views(i).Name
    Next i
    Nakymat.ListIndex = 0
  Else
    Nakymat.Enabled = False
    OView.Enabled = False
  End If

  If FSO.FileExists(Application.Preferences.Files.TempFilePath & "LTVFPlot.TXT") Then
    Set oTiedosto = FSO.OpenTextFile(Application.Preferences.Files.TempFilePath & "LTVFPlot.TXT")
    Tiedot = Split(oTiedosto.ReadLine, vbTab)
    oTiedosto.Close
    On Error Resume Next
    Size.ListIndex = Tiedot(0)
    Pens.ListIndex = Tiedot(1)
    KeskitaTuloste.Value = Tiedot(2)
  
    OPortrait.Value = Tiedot(3)
    OLandscape.Value = Tiedot(4)
    Upsidedown.Value = Tiedot(5)
  
    OLimits.Value = Tiedot(6)
    OExtents.Value = Tiedot(7)
    OView.Value = Tiedot(8)
    Nakymat.ListIndex = Tiedot(9)
    OWindow.Value = Tiedot(10)
    If Tiedot(10) <> "x" Then
      ReDim WLowerLeft(0 To 1)
      ReDim WUpperRight(0 To 1)
      WLowerLeft(0) = Tiedot(11)
      WLowerLeft(1) = Tiedot(12)
      WUpperRight(0) = Tiedot(13)
      WUpperRight(1) = Tiedot(14)
    End If
    Skaalaus.ListIndex = Tiedot(15)
    Scale1.Value = Tiedot(16)
    Scale2.Value = Tiedot(17)
    Err.Clear
    On Error GoTo 0
  End If
End Sub
Private Sub KirjoitaTiedot()
Dim oTiedosto As Scripting.TextStream
  Set oTiedosto = FSO.CreateTextFile(Application.Preferences.Files.TempFilePath & "LTVFPlot.TXT", True)
  'Kokoja kynät
  oTiedosto.Write Size.ListIndex & vbTab
  oTiedosto.Write Pens.ListIndex & vbTab
  oTiedosto.Write KeskitaTuloste.Value & vbTab
  
  oTiedosto.Write OPortrait.Value & vbTab
  oTiedosto.Write OLandscape.Value & vbTab
  oTiedosto.Write Upsidedown.Value & vbTab
  
  oTiedosto.Write OLimits.Value & vbTab
  oTiedosto.Write OExtents.Value & vbTab
  oTiedosto.Write OView.Value & vbTab
  oTiedosto.Write Nakymat.ListIndex & vbTab
  oTiedosto.Write OWindow.Value & vbTab
  If IsEmpty(WLowerLeft) = False Then
    oTiedosto.Write WLowerLeft(0) & vbTab
    oTiedosto.Write WLowerLeft(1) & vbTab
    oTiedosto.Write WUpperRight(0) & vbTab
    oTiedosto.Write WUpperRight(1) & vbTab
  Else
    oTiedosto.Write x & vbTab
    oTiedosto.Write x & vbTab
    oTiedosto.Write x & vbTab
    oTiedosto.Write x & vbTab
  End If
  
  oTiedosto.Write Skaalaus.ListIndex & vbTab
  oTiedosto.Write Scale1.Value & vbTab
  oTiedosto.Write Scale2.Value & vbCrLf
  oTiedosto.Close
  Set oTiedosto = Nothing
End Sub
Private Function AsetaTulostus() As Boolean
Dim Origo(0 To 1) As Double
   Origo(0) = 0
   Origo(1) = 0
   With ActiveDocument.ActiveLayout
     .ConfigName = Tulostin
     If OLimits.Value = True Then
       .PlotType = acLimits
     ElseIf OExtents.Value = True Then
       .PlotType = acExtents
     ElseIf OView.Value = True Then
       .ViewToPlot = Nakymat.Value
       .PlotType = acView
     ElseIf OWindow.Value = True Then
       .SetWindowToPlot WLowerLeft, WUpperRight
       .PlotType = acWindow
     End If
     .PlotOrigin = Origo
     .CenterPlot = KeskitaTuloste.Value
     If Skaalaus.Value = 99 Then
       If IsNumeric(Scale1.Value) And IsNumeric(Scale2.Value) Then
         .SetCustomScale Scale1.Value, Scale2.Value
       Else
         MsgBox "Please enter positive real.", vbInformation, "LTVF Plot"
         If IsNumeric(Scale1.Value) Then
           Scale2.SetFocus
           Scale2.SelStart = 1
           Scale2.SelLength = Len(Scale2.Value)
         Else
           Scale1.SetFocus
           Scale1.SelStart = 1
           Scale1.SelLength = Len(Scale1.Value)
         End If
         AsetaTulostus = False
         Exit Function
       End If
     Else
       .StandardScale = Skaalaus.Value  ' acScaleToFit
     End If
     .PlotWithPlotStyles = True
     .StyleSheet = Pens.Value
     .CanonicalMediaName = Size.Value
     If OPortrait.Value = True Then
       If Upsidedown.Value = True Then
         .PlotRotation = ac180degrees
       Else
         .PlotRotation = ac0degrees
       End If
     Else
       If Upsidedown.Value = True Then
         .PlotRotation = ac270degrees
       Else
         .PlotRotation = ac90degrees
       End If
     End If
     .PaperUnits = acMillimeters
   End With
   ActiveDocument.Regen acActiveViewport
   AsetaTulostus = True
End Function
Private Sub Tulosta_Click()
  If AsetaTulostus Then
    With ActiveDocument.Plot
      .QuietErrorMode = True
      If DWGNo.Value = "" Then   '".00-1.plt"
        MsgBox "Please give proper filename.", vbInformation, "LTVF Plot"
        Exit Sub
      Else
        If Dir(Kansio & Filename.Caption) <> "" Then
          If MsgBox("File """ & Filename.Caption & """ already exists." & vbCrLf & "Do you want to overwrite it?", vbOKCancel, "LTVF Plot") = vbCancel Then
            Exit Sub
          Else
            .PlotToFile Kansio & Filename.Caption
          End If
        Else
          .PlotToFile Kansio & Filename.Caption
        End If
      End If
    End With
  End If
  Unload Me
End Sub
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
