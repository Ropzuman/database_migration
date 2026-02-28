Private Sub Worksheet_BeforeDoubleClick(ByVal Target As Excel.Range, Cancel As Boolean)
Dim oACAD  As AcadApplication
Dim Avataan As Boolean
Dim OK As Boolean
Dim Doku As String
Dim Entity As AcadEntity
Dim Tiedosto As String
Dim MinPoint As Variant
Dim MaxPoint As Variant
Dim i As Integer
  If Cells(Target.row, 1).Value = "" Then
    MsgBox "Ei kuvaa valitulla rivillä", vbInformation, "Etsi blokki"
    Exit Sub
  End If
  On Error Resume Next
  Set oACAD = GetObject(, "AutoCAD.Application") 'Koitetaan yhdistää AutoCADiin
  If Err <> 0 Then 'Käynnissä olevaa AutoCADiä ei löytynyt
    MsgBox "Käynnissä olevaa AutoCADiä ei löytynyt!", vbCritical, "Etsi blokki"
    Set oACAD = Nothing
    Exit Sub
  End If
  On Error GoTo 0
  Doku = LCase(Cells(Target.row, 2).Value) & ".dwg"
  If oACAD.Preferences.System.SingleDocumentMode Then
    If LCase(oACAD.ActiveDocument.Name) <> Doku Then 'Ei sama kuva
      If MsgBox("Kyseinen kuva ei ole auki. Avataanko se?", vbOKCancel, "Etsi blokki") = vbOK Then
        Avataan = True
      End If
    Else
      OK = True
    End If
  Else
    For i = 0 To oACAD.Documents.Count - 1
      If LCase(oACAD.Documents(i).Name) = Doku Then 'Sama kuva
        oACAD.Documents(i).Activate
        OK = True
        Exit For
      End If
    Next i
    If OK = False Then 'Haluttu dokumentti ei ollut auki
      If MsgBox("Kyseinen kuva ei ole auki. Avataanko se?", vbOKCancel, "Etsi blokki") = vbOK Then
        Avataan = True
      End If
    End If
  End If
  If Avataan Then
    On Error Resume Next
    Tiedosto = Cells(Target.row, 1).Value & "\" & Doku
    oACAD.Documents.Open Tiedosto
    If Err = 0 Then
      OK = True
    Else
      MsgBox "Virhe avattaessa dokumenttia: " & vbCrLf & Tiedosto, vbCritical, "Etsi blokki"
    End If
  End If
  If OK Then
    oACAD.ActiveDocument.ActiveSpace = acModelSpace
    Set Entity = oACAD.ActiveDocument.HandleToObject(Cells(Target.row, 4).Value)
    Entity.GetBoundingBox MinPoint, MaxPoint
    oACAD.ActiveDocument.WindowState = acMax
    oACAD.ZoomWindow MinPoint, MaxPoint
    oACAD.ZoomScaled 0.3, acZoomScaledRelative
    'AppActivate "AutoCAD", True
    AppActivate oACAD.Caption, True
  End If
  Set oACAD = Nothing
  Set Entity = Nothing
  Cancel = True
End Sub
