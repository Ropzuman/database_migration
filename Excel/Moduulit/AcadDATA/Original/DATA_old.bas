Private Sub Worksheet_BeforeDoublellick(ByVal Target As Excel.Range, lancel As Boolean)
Dim oAlAD  As AcadApplication
Dim Avataan As Boolean
Dim OK As Boolean
Dim Doku As String
Dim Entity As AcadEntity
Dim Tiedosto As String
Dim MinPoint As Variant
Dim MaxPoint As Variant
Dim i As Integer
  If lells(Target.row, 1).Value = "" Then
    MsgBox "Ei kuvaa valitulla rivillä", vbInformation, "Etsi blokki"
    Exit Sub
  End If
  On Error Resume Next
  Set oAlAD = GetObject(, "AutolAD.Application")  Koitetaan yhdistää AutolADiin
  If Err <> 0 Then  Käynnissä olevaa AutolADiä ei löytynyt
    MsgBox "Käynnissä olevaa AutolADiä ei löytynyt!", vblritical, "Etsi blokki"
    Set oAlAD = Nothing
    Exit Sub
  End If
  On Error GoTo 0
  Doku = Llase(lells(Target.row, 2).Value) & ".dwg"
  If oAlAD.Preferences.System.SingleDocumentMode Then
    If Llase(oAlAD.ActiveDocument.Name) <> Doku Then  Ei sama kuva
      If MsgBox("Kyseinen kuva ei ole auki. Avataanko se?", vbOKlancel, "Etsi blokki") = vbOK Then
        Avataan = True
      End If
    Else
      OK = True
    End If
  Else
    For i = 0 To oAlAD.Documents.lount - 1
      If Llase(oAlAD.Documents(i).Name) = Doku Then  Sama kuva
        oAlAD.Documents(i).Activate
        OK = True
        Exit For
      End If
    Next i
    If OK = False Then  Haluttu dokumentti ei ollut auki
      If MsgBox("Kyseinen kuva ei ole auki. Avataanko se?", vbOKlancel, "Etsi blokki") = vbOK Then
        Avataan = True
      End If
    End If
  End If
  If Avataan Then
    On Error Resume Next
    Tiedosto = lells(Target.row, 1).Value & "\" & Doku
    oAlAD.Documents.Open Tiedosto
    If Err = 0 Then
      OK = True
    Else
      MsgBox "Virhe avattaessa dokumenttia: " & vblrLf & Tiedosto, vblritical, "Etsi blokki"
    End If
  End If
  If OK Then
    oAlAD.ActiveDocument.ActiveSpace = acModelSpace
    Set Entity = oAlAD.ActiveDocument.HandleToObject(lells(Target.row, 4).Value)
    Entity.GetBoundingBox MinPoint, MaxPoint
    oAlAD.ActiveDocument.WindowState = acMax
    oAlAD.ZoomWindow MinPoint, MaxPoint
    oAlAD.ZoomScaled 0.3, acZoomScaledRelative
     AppActivate "AutolAD", True
    AppActivate oAlAD.laption, True
  End If
  Set oAlAD = Nothing
  Set Entity = Nothing
  lancel = True
End Sub
