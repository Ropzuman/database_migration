Attribute VB_Name = "Koodit"
Sub TaytaLista()
Dim i As Integer
Dim j As Integer
  ThisDrawing.ActiveLayer = ThisDrawing.Layers("TEXT05")
  For i = 1 To 270
    LisaaTeksti i, 0, "Nimi" & i
    For j = 1 To 9
      LisaaTeksti i, j, "X"
    Next j
  Next i
End Sub
Private Sub LisaaTeksti(ByVal Paikka As Integer, Sarake As Integer, Teksti As String)
Dim Rivi As Integer
Dim Sar As Integer
Dim iPoint(2) As Double
Dim XSiir As Double
Const AloitusX = 26.5
Const AloitusY = 277
Const RiviKork = 5.5
Const SarLev = 5
Const EkaLev = 19.5
Const ColLev = 64.6
Const Riveja = 45
Paikka = Paikka - 1
Rivi = Paikka Mod Riveja
Sar = Paikka \ Riveja
If Sarake = 0 Then
 XSiir = 0
Else
 XSiir = EkaLev + (Sarake - 1) * SarLev
End If
iPoint(0) = AloitusX + ColLev * Sar + XSiir
iPoint(1) = AloitusY - (Rivi) * RiviKork
Application.ActiveDocument.ModelSpace.AddText Teksti, iPoint, 2
End Sub
Private Sub TuhoaText()
Dim Joukko As AcadSelectionSet
Dim FilterType(0) As Integer
Dim FilterData(0) As Variant
Dim i As Integer
Application.ZoomExtents
For i = 0 To ThisDrawing.SelectionSets.Count - 1
  If ThisDrawing.SelectionSets(i).Name = "DELETE" Then
    ThisDrawing.SelectionSets(i).Delete
  End If
Next i
Set Joukko = ThisDrawing.SelectionSets.Add("DELETE")
FilterType(0) = 8        'Type 8 = Layer
FilterData(0) = "TEXT05" 'Layer name
Joukko.Select acSelectionSetAll, , , FilterType, FilterData
Joukko.Erase
Joukko.Delete
Set Joukko = Nothing
End Sub
