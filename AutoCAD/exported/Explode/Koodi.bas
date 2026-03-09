Attribute VB_Name = "Koodi"
Option Explicit
Sub Hajota()
Dim Joukko As AcadSelectionSet
Dim Point(0 To 2) As Double
Dim FilterType(0) As Integer
Dim FilterData(0) As Variant
Dim i As Integer
Dim Blokki As AcadBlockReference
Dim Valinta As AcadObject
Dim Viesti As String
  FilterType(0) = 0
  FilterData(0) = "INSERT"
  For i = 0 To ThisDrawing.SelectionSets.Count - 1
    If ThisDrawing.SelectionSets(i).Name = "Rajaytys" Then
      ThisDrawing.SelectionSets(i).Delete
      Exit For
    End If
  Next i
  Set Joukko = ThisDrawing.SelectionSets.Add("Rajaytys")
  Joukko.Select acSelectionSetPrevious, , , FilterType, FilterData
  If Joukko.Count = 0 Then
    Joukko.Delete
    ThisDrawing.SendCommand Chr(27) & Chr(27) 'Varmistetaan että ollaan poistuttu komennosta
    Viesti = "Choose Block..."
    On Error Resume Next
    Do
      ThisDrawing.Utility.GetEntity Valinta, Point, Viesti
      If Err <> 0 Then
        Err.Clear
        On Error GoTo 0
        GoTo Loppu
      ElseIf Valinta.ObjectName = "AcDbBlockReference" Then
        Set Blokki = Valinta
        Exit Do
      Else
        Viesti = "Not a block. Choose again..."
      End If
    Loop
    Pura Blokki
  Else
    For i = 0 To Joukko.Count - 1
      Set Blokki = Joukko(i)
      Pura Blokki
    Next i
    Joukko.Delete
  End If
  MsgBox "Block(s) Exploded." & vbCrLf
Loppu:
  Set Joukko = Nothing
  Set Blokki = Nothing
  ThisDrawing.SendCommand "vbaunload explode.dvb "
End Sub
Private Sub Pura(Blokki As AcadBlockReference)
Dim oAttribs As Variant
Dim i As Integer
Dim Teksti As AcadText
Dim Rikottu As Variant
  If Blokki.HasAttributes Then
    oAttribs = Blokki.GetAttributes
    For i = 0 To UBound(oAttribs)
      If oAttribs(i).TextString <> "" Then
        If oAttribs(i).Invisible = False Then
          On Error Resume Next
          Set Teksti = ThisDrawing.ModelSpace.AddText(oAttribs(i).TextString, oAttribs(i).InsertionPoint, oAttribs(i).Height)
          Teksti.Alignment = oAttribs(i).Alignment
          Teksti.Color = oAttribs(i).Color
          Teksti.Layer = oAttribs(i).Layer
          Teksti.TextAlignmentPoint = oAttribs(i).TextAlignmentPoint
          Teksti.Backward = oAttribs(i).Backward
          Teksti.UpsideDown = oAttribs(i).UpsideDown
          Teksti.ScaleFactor = oAttribs(i).ScaleFactor
          Teksti.StyleName = oAttribs(i).StyleName
          Err.Clear
          On Error GoTo 0
        End If
      End If
    Next i
  End If
  Rikottu = Blokki.Explode
  For i = 0 To UBound(Rikottu)
    If Rikottu(i).ObjectName = "AcDbAttributeDefinition" Then
      Rikottu(i).Delete
    End If
  Next i
  Blokki.Delete
End Sub
