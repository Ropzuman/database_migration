Attribute VB_Name = "Module1"
Sub Tasot()
Dim APU As Object
Dim Piste As Variant
Dim oBlock As AcadBlockReference
Dim Blokki As AcadBlock
Dim Objektit As Variant
Dim Nimi As String
Dim Joukko As AcadSelectionSet
On Error GoTo Loppu
ActiveDocument.Utility.GetEntity APU, pististe, "Valitse Blokki"

If APU.ObjectName = "AcDbBlockReference" Then
  Set oBlock = APU
  Nimi = oBlock.Name
  Piste = oBlock.InsertionPoint
  If MsgBox("Haluako muuttaa blokin " & Nimi & " osat aktiiviselle tasolle?", vbYesNo, "Tasot") = vbYes Then
    'Muutetaan blokin elementtien layer
    For i = 0 To ActiveDocument.Blocks(Nimi).Count - 1
      ActiveDocument.Blocks(Nimi).Item(i).Layer = ActiveDocument.ActiveLayer.Name
    Next i
    ActiveDocument.Regen acActiveViewport
  End If
End If
Exit Sub
Loppu:
End Sub

