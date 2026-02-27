Attribute VB_Name = "Koodit"
Sub BlokkienTeko()
'Dim sset As AcadSelectionSet
Dim oEntity As AcadEntity
Dim oBlock As AcadBlockReference
Dim iPoint(2) As Double
Dim Origo(2) As Double
Dim Palaset As Variant
Dim i As Integer
Dim Nimi As String
Dim sset As AcadSelectionSet
i = 1
For Each oEntity In ActiveDocument.ModelSpace
  Debug.Print oEntity.Name
  If oEntity.Name <> "UNITSQ_1" Then
    Set oBlock = oEntity
    If i < 10 Then
      Nimi = "ES0" & i & ".dwg"
    Else
      Nimi = "ES" & i & ".dwg"
    End If
    iPoint(0) = oBlock.InsertionPoint(0)
    iPoint(1) = oBlock.InsertionPoint(1)
    iPoint(2) = oBlock.InsertionPoint(2)
    oBlock.InsertionPoint = Origo
    Palaset = oBlock.Explode
    oBlock.InsertionPoint = iPoint
    Set sset = ActiveDocument.ActiveSelectionSet
    sset.Clear
    sset.AddItems Palaset
    ActiveDocument.Wblock "L:\Projekti\ITATA\Process\Flowsheets\BLOCKS\" & Nimi, sset
  End If
  i = i + 1
Next
End Sub
Sub BlokkienTeko2()
'Dim sset As AcadSelectionSet
Dim oEntity As AcadEntity
Dim Origo(2) As Double
Dim mPoint(2) As Double
Dim VAPoint(2) As Double
Dim OYPoint(2) As Double
Dim Palaset As Variant
Dim i As Integer
Dim Nimi As String
Dim sset As AcadSelectionSet
i = 1
VAPoint(0) = 340
OYPoint(0) = 400
VAPoint(1) = 779
OYPoint(1) = 812.25
mPoint(0) = 358.75
mPoint(1) = 797.5

For i = 1 To 22
  Set oBlock = oEntity
  If i < 10 Then
    Nimi = "I0" & i & ".dwg"
  Else
    Nimi = "I" & i & ".dwg"
  End If
  
  Set sset = ActiveDocument.ActiveSelectionSet
  sset.Clear
  sset.Select acSelectionSetWindow, VAPoint, OYPoint
'ActiveDocument.ModelSpace.AddLine VAPoint, OYPoint
  For Each oEntity In sset
    oEntity.Move mPoint, Origo
  Next
  ActiveDocument.Wblock "L:\Projekti\ITATA\Process\Flowsheets\BLOCKS\" & Nimi, sset
  VAPoint(1) = VAPoint(1) - 33.75
  OYPoint(1) = OYPoint(1) - 33.75
  mPoint(1) = mPoint(1) - 33.75
Next i
End Sub
Sub apuinsert()
Dim i As Integer, j As Integer
Dim Nimi As String
Dim Polku As String
Dim iPoint(2) As Double
Dim iPoint2(2) As Double
Dim oText As AcadText
Dim Attrib As Variant
Dim oBlock As AcadBlockReference
Polku = "L:\Projekti\ITATA\Process\Flowsheets\BLOCKS\"
For i = 1 To 21
  iPoint2(0) = -10
  If i < 10 Then
    Nimi = "V0" & i
  Else
    Nimi = "V" & i
  End If
  iPoint(1) = i * 13 * -1
  iPoint2(1) = i * 13 * -1 ' - 2
  Set oBlock = ActiveDocument.ModelSpace.InsertBlock(iPoint, Polku & Nimi & ".dwg", 1, 1, 1, 0)
  Set oText = ActiveDocument.ModelSpace.AddText(Nimi, iPoint2, 3)
  oText.Alignment = acAlignmentMiddleRight
  oText.TextAlignmentPoint = iPoint2
  
  iPoint2(0) = 12
  Attrib = oBlock.GetAttributes
  For j = 0 To UBound(Attrib)
    If Attrib(j).TagString = "TYPE" Then
      Set oText = ActiveDocument.ModelSpace.AddText(Attrib(j).TextString, iPoint2, 3)
      oText.Alignment = acAlignmentMiddleLeft
      oText.TextAlignmentPoint = iPoint2
      Exit For
    End If
  Next j
Next i
End Sub
Sub koe()
Dim sset As AcadSelectionSet
Dim viiva As AcadLine
  Set sset = ActiveDocument.ActiveSelectionSet
  Set viiva = sset(0)
  Debug.Print viiva.StartPoint(0)
  Debug.Print viiva.StartPoint(1)
End Sub
