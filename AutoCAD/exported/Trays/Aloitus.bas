Attribute VB_Name = "Aloitus"
Public Piirto As Boolean
Public Lev As Integer
Public Puoli As Integer
Public EdViiva As AcadMLine
Public kPisteet() As Double
Public vPisteet() As Double
Public oPisteet() As Double
Public MLinePituus As Double
Public Kulmia As Integer
Public Const PI = 3.14159265358979
Sub Aloita()
 CableTray.Show vbModeless
End Sub
Public Function LaskePituus(Piste1 As Variant, Piste2 As Variant, Optional Elev As Double) As Double
Dim xlen As Double
Dim ylen As Double
Dim zlen As Double
      xlen = Abs(Piste1(0) - Piste2(0))
      ylen = Abs(Piste1(1) - Piste2(1))
      If Elev <> 0 Then
        zlen = Elev
      Else
        zlen = Abs(Piste1(2) - Piste2(2))
      End If
      LaskePituus = Sqr(xlen * xlen + ylen * ylen + zlen * zlen)
End Function
Sub LaskeMLinePituus(Optional Elev As Double)
Dim i As Integer '
Dim Piste1(2) As Double
Dim Piste2(2) As Double
  MLinePituus = 0
  If Not EdViiva Is Nothing Then
    For i = 0 To UBound(EdViiva.Coordinates) - 3 Step 3
        Piste1(0) = EdViiva.Coordinates(i)
        Piste1(1) = EdViiva.Coordinates(i + 1)
        Piste1(2) = EdViiva.Coordinates(i + 2)
        Piste2(0) = EdViiva.Coordinates(i + 3)
        Piste2(1) = EdViiva.Coordinates(i + 4)
        Piste2(2) = EdViiva.Coordinates(i + 5)
        MLinePituus = MLinePituus + LaskePituus(Piste1, Piste2, Elev)
    Next i
    CableTray.VPituus.Value = CLng(MLinePituus)
  End If
End Sub
