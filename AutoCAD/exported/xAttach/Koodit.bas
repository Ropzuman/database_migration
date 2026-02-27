Attribute VB_Name = "Koodit"
Sub PoistaXRef()
Dim i As Long
Dim Joukko As AcadSelectionSet
Dim FilterType(0) As Integer
Dim FilterData(0) As Variant
Dim iPiste As Variant
Dim Nimi As String
Dim xScale As Double, yScale As Double, zScale As Double, Rotation As Double
Dim Loytyi As Boolean
Dim Tied As String
Dim Dokki As AcadDocument
Dim oACAD As AcadApplication
Dim oBlokki As AcadBlockReference

Set oACAD = Application
Set Dokki = ActiveDocument

Tied = Dokki.Path & "\" & Left(Dokki.Name, Len(Dokki.Name) - 4) & "_NoXRef.DWG"

FilterType(0) = 2
For i = 0 To Dokki.SelectionSets.Count - 1
  If Dokki.SelectionSets(i).Name = "XREFSELECT" Then
    Dokki.SelectionSets(i).Delete
    Exit For
  End If
Next i
On Error Resume Next
Do
  Loytyi = False
  Set Joukko = Dokki.SelectionSets.Add("XREFSELECT")
  For i = 0 To Dokki.Blocks.Count - 1
    If Dokki.Blocks(i).IsXRef Then
      Loytyi = True
      Joukko.Clear
      FilterData(0) = Dokki.Blocks(i).Name
      Joukko.Select acSelectionSetAll, , , FilterType, FilterData
      If Joukko.Count > 0 Then
        iPiste = Joukko(0).InsertionPoint
        Nimi = Joukko(0).Path
        xScale = Joukko(0).XScaleFactor
        yScale = Joukko(0).YScaleFactor
        zScale = Joukko(0).ZScaleFactor
        Rotation = Joukko(0).Rotation
        Dokki.Blocks(i).Detach
        If Err = 0 Then
          Joukko.Erase
          Joukko.Delete
          Set Joukko = Nothing
          Dokki.PurgeAll
          Dokki.PurgeAll
          Dokki.Regen acAllViewports
          Dokki.SaveAs Tied
          Dokki.Close False
          Set Dokki = oACAD.Documents.Open(Tied)
          Set oBlokki = Dokki.ModelSpace.InsertBlock(iPiste, Nimi, xScale, yScale, zScale, Rotation)
          If Err = 0 Then
            oBlokki.Explode
            oBlokki.Delete
          Else
            Err.Clear
            Dokki.SendCommand "-insert """ & Nimi & """" & vbCr & "0,0" & vbCr & xScale & vbCr & yScale & vbCr & Rotation & vbCr
          End If
          Dokki.PurgeAll
          Dokki.PurgeAll
          Dokki.AuditInfo True
          Dokki.Save
          Exit For
        Else
          Err.Clear
        End If
      End If
    End If
  Next i
Loop Until Loytyi = False
Set Dokki = Nothing
Set oACAD = Nothing
MsgBox "XRefit poistettu kuvasta ja insertoitu blokkeina", vbInformation, "Poista XRef"
End Sub

