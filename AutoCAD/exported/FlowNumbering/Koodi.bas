Attribute VB_Name = "Koodi"
Sub ManualValveNumbering()
  FrmManValve.Show
End Sub
Sub FlowSubNumebers()
Dim Joukko As AcadSelectionSet
Dim i As Integer
Dim FilterType(0) As Integer
Dim FilterData(0) As Variant
Dim Attribuutit As Variant
Dim Attribuutit2 As Variant
Dim Attrib As AcadAttributeReference
Dim MinPoint As Variant
Dim MaxPoint As Variant
Dim Valinta As Object
Dim IPiste As Variant

FilterType(0) = 2       'ObjecName
FilterData(0) = "UP008"  'Block Names

For i = 0 To ActiveDocument.SelectionSets.Count - 1
  If ActiveDocument.SelectionSets(i).Name = "APUFLOW" Then
    ActiveDocument.SelectionSets(i).Delete
    Exit For
  End If
Next i
Set Joukko = ActiveDocument.SelectionSets.Add("APUFLOW")
Joukko.Select acSelectionSetAll, , , FilterType, FilterData
On Error Resume Next
For i = 0 To Joukko.Count - 1
  Attribuutit = Joukko(i).GetAttributes
  For j = 0 To UBound(Attribuutit)
    If Attribuutit(j).TagString = "CUSTPOS" Then
      Set Attrib = Attribuutit(j)
      Exit For
    End If
  Next j
  If Attrib.TextString = "" Then
    ActiveDocument.Regen acActiveViewport
    Joukko(i).GetBoundingBox MinPoint, MaxPoint
    ZoomWindow MinPoint, MaxPoint
    ZoomScaled 0.3, acZoomScaledRelative
    Joukko(i).Highlight True
    Do
      ActiveDocument.Utility.GetEntity Valinta, IPiste, "Select Block..."
      If Err = 0 Then
       If Valinta.ObjectName = "AcDbBlockReference" Then
         If Valinta.HasAttributes Then
           Attribuutit2 = Valinta.GetAttributes
           For j = 0 To UBound(Attribuutit2)
             If Attribuutit2(j).TagString = "CUSTPOS" Then
               Attrib.TextString = Attribuutit2(j).TextString
               ActiveDocument.Utility.Prompt "Selected number: " & Attribuutit2(j).TextString & vbCrLf
               Exit Do
             End If
           Next j
         End If
       End If
      Else
        Err.Clear
      End If
    Loop
    Joukko(i).Highlight False
  End If
  Set Attrib = Nothing
Next i
End Sub
Sub AutomNumebers()
Dim Joukko As AcadSelectionSet
Dim i As Integer
Dim FilterType(0) As Integer
Dim FilterData(0) As Variant
Dim Attribuutit As Variant
Dim Attribuutit2 As Variant
Dim Attrib As AcadAttributeReference
Dim MinPoint As Variant
Dim MaxPoint As Variant
Dim Valinta As Object
Dim IPiste As Variant
Dim POS As String
Dim FUNC As String

FilterType(0) = 2       'ObjecName
FilterData(0) = "UP070,UP071"  'Block Names

Set Joukko = ActiveDocument.ActiveSelectionSet
Joukko.Clear
Joukko.Select acSelectionSetAll, , , FilterType, FilterData
On Error Resume Next
For i = 0 To Joukko.Count - 1
  Attribuutit = Joukko(i).GetAttributes
  For j = 0 To UBound(Attribuutit)
    If Attribuutit(j).TagString = "SD-POS" Then
      Set Attrib = Attribuutit(j)
      Exit For
    End If
  Next j
  If Attrib.TextString = "" Or Attrib.TextString = "0001" Then
    ActiveDocument.Regen acActiveViewport
    Joukko(i).GetBoundingBox MinPoint, MaxPoint
    ZoomWindow MinPoint, MaxPoint
    ZoomScaled 0.3, acZoomScaledRelative
    Joukko(i).Highlight True
    Do
      ActiveDocument.Utility.GetEntity Valinta, IPiste, "Select Block..."
      If Err = 0 Then
       If Valinta.ObjectName = "AcDbBlockReference" Then
         If Valinta.HasAttributes Then
           Attribuutit2 = Valinta.GetAttributes
           POS = ""
           FUNC = ""
           For j = 0 To UBound(Attribuutit2)
             If Attribuutit2(j).TagString = "CUSTPOS" Then
                POS = Attribuutit2(j).TextString
             ElseIf Attribuutit2(j).TagString = "FUNCTION" Then
                FUNC = Attribuutit2(j).TextString
             End If
           Next j
           If POS <> "" And FUNC <> "" Then
              Attrib.TextString = FUNC & "-" & POS
              ActiveDocument.Utility.Prompt "Selected number: " & FUNC & "-" & POS & vbCrLf
           End If
           Exit Do
         End If
       End If
      Else
        Err.Clear
      End If
    Loop
    Joukko(i).Highlight False
  End If
  Set Attrib = Nothing
Next i
End Sub
Sub AutomValve()
Dim Joukko As AcadSelectionSet
Dim i As Integer, j As Integer, ii As Integer
Dim FilterType(0) As Integer
Dim FilterData(0) As Variant
Dim Attribuutit As Variant
Dim Attribuutit2 As Variant
Dim Attrib As AcadAttributeReference
Dim MinPoint As Variant
Dim MaxPoint As Variant
Dim Valinta As Object
Dim IPiste As Variant
Dim vBlock As AcadBlockReference
Dim uBlock As AcadBlockReference
Dim Exploded As Variant

FilterType(0) = 2       'ObjecName
FilterData(0) = "UP070,UP071,UP009"  'Block Names

Set Joukko = ActiveDocument.ActiveSelectionSet
Joukko.Clear
Joukko.Select acSelectionSetAll, , , FilterType, FilterData
On Error Resume Next
For i = 0 To Joukko.Count - 1
    ActiveDocument.Regen acActiveViewport
    Joukko(i).GetBoundingBox MinPoint, MaxPoint
    ZoomWindow MinPoint, MaxPoint
    ZoomScaled 0.3, acZoomScaledRelative
    Joukko(i).Highlight True
    Set vBlock = Joukko(i)
    Do
      ActiveDocument.Utility.GetEntity Valinta, IPiste, "Select Block..."
      If Err = 0 Then
        If Valinta.ObjectName = "AcDbBlockReference" Then
          If Valinta.HasAttributes Then
            With vBlock
              Set uBlock = ActiveDocument.ModelSpace.InsertBlock(.InsertionPoint, .Name & "B", .XScaleFactor, .YScaleFactor, .ZScaleFactor, .Rotation)
            End With
            Attribuutit = uBlock.GetAttributes
            Attribuutit2 = Valinta.GetAttributes
            For j = 0 To UBound(Attribuutit2)
              For ii = 0 To UBound(Attribuutit)
                If Attribuutit(ii).TagString = Attribuutit2(j).TagString Then
                  Attribuutit(ii).TextString = Attribuutit2(j).TextString
                  Attribuutit(ii).Rotation = Attribuutit2(j).Rotation
                  Attribuutit(ii).Alignment = Attribuutit2(j).Alignment
                  Attribuutit(ii).InsertionPoint = Attribuutit2(j).InsertionPoint
                  Attribuutit(ii).TextAlignmentPoint = Attribuutit2(j).TextAlignmentPoint
                  Attribuutit(ii).Height = Attribuutit2(j).Height
                  Attribuutit(ii).ScaleFactor = Attribuutit2(j).ScaleFactor
                  Exit For
                End If
              Next ii
            Next j
            Exploded = Valinta.Explode
            Valinta.Delete
            For j = 0 To UBound(Exploded)
              If Exploded(j).ObjectName <> "AcDbCircle" Then
                Exploded(j).Delete
              End If
            Next j
            Joukko(i).Delete
            Exit Do
          End If
        End If
      Else
        Err.Clear
      End If
    Loop
    Joukko(i).Highlight False
  Set Attrib = Nothing
Next i
End Sub
