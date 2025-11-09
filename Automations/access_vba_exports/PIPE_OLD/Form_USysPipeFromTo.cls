Option Compare Database
Option Explicit
Dim oACAD As AcadApplication
Dim oDOC As AcadDocument
Dim AttribFrom As AcadAttributeReference
Dim AttribTo As AcadAttributeReference
Dim oBlock As AcadBlockReference
Dim Joukko As AcadSelectionSet
Dim i As Integer
Private Sub BClearFrom_Click()
  AttribFrom.TextString = ""
  TFrom.Value = ""
End Sub

Private Sub BClearTo_Click()
  AttribTo.TextString = ""
  TTo.Value = ""
End Sub
Private Sub BClose_Click()
  DoCmd.Close acForm, "UsysPipeFromTo"
End Sub
Private Sub OpenSource(Teksti As TextBox, oAttrib As AcadAttributeReference)
Dim Kohde As Object
Dim Piste As Variant
Dim oPipeBlock As AcadBlockReference
Dim Attrib As Variant
Dim j As Integer
AppActivate oACAD.Caption, True
On Error Resume Next
oDOC.Utility.GetEntity Kohde, Piste, "Pick pipeline, tank, pump or equipment: "
If Err <> 0 Then
  Err.Clear
Else
  If Kohde.ObjectName = "AcDbBlockReference" Then
    Set oPipeBlock = Kohde
    Select Case LCase(oPipeBlock.EffectiveName)
      Case "pipeline", "pipeline_f", "arapipel", "pipeline_data", "metso_pipe"
         Attrib = oPipeBlock.GetAttributes
         Teksti.Value = AttribToPipeline(Attrib)
      Case Else
        If oPipeBlock.HasAttributes Then
          Attrib = oPipeBlock.GetAttributes
          For j = 0 To UBound(Attrib)
            If Attrib(j).TagString = "POS.NO" Or Attrib(j).TagString = "CUSTPOS" Or Attrib(j).TagString = "SDPOS" Then
              Teksti.Value = Attrib(j).TextString
              Exit For
            End If
          Next j
        End If
    End Select
    oAttrib.TextString = Teksti.Value
  ElseIf Kohde.ObjectName = "AcDbText" Then
    Teksti.Value = Kohde.TextString
    oAttrib.TextString = Teksti.Value
  End If
End If
On Error GoTo 0
AppActivate CurrentDb.Properties("AppTitle").Value

End Sub
Private Function AttribToPipeline(Attrib As Variant) As String
Dim j As Integer
Dim DEP As String
Dim FSB As String
Dim LINE As String

For j = 0 To UBound(Attrib)
  Select Case Attrib(j).TagString
    Case "DEP", "DEPA"
      DEP = Attrib(j).TextString
    Case "FSB", "FLOWSUBSTMAINGROUP"
      FSB = Attrib(j).TextString
    Case "LINE", "LINENO"
      LINE = Attrib(j).TextString
    Case Else
    
  End Select
Next j
AttribToPipeline = DEP & FSB & LINE
'AttribToPipeline = DEP & "-" & FSB & "-" & LINE
End Function

Private Sub BFind_Click()
Do
  i = i + 1
  If i > Joukko.Count - 1 Then
    If MsgBox("Last block!" & vbCrLf & "Start from first?", vbYesNo, "Find next empty") = vbYes Then
       i = 0
    Else
      Exit Do
    End If
  End If
  AvaaBlokki
  If AttribFrom.TextString = "" Or AttribTo.TextString = "" Then Exit Do
Loop
End Sub

Private Sub BNext_Click()
  NextBlock
End Sub
Private Sub NextBlock()
i = i + 1
If i > Joukko.Count - 1 Then
  i = 0
End If
AvaaBlokki
End Sub
Private Sub KirjoitaAtt(Teksti As String, oAttrib As AcadAttributeReference)
If oAttrib Is Nothing Then
  Exit Sub
Else
  oAttrib.TextString = Teksti
End If
End Sub
Private Sub BPickFrom_Click()
  OpenSource TFrom, AttribFrom
  AppActivate CurrentDb.Properties("AppTitle").Value
End Sub

Private Sub BPickTo_Click()
  OpenSource TTo, AttribTo
  AppActivate CurrentDb.Properties("AppTitle").Value
End Sub

Private Sub BPrev_Click()
  PrevBlock
End Sub
Private Sub PrevBlock()
i = i - 1
If i < 0 Then
  i = Joukko.Count - 1
End If
AvaaBlokki
End Sub

'Private Sub BWriteFrom_Click()
'  KirjoitaAtt TFrom.Value, AttribFrom
'End Sub

'Private Sub BWriteTo_Click()
'  KirjoitaAtt TTo.Value, AttribTo
'End Sub

Private Sub Form_Close()
  If Not Joukko Is Nothing Then
    Joukko.Delete
    Set Joukko = Nothing
  End If
  Set oDOC = Nothing
  Set oACAD = Nothing
End Sub

Private Sub Form_Load()
Dim FT(0) As Integer
Dim FD(0) As Variant
Dim Poista() As AcadEntity
Dim j As Integer, L As Integer
Dim Nimi As String
 If Yhdista = False Then
   BClose_Click
   Exit Sub
 Else
  If Err = 0 Then
     FT(0) = 0
     FD(0) = "INSERT"
     TeeJoukko "APUVALINTA", Joukko
     Joukko.Select acSelectionSetAll, , , FT, FD
     For j = 0 To Joukko.Count - 1
       Nimi = UCase(Joukko(j).EffectiveName)
       If Nimi <> "PIPELINE" And Nimi <> "PIPELINE_DATA" And Nimi <> "METSO_PIPE" Then
         ReDim Preserve Poista(L)
         Set Poista(L) = Joukko(j)
         L = L + 1
       End If
     Next j
     If L > 0 Then Joukko.RemoveItems Poista
     If Joukko.Count = 0 Then
       MsgBox "No pipeline blocks found."
       Err.Raise vbObjectError + 1000
     Else
       AvaaBlokki
     End If
   End If
 End If
End Sub

Private Sub TeeJoukko(Nimi As String, VJoukko As AcadSelectionSet)
Dim i As Integer
For i = 0 To oDOC.SelectionSets.Count - 1
  If oDOC.SelectionSets(i).Name = Nimi Then
    oDOC.SelectionSets(i).Delete
    Exit For
  End If
Next i
Set VJoukko = ActiveDocument.SelectionSets.Add(Nimi)
End Sub
Private Sub AvaaBlokki()
Dim Attrib As Variant
Dim MinPoint As Variant
Dim MaxPoint As Variant
Dim j As Integer
Me.Caption = i + 1 & "/" & Joukko.Count
If Not oBlock Is Nothing Then
  oBlock.Highlight False
End If
Set oBlock = Joukko(i)
oBlock.GetBoundingBox MinPoint, MaxPoint
oACAD.ZoomWindow MinPoint, MaxPoint
oACAD.ZoomScaled 0.2, acZoomScaledRelative
oBlock.Highlight True
Attrib = oBlock.GetAttributes
LPipeLine.Caption = AttribToPipeline(Attrib)
For j = 0 To UBound(Attrib)
  If Attrib(j).TagString = "FROM" Then
    Set AttribFrom = Attrib(j)
    TFrom.Value = Attrib(j).TextString
  ElseIf Attrib(j).TagString = "TO" Then
    Set AttribTo = Attrib(j)
    TTo.Value = Attrib(j).TextString
  End If
Next j
End Sub
Private Function Yhdista() As Boolean
  On Error Resume Next
  Set oACAD = GetObject(, "AutoCAD.Application") 'Koitetaan yhdist�� AutoCADiin
  If Err <> 0 Then 'K�ynniss� olevaa AutoCADi� ei l�ytynyt
    MsgBox "K�ynniss� olevaa AutoCADi� ei l�ytynyt!", vbCritical, "Virhe!"
    Set oACAD = Nothing
    Yhdista = False
    Exit Function
  Else
    Set oDOC = oACAD.ActiveDocument
    Yhdista = True
  End If
  On Error GoTo 0
End Function

Private Sub TFrom_AfterUpdate()
  KirjoitaAtt Nz(TFrom.Value), AttribFrom
End Sub

Private Sub TTo_AfterUpdate()
  KirjoitaAtt Nz(TTo.Value), AttribTo
End Sub

