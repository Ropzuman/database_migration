VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_USysPipeToOther"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit
Dim oACAD As AcadApplication
Dim oDOC As AcadDocument
Dim PipeAttrib As AcadAttributeReference
Dim oBlock As AcadBlockReference
Dim Joukko As AcadSelectionSet
Dim i As Integer
Private Sub BClose_Click()
  DoCmd.Close acForm, "UsysPipeToOther"
End Sub
Private Sub OpenPipeline()
Dim Kohde As Object
Dim Piste As Variant
Dim oPipeBlock As AcadBlockReference
Dim Attrib As Variant
Dim Putkilinja As String
Dim OK As Boolean
Dim j As Integer
Dim Nimi As String

AppActivate oACAD.Caption, True
On Error Resume Next
OK = False
Do
  oDOC.Utility.GetEntity Kohde, Piste, "Pick pipeline: "
  If Err <> 0 Then
    Err.Clear
    Exit Do
  Else
    If Kohde.ObjectName = "AcDbBlockReference" Then
      Set oPipeBlock = Kohde
      Nimi = UCase(oPipeBlock.EffectiveName)
      If Nimi = "PIPELINE_F" Or Nimi = "PIPELINE" Or Nimi = "PIPELINE_DATA" Or Nimi = "METSO_PIPE" Or Nimi = "BAHPIPEL" Then
        Attrib = oPipeBlock.GetAttributes
        Putkilinja = AttribToPipeline(Attrib)
        TPipeline.Value = Putkilinja
        PipeAttrib.TextString = Putkilinja
        OK = True
      ElseIf Left(Nimi, 6) = "_PLATE" Then
        Attrib = oPipeBlock.GetAttributes
        TPipeline.Value = Attrib(1).TextString
        PipeAttrib.TextString = Attrib(1).TextString
        OK = True
      End If
    End If
  End If
  If OK Then
    If i < Joukko.Count - 1 Then
      NextBlock
    Else
      Exit Do
    End If
  End If
Loop
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
  If TPipeline.Value = "" Then Exit Do
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

Private Sub BPickPipeline_Click()
  OpenPipeline
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
     TeeJoukko "APUVALINTA", Joukko
     If InStr(OpenArgs, ",") Then
        FT(0) = 2
        FD(0) = OpenArgs
        Joukko.Select acSelectionSetAll, , , FT, FD
     Else
        FT(0) = 0
        FD(0) = "INSERT"
        Joukko.Select acSelectionSetAll, , , FT, FD
        For j = 0 To Joukko.Count - 1
          Nimi = UCase(Joukko(j).EffectiveName)
          If Nimi <> UCase(OpenArgs) Then
            ReDim Preserve Poista(L)
            Set Poista(L) = Joukko(j)
            L = L + 1
          End If
        Next j
        If L > 0 Then Joukko.RemoveItems Poista
     End If
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
For i = 0 To ActiveDocument.SelectionSets.Count - 1
  If ActiveDocument.SelectionSets(i).Name = Nimi Then
    ActiveDocument.SelectionSets(i).Delete
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
Set PipeAttrib = Nothing
For j = 0 To UBound(Attrib)
  If Attrib(j).TagString = "PIPELINE" Or Attrib(j).TagString = "PIPELINENO" Then
    Set PipeAttrib = Attrib(j)
    TPipeline.Value = PipeAttrib.TextString
  End If
Next j
If PipeAttrib Is Nothing Then
  MsgBox "Blokista ei löytynyt 'PIPELINE'-attribuuttia.", vbCritical, "VIRHE"
End If
End Sub
Private Function Yhdista() As Boolean
  On Error Resume Next
  Set oACAD = GetObject(, "AutoCAD.Application") 'Koitetaan yhdistää AutoCADiin
  If Err <> 0 Then 'Käynnissä olevaa AutoCADiä ei löytynyt
    MsgBox "Käynnissä olevaa AutoCADiä ei löytynyt!", vbCritical, "Virhe!"
    Set oACAD = Nothing
    Yhdista = False
    Exit Function
  Else
    Set oDOC = oACAD.ActiveDocument
    Yhdista = True
  End If
  On Error GoTo 0
End Function

Private Sub TPipeline_AfterUpdate()
  KirjoitaAtt Nz(TPipeline.Value), PipeAttrib
End Sub

