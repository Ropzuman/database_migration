VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_USysFlowPickNo_OLD"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Dim oACAD As AcadApplication
Dim oDOC As AcadDocument
Dim oAttrib As AcadAttributeReference
Dim Joukko As AcadSelectionSet
Dim j As Integer
Dim EiLinjaa As Boolean

Private Sub Form_Close()
  If Not Joukko Is Nothing Then
    Joukko.Delete
    Set Joukko = Nothing
  End If
End Sub

Private Sub Form_Load()
Dim i As Integer
Dim FilterType(0) As Integer
Dim FilterData(0) As Variant
If Yhdista = False Then
  DoCmd.Close acForm
  Exit Sub
End If
For i = 0 To oDOC.SelectionSets.Count - 1
  If oDOC.SelectionSets(i).Name = "APUPICK" Then
    oDOC.SelectionSets(i).Delete
    Exit For
  End If
Next i
FilterType(0) = 2
FilterData(0) = "PDUFLOWDATA,ARAFLOW,FLOWBLOCK,BALANCE_DATA"
Set Joukko = oDOC.SelectionSets.Add("APUPICK")
Joukko.Select acSelectionSetAll, , , FilterType, FilterData
If Joukko.Count = 0 Then
  MsgBox "Ei tasearvoblokkeja"
  DoCmd.Close
  Exit Sub
End If
AvaaBlokki
End Sub


Private Sub BAutohaku_Click()
EiLinjaa = False
For j = 0 To Joukko.Count - 1
  AvaaBlokki
  EtsiLinja
Next j
If EiLinjaa Then
  MsgBox "Kaikki linjat on käyty läpi. Yksi tai useampia virtausarvoblokkeja jäi ilman putkilinjanumeroa.", vbOKOnly
  BSeurT_Click
Else
  MsgBox "Kaikki linjat on käyty läpi. Kaikkiin virtausarvoblokkeihin löytyi putkilinjanumero.", vbOKOnly
End If
End Sub
Private Sub BPoimi_Click()
Dim Kohde As Object
Dim Piste As Variant
Dim oBlock As AcadBlockReference
Dim Attrib As Variant
AppActivate oACAD.Caption, True
On Error Resume Next
oDOC.Utility.GetEntity Kohde, Piste, "Poimi putkilinja..."
If Err <> 0 Then
  Err.Clear
Else
  If Kohde.ObjectName = "AcDbBlockReference" Then
    Set oBlock = Kohde
    Select Case LCase(oBlock.EffectiveName)
      Case "pipeline", "pipeline_f", "arapipel", "bahpipel", "metso_pipe"
         Attrib = oBlock.GetAttributes
         TRefNo.Value = AttribToPipeline(Attrib)
      Case Else
        If oBlock.HasAttributes Then
          Attrib = oBlock.GetAttributes
          For j = 0 To UBound(Attrib)
            If Attrib(j).TagString = "POS.NO" Or Attrib(j).TagString = "CUSTPOS" Or Attrib(j).TagString = "SDPOS" Then
              TRefNo.Value = Attrib(j).TextString
              Exit For
            End If
          Next j
        End If
    End Select
    oAttrib.TextString = TRefNo.Value
  End If
End If
AppActivate CurrentDb.Properties("AppTitle").Value, True
On Error GoTo 0
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
    Case "FSB"
      FSB = Attrib(j).TextString
    Case "LINE"
      LINE = Attrib(j).TextString
    Case Else
    
  End Select
Next j
AttribToPipeline = DEP & FSB & LINE
'AttribToPipeline = DEP & "-" & FSB & "-" & LINE
End Function
Private Sub BSeurT_Click()
Dim i As Integer
  Do
    If j < Joukko.Count Then Joukko(j).Highlight False
    j = j + 1
    AvaaBlokki
    If TRefNo.Value = "" Then
      Exit Do
    ElseIf j = Joukko.Count - 1 Then
      MsgBox "Sellaisia virtausarvoblokkeja ei löytynyt, joissa putkilinja olisi tyhjä.", vbInformation
      Exit Do
    End If
  Loop
End Sub
Private Sub BEdellinen_Click()
  Joukko(j).Highlight False
  If j > 0 Then
    j = j - 1
  Else
    j = Joukko.Count - 1
  End If
  AvaaBlokki
End Sub

Private Sub BSeuraava_Click()
  Joukko(j).Highlight False
  If j < Joukko.Count Then
    j = j + 1
  Else
    j = 0
  End If
  AvaaBlokki
End Sub

Private Sub BSulje_Click()
  DoCmd.Close acForm, "USysFlowPickNo"
End Sub

Sub AvaaBlokki()
Dim Attrib As Variant
Dim MinPoint As Variant
Dim MaxPoint As Variant
Dim Loytyi As Boolean
Dim ii As Integer
If j >= Joukko.Count Or j < 0 Then j = 0
Joukko(j).Highlight True

Me.Caption = Joukko(j).Name & ":  " & j + 1 & "/" & Joukko.Count
Joukko(j).GetBoundingBox MinPoint, MaxPoint
oACAD.ZoomCenter Joukko(j).InsertionPoint, 200
'oACAD.ZoomScaled 10, acZoomScaledAbsolute
'oACAD.ZoomWindow MinPoint, MaxPoint
'oACAD.ZoomScaled 0.2, acZoomScaledRelative
Attrib = Joukko(j).GetAttributes
Loytyi = False
For ii = 0 To UBound(Attrib)
  If UCase(Attrib(ii).TagString) = "PIPEREF" Or UCase(Attrib(ii).TagString) = "PIPELINENO" Or UCase(Attrib(ii).TagString) = "PIPE_LINE" Then
    Set oAttrib = Attrib(ii)
    Loytyi = True
    Exit For
  End If
Next ii
If Loytyi = False Then
  MsgBox "Tasearvoblokki ei sisältänyt attribuuttia putkilinjanumerolle.", vbCritical, "FlowPicNo"
  DoCmd.Close
  Exit Sub
End If
TRefNo.Value = oAttrib.TextString
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
Private Sub EtsiLinja()
Dim Poiminta As AcadSelectionSet
Dim i As Integer
Dim aPiste(2) As Double
Dim lPiste(2) As Double
Dim Piste As Variant
Dim FilterType(0) As Integer
Dim FilterData(0) As Variant
Dim Attrib As Variant

Piste = Joukko(j).InsertionPoint
aPiste(0) = Piste(0)
aPiste(1) = Piste(1)
lPiste(0) = aPiste(0) + 56
lPiste(1) = aPiste(1) + 13
For i = 0 To oDOC.SelectionSets.Count - 1
  If oDOC.SelectionSets(i).Name = "APUPICK2" Then
    oDOC.SelectionSets(i).Delete
    Exit For
  End If
Next i
FilterType(0) = 2
FilterData(0) = "PIPELINE,BAHPIPEL,ARAPIPEL"

Set Poiminta = oDOC.SelectionSets.Add("APUPICK2")
Poiminta.Select acSelectionSetCrossing, aPiste, lPiste, FilterType, FilterData
If Poiminta.Count > 0 Then
  Attrib = Poiminta(0).GetAttributes
  TRefNo.Value = AttribToPipeline(Attrib)
  oAttrib.TextString = TRefNo.Value
Else
  oAttrib.TextString = ""
  EiLinjaa = True
End If
Poiminta.Delete
End Sub

Private Sub TRefNo_AfterUpdate()
If Not oAttrib Is Nothing Then
  oAttrib.TextString = Nz(TRefNo.Value)
End If

End Sub
