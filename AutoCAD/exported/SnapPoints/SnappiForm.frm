VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} SnappiForm 
   Caption         =   "Move to Snap"
   ClientHeight    =   2025
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4710
   OleObjectBlob   =   "SnappiForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "SnappiForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub OK_Click()
Dim oEntity As AcadEntity
Dim Joukko As AcadSelectionSet

Set Joukko = ActiveDocument.ActiveSelectionSet
ActiveDocument.SendCommand "SNAP " & VSnap.Value & " "
ActiveDocument.SendCommand "SNAPMODE 1 "
For Each oEntity In Joukko
  Tarkista oEntity
Next
Joukko.Delete
Set Joukko = Nothing
Set oEntity = Nothing
Sulje
End Sub
Private Sub Peruuta_Click()
  Sulje
End Sub
Private Sub Tarkista(oEntity As Object)
Dim oViiva As AcadLine
Dim oBlock As AcadBlockReference
Dim oText As AcadText
Dim oPLine As AcadPolyline
  If oEntity.ObjectName = "AcDbLine" Then
    Set oViiva = oEntity
    oViiva.StartPoint = Snappi(oViiva.StartPoint)
    oViiva.EndPoint = Snappi(oViiva.EndPoint)
  ElseIf oEntity.ObjectName = "AcDbText" Then
    Set oText = oEntity
    If oText.Alignment = 0 Then
      oText.InsertionPoint = Snappi(oText.InsertionPoint)
    Else
      oText.TextAlignmentPoint = Snappi(oText.TextAlignmentPoint)
    End If
  ElseIf oEntity.ObjectName = "AcDbBlockReference" Then
    Set oBlock = oEntity
    oBlock.InsertionPoint = Snappi(oBlock.InsertionPoint)
  ElseIf oEntity.ObjectName = "AcDb2dPolyline" Then
    Set oPLine = oEntity
    oPLine.Coordinates = Snappi(oPLine.Coordinates)
  ElseIf oEntity.ObjectName = "AcDbPolyline" Then
    oEntity.Coordinates = Snappi(oEntity.Coordinates)
  Else
  End If
  Set oViiva = Nothing
  Set oBlock = Nothing
  Set oText = Nothing
  Set oPLine = Nothing
End Sub
Private Function Snappi(Piste As Variant) As Variant
Dim iPoint() As Double
Dim i As Integer
ReDim iPoint(UBound(Piste))
  For i = 0 To UBound(Piste)
    If (Piste(i) * 100) Mod 125 = 0 Then
      iPoint(i) = Val(Piste(i) * 100) / 100
    ElseIf (Piste(i) * 1000) Mod 1250 > 625 Then
      iPoint(i) = (Val(Piste(i) * 100) + 125 - (Val(Piste(i) * 100) Mod 125)) / 100
    Else
      iPoint(i) = (Val(Piste(i) * 100) - (Val(Piste(i) * 100) Mod 125)) / 100
    End If
  Next i
  Snappi = iPoint
End Function
Private Sub UserForm_Initialize()
Dim KSnappiX As Double
Dim KSnappiY As Double
Dim KSNappi As String
  ActiveDocument.ActiveViewport.GetSnapSpacing KSnappiX, KSnappiY
  KSNappi = Replace(CStr(KSnappiX), ",", ".")
  Select Case KSNappi
    Case "1", "1.25", "2", "2.5", "5", "10", "15"
    Case Else
      VSnap.AddItem KSNappi
  End Select
  VSnap.AddItem "1"
  VSnap.AddItem "1.25"
  VSnap.AddItem "2"
  VSnap.AddItem "2.5"
  VSnap.AddItem "5"
  VSnap.AddItem "10"
  VSnap.AddItem "15"
  VSnap.Value = KSNappi
End Sub
