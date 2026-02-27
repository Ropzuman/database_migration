VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} FlowPick 
   Caption         =   "PDUFLOWDATA"
   ClientHeight    =   1350
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3435
   OleObjectBlob   =   "FlowPick.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "FlowPick"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Joukko As AcadSelectionSet
Public oAttrib As AcadAttributeReference
Public j As Integer
Public EiLinjaa As Boolean
Sub EtsiLinja()
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
For i = 0 To ActiveDocument.SelectionSets.Count - 1
  If ActiveDocument.SelectionSets(i).Name = "APUPICK2" Then
    ActiveDocument.SelectionSets(i).Delete
    Exit For
  End If
Next i
FilterType(0) = 2
FilterData(0) = "BAHPIPEL"

Set Poiminta = ActiveDocument.SelectionSets.Add("APUPICK2")
Poiminta.Select acSelectionSetCrossing, aPiste, lPiste, FilterType, FilterData
If Poiminta.Count > 0 Then
  Attrib = Poiminta(0).GetAttributes
  RefNo.Caption = Attrib(1).TextString & "-" & Attrib(3).TextString & "-" & Attrib(2).TextString
  oAttrib.TextString = RefNo.Caption
Else
  oAttrib.TextString = ""
  EiLinjaa = True
End If
Poiminta.Delete
End Sub

Private Sub BAutoHaku_Click()
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

Private Sub BEdellinen_Click()
  j = j - 1
  AvaaBlokki
End Sub
Private Sub BSeuraava_Click()
  j = j + 1
  AvaaBlokki
End Sub

Private Sub BPoimi_Click()
Dim Kohde As Object
Dim Piste As Variant
Dim oBlock As AcadBlockReference
Dim Attrib As Variant
Me.Hide
On Error Resume Next
ActiveDocument.Utility.GetEntity Kohde, Piste, "Poimi putkilinja..."
If Err <> 0 Then
  Err.Clear
Else
  If Kohde.ObjectName = "AcDbBlockReference" Then
    Set oBlock = Kohde
    If oBlock.Name = "BAHPIPEL" Then
      Attrib = oBlock.GetAttributes
      RefNo.Caption = Attrib(1).TextString & "-" & Attrib(3).TextString & "-" & Attrib(2).TextString
      oAttrib.TextString = RefNo.Caption
    End If
  End If
End If
Me.Show
On Error GoTo 0
End Sub
Private Sub BSeurT_Click()
Dim i As Integer
  Do
    j = j + 1
    AvaaBlokki
    If RefNo.Caption = "" Then
      Exit Do
    ElseIf j = Joukko.Count - 1 Then
      MsgBox "Sellaisia virtausarvoblokkeja ei löytynyt, joissa putkilinja olisi tyhjä.", vbInformation
      Exit Do
    End If
  Loop
End Sub
Private Sub BSUlje_Click()
  Unload Me
End Sub
Private Sub UserForm_Initialize()
Dim i As Integer
Dim FilterType(0) As Integer
Dim FilterData(0) As Variant
For i = 0 To ActiveDocument.SelectionSets.Count - 1
  If ActiveDocument.SelectionSets(i).Name = "APUPICK" Then
    ActiveDocument.SelectionSets(i).Delete
    Exit For
  End If
Next i
FilterType(0) = 2
FilterData(0) = "PDUFLOWDATA"
Set Joukko = ActiveDocument.SelectionSets.Add("APUPICK")
Joukko.Select acSelectionSetAll, , , FilterType, FilterData
AvaaBlokki
End Sub
Sub AvaaBlokki()
Dim Attrib As Variant
Dim MinPoint As Variant
Dim MaxPoint As Variant
If j >= Joukko.Count Or j < 0 Then j = 0

Me.Caption = "PDUFLOWDATA " & j + 1 & "/" & Joukko.Count
Joukko(j).GetBoundingBox MinPoint, MaxPoint
Application.ZoomWindow MinPoint, MaxPoint
Application.ZoomScaled 0.2, acZoomScaledRelative
Attrib = Joukko(j).GetAttributes
Set oAttrib = Attrib(0)
RefNo.Caption = oAttrib.TextString
End Sub
Private Sub UserForm_Terminate()
  Joukko.Delete
  Set Joukko = Nothing
End Sub
