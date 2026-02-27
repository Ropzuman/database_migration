VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} PLToLoop 
   Caption         =   "Loop"
   ClientHeight    =   1335
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3750
   OleObjectBlob   =   "PLToLoop.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "PLToLoop"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Joukko As AcadSelectionSet
Public jj As Integer
Public Looppi As String

Private Sub BEdellinen_Click()
  If Joukko Is Nothing Then
    PoimiJoukko
  Else
    Joukko(jj).Highlight False
    jj = jj - 1
    AvaaBlokki
  End If
End Sub
Private Sub BSeuraava_Click()
  If Joukko Is Nothing Then
    PoimiJoukko
  Else
    Joukko(jj).Highlight False
    jj = jj + 1
    AvaaBlokki
  End If
End Sub

Private Sub BPoimi_Click()
Dim Kohde As Object
Dim Piste As Variant
Dim oBlock As AcadBlockReference
Dim Attrib As Variant
Me.Hide
On Error Resume Next
ActiveDocument.Utility.GetEntity Kohde, Piste, "Poimi putkilinja piirille: " & Looppi
If Err <> 0 Then
  Err.Clear
Else
  If Kohde.ObjectName = "AcDbBlockReference" Then
    Set oBlock = Kohde
    If oBlock.Name = "BAHPIPEL" Then
      Attrib = oBlock.GetAttributes
      PLNo.Caption = Attrib(1).TextString & "-" & Attrib(3).TextString & "-" & Attrib(2).TextString & "-" & Attrib(4).TextString & "-" & Attrib(5).TextString
      DataValue(1) = PLNo.Caption
      Joukko(jj).SetXData DataType, DataValue
    End If
  End If
End If
Me.Show vbModeless
On Error GoTo 0
End Sub
Private Sub BSeurT_Click()
Dim i As Integer
Dim Toinen As Boolean
  Toinen = False
  If Joukko Is Nothing Then
    PoimiJoukko
  End If
  Joukko(jj).Highlight False
  Do
    jj = jj + 1
    AvaaBlokki
    If PLNo.Caption = "" Then
      Exit Do
    ElseIf jj = Joukko.Count - 1 Then
      If Toinen = False Then
        If MsgBox("Sellaisia venttiiliblokkeja ei löytynyt, joissa putkilinja olisi tyhjä." & vbCrLf & "Aloitetaanko etsintä alusta?", vbOKCancel, "Etsi seuraava tyhjä") = vbOK Then
          Toinen = True
          jj = -1
        Else
          Exit Do
        End If
      Else
        MsgBox "Sellaisia venttiiliblokkeja ei löytynyt, joissa putkilinja olisi tyhjä.", vbInformation, "Etsi seuraava tyhjä"
        Exit Do
      End If
    End If
  Loop
End Sub
Private Sub BSUlje_Click()
  Unload Me
End Sub
Private Sub PoimiJoukko()
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
FilterData(0) = "UP008,UP009,UP044"
Set Joukko = ActiveDocument.SelectionSets.Add("APUPICK")
Joukko.Select acSelectionSetAll, , , FilterType, FilterData
AvaaBlokki
End Sub
Sub AvaaBlokki()
Dim Attib As Variant
Dim MinPoint As Variant
Dim MaxPoint As Variant
Dim XDataType As Variant
Dim XDataValue As Variant

If jj >= Joukko.Count Or jj < 0 Then jj = 0

Joukko(jj).GetBoundingBox MinPoint, MaxPoint
Application.ZoomWindow MinPoint, MaxPoint
Application.ZoomScaled 0.2, acZoomScaledRelative
Joukko(jj).Highlight True

Joukko(jj).GetXData "PIPELINE", XDataType, XDataValue
Attrib = Joukko(jj).GetAttributes
Looppi = Attrib(1).TextString & "-" & Attrib(3).TextString
Me.Caption = Looppi & "   (" & jj + 1 & "/" & Joukko.Count & ")"

If IsEmpty(XDataValue) Then
  PLNo.Caption = ""
  DataValue(1) = ""
Else
  PLNo.Caption = XDataValue(1)
  DataValue(1) = XDataValue(1)
End If
BPoimi.Enabled = True
'Liikutetaan formia, jotta näyttö päivittyy (korostettu "Highlight" piiri näkyy korostettuna)
Me.Move Me.Left + 1
Me.Move Me.Left - 1
End Sub
Private Sub UserForm_Initialize()
  DataType(0) = 1001
  DataType(1) = 1000
  DataValue(0) = "PIPELINE"
  PoimiJoukko
End Sub

Private Sub UserForm_Terminate()
If Not Joukko Is Nothing Then
  Joukko.Highlight False
  Joukko.Delete
  Set Joukko = Nothing
End If
End Sub
