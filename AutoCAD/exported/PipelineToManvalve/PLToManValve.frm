VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} PLToManValve 
   Caption         =   "ManValve"
   ClientHeight    =   1335
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3750
   OleObjectBlob   =   "PLToManValve.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "PLToManValve"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Joukko As AcadSelectionSet
Public oAttrib As AcadAttributeReference
Public jj As Integer
Public Venttiili As String

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
ActiveDocument.Utility.GetEntity Kohde, Piste, "Poimi putkilinja ventiilille: " & Venttiili
If Err <> 0 Then
  Err.Clear
Else
  If Kohde.ObjectName = "AcDbBlockReference" Then
    Set oBlock = Kohde
    If oBlock.Name = "BAHPIPEL" Then
      Attrib = oBlock.GetAttributes
      PLNo.Caption = Attrib(1).TextString & "-" & Attrib(3).TextString & "-" & Attrib(2).TextString & "-" & Attrib(4).TextString & "-" & Attrib(5).TextString
      oAttrib.TextString = PLNo.Caption
    End If
  End If
End If
Me.Show
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
FilterData(0) = "UP050_B,UP051_B,UP052_B,UP053_B,UP054_B"
Set Joukko = ActiveDocument.SelectionSets.Add("APUPICK")
Joukko.Select acSelectionSetAll, , , FilterType, FilterData
AvaaBlokki
End Sub
Sub AvaaBlokki()
Dim Attrib As Variant
Dim MinPoint As Variant
Dim MaxPoint As Variant

If jj >= Joukko.Count Or jj < 0 Then jj = 0

Joukko(jj).GetBoundingBox MinPoint, MaxPoint
Application.ZoomWindow MinPoint, MaxPoint
Application.ZoomScaled 0.2, acZoomScaledRelative
Joukko(jj).Highlight True
Attrib = Joukko(jj).GetAttributes
Venttiili = Attrib(0).TextString & "-" & Attrib(1).TextString
Me.Caption = Venttiili & " (" & jj + 1 & "/" & Joukko.Count & ")"
Set oAttrib = Attrib(4)
PLNo.Caption = oAttrib.TextString
BPoimi.Enabled = True
Me.Move Me.Left + 1
End Sub
Private Sub CBlocks_Click()
  VaihdaBlokit "UP050", "K:\PROJECTS\L_Projekti\BahiaSul\ProcesDiag\Flowsheets\UP050_B.dwg"
  VaihdaBlokit "UP051", "K:\PROJECTS\L_Projekti\BahiaSul\ProcesDiag\Flowsheets\UP051_B.dwg"
  VaihdaBlokit "UP052", "K:\PROJECTS\L_Projekti\BahiaSul\ProcesDiag\Flowsheets\UP052_B.dwg"
  VaihdaBlokit "UP053", "K:\PROJECTS\L_Projekti\BahiaSul\ProcesDiag\Flowsheets\UP053_B.dwg"
  VaihdaBlokit "UP054", "K:\PROJECTS\L_Projekti\BahiaSul\ProcesDiag\Flowsheets\UP054_B.dwg"
  PoimiJoukko
End Sub
Private Sub VaihdaBlokit(VBlock As String, UBlock As String)
Dim Blokit As AcadSelectionSet
Dim UBlokki As AcadBlockReference
Dim FilterType(0) As Integer
Dim FilterData(0) As Variant
Dim VAttrib As Variant
Dim UAttrib As Variant
Dim UAtt As AcadAttributeReference
Dim VAtt As AcadAttributeReference
Dim i As Integer, ii As Integer, j As Integer
  
  Set Joukko = Nothing
  FilterType(0) = 2
  FilterData(0) = VBlock
  For i = 0 To ActiveDocument.SelectionSets.Count - 1
    If UCase(ActiveDocument.SelectionSets(i).Name) = "CHANGEBLOCK" Then
      ActiveDocument.SelectionSets(i).Delete
      Exit For
    End If
  Next i
  Set Blokit = ActiveDocument.SelectionSets.Add("CHANGEBLOCK")
  Blokit.Select acSelectionSetAll, , , FilterType, FilterData
  For i = 0 To Blokit.Count - 1
    Me.Caption = "Changing Block " & i + 1 & "/" & Blokit.Count
    VAttrib = Blokit(i).GetAttributes
    Set UBlokki = ActiveDocument.ModelSpace.InsertBlock(Blokit(i).InsertionPoint, UBlock, Blokit(i).XScaleFactor, Blokit(i).YScaleFactor, Blokit(i).ZScaleFactor, Blokit(i).Rotation)
    UBlokki.Rotation = Blokit(i).Rotation
    UAttrib = UBlokki.GetAttributes
          
    For ii = 0 To UBound(VAttrib)
      For j = 0 To UBound(UAttrib)
        If VAttrib(ii).TagString = UAttrib(j).TagString Then
          Set VAtt = VAttrib(ii)
          Set UAtt = UAttrib(j)
          UAtt.TextString = VAtt.TextString
          UAtt.InsertionPoint = VAtt.InsertionPoint
          UAtt.Rotation = VAtt.Rotation
          Exit For
        End If
      Next j
    Next ii
  Next i
  Blokit.Erase
  Blokit.Delete
Set VAtt = Nothing
Set UAtt = Nothing
Set Blokit = Nothing
Set UBlokki = Nothing
End Sub

Private Sub UserForm_Terminate()
If Not Joukko Is Nothing Then
  Joukko.Highlight False
  Joukko.Delete
  Set Joukko = Nothing
End If
End Sub
