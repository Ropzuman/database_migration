VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} LineForm 
   Caption         =   "Line Properties"
   ClientHeight    =   2325
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   2865
   OleObjectBlob   =   "LineForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "LineForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim EkaViiva As AcadLine
Dim ViimViiva As AcadLine
Dim Valmis As Boolean
Const Pi = 3.14159265358979
Private Sub NoSymbol_Click()
  Valmis = False
  LineForm.Hide
  If LinkkiFormi Then
    LinkForm.Show
  Else
    MainForm.Show
  End If
End Sub
Private Sub OK_Click() 'Piirret‰‰n nuoli tai muu symboli viivan p‰‰h‰n
Dim Lyhennys As Double
Dim iPoint(2) As Double
Dim LoppuPiste(2) As Double
Dim Rotation As Double
If ARROW Then 'Nuoli
  BName = BLOCKPATH & "ARROW.DWG"
  Lyhennys = 7.5
ElseIf NOTARROW Then 'Nuli ja Not merkki
  BName = BLOCKPATH & "NOTARROW.DWG"
  Lyhennys = 7.5
Else 'Piste
  BName = BLOCKPATH & "POINT.DWG"
  Lyhennys = 0
End If
If LStart Then
  With EkaViiva
    iPoint(0) = .StartPoint(0)
    iPoint(1) = .StartPoint(1)
    Rotation = .Angle + Pi
    LoppuPiste(0) = .StartPoint(0) + Lyhennys * Cos(Rotation - Pi)
    LoppuPiste(1) = .StartPoint(1) + Lyhennys * Sin(Rotation - Pi)
    .StartPoint = LoppuPiste
  End With
Else
  With ViimViiva
    iPoint(0) = .EndPoint(0)
    iPoint(1) = .EndPoint(1)
    Rotation = .Angle
    LoppuPiste(0) = .EndPoint(0) - Lyhennys * Cos(Rotation)
    LoppuPiste(1) = .EndPoint(1) - Lyhennys * Sin(Rotation)
    .EndPoint = LoppuPiste
  End With
End If
ThisDrawing.ModelSpace.InsertBlock iPoint, BName, 1, 1, 1, Rotation
Valmis = False
Set EkaViiva = Nothing
Set ViimViiva = Nothing
LineForm.Hide
  If LinkkiFormi Then
    LinkForm.Show
  Else
    MainForm.Show
  End If
End Sub
Private Sub UserForm_Activate()
Dim iPoint As Variant
Dim sPoint(2) As Double
Dim i As Integer
If Valmis = False Then
  On Error Resume Next
  LineForm.Hide
  ThisDrawing.ActiveLayer = ThisDrawing.Layers("0")
  For i = 0 To ThisDrawing.Layers.Count - 1
    If ThisDrawing.Layers(i).Name = "CANVAS" Then
      ThisDrawing.ActiveLayer = ThisDrawing.Layers(i)
      Exit For
    End If
  Next i
  iPoint = ThisDrawing.Utility.GetPoint(, "Set start point")
  If Err = 0 Then
    sPoint(0) = iPoint(0)
    sPoint(1) = iPoint(1)
    Do
      iPoint = ThisDrawing.Utility.GetPoint(sPoint, "Set end point or end drawing")
      If Err <> 0 Then
        If ViimViiva Is Nothing Then 'Lopetettiin ennen kuin yht‰k‰‰n viiva oli piirretty
          If LinkkiFormi Then
            LinkForm.Show
          Else
            MainForm.Show
          End If
        Else
          Valmis = True
          LineForm.Show
        End If
        Exit Do
      Else
        Set ViimViiva = ThisDrawing.ModelSpace.AddLine(sPoint, iPoint)
        If EkaViiva Is Nothing Then
          Set EkaViiva = ViimViiva
        End If
        sPoint(0) = iPoint(0)
        sPoint(1) = iPoint(1)
      End If
    Loop
  Else 'Keskeytettiin heti
    If LinkkiFormi Then
      LinkForm.Show
    Else
      MainForm.Show
    End If
  End If
  Err.Clear
End If
End Sub
