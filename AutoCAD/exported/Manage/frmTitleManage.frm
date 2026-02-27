VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmTitleManage 
   Caption         =   "Title Manager"
   ClientHeight    =   4590
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6420
   OleObjectBlob   =   "frmTitleManage.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmTitleManage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim AttDesc As String
Dim AttTag As String
Dim AttText As String
Private Sub Attrib_Click()
Dim Tagi As String
Dim oEntity As AcadEntity
Dim Vastaus As VbMsgBoxResult
  
  AttTag = Attrib.List(Attrib.ListIndex, 0)
  AttText = Attrib.List(Attrib.ListIndex, 1)
  If IsNull(Attrib.Value) Or Attrib.Value = "" Then
  Else
    If UCase(DocLista.Value & ".DWG") = UCase(ActiveDocument.Name) Or IsNull(DocLista.Value) Or DocLista.Value = "" Then
      For Each oEntity In ActiveDocument.Blocks(Blokit.Value)
        If oEntity.ObjectName = "AcDbAttributeDefinition" Then
          If oEntity.TagString = AttTag Then
            AttDesc = oEntity.PromptString
            Exit For
          End If
        End If
      Next
    Else
      oDbx.Open ActiveDocument.Path & "\" & DocLista.Value & ".dwg"
      For Each oEntity In oDbx.Blocks(Blokit.Value)
        If oEntity.ObjectName = "AcDbAttributeDefinition" Then
          If oEntity.TagString = AttTag Then
            AttDesc = oEntity.PromptString
            Exit For
          End If
        End If
      Next
    End If
    AttText = InputBox("Anna uusi tieto." & vbCrLf & "Tag: " & AttTag & vbCrLf & "Prompt: " & AttDesc, "Muuta tieto", AttText)
    If AttText <> "" Then
      Vastaus = MsgBox("Muutetaanko tieto kaikkiin kuviin?", vbYesNoCancel, "Muuta tieto")
      If Vastaus = vbYes Then
        For i = 0 To DocLista.ListCount - 1
          MuutaTieto DocLista.List(i)
        Next i
        Attrib.List(Attrib.ListIndex, 1) = AttText
      ElseIf Vastaus = vbNo Then
        If IsNull(DocLista.Value) Or DocLista.Value = "" Then
          MuutaTieto Tied.Caption
        Else
          MuutaTieto DocLista.Value
        End If
        Attrib.List(Attrib.ListIndex, 1) = AttText
      End If
    End If
  End If
End Sub
Private Sub MuutaTieto(Doku As String)
Dim Joukko As AcadSelectionSet
Dim FilterType(0) As Integer
Dim FilterData(0) As Variant
Dim Blokki As AcadBlockReference
Dim Attribuutit As Variant
Dim Nykyinen As Boolean
Dim i As Integer
  
  FilterType(0) = 2
  FilterData(0) = Blokit.Value
  If UCase(Doku & ".DWG") = UCase(ActiveDocument.Name) Then
    Nykyinen = True
    Set Joukko = ActiveDocument.ActiveSelectionSet
    Joukko.Clear
    Joukko.Select acSelectionSetAll, , , FilterType, FilterData
    If Joukko.Count > 0 Then
      Set Blokki = Joukko(0)
    End If
    Set Joukko = Nothing
  Else
    oDbx.Open ActiveDocument.Path & "\" & Doku & ".dwg"
    For i = 0 To oDbx.ModelSpace.Count - 1
      If oDbx.ModelSpace(i).ObjectName = "AcDbBlockReference" Then
        If UCase(oDbx.ModelSpace(i).Name) = UCase(Blokit.Value) Then
          Set Blokki = oDbx.ModelSpace(i)
          Exit For
        End If
      End If
    Next i
  End If
  If Not Blokki Is Nothing Then
    If Blokki.HasAttributes Then
      Attribuutit = Blokki.GetAttributes
      For i = 0 To UBound(Attribuutit)
        If UCase(Attribuutit(i).TagString) = UCase(AttTag) Then
          Attribuutit(i).TextString = AttText
          Exit For
        End If
      Next i
    End If
  End If
  If Nykyinen Then
    ActiveDocument.Save
  Else
    oDbx.SaveAs ActiveDocument.Path & "\" & Doku & ".dwg"
  End If
End Sub
Private Sub Blokit_Change()
Dim Joukko As AcadSelectionSet
Dim FilterType(0) As Integer
Dim FilterData(0) As Variant
Dim Blokki As AcadBlockReference
Dim Attribuutit As Variant
Dim i As Integer

FilterType(0) = 2
FilterData(0) = Blokit.Value
If Blokit.Value <> "" Then
  If UCase(DocLista.Value & ".DWG") = UCase(ActiveDocument.Name) Or IsNull(DocLista.Value) Or DocLista.Value = "" Then
    Set Joukko = ActiveDocument.ActiveSelectionSet
    Joukko.Clear
    Joukko.Select acSelectionSetAll, , , FilterType, FilterData
    If Joukko.Count > 0 Then
      Set Blokki = Joukko(0)
    End If
    Set Joukko = Nothing
  Else
    oDbx.Open ActiveDocument.Path & "\" & DocLista.Value & ".dwg"
    For i = 0 To oDbx.ModelSpace.Count - 1
      If oDbx.ModelSpace(i).ObjectName = "AcDbBlockReference" Then
        If UCase(oDbx.ModelSpace(i).Name) = UCase(Blokit.Value) Then
          Set Blokki = oDbx.ModelSpace(i)
          Exit For
        End If
      End If
    Next i
  End If
  If Not Blokki Is Nothing Then
    Attrib.Clear
    If Blokki.HasAttributes Then
      Attribuutit = Blokki.GetAttributes
      For i = 0 To UBound(Attribuutit)
        Attrib.AddItem Attribuutit(i).TagString
        Attrib.List(i, 1) = Attribuutit(i).TextString
      Next i
    End If
  End If
End If
End Sub
Private Sub DocLista_Click()
Dim Blokki As AcadBlock
Dim i As Integer
Dim Oletus As String
Dim OK As Boolean
Oletus = Blokit.Value
  
  If UCase(DocLista.Value & ".DWG") = UCase(ActiveDocument.Name) Then
    PaivitaBlokit
  Else
    On Error GoTo Loppu
    Blokit.Clear
    oDbx.Open ActiveDocument.Path & "\" & DocLista.Value & ".dwg"
    For Each Blokki In oDbx.Blocks
      If Left(Blokki.Name, 1) <> "*" Then
        Blokit.AddItem Blokki.Name
        If UCase(Blokki.Name) = UCase(Oletus) Then OK = True
      End If
    Next
    If OK Then
      Blokit.Value = Oletus
    Else
      Blokit.Value = Blokit.List(0)
    End If
    Set Blokki = Nothing
  End If
  Exit Sub
Loppu:
  Err.Clear
  MsgBox "Kuva oli jo auki!", vbCritical
End Sub

Private Sub SBNumero_Change()
  PaivitaLista
End Sub

Private Sub SBRevisio_Change()
  PaivitaLista
End Sub

Private Sub Sulje_Click()
  Unload Me
End Sub

Private Sub PaivitaLista()
Dim Nimi As String
  Numero.Caption = Left(Tied.Caption, SBNumero.Value)
  If SBRevisio.Value = 0 Then
    Revisio.Caption = ""
  Else
    Revisio.Caption = Right(Tied.Caption, SBRevisio.Value)
  End If
  DocLista.Clear
  Nimi = Dir(ActiveDocument.Path & "\" & Numero.Caption & "*" & Revisio.Caption & ".dwg")
  Do While Nimi <> ""
    DocLista.AddItem Left(Nimi, Len(Nimi) - 4)
    Nimi = Dir
  Loop
End Sub
Private Sub UserForm_Initialize()
Dim Osoitin As Long
  Tied.Caption = Left(ActiveDocument.Name, Len(ActiveDocument.Name) - 4)
  SBNumero.Min = 1
  SBNumero.Max = Len(Tied.Caption)
  SBNumero.Value = Len(Tied.Caption)
  SBRevisio.Min = 0
  SBRevisio.Max = Len(Tied.Caption) - 1
  SBRevisio.Value = 0
  Osoitin = InStr(Tied.Caption, "_")
  If Osoitin <> 0 Then
    SBNumero.Value = Osoitin - 1
  End If
  PaivitaLista
  PaivitaBlokit
End Sub

Private Sub UserForm_Terminate()
  Set oDbx = Nothing
End Sub
Private Sub PaivitaBlokit()
Dim Blokki As AcadBlock
Dim Oletus As String
Dim OK As Boolean
Oletus = Blokit.Value
  Blokit.Clear
  For Each Blokki In ActiveDocument.Blocks
    If Left(Blokki.Name, 1) <> "*" Then
      Blokit.AddItem Blokki.Name
      If UCase(Blokki.Name) = UCase(Oletus) Then OK = True
    End If
  Next
  If OK Then
    Blokit.Value = Oletus
  Else
    Blokit.Value = Blokit.List(0)
  End If
End Sub
