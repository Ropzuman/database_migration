VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} VaihdaBlock 
   Caption         =   "Change Flowblock"
   ClientHeight    =   480
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4560
   OleObjectBlob   =   "VaihdaBlock.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "VaihdaBlock"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub VBlokki_Change()
Dim i As Integer, j As Integer
For i = 1 To 10
    Controls("NAtr" & i).Clear
    Controls("NAtr" & i).AddItem ""
Next i
For i = 0 To ActiveDocument.Blocks(VBlokki.Value).Count - 1
  If ActiveDocument.Blocks(VBlokki.Value).Item(i).ObjectName = "AcDbAttributeDefinition" Then
    For j = 1 To 10
      Controls("NAtr" & j).AddItem ActiveDocument.Blocks(VBlokki.Value).Item(i).TagString
    Next j
  End If
Next i
For i = 1 To 10
  If Controls("NAtr" & i).Visible = False Then Exit For
  Controls("NAtr" & i).ListIndex = 0
  For j = 0 To Controls("NAtr" & i).ListCount - 1
    If UCase(Controls("Atr" & i).Caption = UCase(Controls("NAtr" & i).List(j))) Then
      Controls("NAtr" & i).ListIndex = j
      Exit For
    End If
  Next j
Next i
End Sub
Private Sub UserForm_Initialize()
Dim i As Integer
Dim OK As Integer
With ActiveDocument.Blocks
  For i = 0 To .Count - 1
    With .Item(i)
      If .Name = "PDUFLOWDATA1" Or .Name = "PDUFLOWMAXDATA1" Or .Name = "PDUFLOWDATA3" Or .Name = "PDUFLOWMAXDATA3" Then
        OK = OK + 1
      End If
    End With
  Next i
End With
If OK < 4 Then
  MsgBox "Insert first blocks:" & vbCrLf & "PDUFLOWDATA1" & vbCrLf & "PDUFLOWMAXDATA1" & vbCrLf & "PDUFLOWDATA3" & vbCrLf & "PDUFLOWMAXDATA3", vbCritical, "Change Flow Blocks"
  Unload Me
End If
End Sub
Private Sub Vaihda_Click()
Dim Valinta As Object
Dim PipeValinta As Object
Dim Tyyppi As String
Dim TransMatrix As Variant, ContextData As Variant
Dim IPiste As Variant
Dim VAttribuutit As Variant
Dim UAttribuutit As Variant
Dim Attribuutit As Variant
Dim UBlokki As AcadBlockReference
Dim UAttrib As AcadAttributeReference
Dim VAttrib As AcadAttributeReference
Dim i As Integer, j As Integer
Dim VanhaBlokki As String
Dim VanhaBlokki2 As String
Dim VanhaMaxBlokki As String
Dim VanhaMaxBlokki2 As String
Dim UusiBlokki As String
Dim UusiBlokki2 As String
Dim UusiMaxBlokki As String
Dim UusiMaxBlokki2 As String
Dim PipelineBlock As String
Dim PipelineAttrib As String
Dim PipelineNo As String
  
  On Error Resume Next
  VanhaBlokki = "PDUFLOWDATA"
  VanhaMaxBlokki = "PDUFLOWMAXDATA"
  VanhaBlokki2 = "PDUFLOWDATA2"
  VanhaMaxBlokki2 = "PDUFLOWMAXDATA2"
  UusiBlokki = "PDUFLOWDATA1"
  UusiMaxBlokki = "PDUFLOWMAXDATA1"
  UusiBlokki2 = "PDUFLOWDATA3"
  UusiMaxBlokki2 = "PDUFLOWMAXDATA3"
  PipelineBlock = "ARAPIPEL"
  PipelineAttrib = "LINE"
  
  
  Viesti.Caption = ""
  Me.Hide
  'Pyydetään käyttäjää valitsemaan Blokki
  ActiveDocument.Utility.GetEntity Valinta, IPiste, "Select flowblock..."
  If Err = 0 Then 'Valinta osui objektiin tai ei painettu Esciä
    Tyyppi = Valinta.ObjectName
    If Tyyppi = "AcDbBlockReference" Then 'Valittiin Blokki
      Select Case UCase(Valinta.Name)
        Case VanhaBlokki
          Set UBlokki = ActiveDocument.ModelSpace.InsertBlock(Valinta.InsertionPoint, UusiBlokki, 1, 1, 1, 0)
        Case VanhaMaxBlokki
          Set UBlokki = ActiveDocument.ModelSpace.InsertBlock(Valinta.InsertionPoint, UusiMaxBlokki, 1, 1, 1, 0)
        Case VanhaBlokki2
          Set UBlokki = ActiveDocument.ModelSpace.InsertBlock(Valinta.InsertionPoint, UusiBlokki2, 1, 1, 1, 0)
        Case VanhaMaxBlokki2
          Set UBlokki = ActiveDocument.ModelSpace.InsertBlock(Valinta.InsertionPoint, UusiMaxBlokki, 1, 1, 1, 0)
        Case Else
          MsgBox "Selected block (" & Valinta.Name & ") can't be changed", vbCritical, "Change flow block"
          GoTo Loppu
      End Select
      UBlokki.Rotation = Valinta.Rotation
'Uusi blokki on insertoitu
      ActiveDocument.Utility.Prompt vbCrLf & "Block " & Valinta.Name & " selected." & vbCrLf
'Valitaan putkilinjanumero
      ActiveDocument.Utility.GetEntity PipeValinta, IPiste, "Select Pipelineblock..."
      If Err = 0 Then
        If PipeValinta.ObjectName = "AcDbBlockReference" Then 'Valittiin Blokki
          If PipeValinta.HasAttributes Then
            Attribuutit = PipeValinta.GetAttributes
            For i = 0 To UBound(Attribuutit)
              If UCase(Attribuutit(i).TagString) = PipelineAttrib Then
                PipelineNo = Attribuutit(i).TextString
                ActiveDocument.Utility.Prompt vbCrLf & "Selected pipeline number " & PipelineNo & vbCrLf
                Exit For
              End If
            Next i
          End If
        End If
      End If
      Err.Clear
      VAttribuutit = Valinta.GetAttributes
      UAttribuutit = UBlokki.GetAttributes
      For i = 0 To UBound(VAttribuutit)
        For j = 0 To UBound(UAttribuutit)
          If UAttribuutit(j).TagString = "PIPELINENO" Then
            UAttribuutit(j).TextString = PipelineNo
          End If
          If VAttribuutit(i).TagString = UAttribuutit(j).TagString Then
            Set VAttrib = VAttribuutit(i)
            Set UAttrib = UAttribuutit(j)
            UAttrib.TextString = VAttribuutit(i).TextString
'            UAttrib.InsertionPoint(0) = VAttrib.InsertionPoint(0)
'            UAttrib.InsertionPoint(1) = VAttrib.InsertionPoint(1)
'            UAttrib.Rotation = VAttrib.Rotation
          End If
        Next j
      Next i
      Valinta.Delete
      Viesti.Caption = "Block changed."
    Else
      Viesti.Caption = "Not a block selected!"
    End If
  Else
    Viesti.Caption = "No changes made"
  End If
Loppu:
  Err.Clear
  On Error GoTo 0
  Set Valinta = Nothing
  Me.Show False
'      If UCase(Left(Valinta.Name, 2)) <> "UP" Then
'        Viesti.Caption = "No UP block selected"
'      Else
'        'Blokki valittu, vaihdetaan se
'        Set UBlokki = ActiveDocument.ModelSpace.InsertBlock(Valinta.InsertionPoint, Blokki.Value, 1, 1, 1, 0)
'        UBlokki.Rotation = Valinta.Rotation
'        If Valinta.HasAttributes Then
'          VAttribuutit = Valinta.GetAttributes
'          UAttribuutit = UBlokki.GetAttributes
'          For i = 0 To UBound(VAttribuutit)
'            If UCase(Valinta.Name) = VBlokki.Value Then 'Jos valittu blokki on sama kuin muutosehdoissa
'              MuutaAttribuutit VAttribuutit, UAttribuutit
'            Else 'Blokille ei ole muutosehtoja, muutetaan vain nimien mukaan
'              For j = 0 To UBound(UAttribuutit)
'                If VAttribuutit(i).TagString = UAttribuutit(j).TagString Then
'                  Set VAttrib = VAttribuutit(i)
'                  Set UAttrib = UAttribuutit(j)
'                  UAttrib.TextString = VAttribuutit(i).TextString
'                  UAttrib.InsertionPoint(0) = VAttrib.InsertionPoint(0)
'                  UAttrib.InsertionPoint(1) = VAttrib.InsertionPoint(1)
'                  UAttrib.Rotation = VAttrib.Rotation
'                  Exit For
'                End If
'              Next j
'            End If
'          Next i
'        End If
'        Viesti.Caption = "Block " & Valinta.Name & " changed"
'        Valinta.Delete
'      End If
'    Else
'      Viesti.Caption = "No block selected"
'    End If
'  End If
'Loppu:
'  If Err <> 0 Then Viesti.Caption = "No changes"
'  Set Valinta = Nothing
'  Err.Clear
'  Me.Show False
End Sub
Private Sub MuutaAttribuutit(Vanhat As Variant, Uudet As Variant)
Dim i As Integer, j As Integer, a As Integer
For i = 0 To UBound(Uudet)
  For j = 1 To 10
    If Controls("Atr" & j).Visible = False Then Exit For
    If UCase(Controls("Atr" & j).Caption) = UCase(Uudet(i).TagString) Then
      If Controls("NAtr" & j).Value <> "" Then
        For a = 0 To UBound(Vanhat)
          If UCase(Vanhat(a).TagString) = UCase(Controls("NAtr" & j).Value) Then
            Uudet(i).TextString = Vanhat(a).TextString
            Exit For
          End If
        Next a
      End If
      Exit For
    End If
  Next j
Next i
End Sub
