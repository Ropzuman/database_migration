VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ChangeBlock 
   Caption         =   "ChangeBlock"
   ClientHeight    =   7320
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6465
   OleObjectBlob   =   "ChangeBlock.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "ChangeBlock"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Korkeus As Double
Private Sub Blokki_Change()
Dim i As Integer
Dim j As Integer
BNimi.Caption = Blokki.Value
For i = 1 To 19
  Controls("Atr" & i).Visible = False
  Controls("EQ" & i).Visible = False
  Controls("NAtr" & i).Visible = False
Next i
For i = 0 To ActiveDocument.Blocks(Blokki.Value).Count - 1
  If ActiveDocument.Blocks(Blokki.Value).Item(i).ObjectName = "AcDbAttributeDefinition" Then
    j = j + 1
    If j < 20 Then
      If j > 2 Then
        Korkeus = 77.25 + (j - 2) * 18
      Else
        Korkeus = 77.25
      End If
      Controls("Atr" & j).Caption = ActiveDocument.Blocks(Blokki.Value).Item(i).TagString
      Controls("Atr" & j).Visible = True
      Controls("EQ" & j).Visible = True
      Controls("NAtr" & j).Visible = True
    End If
  End If
Next i

VaihdaKoko False
End Sub
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
Private Sub Nayta_Click()
  VaihdaKoko True
End Sub
Private Sub VaihdaKoko(Vaihda As Boolean)
  If Nayta.Caption = "->" Then
    If Vaihda Then
      Nayta.Caption = "<-"
      Me.Width = 324
      Me.Height = Korkeus '222.75
      VaihdaKaikki.Enabled = True
    Else
      Me.Width = 142.5
      Me.Height = 77.25
    End If
  Else
    If Vaihda Then
      Nayta.Caption = "->"
      Me.Width = 142.5
      Me.Height = 77.25
      VaihdaKaikki.Enabled = False
    Else
      Me.Width = 324
      Me.Height = Korkeus '222.75
    End If
  End If
End Sub
Private Sub UserForm_Initialize()
Dim i As Integer
Me.Width = 142.5
Me.Height = 77.25
Korkeus = 427.75
For i = 0 To ActiveDocument.Blocks.Count - 1
  If Left(ActiveDocument.Blocks(i).Name, 1) <> "*" And InStr(ActiveDocument.Blocks(i).Name, "$") = 0 Then
    Blokki.AddItem ActiveDocument.Blocks(i).Name
    VBlokki.AddItem ActiveDocument.Blocks(i).Name
  End If
Next i
Blokki.ListIndex = 0
VBlokki.ListIndex = 0
End Sub
Private Sub Vaihda_Click()
Dim Valinta As Object
Dim Tyyppi As String
Dim TransMatrix As Variant, ContextData As Variant
Dim IPiste As Variant
Dim VAttribuutit As Variant
Dim UAttribuutit As Variant
Dim UBlokki As AcadBlockReference
Dim UAttrib As AcadAttributeReference
Dim VAttrib As AcadAttributeReference
Dim i As Integer, j As Integer
On Error GoTo Loppu
  If Blokki.Value = "" Then
    Viesti.Caption = "Select block first"
  End If
  Me.Hide
  'Pyydet‰‰n k‰ytt‰j‰‰ valitsemaan Blokki
  ActiveDocument.Utility.GetEntity Valinta, IPiste, "Choose Block..."
  If Err = 0 Then 'Valinta osui objektiin tai ei painettu Esci‰
    Tyyppi = Valinta.ObjectName
    If Tyyppi = "AcDbBlockReference" Then 'Valittiin Blokki
      If UCase(Left(Valinta.Name, 2)) <> "UP" And False Then
        Viesti.Caption = "No UP block selected"
      Else
        'Blokki valittu, vaihdetaan se
        Set UBlokki = ActiveDocument.ModelSpace.InsertBlock(Valinta.InsertionPoint, Blokki.Value, 1, 1, 1, 0)
        UBlokki.Rotation = Valinta.Rotation
        If Valinta.HasAttributes Then
          VAttribuutit = Valinta.GetAttributes
          UAttribuutit = UBlokki.GetAttributes
          For i = 0 To UBound(VAttribuutit)
            If UCase(Valinta.Name) = VBlokki.Value Then 'Jos valittu blokki on sama kuin muutosehdoissa
              MuutaAttribuutit VAttribuutit, UAttribuutit
            Else 'Blokille ei ole muutosehtoja, muutetaan vain nimien mukaan
              For j = 0 To UBound(UAttribuutit)
                If VAttribuutit(i).TagString = UAttribuutit(j).TagString Then
                  Set VAttrib = VAttribuutit(i)
                  Set UAttrib = UAttribuutit(j)
                  UAttrib.TextString = VAttribuutit(i).TextString
                  UAttrib.InsertionPoint = VAttrib.InsertionPoint
'                  UAttrib.InsertionPoint = VAttrib.InsertionPoint(1)
                  UAttrib.Rotation = VAttrib.Rotation
                  Exit For
                End If
              Next j
            End If
          Next i
        End If
        Viesti.Caption = "Block " & Valinta.Name & " changed"
        Valinta.Delete
      End If
    Else
      Viesti.Caption = "No block selected"
    End If
  End If
Loppu:
  If Err <> 0 Then Viesti.Caption = "No changes"
  Set Valinta = Nothing
  Err.Clear
  Me.Show False
End Sub
Private Sub VaihdaKaikki_Click()
Dim Joukko As AcadSelectionSet
Dim UBlokki As AcadBlockReference
Dim FilterType(0) As Integer
Dim FilterData(0) As Variant
Dim VAttrib As Variant
Dim UAttrib As Variant
Dim i As Integer
If VBlokki.Value = "" Then
  MsgBox "Select new block first!", vbCritical, "Change All"
Else
'  FilterType(0) = 66
'  FilterData(0) = 1
  FilterType(0) = 2
  FilterData(0) = VBlokki.Value
  For i = 0 To ActiveDocument.SelectionSets.Count - 1
    If UCase(ActiveDocument.SelectionSets(i).Name) = "CHANGEBLOCK" Then
      ActiveDocument.SelectionSets(i).Delete
      Exit For
    End If
  Next i
  Set Joukko = ActiveDocument.SelectionSets.Add("CHANGEBLOCK")
  Joukko.Select acSelectionSetAll, , , FilterType, FilterData
  For i = 0 To Joukko.Count - 1
    Viesti.Caption = "Changing Block " & i + 1 & "/" & Joukko.Count
    Me.Repaint

'--Muutos 27.10.2004
    Set UBlokki = ActiveDocument.ModelSpace.InsertBlock(Joukko(i).InsertionPoint, Blokki.Value, 1, 1, 1, 0)
    UBlokki.Rotation = Joukko(i).Rotation
    UBlokki.XScaleFactor = Joukko(i).XScaleFactor
    UBlokki.YScaleFactor = Joukko(i).YScaleFactor
    UBlokki.ZScaleFactor = Joukko(i).ZScaleFactor
'---
    If Joukko(i).HasAttributes Then
      VAttrib = Joukko(i).GetAttributes
      If UBlokki.HasAttributes Then
        UAttrib = UBlokki.GetAttributes
        MuutaAttribuutit VAttrib, UAttrib
      End If
    End If
  Next i
  Joukko.Erase
  Joukko.Delete
End If
Set Joukko = Nothing
Set UBlokki = Nothing
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
