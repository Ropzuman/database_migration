VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} FrmMakeBlocks 
   Caption         =   "Tee blokkeja"
   ClientHeight    =   480
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5655
   OleObjectBlob   =   "FrmMakeBlocks.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "FrmMakeBlocks"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CommandButton1_Click()
Dim Blokki As String
Dim Joukko As AcadSelectionSet
Dim FilterType(0) As Integer
Dim FilterData(0) As Variant
Dim Jarj() As Integer
Dim Luvut() As Double
Dim Ltmp As Double
Dim JarjTmp As Double
Dim OK As Boolean
Dim Piste As Variant
Dim oBlock As AcadBlockReference
Dim i As Integer, j As Integer
Dim Attrib As Variant

  Me.Hide
  Blokki = "P:\24PRO229 Fortum Nuijala\R\Flowsheet\koe.dwg"
  FilterType(0) = 0
  FilterData(0) = "TEXT"
  ActiveDocument.ActiveLayer = ActiveDocument.Layers("0")
  For i = 0 To ActiveDocument.SelectionSets.Count - 1
    If ActiveDocument.SelectionSets(i).Name = "MAKEAPU" Then
      ActiveDocument.SelectionSets(i).Delete
      Exit For
    End If
  Next i
  Set Joukko = ActiveDocument.SelectionSets.Add("MAKEAPU")
  Joukko.Clear
  Joukko.SelectOnScreen FilterType, FilterData
  If Joukko.Count = 0 Then
    MsgBox "Et valinnut yhtään tekstiä!", vbCritical, "Poimi tekstit"
  ElseIf Joukko.Count > 5 Then
    MsgBox "Valitsit enemmän kuin 5 tekstiä. Yritä uudelleen.", vbCritical, "Poimi tekstit"
  Else
    Set oBlock = ActiveDocument.ModelSpace.InsertBlock(Joukko(0).InsertionPoint, Blokki, 1, 1, 1, 0)
    Attrib = oBlock.GetAttributes
    
'Otetaan joukon Y kordinaatit ylös array muuttujaan
    ReDim Luvut(0 To Joukko.Count - 1)
    ReDim Jarj(0 To Joukko.Count - 1)
    For i = 0 To Joukko.Count - 1
      Piste = Joukko(i).InsertionPoint
      Luvut(i) = Piste(1)
      Jarj(i) = i
    Next i
    
'Järjestetään array nousevaan järjestykseen
    Do
      OK = True
      For i = 1 To UBound(Luvut)
        If Luvut(i - 1) < Luvut(i) Then
          Ltmp = Luvut(i - 1)
          Luvut(i - 1) = Luvut(i)
          Luvut(i) = Ltmp
          JarjTmp = Jarj(i - 1)
          Jarj(i - 1) = Jarj(i)
          Jarj(i) = JarjTmp
          OK = False
        End If
      Next i
    Loop Until OK = True
    
'Asetetaan atribuutit arrayn mukaisessa järjestyksessä paikalleen
'
    For i = 0 To Joukko.Count - 2
      Attrib(i).TextString = Joukko(Jarj(i)).TextString
      Attrib(i).InsertionPoint = Joukko(Jarj(i)).InsertionPoint
    Next i
    Attrib(4).TextString = Joukko(Jarj(UBound(Jarj))).TextString
    Attrib(4).InsertionPoint = Joukko(Jarj(UBound(Jarj))).InsertionPoint
    
    Joukko.Erase
    Joukko.Delete
    Set Joukko = Nothing
  End If
  Eiku.Enabled = True
  Me.Show
End Sub
Private Sub CommandButton2_Click()
  Unload Me
End Sub
Private Function SortArray(ByRef TheArray As Variant)
Sorted = False
Do While Not Sorted
    Sorted = True
For X = 0 To UBound(TheArray) - 1
    If TheArray(X) > TheArray(X + 1) Then
        Temp = TheArray(X + 1)
        TheArray(X + 1) = TheArray(X)
        TheArray(X) = Temp
        Sorted = False
    End If
Next X
Loop
End Function
Private Sub CommandButton3_Click()
Dim Joukko As AcadSelectionSet
Dim i As Integer
Dim FilterType(0) As Integer
Dim FilterData(0) As Variant
  For i = 0 To ActiveDocument.SelectionSets.Count - 1
    If ActiveDocument.SelectionSets(i).Name = "MAKEAPU" Then
      ActiveDocument.SelectionSets(i).Delete
      Exit For
    End If
  Next i
  Set Joukko = ActiveDocument.SelectionSets.Add("MAKEAPU")
  FilterType(0) = 2
  FilterData(0) = "POSINFO"
  Application.ZoomExtents
  Joukko.Select acSelectionSetAll, , , FilterType, FilterData
  Joukko.Highlight True
  Joukko.Delete
  Set Joukko = Nothing
  'Liikautetaan ikkunaa, jotta korostukset näkyvät
  Me.Move Me.Left + 1
  Me.Move Me.Left - 1
End Sub
Private Sub Eiku_Click()
  ActiveDocument.SendCommand "_u "
  ActiveDocument.Regen acActiveViewport
  Eiku.Enabled = False
End Sub
