VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} FrmExt 
   Caption         =   "Ext Set"
   ClientHeight    =   4245
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7650
   OleObjectBlob   =   "FrmExt.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "FrmExt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Kanta As New ADODB.Connection
Public Taulu As New ADODB.Recordset
Public oEntity As Object

Private Sub BClearHosts_Click()
Dim i As Integer
Dim Loytyi As Boolean
  If oEntity Is Nothing Then
    MsgBox "Select text or  Attribute first!", vbCritical, "Clear hosts"
  ElseIf MsgBox("This will clear all host information.  Are you sure?", vbOKCancel, "Clear Hosts") = vbOK Then
    Do
      Loytyi = False
      For i = 0 To LXDATA.ListCount - 1
        If LXDATA.List(i) = "1005" Then
          LXDATA.RemoveItem i
          Loytyi = True
          Exit For
        End If
      Next i
    Loop Until Loytyi = False
    WriteXData
  End If
End Sub
Private Sub BClose_Click()
  Unload Me
End Sub
Private Sub PaivitaTaul()
If Dir(TDB.Value) <> "" Then
  If Kanta <> "" Then Kanta.Close
  Kanta.Open "Driver=Microsoft Access Driver (*.mdb);DBQ=" & TDB.Value & ";"
  Kanta.CursorLocation = adUseClient
  Set Taulu = Kanta.OpenSchema(adSchemaTables)
  Set Asetukset = Kanta.Execute("SELECT Type, Setting FROM extSettings ORDER BY Type, Setting")
  LTables.Clear
  Do While Not Taulu.EOF
    If LCase(Left(Taulu.Fields(2).Value, 3)) = "ext" Then
      LTables.AddItem Taulu.Fields(2).Value
    End If
    Taulu.MoveNext
  Loop
End If
End Sub
Private Sub PaivitaSar()
  LFields.Clear
  If LTables.Value <> "" Then
    Set Taulu = Kanta.OpenSchema(adSchemaColumns, Array(Empty, Empty, LTables.Value))
    Do While Not Taulu.EOF
      LFields.AddItem Taulu.Fields(3).Value
      Taulu.MoveNext
    Loop
  End If
End Sub
Private Sub BDebug_Click()
Dim i As Integer, j As Integer
Dim Joukko As AcadSelectionSet
Dim FilterType(0) As Integer
Dim FilterData(0) As Variant
Dim Attrib As Variant
Dim Piste As Variant

Application.ZoomExtents
TeeDEBUG 'Varmistetaan ett‰ DEBUG on olemassa ja ett‰ se on valittu
ClearDEBUG False
FilterType(0) = 0
FilterData(0) = "INSERT,TEXT"
For i = 0 To ActiveDocument.SelectionSets.Count - 1
  If ActiveDocument.SelectionSets(i).Name = "EXTAPU" Then
    ActiveDocument.SelectionSets(i).Delete
  End If
Next i
Set Joukko = ThisDrawing.SelectionSets.Add("EXTAPU")
Joukko.Select acSelectionSetAll, , , FilterType, FilterData
For i = 0 To Joukko.Count - 1
  If Joukko(i).EntityType = 32 Then  'Valinta on teksti
    TarkistaX Joukko(i)
  ElseIf Joukko(i).HasAttributes Then 'Valinta on blokki josta tarkistetaan attribuutit
    Attrib = Joukko(i).GetAttributes
    For j = 0 To UBound(Attrib)
      TarkistaX Attrib(j)
    Next j
  End If
Next i
Joukko.Delete
ActiveDocument.Regen acActiveViewport
End Sub
Private Sub PiirraX(Tark As Variant)
Dim GDataType As Variant
Dim GDataValue As Variant
Dim Piste As AcadPoint
  Tark.GetXData "", GDataType, GDataValue
  If IsEmpty(GDataType) = False Then
    Set Piste = ActiveDocument.ModelSpace.AddPoint(Tark.InsertionPoint)
  End If
  Set Piste = Nothing
End Sub


Private Sub TarkistaX(Tark As Variant)
Dim GDataType As Variant
Dim GDataValue As Variant
Dim oKohde As AcadEntity
Dim Vari As AcColor
Dim i As Integer
  Tark.GetXData "", GDataType, GDataValue
  EkaPoint = Tark.InsertionPoint
  If IsEmpty(GDataType) = False Then
    Vari = acRed
    For i = 0 To UBound(GDataType)
      If GDataType(i) = 1005 Then
        If GDataValue(i) <> "0" Then
          On Error Resume Next
          Set oKohde = ActiveDocument.HandleToObject(GDataValue(i)) 'Haetaan Entity sen handleksen perusteella
          If Err <> 0 Then
            PiirraVirhe
          Else
            TokaPoint = oKohde.InsertionPoint
            PiirraArc Vari
            Vari = acGreen
          End If
          Err.Clear
          On Error GoTo 0
        End If
      End If
    Next i
  End If

End Sub

Private Sub BCDebug_Click()
If MsgBox("Do you want to delete DEBUG layer too?", vbYesNo, "Clear Debug Layer") = vbYes Then
  ClearDEBUG True
Else
  ClearDEBUG False
End If
ActiveDocument.Regen acActiveViewport
End Sub
Private Sub PiirraVirhe()
Dim Piste2(2) As Double
Dim Piste3(2) As Double
Piste2(0) = EkaPoint(0) + 10
Piste2(1) = EkaPoint(1) + 10
Piste3(0) = Piste2(0) + 20
Piste3(1) = Piste2(1)
ActiveDocument.ModelSpace.AddLine EkaPoint, Piste2
ActiveDocument.ModelSpace.AddLine Piste2, Piste3
ActiveDocument.ModelSpace.AddText "ERROR IN HANDLE", Piste2, 3
End Sub
Private Sub ClearDEBUG(Optional Tuhoa As Boolean) 'Muodostaa tai tyhjent‰‰ DEBUG layerin tai tuhoaa sen
Dim Joukko As AcadSelectionSet
Dim FilterType(0) As Integer
Dim FilterData(0) As Variant
Dim i As Integer
Application.ZoomExtents
For i = 0 To ActiveDocument.SelectionSets.Count - 1
  If ActiveDocument.SelectionSets(i).Name = "EXTAPU" Then
    ActiveDocument.SelectionSets(i).Delete
    Exit For
  End If
Next i
Set Joukko = ThisDrawing.SelectionSets.Add("EXTAPU")
FilterType(0) = 8       'Type 8 = Layer
FilterData(0) = "DEBUG" 'Layer name
Joukko.Select acSelectionSetAll, , , FilterType, FilterData
Joukko.Erase
Joukko.Delete
If Tuhoa Then
  ActiveDocument.ActiveLayer = ThisDrawing.Layers("0")
  ActiveDocument.Layers("DEBUG").Delete
End If
Set Joukko = Nothing
End Sub
Private Sub PiirraArc(Vari As AcColor)
'Piirt‰‰ kaaren kahden pisteen v‰lille
Dim cPoint(2) As Double
Dim Sade As Double
Dim Viiva As AcadLine
Dim aViiva As AcadLine
Dim AlkuKulma As Double
Dim LoppuKulma As Double
Dim Kaari As AcadArc
Dim kPiste As Variant
Dim oText As AcadText
Dim oPLine As AcadPolyline
Dim pPisteet(5) As Double
Dim Pituus As Double
If EkaPoint(0) > TokaPoint(0) Then
  cPoint(0) = TokaPoint(0) + (EkaPoint(0) - TokaPoint(0)) / 2
Else
  cPoint(0) = EkaPoint(0) + (TokaPoint(0) - EkaPoint(0)) / 2
End If
If EkaPoint(1) > TokaPoint(1) Then
  cPoint(1) = TokaPoint(1) + (EkaPoint(1) - TokaPoint(1)) / 2
Else
  cPoint(1) = EkaPoint(1) + (TokaPoint(1) - EkaPoint(1)) / 2
End If
  Set Viiva = ActiveDocument.ModelSpace.AddLine(EkaPoint, TokaPoint)
  Pituus = Viiva.Length
  Viiva.Rotate cPoint, PI / 2
  Set aViiva = ActiveDocument.ModelSpace.AddLine(Viiva.EndPoint, TokaPoint)
  LoppuKulma = aViiva.Angle
  aViiva.Delete
  Set aViiva = ActiveDocument.ModelSpace.AddLine(Viiva.EndPoint, EkaPoint)
  AlkuKulma = aViiva.Angle
  Sade = aViiva.Length
  aViiva.Delete
  Set Kaari = ActiveDocument.ModelSpace.AddArc(Viiva.EndPoint, Sade, AlkuKulma, LoppuKulma)
  Kaari.Color = Vari
  pPisteet(0) = TokaPoint(0): pPisteet(1) = TokaPoint(1)
  pPisteet(3) = TokaPoint(0): pPisteet(4) = TokaPoint(1) - 2
  Set oPLine = ActiveDocument.ModelSpace.AddPolyline(pPisteet)
  oPLine.SetWidth 0, 0, 1
  oPLine.Rotate TokaPoint, LoppuKulma - (1.57 / Pituus)
  oPLine.Color = Vari
  Viiva.Delete
  Set oText = Nothing
  Set Kaari = Nothing
  Set Viiva = Nothing
  Set aViiva = Nothing
End Sub
Private Sub TeeDEBUG()
Dim i As Integer
Dim Kerros As AcadLayer
For i = 0 To ThisDrawing.Layers.Count - 1
  If ThisDrawing.Layers(i).Name = "DEBUG" Then
    Set Kerros = ThisDrawing.Layers(i)
    Exit For
  End If
Next i
If Kerros Is Nothing Then
  Set Kerros = ThisDrawing.Layers.Add("DEBUG")
  Kerros.Color = acRed
End If
ThisDrawing.ActiveLayer = Kerros
Set Kerros = Nothing
End Sub
Private Sub AsetaDB()
Dim DataType(1) As Integer
Dim DataValue(1) As Variant
DataType(0) = 1001
DataType(1) = 1000
DataValue(0) = "EXTSET"
DataValue(1) = TDB.Value
If Dir(TDB.Value) <> "" Then
  ActiveDocument.ModelSpace.SetXData DataType, DataValue
  PaivitaTaul
  Else
    MsgBox "Database not found!", vbCritical, "Set Database"
  End If
  Set Taulu = Nothing
End Sub

Private Sub BMarkXData_Click()
Dim i As Integer, j As Integer
Dim Joukko As AcadSelectionSet
Dim FilterType(0) As Integer
Dim FilterData(0) As Variant
Dim Attrib As Variant

Application.ZoomExtents
ActiveDocument.SetVariable "PDMODE", 65
TeeDEBUG 'Varmistetaan ett‰ DEBUG on olemassa ja ett‰ se on valittu
'ClearDEBUG False
FilterType(0) = 0
FilterData(0) = "INSERT,TEXT"
For i = 0 To ActiveDocument.SelectionSets.Count - 1
  If ActiveDocument.SelectionSets(i).Name = "EXTAPU" Then
    ActiveDocument.SelectionSets(i).Delete
  End If
Next i
Set Joukko = ThisDrawing.SelectionSets.Add("EXTAPU")
Joukko.Select acSelectionSetAll, , , FilterType, FilterData
For i = 0 To Joukko.Count - 1
  If Joukko(i).EntityType = 32 Then  'Valinta on teksti
    PiirraX Joukko(i)
  ElseIf Joukko(i).HasAttributes Then 'Valinta on blokki josta tarkistetaan attribuutit
    Attrib = Joukko(i).GetAttributes
    For j = 0 To UBound(Attrib)
      PiirraX Attrib(j)
    Next j
  End If
Next i
Joukko.Delete
ActiveDocument.Regen acActiveViewport
End Sub
Private Sub BSelectHost_Click()
Dim i As Integer
Dim Lev As Double
Dim Kork As Double
Dim TransMatrix As Variant
Dim ContextData As Variant
Dim APUEntity As AcadEntity
Dim iPoint As Variant
Dim Loytyi As Boolean
  If oEntity Is Nothing Then
    MsgBox "Select Attribute first!", vbCritical, "Select hosts"
  ElseIf MsgBox("This will clear previous host information.  Are you sure?", vbOKCancel, "Select hosts") = vbOK Then
    On Error GoTo Loppu
    Lev = Me.Width
    Kork = Me.Height
    Me.Width = 0
    Me.Height = 0
    'Poistetaan ensin vanhat handle tiedot
    Do
      Loytyi = False
      For i = 0 To LXDATA.ListCount - 1
        If LXDATA.List(i) = "1005" Then
          LXDATA.RemoveItem i
          Loytyi = True
          Exit For
        End If
      Next i
    Loop Until Loytyi = False
    Do
      ActiveDocument.Utility.GetSubEntity APUEntity, iPoint, TransMatrix, ContextData, "Select host object or esc to exit."
      If APUEntity.EntityType = 6 Or APUEntity.EntityType = 32 Then
        LXDATA.AddItem "1005", LXDATA.ListCount - 1
        LXDATA.List(LXDATA.ListCount - 2, 1) = APUEntity.Handle
      Else
          ActiveDocument.Utility.Prompt "You didn't pick text or attribute."
      End If
    Loop
Loppu:
    Err.Clear
    On Error GoTo 0
    WriteXData
    Me.Width = Lev
    Me.Height = Kork
  End If
End Sub
Private Sub BSetXdata_Click()
Dim i As Integer
Dim Handlet() As String
  If oEntity Is Nothing Then
    MsgBox "Select text or Attribute first!", vbCritical, "Set XData"
  ElseIf LTables.Value = "" Or LFields.Value = "" Then
    MsgBox "Select table and field first!", vbCritical, "Set XData"
  Else
    ReDim Handlet(0)
    If LXDATA.ListCount <> 0 Then
      For i = 0 To LXDATA.ListCount - 1
        If LXDATA.List(i, 0) = "1005" Then
          If LXDATA.List(i, 1) <> "0" Then
            ReDim Preserve Handlet(UBound(Handlet) + 1)
            Handlet(UBound(Handlet)) = LXDATA.List(i, 1)
          End If
        End If
      Next i
      LXDATA.Clear
    End If
    With LXDATA
      .AddItem "1001"
      .List(.ListCount - 1, 1) = "PIKAS_PIIRIK"
      .AddItem "1002"
      .List(.ListCount - 1, 1) = "{"
      .AddItem "1000"
      .List(.ListCount - 1, 1) = Mid(LTables.Value, 4)
      .AddItem "1000"
      .List(.ListCount - 1, 1) = LFields.Value
      .AddItem "1000"
      .List(.ListCount - 1, 1) = TData1.Value
      .AddItem "1000"
      .List(.ListCount - 1, 1) = TData2.Value
      For i = 1 To UBound(Handlet)
        .AddItem "1005"
        .List(.ListCount - 1, 1) = Handlet(i)
      Next i
      .AddItem "1002"
      .List(.ListCount - 1, 1) = "}"
    End With
    WriteXData
  End If
End Sub
Private Sub WriteXData()
Dim ToXTypes() As Integer
Dim ToXValues() As Variant
Dim i As Integer
  'Lista on valmis p‰ivitet‰‰n se saman tien Attribuutin xdataan
  ReDim ToXTypes(LXDATA.ListCount - 1)
  ReDim ToXValues(LXDATA.ListCount - 1)
  For i = 0 To LXDATA.ListCount - 1
    ToXTypes(i) = CInt(LXDATA.List(i, 0))
    ToXValues(i) = LXDATA.List(i, 1)
  Next i
  oEntity.SetXData ToXTypes, ToXValues
End Sub

Private Sub BValitse_Click()
Dim iPoint As Variant
Dim GDataType As Variant
Dim GDataValue As Variant
Dim i As Integer, j As Integer
Dim Lev As Double
Dim Kork As Double
Dim TransMatrix As Variant
Dim ContextData As Variant
Dim Tyypit As Integer

  On Error GoTo Loppu
  Lev = Me.Width
  Kork = Me.Height
  Me.Width = 0
  Me.Height = 0
  Do
    ActiveDocument.Utility.GetSubEntity oEntity, iPoint, TransMatrix, ContextData, "Select object."
    oEntity.GetXData "", GDataType, GDataValue
    If oEntity.EntityType = 6 Then
      LTag.Caption = oEntity.TagString
      Exit Do
    ElseIf oEntity.EntityType = 32 Then
      LTag.Caption = ""
      Exit Do
    End If
  Loop
  LValue.Caption = oEntity.TextString
  LHandle.Caption = oEntity.Handle
  LXDATA.Clear
  If IsEmpty(GDataType) = False Then
    For i = 0 To UBound(GDataType)
      If GDataType(i) = "1000" Then
        tyyppi = tyyppi + 1
        If tyyppi = 1 Then
           For j = 0 To LTables.ListCount - 1
            If InStr(LTables.List(j), GDataValue(i)) Then
                LTables.Selected(j) = True
                Exit For
            End If
          Next j
        ElseIf GDataType(i + 1) <> "1000" Then
           For j = 0 To LFields.ListCount - 1
            If InStr(LFields.List(j), GDataValue(i)) Then
                LFields.Selected(j) = True
                Exit For
            End If
          Next j
        ElseIf tyyppi = 2 And GDataType(i + 1) = "1000" Then
          TData1.Value = GDataValue(i)
        ElseIf tyyppi = 3 And GDataType(i + 1) = "1000" Then
          TData2.Value = GDataValue(i)
        End If
      End If
      LXDATA.AddItem GDataType(i)
      LXDATA.List(i, 1) = GDataValue(i)
    Next i
  End If
  Me.Width = Lev
  Me.Height = Kork
  Exit Sub
Loppu:
  Me.Width = Lev
  Me.Height = Kork
  LTag.Caption = ""
  LValue.Caption = ""
  LHandle.Caption = ""
  Set oEntity = Nothing
  LXDATA.Clear
End Sub
Private Sub LTables_Change()
  PaivitaSar
End Sub
Private Sub VaihdaSymboli()
'Dim Vastaus As String
'Vastaus = InputBox("Give symbol eg. M 1", "Symbol Name", TData2.Value & " " & TData1.Value)
'If Vastaus <> "" Then
'  If InStr(Vastaus, " ") Then
'    TData2.Value = Left(Vastaus, InStr(Vastaus, " ") - 1)
'    TData1.Value = Mid(Vastaus, InStr(Vastaus, " ") + 1)
'  End If
'End If
  FrmSelect.Show vbModal
End Sub

Private Sub TData1_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
  VaihdaSymboli
End Sub
Private Sub TData2_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
  VaihdaSymboli
End Sub
Private Sub TDB_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
Dim Vastaus As String
Vastaus = InputBox("Give DB name including path.", "Set Database", TDB.Value)
If Vastaus <> "" Then
  TDB.Value = Vastaus
  AsetaDB
End If
End Sub

Private Sub UserForm_Initialize()
Dim DataTypeG As Variant
Dim DataValueG As Variant
  ActiveDocument.ModelSpace.GetXData "EXTSET", DataTypeG, DataValueG
  If IsEmpty(DataTypeG) = False Then
    TDB.Value = DataValueG(1)
    PaivitaTaul
  End If
End Sub
Private Sub UserForm_Terminate()
  Set Kanta = Nothing
  Set Taulu = Nothing
End Sub
