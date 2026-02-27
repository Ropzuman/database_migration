Attribute VB_Name = "Koodit"
Sub StartLoopInst()
Dim DB As New ADODB.Connection 'Tietokanta
Dim Taulu As New ADODB.Recordset 'Taulukko
Dim Tietokanta As String
Dim BName As String
Dim oBlock As AcadBlockReference
Dim Viiva As AcadLine
Dim ipoint As Variant
Dim ePoint As Variant
Dim Teksti As String
Dim Looppi As String
Dim Funktio As String
Dim Selitys As String
Dim Attribuutit As Variant
Dim i As Integer
Dim Piste As String
Tietokanta = "N:\whldata\Projekti\Bohui\SAHKO\DATABASE\Instru.mdb"
BName = "UP079"
'Avataan tietokanta
DB.Open "Driver=Microsoft Access Driver (*.mdb);DBQ=" & Tietokanta & ";"
'Avataan Taulukko
Set Taulu = DB.Execute("Select * From Loops ORDER By LoopNo")
ThisDrawing.ActiveLayer = ThisDrawing.Layers("INST")
Do While Not Taulu.EOF
  On Error Resume Next
  Looppi = Taulu.Fields("LoopNo")
  Funktio = Taulu.Fields("LoopSymb")
  Selitys = Taulu.Fields("DescrTL") & " " & Taulu.Fields("DescrBL")
  Teksti = vbCrLf & Looppi & "-" & Funktio & ": " & Selitys
  ipoint = ThisDrawing.Utility.GetPoint(, Teksti)
  If Err = 0 Then
    Set oBlock = ThisDrawing.ModelSpace.InsertBlock(ipoint, BName, 1, 1, 1, 0)
    Attribuutit = oBlock.GetAttributes
    For i = 0 To UBound(Attribuutit) 'K‰yd‰‰n l‰pi kaikki attribuutit
      With Attribuutit(i)
        If .TagString = "FUNCTION" Then
          .TextString = Funktio
        ElseIf .TagString = "SDPOS" Then
          .TextString = Looppi
        ElseIf .TagString = "DESCR" Then
          .TextString = Selitys
        End If
      End With
    Next i
    Piste = Replace(CStr(ipoint(0)), ",", ".") & "," & Replace(CStr(ipoint(1)), ",", ".")
    'Annetaan komentorivikomento jolla saadaan siirretty‰ blokkia haluttuun paikkaan
    ThisDrawing.SendCommand "(command ""move"" ""last"" """" """ & Piste & """)" & vbCr
    'Tarkistetaan ett‰ blokkia yleens‰ siirettiin
    With oBlock
      If .InsertionPoint(0) = ipoint(0) And .InsertionPoint(1) = ipoint(1) Then 'Ei siirretty joten poistetaan blokki
        .Delete
        oDot.Delete
        GoTo Ohitus
      End If
    End With
'Piirret‰‰n viiva
    Do
      ePoint = ThisDrawing.Utility.GetPoint(ipoint, "Draw line")
      If Err <> 0 Then
        If Viiva Is Nothing Then 'Lopetettiin ennen kuin yht‰k‰‰n viiva oli piirretty
          oBlock.Delete   'Tuhotaan Blokki
          GoTo Ohitus
        End If
        Err.Clear
        Exit Do
      Else
        Set Viiva = ThisDrawing.ModelSpace.AddLine(ipoint, ePoint)
        Viiva.Color = acYellow
        ipoint(0) = ePoint(0)
        ipoint(1) = ePoint(1)
      End If
    Loop
  End If
Ohitus:
  Err.Clear
  Taulu.MoveNext
Loop
End Sub

