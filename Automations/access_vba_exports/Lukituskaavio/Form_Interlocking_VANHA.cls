VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_Interlocking_VANHA"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit
Public BaseDWG As String
Public BasePath As String
Public SavePath As String
Public oBlock As AcadBlockReference

Private Sub Blokki_BeforeUpdate(Cancel As Integer)
Dim TAttrib As String
Dim i As Integer
  TKuva.VALUE = DLookup("[IMAGE]", "SETTINGS", "[VALUE]='" & Blokki.VALUE & "'")
  TAttrib = Blokki.Column(1, Blokki.ListIndex)
  For i = 1 To 5
    Controls("Text_" & i).BackColor = RGB(255, 0, 0)
    Controls("T_IN" & i).BackColor = RGB(255, 0, 0)
    Controls("T_OUT" & i).BackColor = RGB(255, 0, 0)
    
    If InStr(TAttrib, "T" & i) Then
      Controls("Text_" & i).BackColor = RGB(60, 255, 60)
    End If
    If InStr(TAttrib, "I" & i) Then
      Controls("T_IN" & i).BackColor = RGB(60, 255, 60)
    End If
    If InStr(TAttrib, "O" & i) Then
      Controls("T_OUT" & i).BackColor = RGB(60, 255, 60)
    End If
  
  Next i
End Sub

Private Sub Blokki_Change()
If Blokki.VALUE = "transition2" Then
    If Text_2.VALUE = "open" Then
    Text_2.VALUE = Text_2.VALUE & " LIMIT"
    Else
    If Text_2.VALUE = "close" Then
    Text_2.VALUE = Text_2.VALUE & " LIMIT"
    End If
    End If
End If
If Blokki.VALUE = "transition2" And Text_4.VALUE Like "*M0*" Then
Text_2.VALUE = "RUN"
Else
Text_2.VALUE = Text_2.VALUE
End If
If Blokki.VALUE = "To_Sec_V" And T_OUT4.VALUE Like "*reset*" Then
Text_5.VALUE = "R:"
Else
Text_5.VALUE = ""
End If

End Sub

Private Sub BMuokkaa_Click()
'Muokkaa blokkia
If BMuokkaa.Caption = "Muokkaa blokkia" Then
  If LueAttribuutit Then
    BMuokkaa.Caption = "Vie blokkiin"
    BMuokkaa.FontBold = True
  End If
Else
  KirjoitaAttribuutit
  Set oBlock = Nothing
  BMuokkaa.Caption = "Muokkaa blokkia"
  BMuokkaa.FontBold = False
End If
End Sub
Private Function LueAttribuutit() As Boolean
'Luetaan attibuutit blokista formiin
Dim Valinta As AcadObject
Dim Point(0 To 2) As Double
Dim BlockNimi As String
Dim Oikea As Boolean
Dim i As Integer
Dim Asetukset As Recordset
Dim oAttrib As Variant

  AppActivate oACAD.Caption
  oDOC.SendCommand Chr(27) & Chr(27) 'Varmistetaan että ollaan poistuttu komennosta
  oDOC.Utility.GetEntity Valinta, Point, "Choose Device..."
  If Valinta.ObjectName = "AcDbBlockReference" Then
    Set Asetukset = CurrentDb.OpenRecordset("SETTINGS")
    BlockNimi = UCase(Valinta.Name)
    Do While Not Asetukset.EOF 'Tarkistetaan että blokin nimi löytyy luettelosta
      If UCase(Asetukset.Fields("[VALUE]")) = BlockNimi Then   'Löyty
        Oikea = True
        Exit Do
      End If
      Asetukset.MoveNext
    Loop
    If Oikea Then
      If Valinta.HasAttributes Then
        LueAttribuutit = True
        Set oBlock = Valinta
        oAttrib = oBlock.GetAttributes
        TaytaTextBoxit oAttrib
      End If
    End If
  End If
  If LueAttribuutit = False Then
    MsgBox "Virhe blokin valinnassa!", vbOKOnly, "Muokkaa blokkia"
  End If
  AppActivate CurrentDb.Properties("AppTitle").VALUE
End Function
Private Sub KirjoitaAttribuutit()
Dim i As Integer
Dim oAttrib As Variant
'Kirjoitetaan attribuutit formista takaisin blokkiin
If Not oBlock Is Nothing Then
    oAttrib = oBlock.GetAttributes
    For i = 0 To UBound(oAttrib)
      With oAttrib(i)
        Select Case Left(UCase(.TagString), Len(.TagString) - 1)
          Case "TEXT_", "TXT"
            .TextString = Nz(Controls("Text_" & Right(.TagString, 1)))
          Case "IN_"
            .TextString = Nz(Controls("T_IN" & Right(.TagString, 1)))
          Case "OUT_"
            .TextString = Nz(Controls("T_OUT" & Right(.TagString, 1)))
        End Select
      End With
    Next i
End If
End Sub

Private Sub Command7_Click()
'Avaa uusi pohjakuva
  Set oDOC = oACAD.Documents.Open(BasePath & BaseDWG & ".dwg", True)
  oACAD.ZoomExtents
  AppActivate oACAD.Caption
  
End Sub

Private Sub Command89_Click()
'Tyhjennä IN
  Tyhjenna "IN"
End Sub

Private Sub Command90_Click()
'Tyhjennä OUT
  Tyhjenna "OUT"
End Sub

Private Sub CPisteet_Click()
Dim gPoint As Variant
Dim iPoint(2) As Double
Dim BName As String
Dim Piste As String
Dim PointBlock As AcadBlockReference
On Error GoTo Virhe 'Täytyy olla, jos GetPoint ei toimikaan
  
  AppActivate oACAD.Caption
  oDOC.SendCommand Chr(27) & Chr(27) 'Varmistetaan että ollaan poistuttu komennosta
  'Käsketään klikata jotain pistettä
Virhe:

  gPoint = oDOC.Utility.GetPoint(, "Set start point")
  Err.Clear
  On Error GoTo 0
  iPoint(0) = gPoint(0) + 0.00001
  iPoint(1) = gPoint(1) + 0.00001
'---------------------------------------------------------
  Piste = DoubleToPoint(CDbl(iPoint(0))) & "," & DoubleToPoint(CDbl(iPoint(1)))
  BName = BlockPath & "POINT.DWG"
  oDOC.SendCommand Chr(27) & Chr(27) 'Varmistetaan että ollaan poistuttu komennosta
  Set PointBlock = oDOC.ModelSpace.InsertBlock(iPoint, BName, 1, 1, 1, 0)
  oDOC.SendCommand "(command ""move"" ""last"" """" """ & Piste & """)" & vbCr
  With PointBlock
    If .InsertionPoint(0) = iPoint(0) And .InsertionPoint(1) = iPoint(1) Then
      .Delete
    End If
  End With
  Set PointBlock = Nothing
  AppActivate CurrentDb.Properties("AppTitle").VALUE
End Sub

Private Sub CViivanPiirto_Click()
Dim EkaViiva As AcadLine
Dim ViimViiva As AcadLine
Const Pi = 3.14159265358979

Dim Lyhennys As Double
Dim iPoint(2) As Double
Dim LoppuPiste(2) As Double
Dim Rotation As Double
Dim BName As String


Dim Piste As Variant
Dim sPoint(2) As Double
Dim i As Integer
  On Error Resume Next
  
  AppActivate oACAD.Caption
  oDOC.SendCommand Chr(27) & Chr(27) 'Varmistetaan että ollaan poistuttu komennosta
  
  
'  For i = 0 To oDOC.Layers.Count - 1
'    If oDOC.Layers(i).Name = "CANVAS" Then
'      oDOC.ActiveLayer = oDOC.Layers(i)
'      Exit For
'    End If
'  Next i
  Piste = oDOC.Utility.GetPoint(, "Set start point")
  If Err = 0 Then
    sPoint(0) = Piste(0)
    sPoint(1) = Piste(1)
    Do
      Piste = oDOC.Utility.GetPoint(sPoint, "Set end point or end drawing")
      If Err <> 0 Then
        If ViimViiva Is Nothing Then 'Lopetettiin ennen kuin yhtäkään viiva oli piirretty
        Else 'Lopetettiin, piirretään viivan pää
        '------- VIIVAN PÄÄ ---------
            If Tyyppi.VALUE = 1 Then 'Nuoli
              BName = BlockPath & "ARROW.DWG"
              Lyhennys = 7.5
            ElseIf Tyyppi.VALUE = 2 Then 'Nuoli ja Not merkki
              BName = BlockPath & "NOTARROW.DWG"
              Lyhennys = 7.5
            Else 'Piste
              BName = BlockPath & "POINT.DWG"
              Lyhennys = 0
            End If
            If Sijainti.VALUE = 1 Then
              'Ei piirretä viivaa
            ElseIf Sijainti.VALUE = 2 Then 'Piirretään nuoli alkuun
              With EkaViiva
                iPoint(0) = .StartPoint(0)
                iPoint(1) = .StartPoint(1)
                Rotation = .Angle + Pi
                LoppuPiste(0) = .StartPoint(0) + Lyhennys * Cos(Rotation - Pi)
                LoppuPiste(1) = .StartPoint(1) + Lyhennys * Sin(Rotation - Pi)
                .StartPoint = LoppuPiste
              End With
              oDOC.ModelSpace.InsertBlock iPoint, BName, 1, 1, 1, Rotation
            Else 'Piirretään nuoli loppuun
              With ViimViiva
                iPoint(0) = .EndPoint(0)
                iPoint(1) = .EndPoint(1)
                Rotation = .Angle
                LoppuPiste(0) = .EndPoint(0) - Lyhennys * Cos(Rotation)
                LoppuPiste(1) = .EndPoint(1) - Lyhennys * Sin(Rotation)
                .EndPoint = LoppuPiste
              End With
              oDOC.ModelSpace.InsertBlock iPoint, BName, 1, 1, 1, Rotation
            End If
            Set EkaViiva = Nothing
            Set ViimViiva = Nothing
        '------- VIIVAN PÄÄ ---------
        End If
        AppActivate CurrentDb.Properties("AppTitle").VALUE
        Exit Do
      Else
        Set ViimViiva = oDOC.ModelSpace.AddLine(sPoint, Piste)
        If EkaViiva Is Nothing Then
          Set EkaViiva = ViimViiva
        End If
        sPoint(0) = Piste(0)
        sPoint(1) = Piste(1)
      End If
    Loop
  Else 'Keskeytettiin heti
    AppActivate CurrentDb.Properties("AppTitle").VALUE
  End If
  Err.Clear

End Sub

Private Sub Form_Close()
'Vapautetaan objektit ennen sulkemista
  Set oACAD = Nothing
  Set oDOC = Nothing
End Sub

Private Sub Form_Load()

'Yhdistetään ensin AutoCADiin
'Otetaan yhteys AutoCADiin
  On Error Resume Next
  Set oACAD = GetObject(, "AutoCAD.Application") 'Koitetaan yhdistää AutoCADiin
  If Err <> 0 Then 'Käynnissä olevaa AutoCADiä ei löytynyt
    MsgBox "Käynnissä olevaa AutoCADiä ei löytynyt!", vbCritical, "Virhe!"
    Set oACAD = Nothing
    Unload Me
    Exit Sub
  End If
  On Error GoTo 0
  Set oDOC = oACAD.ActiveDocument
  'Haetaan asetukset
  BaseDWG = DLookup("[VALUE]", "SETTINGS", "[SETTING]='BASEDWG'")
  BasePath = DLookup("[VALUE]", "SETTINGS", "[SETTING]='BASEPATH'")
  BlockPath = DLookup("[VALUE]", "SETTINGS", "[SETTING]='BLOCKPATH'")
  SavePath = DLookup("[VALUE]", "SETTINGS", "[SETTING]='SAVEPATH'")
   
  
End Sub

Private Sub Lehdet_Change()
  'If Lehdet <> 3 Then
  If Lehdet = 2 Then
    Blokki.RowSource = "SELECT SETTINGS.VALUE, ATTRIB FROM SETTINGS WHERE (((SETTINGS.SETTING)='BLOCK_S'));"
    Else
    Blokki.RowSource = "SELECT SETTINGS.VALUE, ATTRIB FROM SETTINGS WHERE (((SETTINGS.SETTING)='BLOCK'));"
  End If
  If Lehdet = 0 Then 'Loop
    Lehdet_Sub0.Form.VaihdaTiedot
  ElseIf Lehdet = 1 Then 'Motor
    Lehdet_sub1.Form.VaihdaTiedot
  ElseIf Lehdet = 2 Then 'Sequense
    Lehdet_Sub2.Form.VaihdaTiedot
    Controls("Text_2").VALUE = vbNullString
  ElseIf Lehdet = 3 Then 'Sequense
    TyhjennaKentat
    Blokki.RowSource = "SELECT SETTINGS.VALUE, ATTRIB FROM SETTINGS WHERE (((SETTINGS.SETTING)='BLOCK_O')); "
  End If
  Blokki.Requery
  Blokki.VALUE = ""
  TKuva.VALUE = Nothing
End Sub

Private Sub Lisays_Click()
Dim BName As String
Dim iPoint(2) As Double
Dim gPoint As Variant
Dim Piste As String
On Error GoTo Virhe
'Liitetään ensin blokki kuvaan
  Set oDOC = oACAD.ActiveDocument
  
  If Dir(BlockPath & Blokki.VALUE & ".dwg") = "" Then
    MsgBox "Insertoitavaa blokkia ei löytynyt", vbCritical, "Insertoi"
    Exit Sub
  End If
  
  AppActivate oACAD.Caption
  oDOC.SendCommand Chr(27) & Chr(27) 'Varmistetaan että ollaan poistuttu komennosta
  
'AutoCAD etsii itse pisteen
'Käsketään klikata jotain pistettä
Virhe:
  gPoint = oDOC.Utility.GetPoint(, "Set start point")
  Err.Clear
  On Error GoTo 0
  iPoint(0) = gPoint(0) + 0.00001
  iPoint(1) = gPoint(1) + 0.00001
'---------------------------------------------------------
  Piste = DoubleToPoint(CDbl(iPoint(0))) & "," & DoubleToPoint(CDbl(iPoint(1)))
  BName = BlockPath & Blokki & ".dwg"
  Set oBlock = oDOC.ModelSpace.InsertBlock(iPoint, BName, 1, 1, 1, 0)
  'Kirjoitetaan arvot textboxeista
  KirjoitaAttribuutit
  'Annetaan komentorivikomento jolla saadaan siirrettyä blokkia haluttuun paikkaan
  oDOC.SendCommand "(command ""move"" ""last"" """" """ & Piste & """)" & vbCr
  With oBlock
    If .InsertionPoint(0) = iPoint(0) And .InsertionPoint(1) = iPoint(1) Then
      .Delete
    Else 'Kirjoitetaan tieto insertoinnista taulukkoon
      If Lehdet = 0 Then 'Jos insertoitiin joku muu kuin yleisblokki
        Lehdet_Sub0.Form.USED = True
        Lehdet_Sub0.Requery
      ElseIf Lehdet = 1 Then
        Lehdet_sub1.Form.USED = True
        Lehdet_sub1.Requery
      ElseIf Lehdet = 2 Then
        Lehdet_Sub2.Form.USED = True
        Lehdet_Sub2.Requery
      End If
    End If
  End With
  Set oBlock = Nothing
  AppActivate CurrentDb.Properties("AppTitle").VALUE

End Sub
Private Function DoubleToPoint(ByVal Piste As Double) As String
Dim APU As String
Dim Osoitin As Integer
  APU = CStr(Piste)
  Osoitin = InStr(APU, ",")
  If Osoitin = 0 Then
    DoubleToPoint = APU
  Else
    DoubleToPoint = Left(APU, Osoitin - 1) & "." & Mid(APU, Osoitin + 1)
  End If
End Function
Private Function Arvo(Tieto As Variant) As String
  If IsNull(Tieto) Or Tieto = vbNullString Then
    Arvo = vbNullChar
  Else
    Arvo = Tieto
  End If
End Function
Private Sub TaytaTextBoxit(oAttrib As Variant)
Dim i As Integer
'Tyhjennetään ensin kaikki textBoxit
  TyhjennaKentat
  For i = 0 To UBound(oAttrib)  'Käydään läpi kaikki blokin attribuutit
    With oAttrib(i)
      Select Case Left(UCase(.TagString), Len(.TagString) - 1)
        Case "TXT", "TEXT_"
          Controls("Text_" & Right(.TagString, 1)).VALUE = .TextString
        Case "IN_"
          Controls("T_IN" & Right(.TagString, 1)).VALUE = .TextString
        Case "OUT_"
          Controls("T_OUT" & Right(.TagString, 1)).VALUE = .TextString
       End Select
    End With
  Next i
  
End Sub
Public Sub TyhjennaKentat()
Dim i As Integer
  Controls("Text_1").VALUE = vbNullString
  If Controls("Text_2").VALUE = "RUN" Then
  Controls("Text_2").VALUE = vbNullString
  Else
  End If
  Controls("Text_3").VALUE = vbNullString
  Controls("Text_4").VALUE = vbNullString
  Controls("Text_5").VALUE = vbNullString
End Sub
Public Sub Tyhjenna(Tyyppi As String)
Dim i As Integer
  For i = 1 To 5
    Controls("T_" & Tyyppi & i).VALUE = vbNullString
  Next i
End Sub



