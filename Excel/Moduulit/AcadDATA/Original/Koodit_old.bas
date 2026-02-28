Option Explicit
Public oACAD As AcadApplication          'AutoCad objekti
Public oDOC As AcadDocument        'Dokumentti objekti (oACad.ActiveDocument)
Public OliAuki As Boolean
Public Ver As Integer

Public Sub TuoDATA(Optional Valitut As Boolean, Optional Filtterit As String)   'Tämä tuo datan AutoCADistä Exceliin
' 7.3.2003 - VG
' 27.3.2003 - VG
' 19.1.2004 - VG
' 29.1.2004 - VG -> Attribuuttien nimien ottaminen huomioon
Dim Tyhjenna As Boolean
Dim Listasta As Boolean
Dim VainValitut As Boolean
  Listasta = Aloitus.Lista.Value
  If Listasta = False Then
    If Valitut Then
      VainValitut = True
    Else
      If MsgBox("Poimitaanko vain valitut kohteet?", vbYesNo, "Tuo DATA") = vbYes Then
        VainValitut = True
      End If
    End If
  End If
  Application.ScreenUpdating = False
'Laitetaan otsikot
'Otetaan yhteys AutoCADiin
  On Error Resume Next
  Set oACAD = GetObject(, "AutoCAD.Application") 'Koitetaan yhdistää AutoCADiin
  If Err <> 0 Then 'Käynnissä olevaa AutoCADiä ei löytynyt
    MsgBox "Käynnissä olevaa AutoCADiä ei löytynyt!", vbCritical, "Virhe!"
    Set oACAD = Nothing
    Exit Sub
  End If
  On Error GoTo 0
'Yhteys muodostettu käynnissä olevaan AutoCADiin
Dim Joukko As AcadSelectionSet  'Joukko, jolla valitaan kaikki halutut blokit
Dim BlockArray As Variant       'Array muuttuja Blokkia varten
Dim i As Integer, j As Integer, jj As Integer  'Indeksimuuttuja
Dim FilterType(0) As Integer
Dim FilterData(0) As Variant
Dim Poista() As AcadEntity
Dim L As Integer
Dim EiPoisteta As Boolean

Dim Nimet As String
Dim Blokit As Variant
Dim Blokki As AcadBlockReference

Dim DWGName As String
Dim Hakemisto As String
Dim EkaKerta As Boolean
Dim Rivi As Long
Dim oText As AcadText
Dim oMText As AcadMText
Dim DocRivi As Integer
Dim DocMaara As Integer
Dim Loytyi As Boolean
Dim Filter2 As Boolean
Dim Docmode As Boolean
 
 If Aloitus.Tyhjenna.Value = True Then Tyhjenna = True
 EkaKerta = True
 DATA.Select
 If Tyhjenna Then
    Cells.Select
    Selection.Clear
    Selection.NumberFormat = "@" 'Asetetaan solujen formaatiksi teksti
    Rows("1:1").Font.Bold = True 'Laitetaan otsikot lihavoiduiksi
    Columns("E:F").NumberFormat = "General"
    Range("A1").Select
    Cells(1, 1).Value = "PATH"   '--------------------------------
    Cells(1, 2).Value = "DWG"    '
    Cells(1, 3).Value = "BLOCK"  '
    Cells(1, 4).Value = "HANDLE" ' 7 kpl yhteistä perus otsikkoa
    Cells(1, 5).Value = "XCord"  '
    Cells(1, 6).Value = "YCord"  '
    Cells(1, 7).Value = "Layer"  '--------------------------------
 End If
 If Listasta Then
   TIEDLISTA.Select
   Do
     i = i + 1
     If Cells(i, 1).Value = "" Then
       DocMaara = i - 1
       Exit Do
     End If
   Loop
   DATA.Select
 Else
   DocMaara = 1
 End If
Rivi = 2
Do
  If Cells(Rivi, 1).Value = "" Then Exit Do
  Rivi = Rivi + 1
Loop
Nimet = UCase(Aloitus.Range("D7").Value) 'Blokkien nimet
Blokit = Split(Nimet, ",")
For i = 0 To UBound(Blokit)
  Blokit(i) = Trim(Blokit(i))
Next i
  If Aloitus.Range("D5").Value = "Tekstit" Then
     FilterType(0) = 0
     FilterData(0) = "TEXT,MTEXT,DTEXT"
  ElseIf Aloitus.Range("D5").Value = "Blokit ja tekstit" Then
    FilterType(0) = 0
    FilterData(0) = "TEXT,MTEXT,DTEXT,INSERT"
  Else
    FilterType(0) = 0
    FilterData(0) = "INSERT"
  End If

Docmode = oACAD.Preferences.System.SingleDocumentMode
oACAD.Preferences.System.SingleDocumentMode = False

For DocRivi = 1 To DocMaara
  Application.StatusBar = "Doc: " & DocRivi & "/" & DocMaara
  If Listasta = False Then
    Set oDOC = oACAD.ActiveDocument
    Loytyi = True
  Else
    
    Loytyi = False
    For i = 0 To oACAD.Documents.Count - 1
      If UCase(oACAD.Documents(i).Name) = UCase(Dir(TIEDLISTA.Cells(DocRivi, 1).Value)) Then
        Set oDOC = oACAD.Documents(i)
        Loytyi = True
      End If
    Next i
    If Loytyi = False Then
      Set oDOC = oACAD.Documents.Open(TIEDLISTA.Cells(DocRivi, 1).Value)
    End If
  End If
  'Otetaan tiedoston nimi aktiivisesta dokumentista
  DWGName = Left(oDOC.Name, Len(oDOC.Name) - 4)    'Poistetaan pääte (.dwg)
  Hakemisto = oDOC.Path
  For i = 0 To oDOC.SelectionSets.Count - 1
    If oDOC.SelectionSets(i).Name = "EXCELHAKU" Then
      oDOC.SelectionSets(i).Delete
      Exit For
    End If
  Next i
  Set Joukko = oDOC.SelectionSets.Add("EXCELHAKU")
  If VainValitut Then
     Joukko.Select acSelectionSetPrevious, , , FilterType, FilterData
  Else
    Joukko.Select acSelectionSetAll, , , FilterType, FilterData
  End If
  
'Etsitään joukosta kaikki ne blokit, jotka eivät täytä valintaa
  L = 0
  For j = 0 To Joukko.Count - 1
    If Joukko(j).EntityName = "AcDbBlockReference" Then
      EiPoisteta = False
      For i = 0 To UBound(Blokit)
        If UCase(Joukko(j).EffectiveName) = Blokit(i) Then
          EiPoisteta = True
          Exit For
        ElseIf Blokit(i) = "*" Then
          EiPoisteta = True
        End If
      Next i
      If EiPoisteta = False Then
        ReDim Preserve Poista(L)
        Set Poista(L) = Joukko(j)
        L = L + 1
      End If
    End If
  Next j
  If L > 0 Then Joukko.RemoveItems Poista
  
  If Joukko.Count = 0 Then
    MsgBox "Kuvasta tai valitulta alueelta ei löytynyt tietoja, jotka täyttäisivät ehdon!", vbCritical, "Tuo DATA"
  End If
  For i = 0 To Joukko.Count - 1
    Application.StatusBar = "Luetaan tietoa: " & i + 1 & "/" & Joukko.Count & "  File: " & DWGName
    Cells(Rivi, 1).Value = Hakemisto       'Dokumentin hakemisto ylös
    Cells(Rivi, 2).Value = DWGName         'Dokumentin nimi ylös
    Cells(Rivi, 4).Value = Joukko(i).Handle   'Blokin Handle
    If Joukko(i).EntityName = "AcDbBlockReference" Then
      Set Blokki = Joukko(i)
      Cells(Rivi, 5).Value = Blokki.InsertionPoint(0) 'XCord
      Cells(Rivi, 6).Value = Blokki.InsertionPoint(1) 'YCord
      Cells(Rivi, 7).Value = Blokki.Layer             'Layer
      Cells(Rivi, 3).Value = Blokki.EffectiveName     'Blokin nimi
      Cells(Rivi, 3).ClearNotes                       'Tyhjennetään kommenttti
      Cells(Rivi, 3).AddComment Blokki.Name           'Blokin nimi
      
      If Blokki.HasAttributes Then                    'Tarkistetaan että blokilla on attribuutteja
        BlockArray = Blokki.GetAttributes             'Otetaan atribuutit talteen muuttujaan
        For j = 0 To UBound(BlockArray)
          Cells(1, 8 + j).Value = BlockArray(j).TagString
          Cells(1, 8 + j).ClearNotes
          Cells(1, 8 + j).AddComment Blokki.EffectiveName
          Cells(Rivi, 8 + j).Value = BlockArray(j).TextString
'          Cells(Rivi, OtsS(BlockArray(j).TagString)).Value = BlockArray(j).TextString
        Next j
        Rivi = Rivi + 1
      End If
    Else
      If Joukko(i).EntityName = "AcDbText" Then
        Set oText = Joukko(i)
        Cells(Rivi, 8).Value = oText.TextString
        Cells(Rivi, 5).Value = oText.InsertionPoint(0) 'XCord
        Cells(Rivi, 6).Value = oText.InsertionPoint(1) 'YCord
        Range(Cells(Rivi, 1), Cells(Rivi, 8)).Interior.ColorIndex = 8
      Else
        Set oMText = Joukko(i)
        Cells(Rivi, 8).Value = oMText.TextString
        Cells(Rivi, 5).Value = oMText.InsertionPoint(0) 'XCord
        Cells(Rivi, 6).Value = oMText.InsertionPoint(1) 'YCord
        Range(Cells(Rivi, 1), Cells(Rivi, 8)).Interior.ColorIndex = 8
      End If
      Rivi = Rivi + 1
    End If
  Next i
  If Loytyi = False Then oDOC.Close False
Next DocRivi
oACAD.Visible = True
oACAD.Preferences.System.SingleDocumentMode = Docmode
  Cells.EntireColumn.AutoFit
  'AppActivate "Excel"
  'Tyhjätään muuttujaobjektit
  Application.StatusBar = False 'Palautetaan käyttöön normaali tilatieto Excelin Statusbariin
  Set Blokki = Nothing
  Set Joukko = Nothing
  Set oDOC = Nothing
  'MsgBox "Valmis!"
End Sub
Public Sub VieDATA() 'Tämä vie datan excelistä AutoCADiin
' 3.1.2001 - VG
' 4.6.2002 - VG
' 7.3.2003 - VG
' 27.3.2003 - VG
' 19.1.2004 - VG
' 29.1.2004 - VG -> Attribuuttien nimien ottaminen huomioon
  Dim i As Long, j As Integer
  Ver = acNative ' Ver = 60
  
  If Sheets("Start").Ver2004.Value = True Then
    Ver = ac2004_dwg ' Ver = 24
  ElseIf Sheets("Start").Ver2007.Value = True Then
    Ver = ac2007_dwg ' Ver = 36
  ElseIf Sheets("Start").Ver2010.Value = True Then
    Ver = ac2010_dwg ' Ver = 48
  ElseIf Sheets("Start").Ver2013.Value = True Then
    Ver = ac2013_dwg ' Ver = 60
  End If
'Varmistetaan että data-sheeti on valittuna
  DATA.Select

'Otetaan yhteys AutoCADiin
  On Error Resume Next
  Set oACAD = GetObject(, "AutoCAD.Application") 'Koitetaan yhdistää AutoCADiin
  If Err <> 0 Then 'Käynnissä olevaa AutoCADiä ei löytynyt
    MsgBox "Käynnissä olevaa AutoCADiä ei löytynyt!", vbCritical, "Vie DATA"
    Set oACAD = Nothing
    Exit Sub
  End If
  On Error GoTo 0
  'Yhteys muodostettu käynnissä olevaan AutoCADiin ja oikea kuva on auki
  '---------- Aloitetaan viemään tietoja avoinna olevaan dokumenttiin ----------
  Dim oEntity As Object
  Dim oBlock As AcadBlockReference 'Blokkimuuttuja
  Dim BlockArray As Variant       'Array muuttuja Blockkia varten
  Dim BlockNimi As String         'Blokin nimi, joita etsitään
  Dim DWGName As String
  Dim oText As AcadText
  Dim oMText As AcadMText
  Dim Docmode As Boolean
  Docmode = oACAD.Preferences.System.SingleDocumentMode

  oACAD.Preferences.System.SingleDocumentMode = False
  i = 1
  Do
    i = i + 1
    If Cells(i, 4).Value = "" Then 'Viimeinen rivi Excel lomakkeessa
      If OliAuki = False Then
        If Not oDOC Is Nothing Then
'          If Ver = 0 Then
'            oDOC.SaveAs oDOC.FullName, odoc.
'          Else
            oDOC.SaveAs oDOC.FullName, Ver
            oDOC.Close False
 '         End If
        End If
      End If
      Exit Do
    Else
      If AvaaDoc(i) Then
        Application.StatusBar = "Viedään tietoa blokkiin: " & i - 1
        Set oEntity = oDOC.HandleToObject(Cells(i, 4).Text)
        If oEntity.EntityName = "AcDbBlockReference" Then 'Blokki
          Set oBlock = oEntity
          If oBlock.HasAttributes Then
            BlockArray = oBlock.GetAttributes   'muodostetaan referenssi atribuutteihin
            For j = 0 To UBound(BlockArray)    'Kirjoitetaan blokkiin atribuuttien arvot
              BlockArray(j).TextString = Cells(i, 8 + j).Text
'              BlockArray(j).TextString = Cells(i, EOtsS(BlockArray(j).TagString)).Text
            Next j
          End If
        Else
         If oEntity.EntityName = "AcDbText" Then
           Set oText = oEntity
           oText.TextString = Cells(i, 8).Value
         Else
           Set oMText = oEntity
           oMText.TextString = Cells(i, 8).Value
         End If
        End If
      End If
    End If
  Loop
  Aloitus.Activate
  oACAD.Preferences.System.SingleDocumentMode = Docmode
  AppActivate "Excel"
  Set oEntity = Nothing
  Set oDOC = Nothing
  Set BlockArray = Nothing
  Set oACAD = Nothing
  'MsgBox "Valmis!"
Application.StatusBar = False 'Palautetaan käyttöön normaali tilatieto Excelin Statusbariin
End Sub
Public Sub PoistaBlokit() 'Tämä poistaa valitut blokit AutoCAD kuvasta
  Dim i As Integer, j As Integer
  Dim Docmode As Boolean
  Dim oEntity As Object
  Dim DWGName As String
  Dim Rivi As Range
  Dim RiviNo As Long
  Dim Kaydyt As String

'Varmistetaan että data-sheeti on valittuna
  DATA.Select
'Otetaan yhteys AutoCADiin
  On Error Resume Next
  Set oACAD = GetObject(, "AutoCAD.Application") 'Koitetaan yhdistää AutoCADiin
  If Err <> 0 Then 'Käynnissä olevaa AutoCADiä ei löytynyt
    MsgBox "Käynnissä olevaa AutoCADiä ei löytynyt!", vbCritical, "Vie DATA"
    Set oACAD = Nothing
    Exit Sub
  End If
  On Error GoTo 0
'--- [ VARMISTETAAN ETTÄ DOKUMENTTI ON AUKI ] ---
  Docmode = oACAD.Preferences.System.SingleDocumentMode
  oACAD.Preferences.System.SingleDocumentMode = False
  
  For Each Rivi In Selection.Rows
    If InStr(Kaydyt, "|" & Rivi.row & "|") = 0 Then
      RiviNo = Rivi.row
      Kaydyt = Kaydyt & "|" & RiviNo & "|"
      If AvaaDoc(RiviNo) Then
        Application.StatusBar = "Tuhotaan objektia rivillä: " & Rivi.row
        Set oEntity = oDOC.HandleToObject(Cells(Rivi.row, 4).Text)
        oEntity.Delete
      End If
    End If
  Next
  If OliAuki = False Then
    If Not oDOC Is Nothing Then oDOC.SaveAs oDOC.FullName, Ver
  End If
  oACAD.Preferences.System.SingleDocumentMode = Docmode
  Application.StatusBar = False
  MsgBox "Valitut objektit tuhottiin"
  Set oEntity = Nothing
  Set oDOC = Nothing
  Set oACAD = Nothing
End Sub
Private Function OtsS(Nimi As String) As Integer
Dim i As Integer
Nimi = UCase(Nimi)
i = 7
Do
  If Cells(1, i).Value = "" Then
    Cells(1, i).Value = Nimi
    OtsS = i
    Exit Do
  ElseIf Cells(1, i).Value = Nimi Then
    OtsS = i
    Exit Do
  End If
  i = i + 1
Loop
End Function
Private Function EOtsS(Nimi As String) As Integer
Dim i As Integer
Nimi = UCase(Nimi)
i = 7
Do
  If Cells(1, i).Value = Nimi Then
    EOtsS = i
    Exit Do
  ElseIf Cells(1, i).Value = "" Then
    EOtsS = i
    Exit Do
  End If
  i = i + 1
Loop
End Function
Private Function AvaaDoc(Rivi As Long) As Boolean
Dim Doku As String
Dim EdDoku As String
Dim Tiedosto As String
Dim i As Integer
  Doku = (Cells(Rivi, 2).Value) & ".dwg"
  EdDoku = (Cells(Rivi - 1, 2).Value) & ".dwg"
  Tiedosto = Cells(Rivi, 1).Value & "\" & Doku
  'Tarkistetaan ensin onko haluttu dokumentti aktiivinen dokumentti
  If Not oDOC Is Nothing Then 'Jokin kuva on jo auki
    If LCase(oDOC.Name) = LCase(Doku) Then 'Kuva on käsittelyssä oleva kuva
      AvaaDoc = True
      Exit Function
    ElseIf LCase(oDOC.Name) = LCase(EdDoku) Then 'Edellinen kuva on avoinna oleva kuva
      If OliAuki = False Then
        On Error Resume Next
        oDOC.Close True
        If Err <> 0 Then
          Err.Clear
          MsgBox "Virhe talletettaessa piirustusta: " & oDOC.Name & vbCrLf & "Kuva saattaa olla jollakin auki.", vbCritical, "Vie tiedot"
          On Error GoTo 0
        End If
        On Error GoTo 0
      End If
    End If
  End If
'Haluttu kuva ei ollut jo käsittelyssä
'Tarkistetaan ettei haluttu kuva ole auki AutoCADissä
  OliAuki = False
  For i = 0 To oACAD.Documents.Count - 1
    If LCase(oACAD.Documents(i).Name) = LCase(Doku) Then 'Kuva on auki asetetaan se aktiiviseksi
      OliAuki = True
      oACAD.Documents(i).Activate
      Set oDOC = oACAD.ActiveDocument
      AvaaDoc = True
      Exit Function
    End If
  Next i
'Haluttu kuva ei ollut käsittelyssä eikä auki AutoCADissä, joten avataan se
  On Error Resume Next
  Set oDOC = oACAD.Documents.Open(Tiedosto)
  If Err <> 0 Then
    MsgBox "Virhe avattaessa piirustusta: " & Doku, vbCritical, "Vie tiedot"
    AvaaDoc = False
    Err.Clear
    On Error GoTo 0
    Exit Function
  Else
    AvaaDoc = True
  End If
  On Error GoTo 0
End Function
Sub Numerointi()
Dim Alku As String
Dim Jakso As Integer
Dim Vali As Integer
Dim i As Integer, j As Integer

  Aloitus.Tyhjenna.Value = True
  Aloitus.Nykyinen.Value = True
'Haetaan kuvasta
  TuoDATA True
  Alku = Aloitus.Range("D13").Value
'  Jakso = 8
  Vali = 2
  Cells.Sort Key1:=Range("E2"), Order1:=xlAscending, Key2:=Range("F2"), Order2:=xlDescending, Header:=xlYes, OrderCustom:=1, MatchCase:=False, Orientation:=xlTopToBottom
  i = 2
  j = Val(Alku)
  Do
    If Cells(i, 1).Value = "" Then Exit Do
    Cells(i, 12).Value = LNumero(j, Alku)
    If Right(CStr(j), 1) = "8" Then
      j = j + Vali
    End If
    j = j + 1
    i = i + 1
  Loop
  Aloitus.Range("D13").Value = LNumero(j, Alku)
  VieDATA
End Sub
Private Function LNumero(No As Integer, Alku As String) As String
  LNumero = CStr(No)
  Do
    If Len(LNumero) < Len(Alku) Then
      LNumero = "0" & LNumero
    Else
      Exit Do
    End If
  Loop
End Function

Sub RefNumerointi()
Dim vSivu As Integer
Dim Kirjain As String

Dim i As Integer, j As Integer

  Aloitus.Tyhjenna.Value = True
  Aloitus.Nykyinen.Value = True
  Kirjain = "A"
'Haetaan kuvasta
  TuoDATA True, "REFERENCE"
  vSivu = CInt(Aloitus.Range("D18").Value)
  
  Cells.Sort Key1:=Range("F2"), Order1:=xlDescending, Header:=xlYes, OrderCustom:=1, MatchCase:=False, Orientation:=xlTopToBottom
  i = 2
  Do
    If Cells(i, 1).Value = "" Then Exit Do
    Cells(i, 7).Value = "(" & vSivu & ":" & Kirjain & ") To Page " & vSivu
    Cells(i, 8).Value = "To Page " & vSivu + 1 & "(" & vSivu & ":" & Kirjain & ")"
    i = i + 1
    Kirjain = Chr(Asc(Kirjain) + 1)
  Loop
  Aloitus.Range("D18").Value = vSivu + 1
  VieDATA
End Sub
Function Lisaa(Nro As String, Maara As Long) As String
Dim Pit As Integer
Dim i As Integer
Pit = Len(Nro)
Nro = CStr(Val(Nro) + Maara)
For i = 1 To Pit - Len(Nro)
    Nro = "0" & Nro
Next i
Lisaa = Nro
End Function
Function Yhd(Alue As Range, Optional Merkki As String) As String
Dim Solu As Range, Teksti As String
For Each Solu In Alue
 If Teksti = "" Then
    Teksti = Solu.Value
 Else
    Teksti = Teksti & Merkki & Solu.Value
 End If
Next
Yhd = Teksti
End Function
