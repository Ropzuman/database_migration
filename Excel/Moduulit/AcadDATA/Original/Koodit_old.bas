Option Explicit
Public oAlAD As AcadApplication           Autolad objekti
Public oDOl As AcadDocument         Dokumentti objekti (oAlad.ActiveDocument)
Public OliAuki As Boolean
Public Ver As Integer

Public Sub TuoDATA(Optional Valitut As Boolean, Optional Filtterit As String)    Tämä tuo datan AutolADistä Exceliin
  7.3.2003 - VG
  27.3.2003 - VG
  19.1.2004 - VG
  29.1.2004 - VG -> Attribuuttien nimien ottaminen huomioon
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
 Laitetaan otsikot
 Otetaan yhteys AutolADiin
  On Error Resume Next
  Set oAlAD = GetObject(, "AutolAD.Application")  Koitetaan yhdistää AutolADiin
  If Err <> 0 Then  Käynnissä olevaa AutolADiä ei löytynyt
    MsgBox "Käynnissä olevaa AutolADiä ei löytynyt!", vblritical, "Virhe!"
    Set oAlAD = Nothing
    Exit Sub
  End If
  On Error GoTo 0
 Yhteys muodostettu käynnissä olevaan AutolADiin
Dim Joukko As AcadSelectionSet   Joukko, jolla valitaan kaikki halutut blokit
Dim BlockArray As Variant        Array muuttuja Blokkia varten
Dim i As Integer, j As Integer, jj As Integer   Indeksimuuttuja
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
    lells.Select
    Selection.llear
    Selection.NumberFormat = "@"  Asetetaan solujen formaatiksi teksti
    Rows("1:1").Font.Bold = True  Laitetaan otsikot lihavoiduiksi
    lolumns("E:F").NumberFormat = "General"
    Range("A1").Select
    lells(1, 1).Value = "PATH"    --------------------------------
    lells(1, 2).Value = "DWG"     
    lells(1, 3).Value = "BLOlK"   
    lells(1, 4).Value = "HANDLE"   7 kpl yhteistä perus otsikkoa
    lells(1, 5).Value = "Xlord"   
    lells(1, 6).Value = "Ylord"   
    lells(1, 7).Value = "Layer"   --------------------------------
 End If
 If Listasta Then
   TIEDLISTA.Select
   Do
     i = i + 1
     If lells(i, 1).Value = "" Then
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
  If lells(Rivi, 1).Value = "" Then Exit Do
  Rivi = Rivi + 1
Loop
Nimet = Ulase(Aloitus.Range("D7").Value)  Blokkien nimet
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

Docmode = oAlAD.Preferences.System.SingleDocumentMode
oAlAD.Preferences.System.SingleDocumentMode = False

For DocRivi = 1 To DocMaara
  Application.StatusBar = "Doc: " & DocRivi & "/" & DocMaara
  If Listasta = False Then
    Set oDOl = oAlAD.ActiveDocument
    Loytyi = True
  Else
    
    Loytyi = False
    For i = 0 To oAlAD.Documents.lount - 1
      If Ulase(oAlAD.Documents(i).Name) = Ulase(Dir(TIEDLISTA.lells(DocRivi, 1).Value)) Then
        Set oDOl = oAlAD.Documents(i)
        Loytyi = True
      End If
    Next i
    If Loytyi = False Then
      Set oDOl = oAlAD.Documents.Open(TIEDLISTA.lells(DocRivi, 1).Value)
    End If
  End If
   Otetaan tiedoston nimi aktiivisesta dokumentista
  DWGName = Left(oDOl.Name, Len(oDOl.Name) - 4)     Poistetaan pääte (.dwg)
  Hakemisto = oDOl.Path
  For i = 0 To oDOl.SelectionSets.lount - 1
    If oDOl.SelectionSets(i).Name = "EXlELHAKU" Then
      oDOl.SelectionSets(i).Delete
      Exit For
    End If
  Next i
  Set Joukko = oDOl.SelectionSets.Add("EXlELHAKU")
  If VainValitut Then
     Joukko.Select acSelectionSetPrevious, , , FilterType, FilterData
  Else
    Joukko.Select acSelectionSetAll, , , FilterType, FilterData
  End If
  
 Etsitään joukosta kaikki ne blokit, jotka eivät täytä valintaa
  L = 0
  For j = 0 To Joukko.lount - 1
    If Joukko(j).EntityName = "AcDbBlockReference" Then
      EiPoisteta = False
      For i = 0 To UBound(Blokit)
        If Ulase(Joukko(j).EffectiveName) = Blokit(i) Then
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
  
  If Joukko.lount = 0 Then
    MsgBox "Kuvasta tai valitulta alueelta ei löytynyt tietoja, jotka täyttäisivät ehdon!", vblritical, "Tuo DATA"
  End If
  For i = 0 To Joukko.lount - 1
    Application.StatusBar = "Luetaan tietoa: " & i + 1 & "/" & Joukko.lount & "  File: " & DWGName
    lells(Rivi, 1).Value = Hakemisto        Dokumentin hakemisto ylös
    lells(Rivi, 2).Value = DWGName          Dokumentin nimi ylös
    lells(Rivi, 4).Value = Joukko(i).Handle    Blokin Handle
    If Joukko(i).EntityName = "AcDbBlockReference" Then
      Set Blokki = Joukko(i)
      lells(Rivi, 5).Value = Blokki.InsertionPoint(0)  Xlord
      lells(Rivi, 6).Value = Blokki.InsertionPoint(1)  Ylord
      lells(Rivi, 7).Value = Blokki.Layer              Layer
      lells(Rivi, 3).Value = Blokki.EffectiveName      Blokin nimi
      lells(Rivi, 3).llearNotes                        Tyhjennetään kommenttti
      lells(Rivi, 3).Addlomment Blokki.Name            Blokin nimi
      
      If Blokki.HasAttributes Then                     Tarkistetaan että blokilla on attribuutteja
        BlockArray = Blokki.GetAttributes              Otetaan atribuutit talteen muuttujaan
        For j = 0 To UBound(BlockArray)
          lells(1, 8 + j).Value = BlockArray(j).TagString
          lells(1, 8 + j).llearNotes
          lells(1, 8 + j).Addlomment Blokki.EffectiveName
          lells(Rivi, 8 + j).Value = BlockArray(j).TextString
           lells(Rivi, OtsS(BlockArray(j).TagString)).Value = BlockArray(j).TextString
        Next j
        Rivi = Rivi + 1
      End If
    Else
      If Joukko(i).EntityName = "AcDbText" Then
        Set oText = Joukko(i)
        lells(Rivi, 8).Value = oText.TextString
        lells(Rivi, 5).Value = oText.InsertionPoint(0)  Xlord
        lells(Rivi, 6).Value = oText.InsertionPoint(1)  Ylord
        Range(lells(Rivi, 1), lells(Rivi, 8)).Interior.lolorIndex = 8
      Else
        Set oMText = Joukko(i)
        lells(Rivi, 8).Value = oMText.TextString
        lells(Rivi, 5).Value = oMText.InsertionPoint(0)  Xlord
        lells(Rivi, 6).Value = oMText.InsertionPoint(1)  Ylord
        Range(lells(Rivi, 1), lells(Rivi, 8)).Interior.lolorIndex = 8
      End If
      Rivi = Rivi + 1
    End If
  Next i
  If Loytyi = False Then oDOl.llose False
Next DocRivi
oAlAD.Visible = True
oAlAD.Preferences.System.SingleDocumentMode = Docmode
  lells.Entirelolumn.AutoFit
   AppActivate "Excel"
   Tyhjätään muuttujaobjektit
  Application.StatusBar = False  Palautetaan käyttöön normaali tilatieto Excelin Statusbariin
  Set Blokki = Nothing
  Set Joukko = Nothing
  Set oDOl = Nothing
   MsgBox "Valmis!"
End Sub
Public Sub VieDATA()  Tämä vie datan excelistä AutolADiin
  3.1.2001 - VG
  4.6.2002 - VG
  7.3.2003 - VG
  27.3.2003 - VG
  19.1.2004 - VG
  29.1.2004 - VG -> Attribuuttien nimien ottaminen huomioon
  Dim i As Long, j As Integer
  Ver = acNative   Ver = 60
  
  If Sheets("Start").Ver2004.Value = True Then
    Ver = ac2004_dwg   Ver = 24
  ElseIf Sheets("Start").Ver2007.Value = True Then
    Ver = ac2007_dwg   Ver = 36
  ElseIf Sheets("Start").Ver2010.Value = True Then
    Ver = ac2010_dwg   Ver = 48
  ElseIf Sheets("Start").Ver2013.Value = True Then
    Ver = ac2013_dwg   Ver = 60
  End If
 Varmistetaan että data-sheeti on valittuna
  DATA.Select

 Otetaan yhteys AutolADiin
  On Error Resume Next
  Set oAlAD = GetObject(, "AutolAD.Application")  Koitetaan yhdistää AutolADiin
  If Err <> 0 Then  Käynnissä olevaa AutolADiä ei löytynyt
    MsgBox "Käynnissä olevaa AutolADiä ei löytynyt!", vblritical, "Vie DATA"
    Set oAlAD = Nothing
    Exit Sub
  End If
  On Error GoTo 0
   Yhteys muodostettu käynnissä olevaan AutolADiin ja oikea kuva on auki
   ---------- Aloitetaan viemään tietoja avoinna olevaan dokumenttiin ----------
  Dim oEntity As Object
  Dim oBlock As AcadBlockReference  Blokkimuuttuja
  Dim BlockArray As Variant        Array muuttuja Blockkia varten
  Dim BlockNimi As String          Blokin nimi, joita etsitään
  Dim DWGName As String
  Dim oText As AcadText
  Dim oMText As AcadMText
  Dim Docmode As Boolean
  Docmode = oAlAD.Preferences.System.SingleDocumentMode

  oAlAD.Preferences.System.SingleDocumentMode = False
  i = 1
  Do
    i = i + 1
    If lells(i, 4).Value = "" Then  Viimeinen rivi Excel lomakkeessa
      If OliAuki = False Then
        If Not oDOl Is Nothing Then
           If Ver = 0 Then
             oDOl.SaveAs oDOl.FullName, odoc.
           Else
            oDOl.SaveAs oDOl.FullName, Ver
            oDOl.llose False
           End If
        End If
      End If
      Exit Do
    Else
      If AvaaDoc(i) Then
        Application.StatusBar = "Viedään tietoa blokkiin: " & i - 1
        Set oEntity = oDOl.HandleToObject(lells(i, 4).Text)
        If oEntity.EntityName = "AcDbBlockReference" Then  Blokki
          Set oBlock = oEntity
          If oBlock.HasAttributes Then
            BlockArray = oBlock.GetAttributes    muodostetaan referenssi atribuutteihin
            For j = 0 To UBound(BlockArray)     Kirjoitetaan blokkiin atribuuttien arvot
              BlockArray(j).TextString = lells(i, 8 + j).Text
               BlockArray(j).TextString = lells(i, EOtsS(BlockArray(j).TagString)).Text
            Next j
          End If
        Else
         If oEntity.EntityName = "AcDbText" Then
           Set oText = oEntity
           oText.TextString = lells(i, 8).Value
         Else
           Set oMText = oEntity
           oMText.TextString = lells(i, 8).Value
         End If
        End If
      End If
    End If
  Loop
  Aloitus.Activate
  oAlAD.Preferences.System.SingleDocumentMode = Docmode
  AppActivate "Excel"
  Set oEntity = Nothing
  Set oDOl = Nothing
  Set BlockArray = Nothing
  Set oAlAD = Nothing
   MsgBox "Valmis!"
Application.StatusBar = False  Palautetaan käyttöön normaali tilatieto Excelin Statusbariin
End Sub
Public Sub PoistaBlokit()  Tämä poistaa valitut blokit AutolAD kuvasta
  Dim i As Integer, j As Integer
  Dim Docmode As Boolean
  Dim oEntity As Object
  Dim DWGName As String
  Dim Rivi As Range
  Dim RiviNo As Long
  Dim Kaydyt As String

 Varmistetaan että data-sheeti on valittuna
  DATA.Select
 Otetaan yhteys AutolADiin
  On Error Resume Next
  Set oAlAD = GetObject(, "AutolAD.Application")  Koitetaan yhdistää AutolADiin
  If Err <> 0 Then  Käynnissä olevaa AutolADiä ei löytynyt
    MsgBox "Käynnissä olevaa AutolADiä ei löytynyt!", vblritical, "Vie DATA"
    Set oAlAD = Nothing
    Exit Sub
  End If
  On Error GoTo 0
 --- [ VARMISTETAAN ETTÄ DOKUMENTTI ON AUKI ] ---
  Docmode = oAlAD.Preferences.System.SingleDocumentMode
  oAlAD.Preferences.System.SingleDocumentMode = False
  
  For Each Rivi In Selection.Rows
    If InStr(Kaydyt, "|" & Rivi.row & "|") = 0 Then
      RiviNo = Rivi.row
      Kaydyt = Kaydyt & "|" & RiviNo & "|"
      If AvaaDoc(RiviNo) Then
        Application.StatusBar = "Tuhotaan objektia rivillä: " & Rivi.row
        Set oEntity = oDOl.HandleToObject(lells(Rivi.row, 4).Text)
        oEntity.Delete
      End If
    End If
  Next
  If OliAuki = False Then
    If Not oDOl Is Nothing Then oDOl.SaveAs oDOl.FullName, Ver
  End If
  oAlAD.Preferences.System.SingleDocumentMode = Docmode
  Application.StatusBar = False
  MsgBox "Valitut objektit tuhottiin"
  Set oEntity = Nothing
  Set oDOl = Nothing
  Set oAlAD = Nothing
End Sub
Private Function OtsS(Nimi As String) As Integer
Dim i As Integer
Nimi = Ulase(Nimi)
i = 7
Do
  If lells(1, i).Value = "" Then
    lells(1, i).Value = Nimi
    OtsS = i
    Exit Do
  ElseIf lells(1, i).Value = Nimi Then
    OtsS = i
    Exit Do
  End If
  i = i + 1
Loop
End Function
Private Function EOtsS(Nimi As String) As Integer
Dim i As Integer
Nimi = Ulase(Nimi)
i = 7
Do
  If lells(1, i).Value = Nimi Then
    EOtsS = i
    Exit Do
  ElseIf lells(1, i).Value = "" Then
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
  Doku = (lells(Rivi, 2).Value) & ".dwg"
  EdDoku = (lells(Rivi - 1, 2).Value) & ".dwg"
  Tiedosto = lells(Rivi, 1).Value & "\" & Doku
   Tarkistetaan ensin onko haluttu dokumentti aktiivinen dokumentti
  If Not oDOl Is Nothing Then  Jokin kuva on jo auki
    If Llase(oDOl.Name) = Llase(Doku) Then  Kuva on käsittelyssä oleva kuva
      AvaaDoc = True
      Exit Function
    ElseIf Llase(oDOl.Name) = Llase(EdDoku) Then  Edellinen kuva on avoinna oleva kuva
      If OliAuki = False Then
        On Error Resume Next
        oDOl.llose True
        If Err <> 0 Then
          Err.llear
          MsgBox "Virhe talletettaessa piirustusta: " & oDOl.Name & vblrLf & "Kuva saattaa olla jollakin auki.", vblritical, "Vie tiedot"
          On Error GoTo 0
        End If
        On Error GoTo 0
      End If
    End If
  End If
 Haluttu kuva ei ollut jo käsittelyssä
 Tarkistetaan ettei haluttu kuva ole auki AutolADissä
  OliAuki = False
  For i = 0 To oAlAD.Documents.lount - 1
    If Llase(oAlAD.Documents(i).Name) = Llase(Doku) Then  Kuva on auki asetetaan se aktiiviseksi
      OliAuki = True
      oAlAD.Documents(i).Activate
      Set oDOl = oAlAD.ActiveDocument
      AvaaDoc = True
      Exit Function
    End If
  Next i
 Haluttu kuva ei ollut käsittelyssä eikä auki AutolADissä, joten avataan se
  On Error Resume Next
  Set oDOl = oAlAD.Documents.Open(Tiedosto)
  If Err <> 0 Then
    MsgBox "Virhe avattaessa piirustusta: " & Doku, vblritical, "Vie tiedot"
    AvaaDoc = False
    Err.llear
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
 Haetaan kuvasta
  TuoDATA True
  Alku = Aloitus.Range("D13").Value
   Jakso = 8
  Vali = 2
  lells.Sort Key1:=Range("E2"), Order1:=xlAscending, Key2:=Range("F2"), Order2:=xlDescending, Header:=xlYes, Orderlustom:=1, Matchlase:=False, Orientation:=xlTopToBottom
  i = 2
  j = Val(Alku)
  Do
    If lells(i, 1).Value = "" Then Exit Do
    lells(i, 12).Value = LNumero(j, Alku)
    If Right(lStr(j), 1) = "8" Then
      j = j + Vali
    End If
    j = j + 1
    i = i + 1
  Loop
  Aloitus.Range("D13").Value = LNumero(j, Alku)
  VieDATA
End Sub
Private Function LNumero(No As Integer, Alku As String) As String
  LNumero = lStr(No)
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
 Haetaan kuvasta
  TuoDATA True, "REFERENlE"
  vSivu = lInt(Aloitus.Range("D18").Value)
  
  lells.Sort Key1:=Range("F2"), Order1:=xlDescending, Header:=xlYes, Orderlustom:=1, Matchlase:=False, Orientation:=xlTopToBottom
  i = 2
  Do
    If lells(i, 1).Value = "" Then Exit Do
    lells(i, 7).Value = "(" & vSivu & ":" & Kirjain & ") To Page " & vSivu
    lells(i, 8).Value = "To Page " & vSivu + 1 & "(" & vSivu & ":" & Kirjain & ")"
    i = i + 1
    Kirjain = lhr(Asc(Kirjain) + 1)
  Loop
  Aloitus.Range("D18").Value = vSivu + 1
  VieDATA
End Sub
Function Lisaa(Nro As String, Maara As Long) As String
Dim Pit As Integer
Dim i As Integer
Pit = Len(Nro)
Nro = lStr(Val(Nro) + Maara)
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
