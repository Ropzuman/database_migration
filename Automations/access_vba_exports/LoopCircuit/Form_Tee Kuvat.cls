VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_Tee Kuvat"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit
Dim Ajastin As Long
Dim min As Integer, sec As Integer
Dim OSMODE As Integer
Dim LokiPituus As Long
Dim oAcad As AcadApplication
Public oDoc As AcadDocument
Private Sub Odota(Aika As Integer)
'Tämä aliohjema odottaa määritellyn ajan sekunnin sadasosina
Dim odotus As Double
odotus = CDbl(Timer) + CDbl((Aika / 100))
Do While odotus > Timer
  DoEvents
Loop
End Sub
Private Sub Tila(Tiedosto As String)
  sec = CInt(Timer - Ajastin)
  min = sec \ 60
  sec = sec Mod 60
  SysCmd acSysCmdSetStatus, "Käsittelee tiedostoa: " & Tiedosto & "  Aika: " & min & " min  " & sec & " sek"
End Sub
Private Sub Form_Load()
  Loki.Value = ""
End Sub
Private Sub Form_Unload(Cancel As Integer)
  DoCmd.Hourglass False
  SysCmd acSysCmdClearStatus
End Sub
Private Sub FormClose_Click()
 DoCmd.Close
End Sub
Private Sub HaeTekstit_Click()
Dim Tied As String
Dim Polku As String
Dim Nimi As String
Dim Taulu As Recordset
Dim oEnt As acadentity
Dim oAcad As AutoCAD.AcadApplication

Ajastin = Timer
Set oAcad = CreateObject("AutoCAD.Application")

Polku = PohjaHakem.Value
If Right(Polku, 1) <> "\" Then Polku = Polku & "\"
Tied = Dir(Polku & "*.dwg")
CurrentDb.Execute "DELETE * FROM Blokit WHERE PATH='" & PohjaHakem.Value & "'"
Set Taulu = CurrentDb.OpenRecordset("Blokit")
Do While Tied <> ""
  Nimi = Left(Tied, Len(Tied) - 4)
  If UCase(Left(Tied, 4)) <> "BASE" And UCase(Left(Tied, 8)) <> "BLOCKMAP" Then
    Set oDoc = oAcad.Documents.Open(Polku & Tied)
    Tila Tied
    For Each oEnt In oDoc.ModelSpace
      If oEnt.ObjectName = "AcDbAttributeDefinition" Then
        Taulu.AddNew
        Taulu.Fields("Block").Value = Nimi
        Taulu.Fields("ATAG").Value = oEnt.TagString
        Taulu.Fields("VALUE").Value = oEnt.TextString
        Taulu.Fields("Path").Value = Polku
        Taulu.Fields("VALUE").Value = oEnt.TextString
        Taulu.Update
      End If
    Next
    oDoc.Close False
  End If
  Tied = Dir
Loop
Taulu.Close
oAcad.Quit
Set oAcad = Nothing
MsgBox "Valmis", vbInformation, "Päivitä tekstit"
Set Taulu = Nothing
End Sub
Private Sub HaeValitutTekstit_Click()
Dim Tied As String
Dim Polku As String
Dim Nimi As String
Dim Taulu As Recordset
Dim oEnt As acadentity
Dim Blokki As AcadBlockReference
Dim oAcad As AutoCAD.AcadApplication
Dim oDoc As AutoCAD.AcadDocument
Dim Joukko As AcadSelectionSet
Dim BlockArray As Variant       'Array muuttuja Blokkia varten
Dim i As Integer, j As Integer  'Indeksimuuttuja
Dim FilterType(0) As Integer
Dim FilterData(0) As Variant
Dim FilterType2(0) As Integer
Dim FilterData2(0) As Variant
Dim DWGName As String
Dim Hakemisto As String
Dim Rivi As Integer

          FilterType(0) = 0
          FilterData(0) = "INSERT"

Ajastin = Timer


Set oAcad = GetObject(, "AutoCAD.Application")
Polku = PohjaHakem.Value
If Right(Polku, 1) <> "\" Then Polku = Polku & "\"
Tied = Dir(Polku & "*.dwg")

'CurrentDb.Execute "DELETE * FROM Blokit"

Set Taulu = CurrentDb.OpenRecordset("Blokit")

  Nimi = Left(Tied, Len(Tied) - 4)
    
    Set oDoc = oAcad.ActiveDocument
  DWGName = Left(oDoc.Name, Len(oDoc.Name) - 4)    'Poistetaan pääte (.dwg)
  Hakemisto = oDoc.Path
  Set Joukko = oDoc.ActiveSelectionSet
  Joukko.Clear


     Joukko.Select acSelectionSetPrevious, , , FilterType, FilterData


     If Joukko.Count = 0 Then
       MsgBox "Kuvasta ei löytynyt valittuja blokkeja, jotka täyttäisivät ehdon!", vbCritical, "Tuo DATA"
     End If '     Joukko.Select acSelectionSetPrevious,,,
    
  For i = 0 To Joukko.Count - 1
    If Joukko(i).EntityName = "AcDbBlockReference" Then
      Set Blokki = Joukko(i)
    
    CurrentDb.Execute "DELETE * FROM Blokit WHERE Block='" & Blokki.Name & "'"

     
      If Blokki.HasAttributes Then            'Tarkistetaan että blokilla on attribuutteja
        BlockArray = Blokki.GetAttributes     'Otetaan atribuutit talteen muuttujaan
        For j = 0 To UBound(BlockArray)
        Taulu.AddNew
        Taulu.Fields("Block").Value = Blokki.Name
        Taulu.Fields("ATAG").Value = BlockArray(j).TagString
        Taulu.Fields("VALUE").Value = BlockArray(j).TextString
        Taulu.Fields("Path").Value = Hakemisto
        Taulu.Update
        Next j
        Rivi = Rivi + 1
      End If
      End If
 Next i
    
Taulu.Close
Set oAcad = Nothing
MsgBox "Valmis", vbInformation, "Päivitä tekstit"
Set Taulu = Nothing
End Sub


Private Sub TeeKuvat_Click()
'Tämä avaa kuvan insertoi kuvaan blokin/blokit, vaihtaa attribuuttien tiedot ja tallentaa kuvan
Dim db As Object ' Database
Dim Kuvat As Object 'Recordset
Dim DbBlokit As Object ' Recordset
Dim Tiedot As Object ' Recordset
Dim FSO As New FileSystemObject
Dim Tiedosto As String
Dim Blokki As Object ' AcadBlockReference
Dim Attribuutit As Variant
Dim Paikka As Integer
Dim BNimi As String
Dim i As Integer, a As Integer
Dim Sql As String
Dim OtsTaulukko As String
Dim BlockTaulukko As String
Dim BlockTaulukkoTyhjennys As String
Dim Vastaus As Integer
Dim KuvaNimi As String
Dim oAcad As Object 'AutoCAD.AcadApplication
Dim Os As Long, Os2 As Long
Dim Kysely As String
Dim Burst As Boolean

'Set last_criteria =
last_used = 0

Set db = CurrentDb
Ajastin = Timer
LokiPituus = 0
OtsTaulukko = TOtsTaulukko.Value
Burst = TBurst.Value

BlockTaulukko = TBlockTaulukko.Value
BlockTaulukkoTyhjennys = TBlockTaulukko.Value

Vastaus = MsgBox("Päivitetäänkö " & BlockTaulukkoTyhjennys & "-taulukko?", vbYesNoCancel, "Tee kuvat")
If Vastaus = vbYes Then 'Päivitetään Kuviin1 taulukko
  db.Execute "DELETE * FROM " & BlockTaulukkoTyhjennys
  
  db.Execute "DELETE * FROM Kuviin_Device_Info_Tbl" '"Tyhjennetään aputaulukko" MRu
  Os = 1
  Do
    Os2 = InStr(Os, TKyselyt.Value, vbCrLf)
    If Os2 = 0 Then
      Kysely = Mid(TKyselyt.Value, Os)
      db.Execute Kysely
      LisaaLokiin "- Ajettiin kysely " & Kysely & "'"
      Exit Do
    Else
      Kysely = Mid(TKyselyt.Value, Os, Os2 - Os)
      db.Execute Kysely
      LisaaLokiin "- Ajettiin kysely '" & Kysely & "'"
      Os = Os2 + 2
    End If
  Loop
  
  LisaaLokiin "-Päivitettiin " & BlockTaulukkoTyhjennys & "-taulukko"
ElseIf Vastaus = vbCancel Then
  Set db = Nothing
  Exit Sub
End If
If IsNull(KuvaHakem.Value) Then
  KuvaHakem.Value = PohjaHakem.Value
  LisaaLokiin "-Kuva hakemistoa ei oltu määritelty, talletetaan oletuksena pohjahakemistoon.."
End If
If FSO.FolderExists(PohjaHakem.Value) = False Then 'Hakemisto ei ollut laillinen annetaan virheilmoitus ja keskeytetään
  LisaaLokiin "Virhe: Hakemistoa ei löytynyt (" & KuvaHakem.Value & ")"
  MsgBox "Hakemistoa: " & vbCrLf & PohjaHakem.Value & vbCrLf & " ei löytynyt!" & vbCrLf & "Valitse hakemisto ja yritä uudelleen.", vbCritical, "Tee kuvat"
ElseIf FSO.FolderExists(KuvaHakem.Value) = False Then 'Hakemisto ei ollut laillinen annetaan virheilmoitus ja keskeytetään
  LisaaLokiin "Virhe: Hakemistoa ei löytynyt (" & KuvaHakem.Value & ")"
  MsgBox "Hakemistoa: " & vbCrLf & KuvaHakem.Value & vbCrLf & " ei löytynyt!" & vbCrLf & "Valitse hakemisto ja yritä uudelleen.", vbCritical, "Tee kuvat"
Else
  If Right(PohjaHakem.Value, 1) <> "\" Then PohjaHakem.Value = PohjaHakem.Value & "\"
  If Right(KuvaHakem.Value, 1) <> "\" Then KuvaHakem.Value = KuvaHakem.Value & "\"
  Set Kuvat = db.OpenRecordset("SELECT * FROM " & OtsTaulukko)
  If Kuvat.EOF Then 'Taulukossa ei ollut tarvittavia tietoja, annetaan virhe ja keskeytetään
    LisaaLokiin "Virhe: Taulukosta (" & OtsTaulukko & ") ei löytynyt kuvia"
    MsgBox "Taulukosta/kyselystä """ & OtsTaulukko & """ ei löytyny tietoja!" & vbCrLf & "Valitse taulukko ja yritä uudelleen.", vbCritical, "Tee kuvat"
  Else 'OK hakemisto löytyi ja taulukossa oli tietoja, ruvetaan käymään taulukon kuvia läpi
    Set oAcad = CreateObject("AutoCAD.Application")
    Do While Not Kuvat.EOF
      Tiedosto = PohjaHakem.Value & Kuvat.Fields("BaseDWG").Value
      If LCase(Right(Tiedosto, 4)) <> ".dwg" Then 'Varmistetaan että tiedostonimen pääte on .dwg
        Tiedosto = Tiedosto & ".dwg"
      End If
      If FSO.FileExists(Tiedosto) Then 'Tiedosto löytyi voidaan avata se AutoCADissä
        Tila Kuvat.Fields("FileID").Value
        Set oDoc = oAcad.Documents.Open(Tiedosto)   'Avataan uusi dokumentti
        OSMODE = oDoc.GetVariable("OSMODE") 'Tallennetaan OSNAP-asetus
        oDoc.SetVariable "OSMODE", 0 'Asetetaan OSNAP tila pois päältä
        'Haetaan insertointipaikat kuvasta
        If HaeIPoints(Kuvat.Fields("BaseDWG").Value) Then
          'Vaihdetaan otsikkotaulun tiedot
          If VaihdaOtsikkotiedot(Kuvat, "", "", "") Then 'Jos Otsikkotiedot saatiin vaihdettua OK niin kannattaa mennä eteenpäin
            'Valitaan taulukosta kuvaan insertoitavat blokit
            Sql = "SELECT Block, Blocknumber FROM " & BlockTaulukko & " WHERE FileID='" & Kuvat.Fields("FileID").Value & "' GROUP BY Block, Blocknumber HAVING (Blocknumber Is Not Null);"
            Set DbBlokit = db.OpenRecordset(Sql)
            Do While Not DbBlokit.EOF
              If FSO.FileExists(PohjaHakem.Value & DbBlokit.Fields("Block").Value & ".dwg") Then 'Blokki löytyi
                Sql = "SELECT * FROM " & BlockTaulukko & " WHERE FileID='" & Kuvat.Fields("FileID").Value & "' AND Block='" & DbBlokit.Fields("Block").Value & "'  AND Blocknumber='" & DbBlokit.Fields("Blocknumber").Value & "' ORDER BY AttrTag;"
                Set Tiedot = db.OpenRecordset(Sql)
                Paikka = CInt(DbBlokit.Fields("Blocknumber").Value)
                BNimi = PohjaHakem.Value & DbBlokit.Fields("Block").Value & ".dwg"
                'Insertoidaan blokki (Paikka, Blokin nimi, ScaleX, ScaleY, ScaleZ, Rotation)
                Set Blokki = oDoc.ModelSpace.InsertBlock(Paikat(Paikka).Pisteet, BNimi, 1, 1, 1, 0)
                'Muutetaan blokin attribuutit
                Attribuutit = Blokki.GetAttributes 'Haetaan attribuutit
                For i = 0 To UBound(Attribuutit) 'Käydään läpi kaikki attribuutit
                  Tiedot.FindFirst "AttrTag = '" & Attribuutit(i).TagString & "'" 'Etsitään tietokannasta attribuutin tagi
                  If Tiedot.NoMatch = False Then 'Jos löytyi, niin päivitetään tieto blokkiin
                    If IsNull(Tiedot.Fields("Teksti").Value) = False Then
                      Attribuutit(i).TextString = Tiedot.Fields("Teksti").Value
                    Else
                      Attribuutit(i).TextString = ""
                    End If
                  Else
                    LisaaLokiin "Virhe: Blokin " & DbBlokit.Fields("Block").Value & " attribuutille (" & Attribuutit(i).TagString & ") ei löytynyt arvoa tietokannasta"
                  End If
                Next i
                'Burst
                If Burst Then
                    oDoc.SendCommand "burst" & VBA.vbCr & "last" & VBA.vbCr & VBA.vbCr
                End If
              Else
                LisaaLokiin "Virhe: Blokkia " & DbBlokit.Fields("Block").Value & " ei löytynyt"
              End If
              DbBlokit.MoveNext
            Loop
          End If
        End If
        KuvaNimi = KuvaHakem.Value & Kuvat.Fields("FileID").Value
        If Right(LCase(KuvaNimi), 4) <> ".dwg" Then KuvaNimi = KuvaNimi & ".dwg"
        If FSO.FileExists(KuvaNimi) Then
          FSO.DeleteFile KuvaNimi
        End If
        oDoc.SetVariable "OSMODE", OSMODE 'Palautetaan OSNAP tila siihen joka se oli ennen sen poistamista pois päältä
        oDoc.SaveAs KuvaNimi 'Talletetaan kuva
        oDoc.Close
      Else 'Tiedostoa ei löytynyt
        LisaaLokiin "Virhe: Määriteltyä tiedostoa ei löytynyt (" & Tiedosto & ")"
      End If
      Kuvat.MoveNext
    Loop
    MsgBox "Kuvat tehty", vbOKOnly, "Valmis"
  End If
End If
oAcad.Quit
Loppu:

DoCmd.Hourglass False
SysCmd acSysCmdClearStatus

Set oAcad = Nothing
Set db = Nothing
Set Kuvat = Nothing
Set Tiedot = Nothing
Set FSO = Nothing
Set Blokki = Nothing
Set DbBlokit = Nothing
End Sub
Private Function HaeIPoints(ByVal Kuva As String) As Boolean
Dim Virhe As Boolean
Dim Blokki As AcadBlockReference ' Object
Virhe = False
ReDim Paikat(1 To 1)
Dim iPaikka As Variant
Dim Attribuutit As Variant
Dim Numero As Integer
Dim i As Integer
Dim Joukko As AcadSelectionSet
Dim FilterType(0) As Integer
Dim FilterData(0) As Variant
Dim Paikkablokki As String

Paikkablokki = TPaikkaBlokki.Value

FilterType(0) = 2             'Entityn nimi filtterinä
FilterData(0) = Paikkablokki  'Blokin nimi
Set Joukko = oDoc.ActiveSelectionSet
Joukko.Clear
Joukko.Select acSelectionSetAll, , , FilterType, FilterData
If Joukko.Count > 0 Then Set Blokki = Joukko(0)
      
  If Not Blokki Is Nothing Then
    If Blokki.HasAttributes Then 'Tarkistetaan vielä että blokilla on attribuutteja
      Attribuutit = Blokki.GetAttributes 'Otetaan attribuutit muuttujaan
      For i = 0 To UBound(Attribuutit) 'Käydään läpi attribuutit
        Numero = CInt(Attribuutit(i).TextString) 'Oletetaan että attribuutin teksti on numero ja muunnetaan se numeroksi
        If UBound(Paikat) < Numero Then
          ReDim Preserve Paikat(1 To Numero)     'Varmistetaan että Paikat Arrayssa on tilaa tämän paikan tiedolle
        End If
        iPaikka = Attribuutit(i).InsertionPoint  'Poimitaan attribuutista paikkatieto
        Paikat(Numero).Pisteet(0) = iPaikka(0)   'X:  Laitetaan paikan kordinaatit Paikat Arrayhyn kohtaan jonka attribuutin numero määritteli
        Paikat(Numero).Pisteet(1) = iPaikka(1)   'Y
        Paikat(Numero).Pisteet(2) = iPaikka(2)   'Z
      Next i
    Else
      LisaaLokiin "Virhe: Insertointipaikat sisältävästä blokista ei löytynyt attribuutteja (Kuva:" & Kuva & ")"
      Virhe = True
    End If
  Else
    LisaaLokiin "Virhe: Kuvasta " & Kuva & " ei löytynyt insertointipaikkoja sisältävää blokkia (" & Paikkablokki & ")"
    Virhe = True
  End If
  Set Blokki = Nothing
  HaeIPoints = Not Virhe
End Function

Private Function VaihdaOtsikkotiedot(Taulu As Recordset, TallNimi As String, NextPage As String, CurrPage As String) As Boolean
    Dim Attribuutit As Variant
    Dim Virhe As Boolean
    Dim i As Integer, j As Integer, b As Integer
    Dim Blokki As AcadBlockReference
    Dim tagi As String
    Dim tagiValue As String
    Dim Loytyi As Boolean
    Dim FilterType(0) As Integer
    Dim FilterData(0) As Variant
    Dim Titleblokki As String
    Dim Joukko As AcadSelectionSet
    
    Virhe = False
    
    Titleblokki = TTitleBlokki.Value
  
    FilterType(0) = 2             'Entityn nimi filtterinä
    FilterData(0) = Titleblokki  'Blokin nimi
    Set Joukko = oDoc.ActiveSelectionSet
    Joukko.Clear
    Joukko.Select acSelectionSetAll, , , FilterType, FilterData
    If Joukko.Count > 0 Then
      For b = 0 To Joukko.Count - 1
        Set Blokki = Joukko(b)
        If Not Blokki Is Nothing Then
          If Blokki.HasAttributes Then 'Tarkistetaan vielä että blokilla on attribuutteja
            
            Attribuutit = Blokki.GetAttributes 'Otetaan attribuutit muuttujaan
            
            For i = 0 To UBound(Attribuutit) 'Käydään läpi attribuutit
              tagi = Attribuutit(i).TagString
              tagiValue = Attribuutit(i).TextString
              Loytyi = False
              For j = 0 To Taulu.Fields.Count - 1
   '             If taulu.Fields(j).Name = tagi Then
    
                    If Taulu.Fields(j).Name = tagiValue Then
                    Attribuutit(i).TextString = IIf(IsNull(Taulu.Fields(tagiValue).Value), "", Taulu.Fields(tagiValue).Value)
                           
                    '    If taulu.Fields(j).Name = "CAD_FILE" Then
                   '     Attribuutit(i).TextString = IIf(IsNull(taulu.Fields(TallNimi).Value), "", taulu.Fields(TallNimi).Value)
                   
                   '     If taulu.Fields(j).Name = "CAD_FILE" Then
                   '     Attribuutit(i).TextString = IIf(IsNull(TallNimi), "", TallNimi)
                   '     ElseIf taulu.Fields(j).Name = "Page" Then
                   '     Attribuutit(i).TextString = IIf(IsNull(CurrPage), "", CurrPage)
                   '     ElseIf taulu.Fields(j).Name = "NextPage" Then
                   '     Attribuutit(i).TextString = IIf(IsNull(NextPage), "", NextPage)
        '            Else
       '                Attribuutit(i).TextString = IIf(IsNull(taulu.Fields(tagi).Value), "", taulu.Fields(tagi).Value)
                  '      End If
                    
                     '   End If
                  
                    ElseIf Taulu.Fields(j).Name = tagi Then
                        'If taulu.Fields(j).Name = "CAD_FILE" Then
                            Attribuutit(i).TextString = IIf(IsNull(Taulu.Fields(tagi).Value), "", Taulu.Fields(tagi).Value)
                        'End If
                    
                    Loytyi = True
                    Exit For
                
                    End If
                   
              
              Next j
              If Loytyi = False Then
                LisaaLokiin "- Taulukosta ei löytynyt tietoa otsikkotaulun attribuutille: " & tagiValue
              End If
            Next i
          Else
            LisaaLokiin "Virhe: Otsikkoblokista ei löytynyt attribuutteja (Kuva: " & Taulu.Fields("BaseDWG").Value & ")"
            Virhe = True
          End If
        Else 'Selection set ei sisältänyt blokkia
          LisaaLokiin "Virhe: Kuvasta " & Taulu.Fields("BaseDWG").Value & "ei löytynyt otsikkotietoja sisältävää blokkia (" & Titleblokki & ")"
          Virhe = True
        End If
      Next b
    End If
    Set Blokki = Nothing
    VaihdaOtsikkotiedot = Not Virhe
End Function


Private Sub LisaaLokiin(Tieto As String)
'  Loki.SelLength = 0
  'Loki.SetFocus
 ' If LokiPituus > 0 Then
 '   Loki.Value = Loki.Value & vbCrLf & Tieto
 ' Else
'    Loki.Value = Tieto
'    LokiPituus = 1
'  End If
'  Loki.SelStart = LokiPituus
'  Me.Repaint
End Sub

Private Sub ValHakem_Click()
Dim Hakemisto As String
  Hakemisto = PoimiHakemisto(IIf(IsNull(PohjaHakem.Value), "", PohjaHakem.Value))
  If Hakemisto <> "" Then PohjaHakem.Value = Hakemisto
End Sub
Private Sub ValHakem2_Click()
Dim Hakemisto As String
  Hakemisto = PoimiHakemisto(IIf(IsNull(KuvaHakem.Value), "", KuvaHakem.Value))
  If Hakemisto <> "" Then KuvaHakem.Value = Hakemisto
End Sub
Private Function PoimiHakemisto(AHakem As String) As String
'Tämä valitsee tiedoston hakemistosta
    Dim OpenFile As OPENFILENAME
    Dim lReturn As Long
    Dim Filtteri As String
    Dim TNimi As String
    Dim Otsikko As String
    
    Filtteri = "AutoCAD Drawing (*.dwg)" & Chr(0) & "*.DWG" & Chr(0)
    Otsikko = "Valitse AutoCAD piirustus saadaksesi ylös hakemiston"
    With OpenFile
      .lStructSize = Len(OpenFile)
      .hwndOwner = Me.Hwnd
      .hInstance = 0
      .lpstrFilter = Filtteri
      .nFilterIndex = 1
      .lpstrFile = String(257, 0)
      .nMaxFile = Len(.lpstrFile) - 1
      .lpstrFileTitle = .lpstrFile
      .nMaxFileTitle = .nMaxFile
      .lpstrInitialDir = AHakem
      .lpstrTitle = Otsikko
      .flags = 0
    End With
    lReturn = GetOpenFileName(OpenFile)
    If lReturn = 0 Then 'Painettiin Cancel painiketta. Palautetaan alkuperäinen hakemisto
       PoimiHakemisto = AHakem
    Else 'Otetaan ylös Tiedostonimi ja Hakemisto
       TNimi = Mid(OpenFile.lpstrFileTitle, 1, InStr(OpenFile.lpstrFileTitle, Chr(0)) - 1)
       PoimiHakemisto = Mid(OpenFile.lpstrFile, 1, InStr(OpenFile.lpstrFile, Chr(0)) - 1 - Len(TNimi))
    End If
End Function
