VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_LisääKuviin_ACAD"
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
CurrentDb.Execute "DELETE * FROM Blokit"
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
  Hakemisto = oDoc.PATH
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

Private Sub PaivitaDocRev_Click()

Dim dbsDocuments As Database
Dim rstTable As Recordset
Dim rstDynaset As Recordset
Dim rstSnapshot As Recordset
Dim rstForwardOnly As Recordset
Dim rstLoop As Recordset
Dim prpLoop As Property
Dim SQLQry As String
Dim laskuri As Integer

Dim dbs As DAO.Database
Dim LuoKentat As String
Dim LisaaArvot As String
Dim Kentat As String
Dim KentanArvot As String
Dim u As Integer

Set dbs = CurrentDb

'Avataan Documents-kannan back end -kanta käsittelyä varten.
Set dbsDocuments = OpenDatabase("W:\_DocumentDB\BackEnd\Documents_be.accdb")

'Suoritetaan kysely monesta BE:n taulusta. Kyselyn tuloksen perusteella luodaan uusi
'taulu nimeltä DocumentsDetail.

SQLQry = "SELECT " & _
"ProjectInfo.ProjectNo, ProjectInfo.ProjectName, " & _
"Area.Area, Area.AreaName, DocType.Type, " & _
"DocType.TypeObject, DocType.TypeDesc, " & _
"DocSize.Size, DocSize.SizeName, Documents.DocumentID, Documents.ProjectID, " & _
"Documents.DocName1, Documents.DocName2, Documents.DocName3, Documents.SAVDocNo, Documents.CustDocNo, Documents.DocNo1, Documents.DocNo2, " & _
"Status.Status, Revision.RevAuthor, Revision.RevChecker, Revision.RevApprover, Revision.RevDate, " & _
"Files.File, Files.ReadyPath, Files.WorkPath, Files.DistFolder, " & _
"Revision.Revision, SAVDocNoTbl.SAVDocNoID " & _
"FROM ((((((ProjectInfo INNER JOIN (Area INNER JOIN Documents ON Area.AreaID=Documents.AreaID) ON ProjectInfo.ProjectID=Documents.ProjectID) " & _
"INNER JOIN DocType ON Documents.TypeID=DocType.TypeID) INNER JOIN DocSize ON Documents.SizeID=DocSize.SizeID) " & _
"INNER JOIN Status ON Documents.StatusID=Status.StatusID) INNER JOIN Revision ON Documents.CurrentRevID=Revision.RevID) " & _
"INNER JOIN Files ON Documents.CurrentFileID=Files.FileID) LEFT JOIN SAVDocNoTbl ON Documents.DocumentID=SAVDocNoTbl.DocumentID " & _
"WHERE Documents.ProjectID=" & RefID.Value & ""

'Edellä olevasta kyselystä saadaan recordset rstSnapshot

Set rstSnapshot = dbsDocuments.OpenRecordset(SQLQry, dbOpenSnapshot)

'Jos ko. taulu on olemassa, poistetaan se

If TableExists("DocumentsDetail") = True Then
    dbs.Execute "DROP TABLE DocumentsDetail"
End If

'Haetaan recordsetista taulun luontiin tarvittavat sarakkeiden nimet

LuoKentat = "("

For laskuri = 0 To rstSnapshot.Fields.Count - 2

'Tarvittavat ID-kentät Long -tyypiksi, muut tekstikentäksi
If rstSnapshot.Fields(laskuri).Name = "DocumentID" Or rstSnapshot.Fields(laskuri).Name = "ProjectID" Then
    LuoKentat = LuoKentat & "[" & rstSnapshot.Fields(laskuri).Name & "] long,"
Else
    LuoKentat = LuoKentat & "[" & rstSnapshot.Fields(laskuri).Name & "] TEXT(250),"
End If
Next laskuri

LuoKentat = LuoKentat & "[" & rstSnapshot.Fields(laskuri).Name & "] TEXT(250)"
LuoKentat = LuoKentat & ")"

'Taulun luontikyselyn koostaminen

dbs.Execute "CREATE TABLE DocumentsDetail " & LuoKentat

rstSnapshot.MoveLast
rstSnapshot.MoveFirst

'Käydään recordset läpi ja lisätään siinä olevat tiedot tauluun DocumentsDetail

For u = 1 To rstSnapshot.RecordCount

    Kentat = " ("
    KentanArvot = "("
    
    For laskuri = 0 To rstSnapshot.Fields.Count - 2

        LuoKentat = LuoKentat & "[" & rstSnapshot.Fields(laskuri).Name & "] TEXT(250),"
        Kentat = Kentat & rstSnapshot.Fields(laskuri).Name & ", "
        KentanArvot = KentanArvot & "'" & rstSnapshot.Fields(laskuri) & "',"
    
    Next laskuri
    
    Kentat = Kentat & rstSnapshot.Fields(laskuri).Name
    Kentat = Kentat & ")"
    KentanArvot = KentanArvot & "'" & rstSnapshot.Fields(laskuri) & "'"
    KentanArvot = KentanArvot & ")"
    
    LisaaArvot = "INSERT INTO DocumentsDetail" & Kentat & " VALUES " & KentanArvot
    
    dbs.Execute LisaaArvot
    
    rstSnapshot.MoveNext
Next u

rstSnapshot.Close

'Suoritetaan kysely monesta BE:n taulusta. Kyselyn tuloksen perusteella luodaan uusi
'taulu nimeltä DocumentsRevision.

SQLQry = "SELECT " & _
"Documents.ProjectID, Revision.DocumentID, Revision.Revision, Revision.RevNote, Status.Status, " & _
"Status.Description , Revision.RevDate, Revision.RevAuthor, Revision.RevChecker, Revision.RevApprover, " & _
"Revision.RevID " & _
"FROM Status INNER JOIN (Documents RIGHT JOIN Revision ON Documents.DocumentID = Revision.DocumentID) " & _
"ON Status.StatusID = Revision.StatusID " & _
"WHERE Documents.ProjectID=" & RefID.Value & ""


'Edellä olevasta kyselystä saadaan recordset rstSnapshot

Set rstSnapshot = dbsDocuments.OpenRecordset(SQLQry, dbOpenSnapshot)

'Jos ko. taulu on olemassa, poistetaan se

If TableExists("DocumentsRevision") = True Then
    dbs.Execute "DROP TABLE DocumentsRevision"
End If

'Haetaan recordsetista taulun luontiin tarvittavat sarakkeiden nimet

LuoKentat = "("

For laskuri = 0 To rstSnapshot.Fields.Count - 2

'Tarvittavat ID-kentät Long -tyypiksi, muut tekstikentäksi
If rstSnapshot.Fields(laskuri).Name = "DocumentID" Or rstSnapshot.Fields(laskuri).Name = "ProjectID" Then
    LuoKentat = LuoKentat & "[" & rstSnapshot.Fields(laskuri).Name & "] long,"
Else
    LuoKentat = LuoKentat & "[" & rstSnapshot.Fields(laskuri).Name & "] TEXT(250),"
End If
Next laskuri

LuoKentat = LuoKentat & "[" & rstSnapshot.Fields(laskuri).Name & "] TEXT(250)"
LuoKentat = LuoKentat & ")"

'Taulun luontikyselyn koostaminen

dbs.Execute "CREATE TABLE DocumentsRevision " & LuoKentat

rstSnapshot.MoveLast
rstSnapshot.MoveFirst

'Käydään recordset läpi ja lisätään siinä olevat tiedot tauluun DocumentsDetail

For u = 1 To rstSnapshot.RecordCount

    Kentat = " ("
    KentanArvot = "("
    
    For laskuri = 0 To rstSnapshot.Fields.Count - 2

        LuoKentat = LuoKentat & "[" & rstSnapshot.Fields(laskuri).Name & "] TEXT(250),"
        Kentat = Kentat & rstSnapshot.Fields(laskuri).Name & ", "
        KentanArvot = KentanArvot & "'" & rstSnapshot.Fields(laskuri) & "',"
    
    Next laskuri
    
    Kentat = Kentat & rstSnapshot.Fields(laskuri).Name
    Kentat = Kentat & ")"
    KentanArvot = KentanArvot & "'" & rstSnapshot.Fields(laskuri) & "'"
    KentanArvot = KentanArvot & ")"
    
    LisaaArvot = "INSERT INTO DocumentsRevision" & Kentat & " VALUES " & KentanArvot
    
    dbs.Execute LisaaArvot
    
    rstSnapshot.MoveNext
Next u




      
      
      
     

      
      
      
      dbsDocuments.Close

End Sub

Public Function TableExists(sTable As String) As Boolean
    
    Dim DB As Database
    Dim tbl As TableDef
    Set DB = CurrentDb()
    
    TableExists = False
    
    For Each tbl In DB.TableDefs
        If tbl.Name = sTable Then TableExists = True
    Next tbl
    
End Function

Private Sub TeeKuvat_Click()
'Tämä avaa kuvan insertoi kuvaan blokin/blokit, vaihtaa attribuuttien tiedot ja tallentaa kuvan
Dim DB As Object ' Database
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
Dim sql As String
Dim OtsTaulukko As String
Dim BlockTaulukko As String
Dim BlockTaulukkoTyhjennys As String
Dim Vastaus As Integer
Dim KuvaNimi As String
Dim oAcad As Object 'AutoCAD.AcadApplication
Dim OS As Long, OS2 As Long
Dim Kysely As String

Set DB = CurrentDb
Ajastin = Timer
LokiPituus = 0
OtsTaulukko = TOtsTaulukko.Value

If Translate.Value = True Then
    OtsTaulukko = OtsTaulukko & "_" & TLang.Value
End If

BlockTaulukko = TBlockTaulukko.Value
BlockTaulukkoTyhjennys = TBlockTaulukko.Value

DoCmd.SetWarnings False 'MRu Lisätty 26.04.2011
DoCmd.OpenQuery ("RevisionPlaceQry")
DoCmd.SetWarnings True

Vastaus = MsgBox("Päivitetäänkö " & BlockTaulukkoTyhjennys & "-taulukko?", vbYesNoCancel, "Tee kuvat")
If Vastaus = vbYes Then 'Päivitetään Kuviin1 taulukko
  DB.Execute "DELETE * FROM " & BlockTaulukkoTyhjennys
  OS = 1
  Do
    OS2 = InStr(OS, TKyselyt.Value, vbCrLf)
    If OS2 = 0 Then
      Kysely = Mid(TKyselyt.Value, OS)
      DB.Execute Kysely
      LisaaLokiin "- Ajettiin kysely " & Kysely & "'"
      Exit Do
    Else
      Kysely = Mid(TKyselyt.Value, OS, OS2 - OS)
      DB.Execute Kysely
      LisaaLokiin "- Ajettiin kysely '" & Kysely & "'"
      OS = OS2 + 2
    End If
  Loop
  If CRef = True Then
    LuoReferenssit
  End If
  LisaaLokiin "-Päivitettiin " & BlockTaulukkoTyhjennys & "-taulukko"

ElseIf Vastaus = vbCancel Then
  Set DB = Nothing
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
  Set Kuvat = DB.OpenRecordset("SELECT * FROM " & OtsTaulukko)
  
  If Kuvat.EOF Then 'Taulukossa ei ollut tarvittavia tietoja, annetaan virhe ja keskeytetään
    LisaaLokiin "Virhe: Taulukosta (" & OtsTaulukko & ") ei löytynyt kuvia"
    MsgBox "Taulukosta/kyselystä """ & OtsTaulukko & """ ei löytyny tietoja!" & vbCrLf & "Valitse taulukko ja yritä uudelleen.", vbCritical, "Tee kuvat"
  
  Else 'OK hakemisto löytyi ja taulukossa oli tietoja, ruvetaan käymään taulukon kuvia läpi
       
       'Otetaan yhteys AutoCADiin
  On Error Resume Next
  Set oAcad = GetObject(, "AutoCAD.Application") 'Koitetaan yhdistää AutoCADiin
  If Err <> 0 Then 'Käynnissä olevaa AutoCADiä ei löytynyt
    'MsgBox "Käynnissä olevaa AutoCADiä ei löytynyt!", vbCritical, "Virhe!"
    Set oAcad = CreateObject("AutoCAD.Application")
    'Set oAcad = Nothing
    'Exit Sub
  End If
  On Error GoTo 0
    
    oAcad.Application.Visible = True    ' MRU 11.10.2011 !!!!!!!!!!!!!!!!!!!!!!!!!!!!
    
    Do While Not Kuvat.EOF
      Tiedosto = PohjaHakem.Value & Kuvat.Fields("BaseDWG").Value
      
      If LCase(Right(Tiedosto, 4)) <> ".dwg" Then 'Varmistetaan että tiedostonimen pääte on .dwg
        Tiedosto = Tiedosto & ".dwg"
      End If
      
      If FSO.FileExists(Tiedosto) Then 'Tiedosto löytyi voidaan avata se AutoCADissä
        Tila Kuvat.Fields("FileID").Value
        
        Set oDoc = oAcad.Documents.Open(Tiedosto)   'Avataan uusi dokumentti
        OSMODE = oDoc.GetVariable("OSMODE")
        oDoc.SetVariable "OSMODE", 0 'Asetetaan OSNAP tila pois päältä
        
        
        
        
        'Haetaan insertointipaikat kuvasta
        If HaeIPoints(Kuvat.Fields("BaseDWG").Value) Then
          
          'Vaihdetaan otsikkotaulun tiedot
          If VaihdaOtsikkotiedot(Kuvat, "", "", "") Then 'Jos Otsikkotiedot saatiin vaihdettua OK niin kannattaa mennä eteenpäin
            
            'Valitaan taulukosta kuvaan insertoitavat blokit
            sql = "SELECT Block, Blocknumber FROM " & BlockTaulukko & " WHERE FileID='" & Kuvat.Fields("FileID").Value & "' GROUP BY Block, Blocknumber HAVING (Blocknumber Is Not Null);"
            Set DbBlokit = DB.OpenRecordset(sql)
            
            Do While Not DbBlokit.EOF
              
              If FSO.FileExists(PohjaHakem.Value & DbBlokit.Fields("Block").Value & ".dwg") Then 'Blokki löytyi
                sql = "SELECT * FROM " & BlockTaulukko & " WHERE FileID='" & Kuvat.Fields("FileID").Value & "' AND Block='" & DbBlokit.Fields("Block").Value & "'  AND Blocknumber='" & DbBlokit.Fields("Blocknumber").Value & "' ORDER BY AttrTag;"
                Set Tiedot = DB.OpenRecordset(sql)
                Paikka = CInt(DbBlokit.Fields("Blocknumber").Value)
                If Translate.Value = True Then 'Tehdään eri kieliset kuvat
                    BNimi = PohjaHakem.Value & DbBlokit.Fields("Block").Value & "_" & TLang.Value & ".dwg" 'Lisätään blokin nimeen kielen tunnus, esim. _RU
                    If Dir(BNimi, vbDirectory) = vbNullString Then 'Tarkistetaan löytyykö kyseistä blokkia
                        BNimi = PohjaHakem.Value & DbBlokit.Fields("Block").Value & ".dwg" 'Jos blokkia ei löydy, käytetään alkuperäistä nimeä
                    End If
                Else
                    BNimi = PohjaHakem.Value & DbBlokit.Fields("Block").Value & ".dwg" 'Lisättävän blokin nimi
                End If
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
                    'LisaaLokiin "Virhe: Blokin " & DbBlokit.Fields("Block").Value & " attribuutille (" & Attribuutit(i).TagString & ") ei löytynyt arvoa tietokannasta"
                  End If
                
                Next i
              
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
        
        oAcad.Application.ZoomExtents
        oAcad.Application.ActiveDocument.PurgeAll

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
'oAcad.Quit
Loppu:
Set oAcad = Nothing
Set DB = Nothing
Set Kuvat = Nothing
Set Tiedot = Nothing
Set FSO = Nothing
Set Blokki = Nothing
Set DbBlokit = Nothing
End Sub
Private Sub LuoReferenssit()
'Luodaan referenssilinkin 2.3.2015 - VG
Dim DB As Object ' Database
Dim Tiedot As Object 'Recordset
Dim sSQL As String
Dim RNro As Integer
Dim E_FileId As String
Dim E_Block As String
Dim E_Blocknumber As String

Set DB = CurrentDb
Set Tiedot = DB.OpenRecordset("SELECT * FROM Kuviin1_Ref_PAIVITA")

'Päivitetaan ensin tuleeviin attribuutteihin mistä sivulta tagi tulee
Do While Not Tiedot.EOF
  sSQL = "UPDATE Kuviin1 SET Kuviin1.Teksti = '" & Tiedot("UusiTeksti") & "' "
  sSQL = sSQL & "WHERE (((FileID)='" & Tiedot("FileID") & "') "
  sSQL = sSQL & "AND ((Block)='" & Tiedot("Block") & "') "
  sSQL = sSQL & "AND ((Blocknumber)='" & Tiedot("BlockNumber") & "') "
  sSQL = sSQL & "AND ((AttrTag)='" & Tiedot("AttrTag") & "'));"
  DB.Execute sSQL
  Tiedot.MoveNext
Loop
'Lisätään sitten lähteviin attribuutteihin, mihin piiriin ja minne sivulle tieto meneesivulle
Tiedot.Close
Set Tiedot = DB.OpenRecordset("SELECT * FROM Kuviin1_Ref_LISAA")

'Päivitetaan ensin tuleeviin attribuutteihin mistä sivulta tagi tulee
Do While Not Tiedot.EOF
 If E_FileId = Tiedot("FileId") And E_Block = Tiedot("Block") And E_Blocknumber = Tiedot("Blocknumber") Then
   RNro = RNro + 1
 Else
   RNro = 1
 End If
 sSQL = "INSERT INTO Kuviin1 ( FileID, Block, Blocknumber, AttrTag, Teksti ) "
 sSQL = sSQL & "Values ('" & Tiedot("FileID") & "', '" & Tiedot("Block") & "','" & Tiedot("Blocknumber") & "','FTO" & RNro & "','" & Tiedot("Teksti") & "')"
 
 DB.Execute sSQL
  
 E_FileId = Tiedot("FileId")
 E_Block = Tiedot("Block")
 E_Blocknumber = Tiedot("Blocknumber")
  
  
  Tiedot.MoveNext
Loop

Set Tiedot = Nothing
Set DB = Nothing
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
'PaikkaBlokki = "POINTBAR"

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
Private Function VaihdaOtsikkotiedot(Taulu As Recordset, TallennusNimi As String, NextPage As String, CurrPage As String) As Boolean
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
    'TitleBlokki = "TITLE"

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
                       ' If taulu.Fields(j).Name = "CAD_FILE" Then
                            Attribuutit(i).TextString = IIf(IsNull(Taulu.Fields(tagi).Value), "", Taulu.Fields(tagi).Value)
                       ' End If
                    
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
'  Loki.SetFocus
'  If LokiPituus > 0 Then
'    Loki.VALUE = Loki.VALUE & vbCrLf & Tieto
'  Else
'    Loki.VALUE = Tieto
'    LokiPituus = 1
'  End If
' Loki.SelStart = LokiPituus
' Me.Repaint
End Sub

Private Sub TOtsTaulukko_BeforeUpdate(Cancel As Integer)

End Sub

Private Sub TTitleBlokki_BeforeUpdate(Cancel As Integer)

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

Private Sub TeeEtusivu_Click()
'Tämä avaa kuvan insertoi kuvaan blokin/blokit, vaihtaa attribuuttien tiedot ja tallentaa kuvan
Dim DB As Object ' Database
Dim Kuvat As Object 'Recordset
Dim DbBlokit As Object ' Recordset
Dim Tiedot As Object ' Recordset
Dim FSO As New FileSystemObject
Dim Tiedosto As String
Dim Blokki As Object ' AcadBlockReference
Dim Attribuutit As Variant
Dim Paikka As Integer
Dim BNimi As String
Dim i As Integer, a As Integer, k As Integer, m As Integer, n As Integer
Dim sql As String
Dim OtsTaulukko As String
Dim BlockTaulukko As String
Dim BlockTaulukkoTyhjennys As String
Dim Vastaus As Integer
Dim KuvaNimi As String
Dim oAcad As Object 'AutoCAD.AcadApplication
Dim OS As Long, OS2 As Long
Dim Kysely As String
Dim Valitallennus As Variant
Dim TallennusNimi As String
Dim NextPage As String
Dim CurrPage As String
Dim FirstLoopNo As String
Dim Valiotsikkotallennus As Variant

ReDim Paikat(1 To 1)

FirstLoopNo = ""

Set DB = CurrentDb
Ajastin = Timer
LokiPituus = 0
OtsTaulukko = TFrontOtsTaulukko.Value



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

Set oAcad = CreateObject("AutoCAD.Application")

Tiedosto = PohjaHakem.Value & FrontBase.Value
If LCase(Right(Tiedosto, 4)) <> ".dwg" Then 'Varmistetaan että tiedostonimen pääte on .dwg
    Tiedosto = Tiedosto & ".dwg"
End If
    
If FSO.FileExists(Tiedosto) Then 'Tiedosto löytyi voidaan avata se AutoCADissä
    Set oDoc = oAcad.Documents.Open(Tiedosto) 'Avataan uusi dokumentti
    oDoc.SetVariable "OSMODE", 0 'Asetetaan OSNAP tila pois päältä

   
   
   'Suoritetaan kysely kannan kyselystä firstfunc
    sql = "SELECT looppisymb, FREV, LoopNo FROM firstfunc;" ' ORDER BY LoopNo;"
    Set DbBlokit = DB.OpenRecordset(sql)
    
    k = 1
    m = 0
    n = 0
     
'Käydään edellämainitussa kyselyssä olevat rivit läpi
    Set Kuvat = DB.OpenRecordset("SELECT * FROM " & OtsTaulukko & " WHERE DocumentID=" & RefDocumenID.Value)

    Do While Not DbBlokit.EOF
    If FirstLoopNo = "" Then
        FirstLoopNo = DbBlokit.Fields("LoopNo").Value
    End If

        If FSO.FileExists(PohjaHakem.Value & "LoopDet.dwg") Then 'Blokki löytyi
            
'Insertoitava blokki
            BNimi = PohjaHakem.Value & "LoopDet.dwg"
            
            If (k Mod Rows.Value) = 0 Then ' jos k on jaollinen 40:llä, vaihdetaan saraketta ja alustetaan k
                m = m + 1
                k = 1
            End If
            
'JOS tulee rivimäärä täyteen, niin vaihdetaan saraketta
            
            If (k Mod Rows.Value) = 0 Or k = 1 Then ' jos k on jaollinen annetulla rivimäärällä, vaihdetaan saraketta ja alustetaan k
            
'JOS sarakemäärä tulee täyteen, tallennetaan nykyinen kuva ja avataan uusi käsittelyä varten

'Tiedosto = PohjaHakem.Value & FrontBase.Value
                If LCase(Right(Kuvat.Fields("DWG_No"), 4)) = ".dwg" Then 'Poistetaan tiedostonimen pääte .dwg
                    TallennusNimi = Left(Kuvat.Fields("DWG_No"), Len(Kuvat.Fields("DWG_No")) - 4)
                '    Tiedosto = Tiedosto & ".dwg"
                Else
                    TallennusNimi = Kuvat.Fields("DWG_NO")
                End If

            '  If Columns.Value = m Then
      '              If n = 0 Then
     '                   TallennusNimi = TallennusNimi & "_VI"
            '            CurrPage = "VI"
           '             NextPage = "VII"
          '          ElseIf n = 1 Then
         '               TallennusNimi = TallennusNimi & "_VII"
        '                CurrPage = "VII"
       '                 NextPage = "VIII"
          '          ElseIf n = 2 Then
          '              TallennusNimi = TallennusNimi & "_VIII"
         '               CurrPage = "VIII"
           '             NextPage = "IX"
          '          ElseIf n = 3 Then
         '               TallennusNimi = TallennusNimi & "_IX"
        '                CurrPage = "IX"
                        NextPage = "X"
       '             ElseIf n = 4 Then
     '                   TallennusNimi = TallennusNimi & "_X"
    '                    CurrPage = "X"
   '                     NextPage = "XI"
      '              ElseIf n = 5 Then
    '                    TallennusNimi = TallennusNimi & "_XI"
    '                    CurrPage = "XI"
     '                   NextPage = "XII"
      '              ElseIf n = 6 Then
     '                   TallennusNimi = TallennusNimi & "_XII"
    '                    CurrPage = "XII"
   '                     NextPage = "XIII"
  '                  End If

                If Columns.Value = m Then
                    n = n + 1
                    m = 0
                 
           If n = 0 Then
                        TallennusNimi = TallennusNimi & "_VI"
                        CurrPage = "VI"
                        NextPage = "VII"
                    ElseIf n = 1 Then
                        TallennusNimi = TallennusNimi & "_VII"
                        CurrPage = "VII"
                        NextPage = "VIII"
                    ElseIf n = 2 Then
                        TallennusNimi = TallennusNimi & "_VIII"
                        CurrPage = "VIII"
                        NextPage = "IX"
                    ElseIf n = 3 Then
                        TallennusNimi = TallennusNimi & "_IX"
                        CurrPage = "IX"
                        NextPage = "X"
                    ElseIf n = 4 Then
                        TallennusNimi = TallennusNimi & "_X"
                        CurrPage = "X"
                        NextPage = "XI"
                    ElseIf n = 5 Then
                        TallennusNimi = TallennusNimi & "_XI"
                        CurrPage = "XI"
                        NextPage = "XII"
                    ElseIf n = 6 Then
                        TallennusNimi = TallennusNimi & "_XII"
                        CurrPage = "XII"
                        NextPage = "XIII"
                    End If
                 
                 
                 
                 Valiotsikkotallennus = VaihdaOtsikkotiedot(Kuvat, TallennusNimi & ".dwg", NextPage, CurrPage)
                 Valitallennus = TallennaUusiKuva(TallennusNimi, False, oAcad)
   
                
             '   Exit Do
                
                End If

'Insertoidaan Sarakkeille otsikot
            
                ReDim Preserve Paikat(1 To k) 'Varmistetaan että Paikat Arrayssa on tilaa tämän paikan tiedolle
                Paikat(k).Pisteet(0) = XCoord.Value + m * 60 '60 on sarakkeen leveys, muuta myös alla olevaan vastaavaan kohtaan
                Paikat(k).Pisteet(1) = YCoord.Value - k * 5 '5 = ns. riviväli, muuta myös alla olevaan vastaavaan kohtaan
                Paikat(k).Pisteet(2) = 0 'Z
                 
                Set Blokki = oDoc.ModelSpace.InsertBlock(Paikat(k).Pisteet, BNimi, 1, 1, 1, 0)
                Attribuutit = Blokki.GetAttributes 'Haetaan attribuutit
                Attribuutit(0).TextString = "LOOP"
                Attribuutit(1).TextString = "REVISION"
            
            End If

'Insertoidaan rivitietoja.
            ReDim Preserve Paikat(1 To k) 'Varmistetaan että Paikat Arrayssa on tilaa tämän paikan tiedolle
            Paikat(k).Pisteet(0) = XCoord.Value + m * 60 '60 on sarakkeen leveys
            Paikat(k).Pisteet(1) = (YCoord.Value - 5) - k * 5 '5 = ns. riviväli
            Paikat(k).Pisteet(2) = 0 'Z
             
            Set Blokki = oDoc.ModelSpace.InsertBlock(Paikat(k).Pisteet, BNimi, 1, 1, 1, 0)
             'Muutetaan blokin attribuutit
            Attribuutit = Blokki.GetAttributes 'Haetaan attribuutit
            
            Attribuutit(0).TextString = DbBlokit.Fields("looppisymb").Value
            Attribuutit(1).TextString = DbBlokit.Fields("FREV").Value
         
        Else
            LisaaLokiin "Virhe: Blokkia " & DbBlokit.Fields("Block").Value & " ei löytynyt"
        End If
        
        k = k + 1
        
    DbBlokit.MoveNext
    Loop

'End If

'Tallennetaan olemassa oleva kuva, eikä jatketa
        
 '---       If n = 0 Then
 '           Tallennusnimi = Tallennusnimi & "_VI"
 '           CurrPage = "VI"
 '       ElseIf n = 1 Then
  '          Tallennusnimi = Tallennusnimi & "_VII"
  '          CurrPage = "VII"
  '      ElseIf n = 2 Then
  '          Tallennusnimi = Tallennusnimi & "_VIII"
  '          CurrPage = "VIII"
  '      ElseIf n = 3 Then
   '         Tallennusnimi = Tallennusnimi & "_IX"
   '         CurrPage = "IX"
   '     ElseIf n = 4 Then
   '         Tallennusnimi = Tallennusnimi & "_X"
   '         CurrPage = "X"
    '    ElseIf n = 5 Then
     '       Tallennusnimi = Tallennusnimi & "_XI"
     '       CurrPage = "XI"
     '   ElseIf n = 6 Then
     '       Tallennusnimi = Tallennusnimi & "_XII"
     '       CurrPage = "XII"
     '   End If

'        Set Kuvat = db.OpenRecordset("SELECT * FROM " & OtsTaulukko & " WHERE DocumentID=" & RefDocumenID.Value)
   '    If VaihdaOtsikkotiedot(Kuvat, Tallennusnimi, FirstLoopNo, CurrPage) Then 'Jos Otsikkotiedot saatiin vaihdettua OK niin kannattaa mennä eteenpäin
   '         Valitallennus = TallennaUusiKuva(Tallennusnimi, True, oAcad)
   '     End If

  
 
'    Valitallennus = TallennaUusiKuva(TallennusNimi, True, oAcad)
        
'     If n = 0 Then
 '        Valitallennus = TallennaUusiKuva("VI", True, oAcad)
'                ElseIf n = 1 Then
  '         Valitallennus = TallennaUusiKuva("VII", True, oAcad)
 '       ElseIf n = 2 Then
  '          Valitallennus = TallennaUusiKuva("VIII", True, oAcad)
'        ElseIf n = 3 Then
'            Valitallennus = TallennaUusiKuva("IX", True, oAcad)
'        ElseIf n = 4 Then
'            Valitallennus = TallennaUusiKuva("X", True, oAcad)
'        ElseIf n = 5 Then
'            Valitallennus = TallennaUusiKuva("XI", True, oAcad)
'        ElseIf n = 6 Then
 '           Valitallennus = TallennaUusiKuva("XI", True, oAcad)
   '    End If


    Else 'Tiedostoa ei löytynyt
        LisaaLokiin "Virhe: Määriteltyä tiedostoa ei löytynyt (" & Tiedosto & ")"
    End If

End If
 
     If VaihdaOtsikkotiedot(Kuvat, TallennusNimi & ".dwg", NextPage, CurrPage) Then 'Jos Otsikkotiedot saatiin vaihdettua OK niin kannattaa mennä eteenpäin
        Valitallennus = TallennaUusiKuva(TallennusNimi, True, oAcad)
    End If

 '  oDoc.SetVariable "OSMODE", OSMODE 'Palautetaan OSNAP tila siihen joka se oli ennen sen poistamista pois päältä
'        oDoc.SaveAs KuvaNimi 'Talletetaan kuva
 '       oDoc.Close

MsgBox "Revisiosivu(t) tehty", vbOKOnly, "Valmis"

oAcad.Quit
Loppu:
Set oAcad = Nothing
Set DB = Nothing
Set Kuvat = Nothing
Set Tiedot = Nothing
Set FSO = Nothing
Set Blokki = Nothing
Set DbBlokit = Nothing
End Sub

Private Function TallennaUusiKuva(TallNimi As String, Loppu As Boolean, oAcad As Object)
        
Dim FSO As New FileSystemObject
Dim KuvaNimi As String
Dim Tiedosto As String
'Dim oAcad As Object 'AutoCAD.AcadApplication

        KuvaNimi = KuvaHakem.Value & TallNimi '& "xxx" ' & ".dwg"
        
        If Right(LCase(KuvaNimi), 4) <> ".dwg" Then
        KuvaNimi = KuvaNimi & ".dwg"
        End If
        
        If FSO.FileExists(KuvaNimi) Then
          FSO.DeleteFile KuvaNimi
        End If
        
        oDoc.SetVariable "OSMODE", OSMODE 'Palautetaan OSNAP tila siihen joka se oli ennen sen poistamista pois päältä
        oDoc.SaveAs KuvaNimi 'Talletetaan kuva
        oDoc.Close
        'End If
        
If Loppu = False Then
        
    Tiedosto = PohjaHakem.Value & FrontBase.Value
    If LCase(Right(Tiedosto, 4)) <> ".dwg" Then 'Varmistetaan että tiedostonimen pääte on .dwg
        Tiedosto = Tiedosto & ".dwg"
    End If
    
    If FSO.FileExists(Tiedosto) Then 'Tiedosto löytyi voidaan avata se AutoCADissä
Set oDoc = Nothing
        Set oDoc = oAcad.Documents.Open(Tiedosto)   'Avataan uusi dokumentti
        oDoc.SetVariable "OSMODE", 0 'Asetetaan OSNAP tila pois päältä
    End If
End If

End Function
Private Sub HaeData()

Dim Yhteys As String

    Yhteys = "ODBC;DBQ=" & Kanta & ";Driver={Microsoft Access Driver (*.mdb)}"
'Set TAULUKKO = ActiveSheet.QueryTables.Add(Connection:=Yhteys, Destination:=Range("A1"))

End Sub
