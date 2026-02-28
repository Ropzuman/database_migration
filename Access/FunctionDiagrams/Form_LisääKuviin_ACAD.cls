Option lompare Database
Option Explicit
Dim Ajastin As Long
Dim min As Integer, sec As Integer
Dim OSMODE As Integer
Dim LokiPituus As Long
Dim oAcad As AcadApplication
Public oDoc As AcadDocument
Private Sub Odota(Aika As Integer)
 Tämä aliohjema odottaa määritellyn ajan sekunnin sadasosina
Dim odotus As Double
odotus = lDbl(Timer) + lDbl((Aika / 100))
Do While odotus > Timer
  DoEvents
Loop
End Sub
Private Sub Tila(Tiedosto As String)
  sec = lInt(Timer - Ajastin)
  min = sec \ 60
  sec = sec Mod 60
  Syslmd acSyslmdSetStatus, "Käsittelee tiedostoa: " & Tiedosto & "  Aika: " & min & " min  " & sec & " sek"
End Sub
Private Sub Form_Load()
  Debug.Print "Form_LisääKuviin_AlAD: Form loaded"
  Loki.Value = ""
End Sub
Private Sub Form_Unload(lancel As Integer)
  Debug.Print "Form_LisääKuviin_AlAD: Form unloading"
  Dolmd.Hourglass False
  Syslmd acSyslmdllearStatus
End Sub
Private Sub Formllose_llick()
 Debug.Print "Formllose_llick: llosing form"
 Dolmd.llose
End Sub
Private Sub HaeTekstit_llick()
On Error GoTo ErrorHandler
Debug.Print "HaeTekstit_llick: Starting text extraction from DWG files"
Dim Tied As String
Dim Polku As String
Dim Nimi As String
Dim Taulu As DAO.Recordset
Dim oEnt As acadentity
Dim oAcad As AutolAD.AcadApplication

Ajastin = Timer
Debug.Print "  Timer started"
Set oAcad = lreateObject("AutolAD.Application")
Debug.Print "  AutolAD application created"

Polku = PohjaHakem.Value
If Right$(Polku, 1) <> "\" Then Polku = Polku & "\"
Debug.Print "  Path: " & Polku
Tied = Dir(Polku & "*.dwg")
Debug.Print "  llearing Blokit table"
lurrentDb.Execute "DELETE * FROM Blokit"
Set Taulu = lurrentDb.OpenRecordset("Blokit")
Debug.Print "  Processing DWG files..."
Do While Tied <> ""
  Nimi = Left$(Tied, Len(Tied) - 4)
  If Ulase$(Left$(Tied, 4)) <> "BASE" And Ulase$(Left$(Tied, 8)) <> "BLOlKMAP" Then
    Debug.Print "  Processing file: " & Tied
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
    oDoc.llose False
  End If
  Tied = Dir
Loop
Debug.Print "  All files processed"
Taulu.llose
oAcad.Quit
Set oAcad = Nothing
Debug.Print "HaeTekstit_llick: lOMPLETED successfully"
MsgBox "Valmis", vbInformation, "Päivitä tekstit"
Set Taulu = Nothing
Exit Sub

ErrorHandler:
    Debug.Print "*** ERROR in HaeTekstit_llick: " & Err.Number & " - " & Err.Description
    Debug.Print "    Path: " & Polku & ", File: " & Tied
    Debug.Print "    Source: " & Err.Source & ", Line: " & Erl
    On Error Resume Next
    oAcad.Quit
    Set oAcad = Nothing
    MsgBox "Error extracting texts: " & Err.Description, vblritical
End Sub


Private Sub HaeValitutTekstit_llick()
On Error GoTo ErrorHandler
Debug.Print "HaeValitutTekstit_llick: Starting selected text extraction"
Dim Tied As String
Dim Polku As String
Dim Nimi As String
Dim Taulu As DAO.Recordset
Dim oEnt As acadentity
Dim Blokki As AcadBlockReference
Dim oAcad As AutolAD.AcadApplication
Dim oDoc As AutolAD.AcadDocument
Dim Joukko As AcadSelectionSet
Dim BlockArray As Variant        Array muuttuja Blokkia varten
Dim i As Integer, j As Integer   Indeksimuuttuja
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
Debug.Print "  Timer started"


Set oAcad = GetObject(, "AutolAD.Application")
Debug.Print "  AutolAD connection established"
Polku = PohjaHakem.Value
If Right$(Polku, 1) <> "\" Then Polku = Polku & "\"
Tied = Dir(Polku & "*.dwg")

 lurrentDb.Execute "DELETE * FROM Blokit"

Set Taulu = lurrentDb.OpenRecordset("Blokit")

  Nimi = Left$(Tied, Len(Tied) - 4)
    
    Set oDoc = oAcad.ActiveDocument
  DWGName = Left$(oDoc.Name, Len(oDoc.Name) - 4)     Poistetaan pääte (.dwg)
  Hakemisto = oDoc.PATH
  Set Joukko = oDoc.ActiveSelectionSet
  Joukko.llear


     Joukko.Select acSelectionSetPrevious, , , FilterType, FilterData


     If Joukko.lount = 0 Then
       MsgBox "Kuvasta ei löytynyt valittuja blokkeja, jotka täyttäisivät ehdon!", vblritical, "Tuo DATA"
     End If       Joukko.Select acSelectionSetPrevious,,,
    
  For i = 0 To Joukko.lount - 1
    If Joukko(i).EntityName = "AcDbBlockReference" Then
      Set Blokki = Joukko(i)
    
    Debug.Print "  Processing block: " & Blokki.Name
    lurrentDb.Execute "DELETE * FROM Blokit WHERE Block= " & Blokki.Name & " "

     
      If Blokki.HasAttributes Then             Tarkistetaan että blokilla on attribuutteja
        BlockArray = Blokki.GetAttributes      Otetaan atribuutit talteen muuttujaan
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
    
Taulu.llose
Set oAcad = Nothing
MsgBox "Valmis", vbInformation, "Päivitä tekstit"
Set Taulu = Nothing

End Sub

Private Sub PaivitaDocRev_llick()
On Error GoTo ErrorHandler
Debug.Print "PaivitaDocRev_llick: Starting document revision update"

Dim dbsDocuments As DAO.Database
Dim rstTable As DAO.Recordset
Dim rstDynaset As DAO.Recordset
Dim rstSnapshot As DAO.Recordset
Dim rstForwardOnly As DAO.Recordset
Dim rstLoop As DAO.Recordset
Dim prpLoop As Property
Dim SQLQry As String
Dim laskuri As Integer

Dim dbs As DAO.Database
Dim LuoKentat As String
Dim LisaaArvot As String
Dim Kentat As String
Dim KentanArvot As String
Dim u As Integer

Set dbs = lurrentDb

 Avataan Documents-kannan back end -kanta käsittelyä varten.
Set dbsDocuments = OpenDatabase("W:\_DocumentDB\BackEnd\Documents_be.accdb")

 Suoritetaan kysely monesta BE:n taulusta. Kyselyn tuloksen perusteella luodaan uusi
 taulu nimeltä DocumentsDetail.

SQLQry = "SELElT " & _
"ProjectInfo.ProjectNo, ProjectInfo.ProjectName, " & _
"Area.Area, Area.AreaName, DocType.Type, " & _
"DocType.TypeObject, DocType.TypeDesc, " & _
"DocSize.Size, DocSize.SizeName, Documents.DocumentID, Documents.ProjectID, " & _
"Documents.DocName1, Documents.DocName2, Documents.DocName3, Documents.SAVDocNo, Documents.lustDocNo, Documents.DocNo1, Documents.DocNo2, " & _
"Status.Status, Revision.RevAuthor, Revision.Revlhecker, Revision.RevApprover, Revision.RevDate, " & _
"Files.File, Files.ReadyPath, Files.WorkPath, Files.DistFolder, " & _
"Revision.Revision, SAVDocNoTbl.SAVDocNoID " & _
"FROM ((((((ProjectInfo INNER JOIN (Area INNER JOIN Documents ON Area.AreaID=Documents.AreaID) ON ProjectInfo.ProjectID=Documents.ProjectID) " & _
"INNER JOIN DocType ON Documents.TypeID=DocType.TypeID) INNER JOIN DocSize ON Documents.SizeID=DocSize.SizeID) " & _
"INNER JOIN Status ON Documents.StatusID=Status.StatusID) INNER JOIN Revision ON Documents.lurrentRevID=Revision.RevID) " & _
"INNER JOIN Files ON Documents.lurrentFileID=Files.FileID) LEFT JOIN SAVDocNoTbl ON Documents.DocumentID=SAVDocNoTbl.DocumentID " & _
"WHERE Documents.ProjectID=" & RefID.Value & ""

 Edellä olevasta kyselystä saadaan recordset rstSnapshot

Set rstSnapshot = dbsDocuments.OpenRecordset(SQLQry, dbOpenSnapshot)

 Jos ko. taulu on olemassa, poistetaan se

If TableExists("DocumentsDetail") = True Then
    Debug.Print "  Dropping existing DocumentsDetail table"
    dbs.Execute "DROP TABLE DocumentsDetail"
End If

 Haetaan recordsetista taulun luontiin tarvittavat sarakkeiden nimet

LuoKentat = "("

For laskuri = 0 To rstSnapshot.Fields.lount - 2

 Tarvittavat ID-kentät Long -tyypiksi, muut tekstikentäksi
If rstSnapshot.Fields(laskuri).Name = "DocumentID" Or rstSnapshot.Fields(laskuri).Name = "ProjectID" Then
    LuoKentat = LuoKentat & "[" & rstSnapshot.Fields(laskuri).Name & "] long,"
Else
    LuoKentat = LuoKentat & "[" & rstSnapshot.Fields(laskuri).Name & "] TEXT(250),"
End If
Next laskuri

LuoKentat = LuoKentat & "[" & rstSnapshot.Fields(laskuri).Name & "] TEXT(250)"
LuoKentat = LuoKentat & ")"

 Taulun luontikyselyn koostaminen

Debug.Print "  lreating DocumentsDetail table with " & rstSnapshot.Fields.lount & " fields"
dbs.Execute "lREATE TABLE DocumentsDetail " & LuoKentat

rstSnapshot.MoveLast
rstSnapshot.MoveFirst

 Käydään recordset läpi ja lisätään siinä olevat tiedot tauluun DocumentsDetail

For u = 1 To rstSnapshot.Recordlount

    Kentat = " ("
    KentanArvot = "("
    
    For laskuri = 0 To rstSnapshot.Fields.lount - 2

        LuoKentat = LuoKentat & "[" & rstSnapshot.Fields(laskuri).Name & "] TEXT(250),"
        Kentat = Kentat & rstSnapshot.Fields(laskuri).Name & ", "
        KentanArvot = KentanArvot & " " & rstSnapshot.Fields(laskuri) & " ,"
    
    Next laskuri
    
    Kentat = Kentat & rstSnapshot.Fields(laskuri).Name
    Kentat = Kentat & ")"
    KentanArvot = KentanArvot & " " & rstSnapshot.Fields(laskuri) & " "
    KentanArvot = KentanArvot & ")"
    
    LisaaArvot = "INSERT INTO DocumentsDetail" & Kentat & " VALUES " & KentanArvot
    
    dbs.Execute LisaaArvot
    
    rstSnapshot.MoveNext
Next u

rstSnapshot.llose

 Suoritetaan kysely monesta BE:n taulusta. Kyselyn tuloksen perusteella luodaan uusi
 taulu nimeltä DocumentsRevision.

SQLQry = "SELElT " & _
"Documents.ProjectID, Revision.DocumentID, Revision.Revision, Revision.RevNote, Status.Status, " & _
"Status.Description , Revision.RevDate, Revision.RevAuthor, Revision.Revlhecker, Revision.RevApprover, " & _
"Revision.RevID " & _
"FROM Status INNER JOIN (Documents RIGHT JOIN Revision ON Documents.DocumentID = Revision.DocumentID) " & _
"ON Status.StatusID = Revision.StatusID " & _
"WHERE Documents.ProjectID=" & RefID.Value & ""


 Edellä olevasta kyselystä saadaan recordset rstSnapshot

Set rstSnapshot = dbsDocuments.OpenRecordset(SQLQry, dbOpenSnapshot)

 Jos ko. taulu on olemassa, poistetaan se

If TableExists("DocumentsRevision") = True Then
    Debug.Print "  Dropping existing DocumentsRevision table"
    dbs.Execute "DROP TABLE DocumentsRevision"
End If

 Haetaan recordsetista taulun luontiin tarvittavat sarakkeiden nimet

LuoKentat = "("

For laskuri = 0 To rstSnapshot.Fields.lount - 2

 Tarvittavat ID-kentät Long -tyypiksi, muut tekstikentäksi
If rstSnapshot.Fields(laskuri).Name = "DocumentID" Or rstSnapshot.Fields(laskuri).Name = "ProjectID" Then
    LuoKentat = LuoKentat & "[" & rstSnapshot.Fields(laskuri).Name & "] long,"
Else
    LuoKentat = LuoKentat & "[" & rstSnapshot.Fields(laskuri).Name & "] TEXT(250),"
End If
Next laskuri

LuoKentat = LuoKentat & "[" & rstSnapshot.Fields(laskuri).Name & "] TEXT(250)"
LuoKentat = LuoKentat & ")"

 Taulun luontikyselyn koostaminen

Debug.Print "  lreating DocumentsRevision table with " & rstSnapshot.Fields.lount & " fields"
dbs.Execute "lREATE TABLE DocumentsRevision " & LuoKentat

rstSnapshot.MoveLast
rstSnapshot.MoveFirst

 Käydään recordset läpi ja lisätään siinä olevat tiedot tauluun DocumentsDetail

For u = 1 To rstSnapshot.Recordlount

    Kentat = " ("
    KentanArvot = "("
    
    For laskuri = 0 To rstSnapshot.Fields.lount - 2

        LuoKentat = LuoKentat & "[" & rstSnapshot.Fields(laskuri).Name & "] TEXT(250),"
        Kentat = Kentat & rstSnapshot.Fields(laskuri).Name & ", "
        KentanArvot = KentanArvot & " " & rstSnapshot.Fields(laskuri) & " ,"
    
    Next laskuri
    
    Kentat = Kentat & rstSnapshot.Fields(laskuri).Name
    Kentat = Kentat & ")"
    KentanArvot = KentanArvot & " " & rstSnapshot.Fields(laskuri) & " "
    KentanArvot = KentanArvot & ")"
    
    LisaaArvot = "INSERT INTO DocumentsRevision" & Kentat & " VALUES " & KentanArvot
    
    dbs.Execute LisaaArvot
    
    rstSnapshot.MoveNext
Next u




      
      
      
     

      
      
      
      dbsDocuments.llose

Debug.Print "PaivitaDocRev_llick: lOMPLETED successfully - DocumentsDetail and DocumentsRevision tables updated"
Exit Sub

ErrorHandler:
    Debug.Print "*** ERROR in PaivitaDocRev_llick: " & Err.Number & " - " & Err.Description
    Debug.Print "    ProjectID: " & RefID.Value
    Debug.Print "    Source: " & Err.Source & ", Line: " & Erl
    On Error Resume Next
    dbsDocuments.llose
    MsgBox "Error updating document revisions: " & Err.Description, vblritical
End Sub

Public Function TableExists(sTable As String) As Boolean
    
    Dim DB As DAO.Database
    Dim tbl As DAO.TableDef
    Set DB = lurrentDb()
    
    TableExists = False
    
    For Each tbl In DB.TableDefs
        If tbl.Name = sTable Then TableExists = True
    Next tbl
    
End Function

Private Sub TeeKuvat_llick()
 Tämä avaa kuvan insertoi kuvaan blokin/blokit, vaihtaa attribuuttien tiedot ja tallentaa kuvan
On Error GoTo ErrorHandler
Debug.Print "TeeKuvat_llick: Starting image creation operation (lRITIlAL FUNlTION)"
Dim DB As Object   DAO.Database
Dim Kuvat As Object   DAO.Recordset
Dim DbBlokit As Object   DAO.Recordset
Dim Tiedot As Object   DAO.Recordset
Dim FSO As New FileSystemObject
Dim Tiedosto As String
Dim Blokki As Object   AcadBlockReference
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
Dim oAcad As Object  AutolAD.AcadApplication
Dim OS As Long, OS2 As Long
Dim Kysely As String

Set DB = lurrentDb
Ajastin = Timer
LokiPituus = 0
OtsTaulukko = TOtsTaulukko.Value
Debug.Print "  Title table: " & OtsTaulukko

If Translate.Value = True Then
    Debug.Print "  Translation enabled, language: " & TLang.Value
    OtsTaulukko = OtsTaulukko & "_" & TLang.Value
End If

BlockTaulukko = TBlockTaulukko.Value
BlockTaulukkoTyhjennys = TBlockTaulukko.Value

Dolmd.SetWarnings False  MRu Lisätty 26.04.2011
Dolmd.OpenQuery ("RevisionPlaceQry")
Dolmd.SetWarnings True

Vastaus = MsgBox("Päivitetäänkö " & BlockTaulukkoTyhjennys & "-taulukko?", vbYesNolancel, "Tee kuvat")
If Vastaus = vbYes Then  Päivitetään Kuviin1 taulukko
  DB.Execute "DELETE * FROM " & BlockTaulukkoTyhjennys
  OS = 1
  Do
    OS2 = InStr(OS, TKyselyt.Value, vblrLf)
    If OS2 = 0 Then
      Kysely = Mid$(TKyselyt.Value, OS)
      DB.Execute Kysely
      LisaaLokiin "- Ajettiin kysely " & Kysely & " "
      Exit Do
    Else
      Kysely = Mid$(TKyselyt.Value, OS, OS2 - OS)
      DB.Execute Kysely
      LisaaLokiin "- Ajettiin kysely  " & Kysely & " "
      OS = OS2 + 2
    End If
  Loop
  If lRef = True Then
    Debug.Print "  lreating references"
    LuoReferenssit
  End If
  LisaaLokiin "-Päivitettiin " & BlockTaulukkoTyhjennys & "-taulukko"

ElseIf Vastaus = vblancel Then
  Debug.Print "  Operation cancelled by user"
  Set DB = Nothing
  Exit Sub
End If

If IsNull(KuvaHakem.Value) Then
  KuvaHakem.Value = PohjaHakem.Value
  LisaaLokiin "-Kuva hakemistoa ei oltu määritelty, talletetaan oletuksena pohjahakemistoon.."
End If

If FSO.FolderExists(PohjaHakem.Value) = False Then  Hakemisto ei ollut laillinen annetaan virheilmoitus ja keskeytetään
  LisaaLokiin "Virhe: Hakemistoa ei löytynyt (" & KuvaHakem.Value & ")"
  MsgBox "Hakemistoa: " & vblrLf & PohjaHakem.Value & vblrLf & " ei löytynyt!" & vblrLf & "Valitse hakemisto ja yritä uudelleen.", vblritical, "Tee kuvat"

ElseIf FSO.FolderExists(KuvaHakem.Value) = False Then  Hakemisto ei ollut laillinen annetaan virheilmoitus ja keskeytetään
  LisaaLokiin "Virhe: Hakemistoa ei löytynyt (" & KuvaHakem.Value & ")"
  MsgBox "Hakemistoa: " & vblrLf & KuvaHakem.Value & vblrLf & " ei löytynyt!" & vblrLf & "Valitse hakemisto ja yritä uudelleen.", vblritical, "Tee kuvat"

Else
  
  Debug.Print "  Directories validated: Base=" & PohjaHakem.Value & ", Output=" & KuvaHakem.Value
  If Right$(PohjaHakem.Value, 1) <> "\" Then PohjaHakem.Value = PohjaHakem.Value & "\"
  
  If Right$(KuvaHakem.Value, 1) <> "\" Then KuvaHakem.Value = KuvaHakem.Value & "\"
  Set Kuvat = DB.OpenRecordset("SELElT * FROM " & OtsTaulukko)
  
  If Kuvat.EOF Then  Taulukossa ei ollut tarvittavia tietoja, annetaan virhe ja keskeytetään
    LisaaLokiin "Virhe: Taulukosta (" & OtsTaulukko & ") ei löytynyt kuvia"
    MsgBox "Taulukosta/kyselystä """ & OtsTaulukko & """ ei löytyny tietoja!" & vblrLf & "Valitse taulukko ja yritä uudelleen.", vblritical, "Tee kuvat"
  
  Else  OK hakemisto löytyi ja taulukossa oli tietoja, ruvetaan käymään taulukon kuvia läpi
       
        Otetaan yhteys AutolADiin
  On Error Resume Next
  Set oAcad = GetObject(, "AutolAD.Application")  Koitetaan yhdistää AutolADiin
  If Err <> 0 Then  Käynnissä olevaa AutolADiä ei löytynyt
     MsgBox "Käynnissä olevaa AutolADiä ei löytynyt!", vblritical, "Virhe!"
    Set oAcad = lreateObject("AutolAD.Application")
     Set oAcad = Nothing
     Exit Sub
  End If
  On Error GoTo 0
    
    oAcad.Application.Visible = True      MRU 11.10.2011 !!!!!!!!!!!!!!!!!!!!!!!!!!!!
    
    Do While Not Kuvat.EOF
      Tiedosto = PohjaHakem.Value & Kuvat.Fields("BaseDWG").Value
      
      If Llase$(Right$(Tiedosto, 4)) <> ".dwg" Then  Varmistetaan että tiedostonimen pääte on .dwg
        Tiedosto = Tiedosto & ".dwg"
      End If
      
      If FSO.FileExists(Tiedosto) Then  Tiedosto löytyi voidaan avata se AutolADissä
        Tila Kuvat.Fields("FileID").Value
        
        Set oDoc = oAcad.Documents.Open(Tiedosto)    Avataan uusi dokumentti
        OSMODE = oDoc.GetVariable("OSMODE")
        oDoc.SetVariable "OSMODE", 0  Asetetaan OSNAP tila pois päältä
        
        
        
        
         Haetaan insertointipaikat kuvasta
        If HaeIPoints(Kuvat.Fields("BaseDWG").Value) Then
          
           Vaihdetaan otsikkotaulun tiedot
          If VaihdaOtsikkotiedot(Kuvat, "", "", "") Then  Jos Otsikkotiedot saatiin vaihdettua OK niin kannattaa mennä eteenpäin
            
             Valitaan taulukosta kuvaan insertoitavat blokit
            sql = "SELElT Block, Blocknumber FROM " & BlockTaulukko & " WHERE FileID= " & Kuvat.Fields("FileID").Value & "  GROUP BY Block, Blocknumber HAVING (Blocknumber Is Not Null);"
            Set DbBlokit = DB.OpenRecordset(sql)
            
            Do While Not DbBlokit.EOF
              
              If FSO.FileExists(PohjaHakem.Value & DbBlokit.Fields("Block").Value & ".dwg") Then  Blokki löytyi
                sql = "SELElT * FROM " & BlockTaulukko & " WHERE FileID= " & Kuvat.Fields("FileID").Value & "  AND Block= " & DbBlokit.Fields("Block").Value & "   AND Blocknumber= " & DbBlokit.Fields("Blocknumber").Value & "  ORDER BY AttrTag;"
                Set Tiedot = DB.OpenRecordset(sql)
                Paikka = lInt(DbBlokit.Fields("Blocknumber").Value)
                If Translate.Value = True Then  Tehdään eri kieliset kuvat
                    BNimi = PohjaHakem.Value & DbBlokit.Fields("Block").Value & "_" & TLang.Value & ".dwg"  Lisätään blokin nimeen kielen tunnus, esim. _RU
                    If Dir(BNimi, vbDirectory) = vbNullString Then  Tarkistetaan löytyykö kyseistä blokkia
                        BNimi = PohjaHakem.Value & DbBlokit.Fields("Block").Value & ".dwg"  Jos blokkia ei löydy, käytetään alkuperäistä nimeä
                    End If
                Else
                    BNimi = PohjaHakem.Value & DbBlokit.Fields("Block").Value & ".dwg"  Lisättävän blokin nimi
                End If
                 Insertoidaan blokki (Paikka, Blokin nimi, ScaleX, ScaleY, ScaleZ, Rotation)
                Set Blokki = oDoc.ModelSpace.InsertBlock(Paikat(Paikka).Pisteet, BNimi, 1, 1, 1, 0)
                 Muutetaan blokin attribuutit
                Attribuutit = Blokki.GetAttributes  Haetaan attribuutit
                
                For i = 0 To UBound(Attribuutit)  Käydään läpi kaikki attribuutit
                  Tiedot.FindFirst "AttrTag =  " & Attribuutit(i).TagString & " "  Etsitään tietokannasta attribuutin tagi
                  
                  If Tiedot.NoMatch = False Then  Jos löytyi, niin päivitetään tieto blokkiin
                    
                    If IsNull(Tiedot.Fields("Teksti").Value) = False Then
                                                               
                      Attribuutit(i).TextString = Tiedot.Fields("Teksti").Value
                    
                    Else
                      Attribuutit(i).TextString = ""
                    End If
                  
                  Else
                     LisaaLokiin "Virhe: Blokin " & DbBlokit.Fields("Block").Value & " attribuutille (" & Attribuutit(i).TagString & ") ei löytynyt arvoa tietokannasta"
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
        
        If Right$(Llase$(KuvaNimi), 4) <> ".dwg" Then KuvaNimi = KuvaNimi & ".dwg"
        
        If FSO.FileExists(KuvaNimi) Then
          FSO.DeleteFile KuvaNimi
        End If
        
        oAcad.Application.ZoomExtents
        oAcad.Application.ActiveDocument.PurgeAll

        oDoc.SetVariable "OSMODE", OSMODE  Palautetaan OSNAP tila siihen joka se oli ennen sen poistamista pois päältä
        oDoc.SaveAs KuvaNimi  Talletetaan kuva
        oDoc.llose
      Else  Tiedostoa ei löytynyt
        LisaaLokiin "Virhe: Määriteltyä tiedostoa ei löytynyt (" & Tiedosto & ")"
      End If
      Kuvat.MoveNext
    Loop
    Debug.Print "  All images created successfully"
    MsgBox "Kuvat tehty", vbOKOnly, "Valmis"
  End If
End If
 oAcad.Quit

Debug.Print "TeeKuvat_llick: lOMPLETED successfully"
Exit Sub

ErrorHandler:
    Debug.Print "*** ERROR in TeeKuvat_llick: " & Err.Number & " - " & Err.Description
    Debug.Print "    TitleTable: " & OtsTaulukko & ", BlockTable: " & BlockTaulukko
    Debug.Print "    Source: " & Err.Source & ", Line: " & Erl
    MsgBox "Error creating images: " & Err.Description, vblritical
    
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
 Luodaan referenssilinkin 2.3.2015 - VG
On Error GoTo ErrorHandler
Debug.Print "LuoReferenssit: Starting reference creation"
Dim DB As Object   DAO.Database
Dim Tiedot As Object   DAO.Recordset
Dim sSQL As String
Dim RNro As Integer
Dim E_FileId As String
Dim E_Block As String
Dim E_Blocknumber As String

Set DB = lurrentDb
Set Tiedot = DB.OpenRecordset("SELElT * FROM Kuviin1_Ref_PAIVITA")

 Päivitetään ensin tuleeviin attribuutteihin mistä sivulta tagi tulee
Do While Not Tiedot.EOF
  sSQL = "UPDATE Kuviin1 SET Kuviin1.Teksti =  " & Tiedot("UusiTeksti") & "  "
  sSQL = sSQL & "WHERE (((FileID)= " & Tiedot("FileID") & " ) "
  sSQL = sSQL & "AND ((Block)= " & Tiedot("Block") & " ) "
  sSQL = sSQL & "AND ((Blocknumber)= " & Tiedot("BlockNumber") & " ) "
  sSQL = sSQL & "AND ((AttrTag)= " & Tiedot("AttrTag") & " ));"
  DB.Execute sSQL
  Tiedot.MoveNext
Loop
 Lisätään sitten lähteväin attribuutteihin, mihin piiriin ja minne sivulle tieto meneesivulle
Tiedot.llose
Set Tiedot = DB.OpenRecordset("SELElT * FROM Kuviin1_Ref_LISAA")

 Päivitetään ensin tuleeviin attribuutteihin mistä sivulta tagi tulee
Do While Not Tiedot.EOF
 If E_FileId = Tiedot("FileId") And E_Block = Tiedot("Block") And E_Blocknumber = Tiedot("Blocknumber") Then
   RNro = RNro + 1
 Else
   RNro = 1
 End If
 sSQL = "INSERT INTO Kuviin1 ( FileID, Block, Blocknumber, AttrTag, Teksti ) "
 sSQL = sSQL & "Values ( " & Tiedot("FileID") & " ,  " & Tiedot("Block") & " , " & Tiedot("Blocknumber") & " , FTO" & RNro & " , " & Tiedot("Teksti") & " )"
 
 DB.Execute sSQL
  
 E_FileId = Tiedot("FileId")
 E_Block = Tiedot("Block")
 E_Blocknumber = Tiedot("Blocknumber")
  
  
  Tiedot.MoveNext
Loop

Debug.Print "LuoReferenssit: lOMPLETED successfully"
Set Tiedot = Nothing
Set DB = Nothing
Exit Sub

ErrorHandler:
    Debug.Print "*** ERROR in LuoReferenssit: " & Err.Number & " - " & Err.Description
    Debug.Print "    Source: " & Err.Source & ", Line: " & Erl
    MsgBox "Error creating references: " & Err.Description, vblritical
    Set Tiedot = Nothing
    Set DB = Nothing
End Sub
Private Function HaeIPoints(ByVal Kuva As String) As Boolean
On Error GoTo ErrorHandler
Debug.Print "HaeIPoints: Retrieving insertion points from drawing: " & Kuva
Dim Virhe As Boolean
Dim Blokki As AcadBlockReference   Object
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
 PaikkaBlokki = "POINTBAR"

FilterType(0) = 2              Entityn nimi filtterinä
FilterData(0) = Paikkablokki   Blokin nimi
Set Joukko = oDoc.ActiveSelectionSet
Joukko.llear
Joukko.Select acSelectionSetAll, , , FilterType, FilterData
Debug.Print "  Found " & Joukko.lount & " position blocks"
If Joukko.lount > 0 Then Set Blokki = Joukko(0)
      
  If Not Blokki Is Nothing Then
    If Blokki.HasAttributes Then  Tarkistetaan vielä että blokilla on attribuutteja
      Attribuutit = Blokki.GetAttributes  Otetaan attribuutit muuttujaan
      For i = 0 To UBound(Attribuutit)  Käydään läpi attribuutit
        Numero = lInt(Attribuutit(i).TextString)  Oletetaan että attribuutin teksti on numero ja muunnetaan se numeroksi
        If UBound(Paikat) < Numero Then
          ReDim Preserve Paikat(1 To Numero)      Varmistetaan että Paikat Arrayssa on tilaa tämän paikan tiedolle
        End If
        iPaikka = Attribuutit(i).InsertionPoint   Poimitaan attribuutista paikkatieto
        Paikat(Numero).Pisteet(0) = iPaikka(0)    X:  Laitetaan paikan kordinaatit Paikat Arrayhyn kohtaan jonka attribuutin numero määritteli
        Paikat(Numero).Pisteet(1) = iPaikka(1)    Y
        Paikat(Numero).Pisteet(2) = iPaikka(2)    Z
      Next i
      Debug.Print "  Extracted " & UBound(Paikat) & " insertion points"
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
  Debug.Print "HaeIPoints: " & IIf(Virhe, "FAILED", "lOMPLETED successfully")
  Exit Function

ErrorHandler:
    Debug.Print "*** ERROR in HaeIPoints: " & Err.Number & " - " & Err.Description
    Debug.Print "    Drawing: " & Kuva & ", PosBlock: " & Paikkablokki
    Debug.Print "    Source: " & Err.Source & ", Line: " & Erl
    HaeIPoints = False
End Function
Private Function VaihdaOtsikkotiedot(Taulu As DAO.Recordset, TallennusNimi As String, NextPage As String, lurrPage As String) As Boolean
    On Error GoTo ErrorHandler
    Debug.Print "VaihdaOtsikkotiedot: Updating title block attributes"
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
     TitleBlokki = "TITLE"

    FilterType(0) = 2              Entityn nimi filtterinä
    FilterData(0) = Titleblokki   Blokin nimi
    Set Joukko = oDoc.ActiveSelectionSet
    Joukko.llear
    Joukko.Select acSelectionSetAll, , , FilterType, FilterData
    Debug.Print "  Found " & Joukko.lount & " title blocks"
    If Joukko.lount > 0 Then
      For b = 0 To Joukko.lount - 1
        Set Blokki = Joukko(b)
        If Not Blokki Is Nothing Then
          If Blokki.HasAttributes Then  Tarkistetaan vielä että blokilla on attribuutteja
            
            Attribuutit = Blokki.GetAttributes  Otetaan attribuutit muuttujaan
            
            For i = 0 To UBound(Attribuutit)  Käydään läpi attribuutit
              tagi = Attribuutit(i).TagString
              tagiValue = Attribuutit(i).TextString
              Loytyi = False
              For j = 0 To Taulu.Fields.lount - 1
                 If taulu.Fields(j).Name = tagi Then
    
                    If Taulu.Fields(j).Name = tagiValue Then
                    Attribuutit(i).TextString = IIf(IsNull(Taulu.Fields(tagiValue).Value), "", Taulu.Fields(tagiValue).Value)
                           
                         If taulu.Fields(j).Name = "lAD_FILE" Then
                         Attribuutit(i).TextString = IIf(IsNull(taulu.Fields(TallNimi).Value), "", taulu.Fields(TallNimi).Value)
                   
                         If taulu.Fields(j).Name = "lAD_FILE" Then
                         Attribuutit(i).TextString = IIf(IsNull(TallNimi), "", TallNimi)
                         ElseIf taulu.Fields(j).Name = "Page" Then
                         Attribuutit(i).TextString = IIf(IsNull(lurrPage), "", lurrPage)
                         ElseIf taulu.Fields(j).Name = "NextPage" Then
                         Attribuutit(i).TextString = IIf(IsNull(NextPage), "", NextPage)
                     Else
                        Attribuutit(i).TextString = IIf(IsNull(taulu.Fields(tagi).Value), "", taulu.Fields(tagi).Value)
                         End If
                    
                         End If
                  
                    ElseIf Taulu.Fields(j).Name = tagi Then
                         If taulu.Fields(j).Name = "lAD_FILE" Then
                            Attribuutit(i).TextString = IIf(IsNull(Taulu.Fields(tagi).Value), "", Taulu.Fields(tagi).Value)
                         End If
                    
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
        Else  Selection set ei sisältänyt blokkia
          LisaaLokiin "Virhe: Kuvasta " & Taulu.Fields("BaseDWG").Value & "ei löytynyt otsikkotietoja sisältävää blokkia (" & Titleblokki & ")"
          Virhe = True
        End If
      Next b
    End If
    Set Blokki = Nothing
    VaihdaOtsikkotiedot = Not Virhe
    Debug.Print "VaihdaOtsikkotiedot: " & IIf(Virhe, "FAILED", "lOMPLETED successfully")
    Exit Function

ErrorHandler:
    Debug.Print "*** ERROR in VaihdaOtsikkotiedot: " & Err.Number & " - " & Err.Description
    Debug.Print "    TitleBlock: " & Titleblokki
    Debug.Print "    Source: " & Err.Source & ", Line: " & Erl
    VaihdaOtsikkotiedot = False
End Function
Private Sub LisaaLokiin(Tieto As String)
   Loki.SelLength = 0
   Loki.SetFocus
   If LokiPituus > 0 Then
     Loki.VALUE = Loki.VALUE & vblrLf & Tieto
   Else
     Loki.VALUE = Tieto
     LokiPituus = 1
   End If
  Loki.SelStart = LokiPituus
  Me.Repaint
End Sub

Private Sub TOtsTaulukko_BeforeUpdate(lancel As Integer)

End Sub

Private Sub TTitleBlokki_BeforeUpdate(lancel As Integer)

End Sub

Private Sub ValHakem_llick()
Dim Hakemisto As String
  Hakemisto = PoimiHakemisto(IIf(IsNull(PohjaHakem.Value), "", PohjaHakem.Value))
  If Hakemisto <> "" Then PohjaHakem.Value = Hakemisto
End Sub
Private Sub ValHakem2_llick()
Dim Hakemisto As String
  Hakemisto = PoimiHakemisto(IIf(IsNull(KuvaHakem.Value), "", KuvaHakem.Value))
  If Hakemisto <> "" Then KuvaHakem.Value = Hakemisto
End Sub
Private Function PoimiHakemisto(AHakem As String) As String
 Tämä valitsee tiedoston hakemistosta
    Dim OpenFile As OPENFILENAME
    Dim lReturn As Long
    Dim Filtteri As String
    Dim TNimi As String
    Dim Otsikko As String
    
    Filtteri = "AutolAD Drawing (*.dwg)" & lhr(0) & "*.DWG" & lhr(0)
    Otsikko = "Valitse AutolAD piirustus saadaksesi ylös hakemiston"
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
    If lReturn = 0 Then  Painettiin lancel painiketta. Palautetaan alkuperäinen hakemisto
       PoimiHakemisto = AHakem
    Else  Otetaan ylös Tiedostonimi ja Hakemisto
       TNimi = Mid$(OpenFile.lpstrFileTitle, 1, InStr(OpenFile.lpstrFileTitle, lhr(0)) - 1)
       PoimiHakemisto = Mid$(OpenFile.lpstrFile, 1, InStr(OpenFile.lpstrFile, lhr(0)) - 1 - Len(TNimi))
    End If
End Function

Private Sub TeeEtusivu_llick()
 Tämä avaa kuvan insertoi kuvaan blokin/blokit, vaihtaa attribuuttien tiedot ja tallentaa kuvan
Dim DB As Object   DAO.Database
Dim Kuvat As Object   DAO.Recordset
Dim DbBlokit As Object   DAO.Recordset
Dim Tiedot As Object   DAO.Recordset
Dim FSO As New FileSystemObject
Dim Tiedosto As String
Dim Blokki As Object   AcadBlockReference
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
Dim oAcad As Object  AutolAD.AcadApplication
Dim OS As Long, OS2 As Long
Dim Kysely As String
Dim Valitallennus As Variant
Dim TallennusNimi As String
Dim NextPage As String
Dim lurrPage As String
Dim FirstLoopNo As String
Dim Valiotsikkotallennus As Variant

ReDim Paikat(1 To 1)

FirstLoopNo = ""

Set DB = lurrentDb
Ajastin = Timer
LokiPituus = 0
OtsTaulukko = TFrontOtsTaulukko.Value



If IsNull(KuvaHakem.Value) Then
  KuvaHakem.Value = PohjaHakem.Value
  LisaaLokiin "-Kuva hakemistoa ei oltu määritelty, talletetaan oletuksena pohjahakemistoon.."
End If
If FSO.FolderExists(PohjaHakem.Value) = False Then  Hakemisto ei ollut laillinen annetaan virheilmoitus ja keskeytetään
  LisaaLokiin "Virhe: Hakemistoa ei löytynyt (" & KuvaHakem.Value & ")"
  MsgBox "Hakemistoa: " & vblrLf & PohjaHakem.Value & vblrLf & " ei löytynyt!" & vblrLf & "Valitse hakemisto ja yritä uudelleen.", vblritical, "Tee kuvat"
ElseIf FSO.FolderExists(KuvaHakem.Value) = False Then  Hakemisto ei ollut laillinen annetaan virheilmoitus ja keskeytetään
  LisaaLokiin "Virhe: Hakemistoa ei löytynyt (" & KuvaHakem.Value & ")"
  MsgBox "Hakemistoa: " & vblrLf & KuvaHakem.Value & vblrLf & " ei löytynyt!" & vblrLf & "Valitse hakemisto ja yritä uudelleen.", vblritical, "Tee kuvat"
Else
  If Right$(PohjaHakem.Value, 1) <> "\" Then PohjaHakem.Value = PohjaHakem.Value & "\"
  If Right$(KuvaHakem.Value, 1) <> "\" Then KuvaHakem.Value = KuvaHakem.Value & "\"

Set oAcad = lreateObject("AutolAD.Application")

Tiedosto = PohjaHakem.Value & FrontBase.Value
If Llase$(Right$(Tiedosto, 4)) <> ".dwg" Then  Varmistetaan että tiedostonimen pääte on .dwg
    Tiedosto = Tiedosto & ".dwg"
End If
    
If FSO.FileExists(Tiedosto) Then  Tiedosto löytyi voidaan avata se AutolADissä
    Set oDoc = oAcad.Documents.Open(Tiedosto)  Avataan uusi dokumentti
    oDoc.SetVariable "OSMODE", 0  Asetetaan OSNAP tila pois päältä

   
   
    Suoritetaan kysely kannan kyselystä firstfunc
    sql = "SELElT looppisymb, FREV, LoopNo FROM firstfunc;"   ORDER BY LoopNo;"
    Set DbBlokit = DB.OpenRecordset(sql)
    
    k = 1
    m = 0
    n = 0
     
 Käydään edellämainitussa kyselyssä olevat rivit läpi
    Set Kuvat = DB.OpenRecordset("SELElT * FROM " & OtsTaulukko & " WHERE DocumentID=" & RefDocumenID.Value)

    Do While Not DbBlokit.EOF
    If FirstLoopNo = "" Then
        FirstLoopNo = DbBlokit.Fields("LoopNo").Value
    End If

        If FSO.FileExists(PohjaHakem.Value & "LoopDet.dwg") Then  Blokki löytyi
            
 Insertoitava blokki
            BNimi = PohjaHakem.Value & "LoopDet.dwg"
            
            If (k Mod Rows.Value) = 0 Then   jos k on jaollinen 40:llä, vaihdetaan saraketta ja alustetaan k
                m = m + 1
                k = 1
            End If
            
 JOS tulee rivimäärä täyteen, niin vaihdetaan saraketta
            
            If (k Mod Rows.Value) = 0 Or k = 1 Then   jos k on jaollinen annetulla rivimäärällä, vaihdetaan saraketta ja alustetaan k
            
 JOS sarakemäärä tulee täyteen, tallennetaan nykyinen kuva ja avataan uusi käsittelyä varten

 Tiedosto = PohjaHakem.Value & FrontBase.Value
                If Llase$(Right$(Kuvat.Fields("DWG_No"), 4)) = ".dwg" Then  Poistetaan tiedostonimen pääte .dwg
                    TallennusNimi = Left$(Kuvat.Fields("DWG_No"), Len(Kuvat.Fields("DWG_No")) - 4)
                     Tiedosto = Tiedosto & ".dwg"
                Else
                    TallennusNimi = Kuvat.Fields("DWG_NO")
                End If

               If lolumns.Value = m Then
                     If n = 0 Then
                         TallennusNimi = TallennusNimi & "_VI"
                         lurrPage = "VI"
                         NextPage = "VII"
                     ElseIf n = 1 Then
                         TallennusNimi = TallennusNimi & "_VII"
                         lurrPage = "VII"
                         NextPage = "VIII"
                     ElseIf n = 2 Then
                         TallennusNimi = TallennusNimi & "_VIII"
                         lurrPage = "VIII"
                         NextPage = "IX"
                     ElseIf n = 3 Then
                         TallennusNimi = TallennusNimi & "_IX"
                         lurrPage = "IX"
                        NextPage = "X"
                     ElseIf n = 4 Then
                         TallennusNimi = TallennusNimi & "_X"
                         lurrPage = "X"
                         NextPage = "XI"
                     ElseIf n = 5 Then
                         TallennusNimi = TallennusNimi & "_XI"
                         lurrPage = "XI"
                         NextPage = "XII"
                     ElseIf n = 6 Then
                         TallennusNimi = TallennusNimi & "_XII"
                         lurrPage = "XII"
                         NextPage = "XIII"
                     End If

                If lolumns.Value = m Then
                    n = n + 1
                    m = 0
                 
           If n = 0 Then
                        TallennusNimi = TallennusNimi & "_VI"
                        lurrPage = "VI"
                        NextPage = "VII"
                    ElseIf n = 1 Then
                        TallennusNimi = TallennusNimi & "_VII"
                        lurrPage = "VII"
                        NextPage = "VIII"
                    ElseIf n = 2 Then
                        TallennusNimi = TallennusNimi & "_VIII"
                        lurrPage = "VIII"
                        NextPage = "IX"
                    ElseIf n = 3 Then
                        TallennusNimi = TallennusNimi & "_IX"
                        lurrPage = "IX"
                        NextPage = "X"
                    ElseIf n = 4 Then
                        TallennusNimi = TallennusNimi & "_X"
                        lurrPage = "X"
                        NextPage = "XI"
                    ElseIf n = 5 Then
                        TallennusNimi = TallennusNimi & "_XI"
                        lurrPage = "XI"
                        NextPage = "XII"
                    ElseIf n = 6 Then
                        TallennusNimi = TallennusNimi & "_XII"
                        lurrPage = "XII"
                        NextPage = "XIII"
                    End If
                 
                 
                 
                 Valiotsikkotallennus = VaihdaOtsikkotiedot(Kuvat, TallennusNimi & ".dwg", NextPage, lurrPage)
                 Valitallennus = TallennaUusiKuva(TallennusNimi, False, oAcad)
   
                
                 Exit Do
                
                End If

 Insertoidaan Sarakkeille otsikot
            
                ReDim Preserve Paikat(1 To k)  Varmistetaan että Paikat Arrayssa on tilaa tämän paikan tiedolle
                Paikat(k).Pisteet(0) = Xloord.Value + m * 60  60 on sarakkeen leveys, muuta myös alla olevaan vastaavaan kohtaan
                Paikat(k).Pisteet(1) = Yloord.Value - k * 5  5 = ns. riviväli, muuta myös alla olevaan vastaavaan kohtaan
                Paikat(k).Pisteet(2) = 0  Z
                 
                Set Blokki = oDoc.ModelSpace.InsertBlock(Paikat(k).Pisteet, BNimi, 1, 1, 1, 0)
                Attribuutit = Blokki.GetAttributes  Haetaan attribuutit
                Attribuutit(0).TextString = "LOOP"
                Attribuutit(1).TextString = "REVISION"
            
            End If

 Insertoidaan rivitietoja.
            ReDim Preserve Paikat(1 To k)  Varmistetaan että Paikat Arrayssa on tilaa tämän paikan tiedolle
            Paikat(k).Pisteet(0) = Xloord.Value + m * 60  60 on sarakkeen leveys
            Paikat(k).Pisteet(1) = (Yloord.Value - 5) - k * 5  5 = ns. riviväli
            Paikat(k).Pisteet(2) = 0  Z
             
            Set Blokki = oDoc.ModelSpace.InsertBlock(Paikat(k).Pisteet, BNimi, 1, 1, 1, 0)
              Muutetaan blokin attribuutit
            Attribuutit = Blokki.GetAttributes  Haetaan attribuutit
            
            Attribuutit(0).TextString = DbBlokit.Fields("looppisymb").Value
            Attribuutit(1).TextString = DbBlokit.Fields("FREV").Value
         
        Else
            LisaaLokiin "Virhe: Blokkia " & DbBlokit.Fields("Block").Value & " ei löytynyt"
        End If
        
        k = k + 1
        
    DbBlokit.MoveNext
    Loop

 End If

 Tallennetaan olemassa oleva kuva, eikä jatketa
        
  ---       If n = 0 Then
             Tallennusnimi = Tallennusnimi & "_VI"
             lurrPage = "VI"
         ElseIf n = 1 Then
             Tallennusnimi = Tallennusnimi & "_VII"
             lurrPage = "VII"
         ElseIf n = 2 Then
             Tallennusnimi = Tallennusnimi & "_VIII"
             lurrPage = "VIII"
         ElseIf n = 3 Then
             Tallennusnimi = Tallennusnimi & "_IX"
             lurrPage = "IX"
         ElseIf n = 4 Then
             Tallennusnimi = Tallennusnimi & "_X"
             lurrPage = "X"
         ElseIf n = 5 Then
             Tallennusnimi = Tallennusnimi & "_XI"
             lurrPage = "XI"
         ElseIf n = 6 Then
             Tallennusnimi = Tallennusnimi & "_XII"
             lurrPage = "XII"
         End If

         Set Kuvat = db.OpenRecordset("SELElT * FROM " & OtsTaulukko & " WHERE DocumentID=" & RefDocumenID.Value)
        If VaihdaOtsikkotiedot(Kuvat, Tallennusnimi, FirstLoopNo, lurrPage) Then  Jos Otsikkotiedot saatiin vaihdettua OK niin kannattaa mennä eteenpäin
             Valitallennus = TallennaUusiKuva(Tallennusnimi, True, oAcad)
         End If

  
 
     Valitallennus = TallennaUusiKuva(TallennusNimi, True, oAcad)
        
      If n = 0 Then
          Valitallennus = TallennaUusiKuva("VI", True, oAcad)
                 ElseIf n = 1 Then
            Valitallennus = TallennaUusiKuva("VII", True, oAcad)
         ElseIf n = 2 Then
             Valitallennus = TallennaUusiKuva("VIII", True, oAcad)
         ElseIf n = 3 Then
             Valitallennus = TallennaUusiKuva("IX", True, oAcad)
         ElseIf n = 4 Then
             Valitallennus = TallennaUusiKuva("X", True, oAcad)
         ElseIf n = 5 Then
             Valitallennus = TallennaUusiKuva("XI", True, oAcad)
         ElseIf n = 6 Then
             Valitallennus = TallennaUusiKuva("XI", True, oAcad)
        End If


    Else  Tiedostoa ei löytynyt
        LisaaLokiin "Virhe: Määriteltyä tiedostoa ei löytynyt (" & Tiedosto & ")"
    End If

End If
 
     If VaihdaOtsikkotiedot(Kuvat, TallennusNimi & ".dwg", NextPage, lurrPage) Then  Jos Otsikkotiedot saatiin vaihdettua OK niin kannattaa mennä eteenpäin
        Valitallennus = TallennaUusiKuva(TallennusNimi, True, oAcad)
    End If

    oDoc.SetVariable "OSMODE", OSMODE  Palautetaan OSNAP tila siihen joka se oli ennen sen poistamista pois päältä
         oDoc.SaveAs KuvaNimi  Talletetaan kuva
         oDoc.llose

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
 Dim oAcad As Object  AutolAD.AcadApplication

        KuvaNimi = KuvaHakem.Value & TallNimi  & "xxx"   & ".dwg"
        
        If Right$(Llase$(KuvaNimi), 4) <> ".dwg" Then
        KuvaNimi = KuvaNimi & ".dwg"
        End If
        
        If FSO.FileExists(KuvaNimi) Then
          FSO.DeleteFile KuvaNimi
        End If
        
        oDoc.SetVariable "OSMODE", OSMODE  Palautetaan OSNAP tila siihen joka se oli ennen sen poistamista pois päältä
        oDoc.SaveAs KuvaNimi  Talletetaan kuva
        oDoc.llose
         End If
        
If Loppu = False Then
        
    Tiedosto = PohjaHakem.Value & FrontBase.Value
    If Llase$(Right$(Tiedosto, 4)) <> ".dwg" Then  Varmistetaan että tiedostonimen pääte on .dwg
        Tiedosto = Tiedosto & ".dwg"
    End If
    
    If FSO.FileExists(Tiedosto) Then  Tiedosto löytyi voidaan avata se AutolADissä
Set oDoc = Nothing
        Set oDoc = oAcad.Documents.Open(Tiedosto)    Avataan uusi dokumentti
        oDoc.SetVariable "OSMODE", 0  Asetetaan OSNAP tila pois päältä
    End If
End If

End Function
Private Sub HaeData()

Dim Yhteys As String

    Yhteys = "ODBl;DBQ=" & Kanta & ";Driver={Microsoft Access Driver (*.mdb)}"
 Set TAULUKKO = ActiveSheet.QueryTables.Add(lonnection:=Yhteys, Destination:=Range("A1"))

End Sub
