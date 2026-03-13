Option Compare Database
Option Explicit
' Sleep API — estää busy wait -kuormituksen CPU:lla (VBA7/64-bit yhteensopiva)
#If VBA7 Then
    Private Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
#Else
    Private Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
#End If
'==========================================================================
' LOMAKE   : Form_LisääKuviin_ACAD
' SOVELLUS : FunctionDiagrams — AutoCAD-piirustusautomaatio
' KUVAUS   : Kattavin lomake funktiokaaviomoduulissa. Avaa AutoCAD-
'            piirustuksia (CreateObject/GetObject, late binding), insertoi
'            blokkeja, päivittää otsikkotietuetiedot ja tallentaa valmiit
'            piirustukset. HaeTekstit_Click / HaeValitutTekstit_Click
'            poimivat attribuuttitiedot DWG-tiedostoista Blokit-tauluun.
'            TeeKuvat_Click ajaa koko tuotantoajon. PaivitaDocRev_Click
'            synkronoi Documents-kannan revision- ja detaljitiedot.
'            PoimiHakemisto käyttää OPENFILENAME-dialogia (General.bas)
'            hakemiston valintaan.
' DIPLE    : Option Explicit; kaikki AutoCAD-objektit late binding (Object);
'            LisääLokiin-lokifunktio jokaisessa haarassa; Set = Nothing
'            kaikkiin virheenkäsittelijöihin.
' PÄIVITETTY: 2026-03-03
'==========================================================================
' AutoCAD SelectionSet-tyypin vakiot (late binding - ei vaadi kirjastoviitettä)
' Lähde: AutoCAD ActiveX/VBA Reference (AcSelect-luettelointi)
Private Const acSelectionSetAll As Integer = 5           ' Valitse kaikki entiteetit
Private Const acSelectionSetPrevious As Integer = 4     ' Edellinen valintajoukko
Dim Ajastin As Long
Dim min As Integer, sec As Integer
Dim OSMODE As Integer
Dim LokiPituus As Long
Dim oAcad As Object '(AcadApplication - muutettu late binding 64-bit yhteensopivuus)
Public oDoc As Object '(AcadDocument - muutettu late binding 64-bit yhteensopivuus)

' 3D-lisäyspiste AutoCAD InsertBlock-kutsulle (koordinaatit X, Y, Z)
Private Type AcadInsertPiste
    Pisteet(0 To 2) As Double
End Type
' Moduulitason insertointipistemasiiivi — jaettu HaeIPoints-, TeeKuvat- ja TeeEtusivu-aliohjelmien välillä
Private Paikat() As AcadInsertPiste

Private Sub Odota(Aika As Integer)
    ' Sleep vapauttaa CPU:n odotusajalla — busy wait korvattu kernel-tasoisella odotuksella
    DoEvents ' Käsitellään mahdolliset taustaviestit ennen odotusta
    Sleep CLng(Aika) * 10
End Sub
Private Sub Tila(Tiedosto As String)
  sec = CInt(Timer - Ajastin)
  min = sec \ 60
  sec = sec Mod 60
  SysCmd acSysCmdSetStatus, "Käsittelee tiedostoa: " & Tiedosto & "  Aika: " & min & " min  " & sec & " sek"
End Sub
Private Sub Form_Load()
  Debug.Print "Form_LisääKuviin_ACAD: Form loaded"
  Me.Loki.Value = ""
End Sub
Private Sub Form_Unload(Cancel As Integer)
  Debug.Print "Form_LisääKuviin_ACAD: Form unloading"
  DoCmd.Hourglass False
  SysCmd acSysCmdClearStatus
End Sub
Private Sub FormClose_Click()
 Debug.Print "FormClose_Click: Closing form"
 DoCmd.Close
End Sub
Private Sub HaeTekstit_Click()
On Error GoTo ErrorHandler
Debug.Print "HaeTekstit_Click: Starting text extraction from DWG files"
Dim Tied As String
Dim Polku As String
Dim Nimi As String
Dim Taulu As DAO.Recordset
Dim oEnt As Object '(acadentity - muutettu late binding)
Dim oAcad As Object '(AutoCAD.AcadApplication - muutettu late binding)
Dim db As DAO.Database

Ajastin = Timer
Debug.Print "  Timer started"

' Yritetään ensin liittyä jo auki olevaan AutoCAD-instanssiin — nopeampaa ja kevyempää
On Error Resume Next
Set oAcad = GetObject(, "AutoCAD.Application")
If Err.Number <> 0 Then
    ' AutoCAD ei ole auki — avataan uusi instanssi
    Err.Clear
    Set oAcad = CreateObject("AutoCAD.Application")
End If
On Error GoTo ErrorHandler
Debug.Print "  AutoCAD application ready"

Polku = Me.PohjaHakem.Value
If Right$(Polku, 1) <> "\" Then Polku = Polku & "\"
Debug.Print "  Path: " & Polku
Tied = Dir(Polku & "*.dwg")
Debug.Print "  Clearing Blokit table for path: " & Polku
' Poistetaan vain kyseisen hakemiston blokit — suojaa muiden projektien data
Set db = CurrentDb
db.Execute "DELETE * FROM Blokit WHERE Path='" & Replace(Polku, "'", "''") & "'", dbFailOnError
Set Taulu = db.OpenRecordset("Blokit", dbOpenDynaset)
Debug.Print "  Processing DWG files..."
Do While Tied <> ""
  Nimi = Left$(Tied, Len(Tied) - 4)
  If UCase$(Left$(Tied, 4)) <> "BASE" And UCase$(Left$(Tied, 8)) <> "BLOCKMAP" Then
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
        Taulu.Update
      End If
    Next
    oDoc.Close False
  End If
  Tied = Dir
Loop
Debug.Print "  All files processed"

Loppu:
On Error Resume Next
If Not Taulu Is Nothing Then
    Taulu.Close
    Set Taulu = Nothing
End If
Set db = Nothing
' AutoCAD-istuntoa ei suljeta tässä, jotta käyttäjä voi jatkaa työskentelyä tekstien haun jälkeen.
Set oAcad = Nothing
On Error GoTo 0
Debug.Print "HaeTekstit_Click: COMPLETED"
MsgBox "Valmis", vbInformation, "Päivitä tekstit"
Exit Sub

ErrorHandler:
    Debug.Print "*** ERROR in HaeTekstit_Click: " & Err.Number & " - " & Err.Description
    Debug.Print "    Path: " & Polku & ", File: " & Tied
    Debug.Print "    Source: " & Err.Source & ", Line: " & Erl
    MsgBox "Virhe tekstien poiminnassa: " & Err.Description, vbCritical
    Resume Loppu
End Sub


Private Sub HaeValitutTekstit_Click()
On Error GoTo ErrorHandler
Debug.Print "HaeValitutTekstit_Click: Starting selected text extraction"
Dim Tied As String
Dim Polku As String
Dim Taulu As DAO.Recordset
Dim Blokki As Object '(AcadBlockReference - muutettu late binding)
Dim oAcad As Object '(AutoCAD.AcadApplication - muutettu late binding)
Dim oDoc As Object '(AutoCAD.AcadDocument - muutettu late binding)
Dim Joukko As Object '(AcadSelectionSet - muutettu late binding)
Dim BlockArray As Variant       'Array muuttuja Blokkia varten
Dim i As Integer, j As Integer  'Indeksimuuttuja
Dim FilterType(0) As Integer
Dim FilterData(0) As Variant
Dim Hakemisto As String

          FilterType(0) = 0
          FilterData(0) = "INSERT"

Ajastin = Timer
Debug.Print "  Timer started"


Set oAcad = GetObject(, "AutoCAD.Application")
Debug.Print "  AutoCAD connection established"
Polku = Me.PohjaHakem.Value
If Right$(Polku, 1) <> "\" Then Polku = Polku & "\"
Tied = Dir(Polku & "*.dwg")

'CurrentDb.Execute "DELETE * FROM Blokit"

Set Taulu = CurrentDb.OpenRecordset("Blokit")

  Set oDoc = oAcad.ActiveDocument
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
    
    Debug.Print "  Processing block: " & Blokki.Name
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
      End If
      End If
 Next i
    
Taulu.Close
Set oAcad = Nothing
MsgBox "Valmis", vbInformation, "Päivitä tekstit"
Set Taulu = Nothing

Debug.Print "HaeValitutTekstit_Click: COMPLETED successfully"
Exit Sub

ErrorHandler:
    Debug.Print "*** ERROR in HaeValitutTekstit_Click: " & Err.Number & " - " & Err.Description
    Debug.Print "    Source: " & Err.Source & ", Line: " & Erl
    On Error Resume Next
    If Not Taulu Is Nothing Then Taulu.Close
    Set oAcad = Nothing
    Set Taulu = Nothing
    MsgBox "Virhe tekstien haussa: " & Err.Description, vbCritical
End Sub

Private Sub PaivitaDocRev_Click()
On Error GoTo ErrorHandler
Debug.Print "PaivitaDocRev_Click: Starting document revision update"

Dim dbsDocuments As DAO.Database
Dim rstSnapshot As DAO.Recordset
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
"WHERE Documents.ProjectID=" & Me.RefID.Value & ""

'Edellä olevasta kyselystä saadaan recordset rstSnapshot

Set rstSnapshot = dbsDocuments.OpenRecordset(SQLQry, dbOpenSnapshot)

'Jos ko. taulu on olemassa, poistetaan se

If TableExists("DocumentsDetail") = True Then
    Debug.Print "  Dropping existing DocumentsDetail table"
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

Debug.Print "  Creating DocumentsDetail table with " & rstSnapshot.Fields.Count & " fields"
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
        ' Suojataan heittomerkki: kenttäarvo ulkoisesta kannasta voi sisältää yksinkertaisen lainausmerkin
        KentanArvot = KentanArvot & "'" & Replace(CStr(rstSnapshot.Fields(laskuri)), "'", "''") & "',"
    
    Next laskuri
    
    Kentat = Kentat & rstSnapshot.Fields(laskuri).Name
    Kentat = Kentat & ")"
    KentanArvot = KentanArvot & "'" & Replace(CStr(rstSnapshot.Fields(laskuri)), "'", "''") & "'"
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
"WHERE Documents.ProjectID=" & Me.RefID.Value & ""


'Edellä olevasta kyselystä saadaan recordset rstSnapshot

Set rstSnapshot = dbsDocuments.OpenRecordset(SQLQry, dbOpenSnapshot)

'Jos ko. taulu on olemassa, poistetaan se

If TableExists("DocumentsRevision") = True Then
    Debug.Print "  Dropping existing DocumentsRevision table"
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

Debug.Print "  Creating DocumentsRevision table with " & rstSnapshot.Fields.Count & " fields"
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
        ' Suojataan heittomerkki: kenttäarvo ulkoisesta kannasta voi sisältää yksinkertaisen lainausmerkin
        KentanArvot = KentanArvot & "'" & Replace(CStr(rstSnapshot.Fields(laskuri)), "'", "''") & "',"
    
    Next laskuri
    
    Kentat = Kentat & rstSnapshot.Fields(laskuri).Name
    Kentat = Kentat & ")"
    KentanArvot = KentanArvot & "'" & Replace(CStr(rstSnapshot.Fields(laskuri)), "'", "''") & "'"
    KentanArvot = KentanArvot & ")"
    
    LisaaArvot = "INSERT INTO DocumentsRevision" & Kentat & " VALUES " & KentanArvot
    
    dbs.Execute LisaaArvot
    
    rstSnapshot.MoveNext
Next u




      
      
      
     

      
      
      
      rstSnapshot.Close
      Set rstSnapshot = Nothing
      dbsDocuments.Close
      Set dbsDocuments = Nothing
      Set dbs = Nothing

Debug.Print "PaivitaDocRev_Click: COMPLETED successfully - DocumentsDetail and DocumentsRevision tables updated"
Exit Sub

ErrorHandler:
    Debug.Print "*** ERROR in PaivitaDocRev_Click: " & Err.Number & " - " & Err.Description
    Debug.Print "    ProjectID: " & Me.RefID.Value
    Debug.Print "    Source: " & Err.Source & ", Line: " & Erl
    On Error Resume Next
    If Not rstSnapshot Is Nothing Then rstSnapshot.Close
    Set rstSnapshot = Nothing
    If Not dbsDocuments Is Nothing Then dbsDocuments.Close
    Set dbsDocuments = Nothing
    Set dbs = Nothing
    MsgBox "Virhe dokumenttirevision päivityksessä: " & Err.Description, vbCritical
End Sub

Public Function TableExists(sTable As String) As Boolean
    
    Dim DB As DAO.Database
    Dim tbl As DAO.TableDef
    Set DB = CurrentDb()
    
    TableExists = False
    
    For Each tbl In DB.TableDefs
        If tbl.Name = sTable Then TableExists = True
    Next tbl
    
End Function

Private Sub TeeKuvat_Click()
'Tämä avaa kuvan insertoi kuvaan blokin/blokit, vaihtaa attribuuttien tiedot ja tallentaa kuvan
On Error GoTo ErrorHandler
Debug.Print "TeeKuvat_Click: Starting image creation operation (CRITICAL FUNCTION)"
Dim DB As Object ' DAO.Database
Dim Kuvat As Object ' DAO.Recordset
Dim DbBlokit As Object ' DAO.Recordset
Dim Tiedot As Object ' DAO.Recordset
Dim FSO As New FileSystemObject
Dim Tiedosto As String
Dim Blokki As Object ' AcadBlockReference
Dim Attribuutit As Variant
Dim Paikka As Integer
Dim BNimi As String
Dim i As Integer
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
OtsTaulukko = Me.TOtsTaulukko.Value
Debug.Print "  Title table: " & OtsTaulukko

If Me.Translate.Value = True Then
    Debug.Print "  Translation enabled, language: " & Me.TLang.Value
    OtsTaulukko = OtsTaulukko & "_" & Me.TLang.Value
End If

BlockTaulukko = Me.TBlockTaulukko.Value
BlockTaulukkoTyhjennys = Me.TBlockTaulukko.Value

DoCmd.SetWarnings False 'MRu Lisätty 26.04.2011
DoCmd.OpenQuery ("RevisionPlaceQry")
DoCmd.SetWarnings True

Vastaus = MsgBox("Päivitetäänkö " & BlockTaulukkoTyhjennys & "-taulukko?", vbYesNoCancel, "Tee kuvat")
If Vastaus = vbYes Then 'Päivitetään Kuviin1 taulukko
  DB.Execute "DELETE * FROM " & BlockTaulukkoTyhjennys
  OS = 1
  Do
    OS2 = InStr(OS, Me.TKyselyt.Value, vbCrLf)
    If OS2 = 0 Then
      Kysely = Mid$(Me.TKyselyt.Value, OS)
      DB.Execute Kysely
      LisaaLokiin "- Ajettiin kysely " & Kysely & "'"
      Exit Do
    Else
      Kysely = Mid$(Me.TKyselyt.Value, OS, OS2 - OS)
      DB.Execute Kysely
      LisaaLokiin "- Ajettiin kysely '" & Kysely & "'"
      OS = OS2 + 2
    End If
  Loop
  If Me.CRef = True Then
    Debug.Print "  Creating references"
    LuoReferenssit
  End If
  LisaaLokiin "-Päivitettiin " & BlockTaulukkoTyhjennys & "-taulukko"

ElseIf Vastaus = vbCancel Then
  Debug.Print "  Operation cancelled by user"
  Set DB = Nothing
  Exit Sub
End If

If IsNull(Me.KuvaHakem.Value) Then
  Me.KuvaHakem.Value = Me.PohjaHakem.Value
  LisaaLokiin "-Kuva hakemistoa ei oltu määritelty, talletetaan oletuksena pohjahakemistoon.."
End If

If FSO.FolderExists(Me.PohjaHakem.Value) = False Then 'Hakemisto ei ollut laillinen annetaan virheilmoitus ja keskeytetään
  LisaaLokiin "Virhe: Hakemistoa ei löytynyt (" & Me.KuvaHakem.Value & ")"
  MsgBox "Hakemistoa: " & vbCrLf & Me.PohjaHakem.Value & vbCrLf & " ei löytynyt!" & vbCrLf & "Valitse hakemisto ja yritä uudelleen.", vbCritical, "Tee kuvat"

ElseIf FSO.FolderExists(Me.KuvaHakem.Value) = False Then 'Hakemisto ei ollut laillinen annetaan virheilmoitus ja keskeytetään
  LisaaLokiin "Virhe: Hakemistoa ei löytynyt (" & Me.KuvaHakem.Value & ")"
  MsgBox "Hakemistoa: " & vbCrLf & Me.KuvaHakem.Value & vbCrLf & " ei löytynyt!" & vbCrLf & "Valitse hakemisto ja yritä uudelleen.", vbCritical, "Tee kuvat"

Else
  
  Debug.Print "  Directories validated: Base=" & Me.PohjaHakem.Value & ", Output=" & Me.KuvaHakem.Value
  If Right$(Me.PohjaHakem.Value, 1) <> "\" Then Me.PohjaHakem.Value = Me.PohjaHakem.Value & "\"
  
  If Right$(Me.KuvaHakem.Value, 1) <> "\" Then Me.KuvaHakem.Value = Me.KuvaHakem.Value & "\"
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
      Tiedosto = Me.PohjaHakem.Value & Kuvat.Fields("BaseDWG").Value
      
      If LCase$(Right$(Tiedosto, 4)) <> ".dwg" Then 'Varmistetaan että tiedostonimen pääte on .dwg
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
              
              If FSO.FileExists(Me.PohjaHakem.Value & DbBlokit.Fields("Block").Value & ".dwg") Then 'Blokki löytyi
                sql = "SELECT * FROM " & BlockTaulukko & " WHERE FileID='" & Kuvat.Fields("FileID").Value & "' AND Block='" & DbBlokit.Fields("Block").Value & "'  AND Blocknumber='" & DbBlokit.Fields("Blocknumber").Value & "' ORDER BY AttrTag;"
                Set Tiedot = DB.OpenRecordset(sql)
                Paikka = CInt(DbBlokit.Fields("Blocknumber").Value)
                If Me.Translate.Value = True Then 'Tehdään eri kieliset kuvat
                    BNimi = Me.PohjaHakem.Value & DbBlokit.Fields("Block").Value & "_" & Me.TLang.Value & ".dwg" 'Lisätään blokin nimeen kielen tunnus, esim. _RU
                    If Dir(BNimi, vbDirectory) = vbNullString Then 'Tarkistetaan löytyykö kyseistä blokkia
                        BNimi = Me.PohjaHakem.Value & DbBlokit.Fields("Block").Value & ".dwg" 'Jos blokkia ei löydy, käytetään alkuperäistä nimeä
                    End If
                Else
                    BNimi = Me.PohjaHakem.Value & DbBlokit.Fields("Block").Value & ".dwg" 'Lisättävän blokin nimi
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
        
        KuvaNimi = Me.KuvaHakem.Value & Kuvat.Fields("FileID").Value
        
        If Right$(LCase$(KuvaNimi), 4) <> ".dwg" Then KuvaNimi = KuvaNimi & ".dwg"
        
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
    Debug.Print "  All images created successfully"
    MsgBox "Kuvat tehty", vbOKOnly, "Valmis"
  End If
End If
'oAcad.Quit

Debug.Print "TeeKuvat_Click: COMPLETED successfully"
GoTo Loppu   ' Siirrytään siivouslohkoon — resurssit vapautetaan aina

ErrorHandler:
    Debug.Print "*** ERROR in TeeKuvat_Click: " & Err.Number & " - " & Err.Description
    Debug.Print "    TitleTable: " & OtsTaulukko & ", BlockTable: " & BlockTaulukko
    Debug.Print "    Source: " & Err.Source & ", Line: " & Erl
    MsgBox "Virhe kuvien luonnissa: " & Err.Description, vbCritical
    Resume Loppu   ' Ohjataan siivouslohkoon — AutoCAD-prosessi ei jää taustalle
  
Loppu:
    ' Suojataan siivouslohko — vapautetaan kaikki resurssit riippumatta siitä, missä virhe tapahtui
    On Error Resume Next
  ' AutoCAD-istuntoa ei suljeta tässä, jotta käyttäjä voi jatkaa työskentelyä kuvien luonnin jälkeen.
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
On Error GoTo ErrorHandler
Debug.Print "LuoReferenssit: Starting reference creation"
Dim DB As Object ' DAO.Database
Dim Tiedot As Object ' DAO.Recordset
Dim sSQL As String
Dim RNro As Integer
Dim E_FileId As String
Dim E_Block As String
Dim E_Blocknumber As String

Set DB = CurrentDb
Set Tiedot = DB.OpenRecordset("SELECT * FROM Kuviin1_Ref_PAIVITA")

'Päivitetään ensin tuleeviin attribuutteihin mistä sivulta tagi tulee
Do While Not Tiedot.EOF
  sSQL = "UPDATE Kuviin1 SET Kuviin1.Teksti = '" & Tiedot("UusiTeksti") & "' "
  sSQL = sSQL & "WHERE (((FileID)='" & Tiedot("FileID") & "') "
  sSQL = sSQL & "AND ((Block)='" & Tiedot("Block") & "') "
  sSQL = sSQL & "AND ((Blocknumber)='" & Tiedot("BlockNumber") & "') "
  sSQL = sSQL & "AND ((AttrTag)='" & Tiedot("AttrTag") & "'));"
  DB.Execute sSQL
  Tiedot.MoveNext
Loop
'Lisätään sitten lähteväin attribuutteihin, mihin piiriin ja minne sivulle tieto meneesivulle
Tiedot.Close
Set Tiedot = DB.OpenRecordset("SELECT * FROM Kuviin1_Ref_LISAA")

'Päivitetään ensin tuleeviin attribuutteihin mistä sivulta tagi tulee
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

Debug.Print "LuoReferenssit: COMPLETED successfully"
Set Tiedot = Nothing
Set DB = Nothing
Exit Sub

ErrorHandler:
    Debug.Print "*** ERROR in LuoReferenssit: " & Err.Number & " - " & Err.Description
    Debug.Print "    Source: " & Err.Source & ", Line: " & Erl
    MsgBox "Virhe referenssien luonnissa: " & Err.Description, vbCritical
    Set Tiedot = Nothing
    Set DB = Nothing
End Sub
Private Function HaeIPoints(ByVal Kuva As String) As Boolean
On Error GoTo ErrorHandler
Debug.Print "HaeIPoints: Retrieving insertion points from drawing: " & Kuva
Dim Virhe As Boolean
Dim Blokki As Object '(AcadBlockReference - muutettu late binding)
Virhe = False
ReDim Paikat(1 To 1)
Dim iPaikka As Variant
Dim Attribuutit As Variant
Dim Numero As Integer
Dim i As Integer
Dim Joukko As Object '(AcadSelectionSet - muutettu late binding)
Dim FilterType(0) As Integer
Dim FilterData(0) As Variant
Dim Paikkablokki As String

Paikkablokki = Me.TPaikkaBlokki.Value
'PaikkaBlokki = "POINTBAR"

FilterType(0) = 2             'Entityn nimi filtterinä
FilterData(0) = Paikkablokki  'Blokin nimi
Set Joukko = oDoc.ActiveSelectionSet
Joukko.Clear
Joukko.Select acSelectionSetAll, , , FilterType, FilterData
Debug.Print "  Found " & Joukko.Count & " position blocks"
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
  Debug.Print "HaeIPoints: " & IIf(Virhe, "FAILED", "COMPLETED successfully")
  Exit Function

ErrorHandler:
    Debug.Print "*** ERROR in HaeIPoints: " & Err.Number & " - " & Err.Description
    Debug.Print "    Drawing: " & Kuva & ", PosBlock: " & Paikkablokki
    Debug.Print "    Source: " & Err.Source & ", Line: " & Erl
    HaeIPoints = False
End Function
Private Function VaihdaOtsikkotiedot(Taulu As DAO.Recordset, TallennusNimi As String, NextPage As String, CurrPage As String) As Boolean
    On Error GoTo ErrorHandler
    Debug.Print "VaihdaOtsikkotiedot: Updating title block attributes"
    Dim Attribuutit As Variant
    Dim Virhe As Boolean
    Dim i As Integer, j As Integer, b As Integer
    Dim Blokki As Object '(AcadBlockReference - muutettu late binding)
    Dim tagi As String
    Dim tagiValue As String
    Dim Loytyi As Boolean
    Dim FilterType(0) As Integer
    Dim FilterData(0) As Variant
    Dim Titleblokki As String
    Dim Joukko As Object '(AcadSelectionSet - muutettu late binding)
    
    Virhe = False
    
    Titleblokki = Me.TTitleBlokki.Value
    'TitleBlokki = "TITLE"

    FilterType(0) = 2             'Entityn nimi filtterinä
    FilterData(0) = Titleblokki  'Blokin nimi
    Set Joukko = oDoc.ActiveSelectionSet
    Joukko.Clear
    Joukko.Select acSelectionSetAll, , , FilterType, FilterData
    Debug.Print "  Found " & Joukko.Count & " title blocks"
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
    Debug.Print "VaihdaOtsikkotiedot: " & IIf(Virhe, "FAILED", "COMPLETED successfully")
    Exit Function

ErrorHandler:
    Debug.Print "*** ERROR in VaihdaOtsikkotiedot: " & Err.Number & " - " & Err.Description
    Debug.Print "    TitleBlock: " & Titleblokki
    Debug.Print "    Source: " & Err.Source & ", Line: " & Erl
    VaihdaOtsikkotiedot = False
End Function
Private Sub LisaaLokiin(Tieto As String)
  ' Kirjoitetaan loki-ruutuun ja Debug Windowiin
  Debug.Print Format(Now, "hh:mm:ss") & " [LisaaLokiin] " & Tieto
  If LokiPituus > 0 Then
    Me.Loki.Value = Me.Loki.Value & vbCrLf & Tieto
  Else
    Me.Loki.Value = Tieto
    LokiPituus = 1
  End If
  ' SelStart vaatii fokuksen — ohitetaan virhe 2185, jos kontrolli ei ole aktiivinen (kosmeettinen vieritys)
  On Error Resume Next
  Me.Loki.SelStart = Len(Me.Loki.Value)
  On Error GoTo 0
  Me.Repaint
End Sub

Private Sub TOtsTaulukko_BeforeUpdate(Cancel As Integer)

End Sub

Private Sub ValHakem_Click()
Dim Hakemisto As String
  Hakemisto = PoimiHakemisto(IIf(IsNull(Me.PohjaHakem.Value), "", Me.PohjaHakem.Value))
  If Hakemisto <> "" Then Me.PohjaHakem.Value = Hakemisto
End Sub
Private Sub ValHakem2_Click()
Dim Hakemisto As String
  Hakemisto = PoimiHakemisto(IIf(IsNull(Me.KuvaHakem.Value), "", Me.KuvaHakem.Value))
  If Hakemisto <> "" Then Me.KuvaHakem.Value = Hakemisto
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
       TNimi = Mid$(OpenFile.lpstrFileTitle, 1, InStr(OpenFile.lpstrFileTitle, Chr(0)) - 1)
       PoimiHakemisto = Mid$(OpenFile.lpstrFile, 1, InStr(OpenFile.lpstrFile, Chr(0)) - 1 - Len(TNimi))
    End If
End Function

Private Sub TeeEtusivu_Click()
'Tämä avaa kuvan insertoi kuvaan blokin/blokit, vaihtaa attribuuttien tiedot ja tallentaa kuvan
Dim DB As Object ' DAO.Database
Dim Kuvat As Object ' DAO.Recordset
Dim DbBlokit As Object ' DAO.Recordset
Dim FSO As New FileSystemObject
Dim Tiedosto As String
Dim Blokki As Object ' AcadBlockReference
Dim Attribuutit As Variant
Dim BNimi As String
Dim i As Integer, k As Integer, m As Integer, n As Integer
Dim sql As String
Dim OtsTaulukko As String
Dim oAcad As Object 'AutoCAD.AcadApplication
Dim Valitallennus As Variant
Dim TallennusNimi As String
Dim NextPage As String
Dim CurrPage As String
ReDim Paikat(1 To 1)

Set DB = CurrentDb
Ajastin = Timer
LokiPituus = 0
OtsTaulukko = Me.TOtsTaulukko.Value



If IsNull(Me.KuvaHakem.Value) Then
  Me.KuvaHakem.Value = Me.PohjaHakem.Value
  LisaaLokiin "-Kuva hakemistoa ei oltu määritelty, talletetaan oletuksena pohjahakemistoon.."
End If
If FSO.FolderExists(Me.PohjaHakem.Value) = False Then 'Hakemisto ei ollut laillinen annetaan virheilmoitus ja keskeytetään
  LisaaLokiin "Virhe: Hakemistoa ei löytynyt (" & Me.KuvaHakem.Value & ")"
  MsgBox "Hakemistoa: " & vbCrLf & Me.PohjaHakem.Value & vbCrLf & " ei löytynyt!" & vbCrLf & "Valitse hakemisto ja yritä uudelleen.", vbCritical, "Tee kuvat"
ElseIf FSO.FolderExists(Me.KuvaHakem.Value) = False Then 'Hakemisto ei ollut laillinen annetaan virheilmoitus ja keskeytetään
  LisaaLokiin "Virhe: Hakemistoa ei löytynyt (" & Me.KuvaHakem.Value & ")"
  MsgBox "Hakemistoa: " & vbCrLf & Me.KuvaHakem.Value & vbCrLf & " ei löytynyt!" & vbCrLf & "Valitse hakemisto ja yritä uudelleen.", vbCritical, "Tee kuvat"
Else
  If Right$(Me.PohjaHakem.Value, 1) <> "\" Then Me.PohjaHakem.Value = Me.PohjaHakem.Value & "\"
  If Right$(Me.KuvaHakem.Value, 1) <> "\" Then Me.KuvaHakem.Value = Me.KuvaHakem.Value & "\"

Set oAcad = CreateObject("AutoCAD.Application")

Tiedosto = Me.PohjaHakem.Value & Me.FrontBase.Value
If LCase$(Right$(Tiedosto, 4)) <> ".dwg" Then 'Varmistetaan että tiedostonimen pääte on .dwg
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
    Set Kuvat = DB.OpenRecordset("SELECT * FROM " & OtsTaulukko & " WHERE DocumentID=" & Me.RefID.Value)

    Do While Not DbBlokit.EOF
        If FSO.FileExists(Me.PohjaHakem.Value & "LoopDet.dwg") Then 'Blokki löytyi
            
'Insertoitava blokki
            BNimi = Me.PohjaHakem.Value & "LoopDet.dwg"
            
            If (k Mod Me.Rows.Value) = 0 Then ' jos k on jaollinen 40:llä, vaihdetaan saraketta ja alustetaan k
                m = m + 1
                k = 1
            End If
            
'JOS tulee rivimäärä täyteen, niin vaihdetaan saraketta
            
            If (k Mod Me.Rows.Value) = 0 Or k = 1 Then ' jos k on jaollinen annetulla rivimäärällä, vaihdetaan saraketta ja alustetaan k
            
'JOS sarakemäärä tulee täyteen, tallennetaan nykyinen kuva ja avataan uusi käsittelyä varten

'Tiedosto = PohjaHakem.Value & FrontBase.Value
                If LCase$(Right$(Kuvat.Fields("DWG_No"), 4)) = ".dwg" Then 'Poistetaan tiedostonimen pääte .dwg
                    TallennusNimi = Left$(Kuvat.Fields("DWG_No"), Len(Kuvat.Fields("DWG_No")) - 4)
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
                        'NextPage = "X" ' Poistettu: alla oleva If/ElseIf lohko asettaa oikean arvon
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

                If Me.Columns.Value = m Then
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
                 
                 
                 
                 Call VaihdaOtsikkotiedot(Kuvat, TallennusNimi & ".dwg", NextPage, CurrPage)
                 Valitallennus = TallennaUusiKuva(TallennusNimi, False, oAcad)
   
                
             '   Exit Do
                
                End If

'Insertoidaan Sarakkeille otsikot
            
                ReDim Preserve Paikat(1 To k) 'Varmistetaan että Paikat Arrayssa on tilaa tämän paikan tiedolle
                Paikat(k).Pisteet(0) = Me.XCoord.Value + m * 60 '60 on sarakkeen leveys, muuta myös alla olevaan vastaavaan kohtaan
                Paikat(k).Pisteet(1) = Me.YCoord.Value - k * 5 '5 = ns. riviväli, muuta myös alla olevaan vastaavaan kohtaan
                Paikat(k).Pisteet(2) = 0 'Z
                 
                Set Blokki = oDoc.ModelSpace.InsertBlock(Paikat(k).Pisteet, BNimi, 1, 1, 1, 0)
                Attribuutit = Blokki.GetAttributes 'Haetaan attribuutit
                Attribuutit(0).TextString = "LOOP"
                Attribuutit(1).TextString = "REVISION"
            
            End If

'Insertoidaan rivitietoja.
            ReDim Preserve Paikat(1 To k) 'Varmistetaan että Paikat Arrayssa on tilaa tämän paikan tiedolle
            Paikat(k).Pisteet(0) = Me.XCoord.Value + m * 60 '60 on sarakkeen leveys
            Paikat(k).Pisteet(1) = (Me.YCoord.Value - 5) - k * 5 '5 = ns. riviväli
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

' AutoCAD-istuntoa ei suljeta tässä, jotta käyttäjä voi jatkaa työskentelyä revisiosivujen luonnin jälkeen.
Loppu:
Set oAcad = Nothing
Set DB = Nothing
Set Kuvat = Nothing
Set FSO = Nothing
Set Blokki = Nothing
Set DbBlokit = Nothing
End Sub

Private Function TallennaUusiKuva(TallNimi As String, Loppu As Boolean, oAcad As Object)
        
Dim FSO As New FileSystemObject
Dim KuvaNimi As String
Dim Tiedosto As String
'Dim oAcad As Object 'AutoCAD.AcadApplication

        KuvaNimi = Me.KuvaHakem.Value & TallNimi '& "xxx" ' & ".dwg"
        
        If Right$(LCase$(KuvaNimi), 4) <> ".dwg" Then
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
        
    Tiedosto = Me.PohjaHakem.Value & Me.FrontBase.Value
    If LCase$(Right$(Tiedosto, 4)) <> ".dwg" Then 'Varmistetaan että tiedostonimen pääte on .dwg
        Tiedosto = Tiedosto & ".dwg"
    End If
    
    If FSO.FileExists(Tiedosto) Then 'Tiedosto löytyi voidaan avata se AutoCADissä
Set oDoc = Nothing
        Set oDoc = oAcad.Documents.Open(Tiedosto)   'Avataan uusi dokumentti
        oDoc.SetVariable "OSMODE", 0 'Asetetaan OSNAP tila pois päältä
    End If
End If

End Function