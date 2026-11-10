Option Compare Database
Option Explicit
' Sleep API estää busy wait -kuormituksen CPU:lla (VBA7/64-bit yhteensopiva)
#If VBA7 Then
    Private Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
#Else
    Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
#End If
'==========================================================================
' LOMAKE   : Form_Lisää Kuviin_ACAD
' SOVELLUS : FunctionDiagrams — AutoCAD-piirustusautomaatio
' KUVAUS   : Kattavin lomake funktiokaaviomoduulissa. Avaa AutoCAD-
'            piirustuksia (CreateObject/GetObject, late binding), insertoi blokit,
'            päivittää otsikkotiedot ja tallentaa valmiit
'            piirustukset. HaeTekstit_Click / HaeValitutTekstit_Click
'            poimivat attribuuttitiedot DWG-tiedostoista Blokit-tauluun.
'            TeeKuvat_Click ajaa koko tuotantoajon. PaivitaDocRev_Click
'            synkronoi Documents-kannan revision- ja detaljitiedot.
'            PoimiHakemisto käyttää OPENFILENAME-dialogia (General.bas)
'            hakemiston valintaan.
' HUOM    : Option Explicit; kaikki AutoCAD-objektit late binding (Object);
'            Lisää Lokiin-lokifunktio jokaisessa haarassa; Set = Nothing
'            kaikkiin virheenkäsittelijöihin.
' PÄIVITETTY: 2026-03-03
'==========================================================================
 ' AutoCAD SelectionSet-tyypin vakiot (late binding - ei vaadi kirjastiviitettä)
 ' Lähde: AutoCAD ActiveX/VBA Reference (AcSelect-luettelointi)
Private Const acSelectionSetAll As Integer = 5           ' Valitse kaikki entiteetit
Private Const acSelectionSetPrevious As Integer = 3     ' Edellinen valintajoukko
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
  SysCmd acSysCmdSetStatus, "käsittelee tiedostoa: " & Tiedosto & "  Aika: " & min & " min  " & sec & " sek"
End Sub
Private Sub Form_Load()
  Debug.Print "Form_Lisää Kuviin_ACAD: Lomake avattu"
  Me.Loki.Value = ""
End Sub
Private Sub Form_Unload(Cancel As Integer)
  Debug.Print "Form_Lisää Kuviin_ACAD: Lomaketta suljetaan"
  DoCmd.Hourglass False
  SysCmd acSysCmdClearStatus
End Sub
Private Sub FormClose_Click()
 Debug.Print "FormClose_Click: Lomake suljetaan"
 DoCmd.Close
End Sub
Private Sub HaeTekstit_Click()
On Error GoTo ErrorHandler
Debug.Print "HaeTekstit_Click: DWG-tiedostojen tekstien poiminta alkaa"
Dim Tied As String
Dim Polku As String
Dim Nimi As String
Dim Taulu As DAO.Recordset
Dim oEnt As Object '(acadentity - muutettu late binding)
Dim oAcad As Object '(AutoCAD.AcadApplication - muutettu late binding)
Dim DB As DAO.Database
Dim Epaonnistui As Boolean

Ajastin = Timer
Debug.Print "  Ajastin käynnistetty"

 ' Yritetään ensin liittyä jo auki olevaan AutoCAD-instanssiin, nopeampaa ja kevyempää
On Error Resume Next
Set oAcad = GetObject(, "AutoCAD.Application")
If Err.Number <> 0 Then
    ' AutoCAD ei ole auki, avataan uusi instanssi
    Err.Clear
    Set oAcad = CreateObject("AutoCAD.Application")
End If
On Error GoTo ErrorHandler
Debug.Print "  AutoCAD-yhteys käytettävissä"

Polku = Me.PohjaHakem.Value
If Right$(Polku, 1) <> "\" Then Polku = Polku & "\"
Debug.Print "  Polku: " & Polku
Tied = Dir(Polku & "*.dwg")
Debug.Print "  Tyhjennetään Blokit-taulun tämän polun tiedot: " & Polku
' Poistetaan vain kyseisen hakemiston blokit, suojaa muiden projektien data
Set DB = CurrentDb
DB.Execute "DELETE * FROM Blokit WHERE Path='" & Replace(Polku, "'", "''") & "'", dbFailOnError
Set Taulu = DB.OpenRecordset("Blokit", dbOpenDynaset)
Debug.Print "  Käsitellään DWG-tiedostoja..."
Do While Tied <> ""
  Nimi = Left$(Tied, Len(Tied) - 4)
  If UCase$(Left$(Tied, 4)) <> "BASE" And UCase$(Left$(Tied, 8)) <> "BLOCKMAP" Then
    Debug.Print "  Käsitellään tiedostoa: " & Tied
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
Debug.Print "  Kaikki tiedostot käsitelty"

Loppu:
On Error Resume Next
If Not Taulu Is Nothing Then
    Taulu.Close
    Set Taulu = Nothing
End If
Set DB = Nothing
 ' AutoCAD-istuntoa ei suljeta tässä, jotta käyttäjä voi jatkaa työskentelyä tekstien haun jälkeen.
Set oAcad = Nothing
On Error GoTo 0
Debug.Print "HaeTekstit_Click: VALMIS"
If Not Epaonnistui Then MsgBox "Valmis", vbInformation, "päivitä tekstit"
Exit Sub

ErrorHandler:
    Debug.Print "*** VIRHE aliohjelmassa HaeTekstit_Click: " & Err.Number & " - " & Err.Description
    Debug.Print "    Polku: " & Polku & ", File: " & Tied
    Debug.Print "    Lähde: " & Err.Source & ", Rivi: " & Erl
    MsgBox "Virhe tekstien poiminnassa: " & Err.Description, vbCritical
    Epaonnistui = True
    SuljeAcadDokumentti False
    Resume Loppu
End Sub

Private Sub HaeValitutTekstit_Click()
    ' Virheen käsittely ja lokitus
    On Error GoTo ErrorHandler
    Debug.Print "HaeValitutTekstit_Click: Valittujen tekstien poiminta alkaa"

    Dim Tied As String
    Dim Polku As String
    Dim Taulu As DAO.Recordset
    Dim Blokki As Object '(AcadBlockReference)
    Dim oAcad As Object '(AutoCAD.AcadApplication)
    Dim oDoc As Object '(AutoCAD.AcadDocument)
    Dim Joukko As Object '(AcadSelectionSet)
    Dim BlockArray As Variant
    Dim i As Integer, j As Integer
    Dim FilterType(0) As Integer
    Dim FilterData(0) As Variant
    Dim Hakemisto As String
    Dim Ajastin As Double
    Dim Valmis As Boolean

    ' Suodatin: haetaan vain blokkeja (INSERT)
    FilterType(0) = 0
    FilterData(0) = "INSERT"

    Ajastin = Timer
    Debug.Print "  Ajastin käynnistetty"

    ' Yhteys AutoCADiin
    Set oAcad = GetObject(, "AutoCAD.Application")
    Debug.Print "  AutoCAD-yhteys avattu"
    
    Polku = Me.PohjaHakem.Value
    If Right$(Polku, 1) <> "\" Then Polku = Polku & "\"
    Tied = Dir(Polku & "*.dwg")

    ' Avataan tietokantataulu
    Set Taulu = CurrentDb.OpenRecordset("Blokit")

    Set oDoc = oAcad.ActiveDocument
    Hakemisto = oDoc.PATH
    
    ' Hallitaan valintajoukkoa
    On Error Resume Next
    Set Joukko = oDoc.SelectionSets.Item("TempJoukko")
    If Err.Number <> 0 Then
        Set Joukko = oDoc.SelectionSets.Add("TempJoukko")
    Else
        Joukko.Clear
    End If
    On Error GoTo ErrorHandler

    ' Valitaan aiemmin valitut objektit
    Joukko.Select acSelectionSetPrevious, , , FilterType, FilterData

    If Joukko.Count = 0 Then
        MsgBox "Kuvasta ei löytynyt valittuja blokkeja, jotka täyttisivät ehdon!", vbCritical, "Tuo DATA"
        GoTo CleanUp
    End If
    
    Debug.Print "  Löytyi " & Joukko.Count & " blokkia. Käsitellään..."

    ' Käydään läpi valitut blokit
    For i = 0 To Joukko.Count - 1
        If Joukko(i).EntityName = "AcDbBlockReference" Then
            Set Blokki = Joukko(i)
            
            Debug.Print "  Käsitellään blokkia: " & Blokki.Name
            
            ' Poistetaan vanhat tiedot kyseisestä blokista (jos tarpeen)
            CurrentDb.Execute "DELETE * FROM Blokit WHERE Block='" & Replace(Blokki.Name, "'", "''") & "'", dbFailOnError

            If Blokki.HasAttributes Then
                BlockArray = Blokki.GetAttributes
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
    Valmis = True

CleanUp:
    On Error Resume Next
    If Not Taulu Is Nothing Then Taulu.Close
    If Not Joukko Is Nothing Then Joukko.Delete ' Poistetaan tilapäinen valintajoukko
    Set Joukko = Nothing
    Set Taulu = Nothing
    Set Blokki = Nothing
    Set oDoc = Nothing
    Set oAcad = Nothing
    On Error GoTo 0

    If Valmis Then
        MsgBox "Valmis! Aikaa kului " & Round(Timer - Ajastin, 2) & " sekuntia.", vbInformation, "päivitä tekstit"
        Debug.Print "HaeValitutTekstit_Click: VALMIS onnistuneesti"
    End If
    Exit Sub

ErrorHandler:
    Debug.Print "VIRHE aliohjelmassa HaeValitutTekstit_Click: " & Err.Number & " - " & Err.Description
    MsgBox "Virhe tapahtui: " & Err.Description, vbExclamation, "Virhe " & Err.Number
    Resume CleanUp
End Sub

Private Sub PaivitaDocRev_Click()
On Error GoTo ErrorHandler
Debug.Print "PaivitaDocRev_Click: Dokumenttirevision päivitys alkaa"

Dim dbsDocuments As DAO.Database
Dim dbs As DAO.Database
Dim SQLQry As String
Dim BePolku As String

If IsNull(Me.RefID.Value) Then
    MsgBox "Valitse ensin projekti (RefID puuttuu).", vbExclamation, "Päivitä dokumenttitiedot"
    Exit Sub
End If

' Back end -kannan polku luetaan SETTINGS-taulusta, oletuksena käytetään aiempaa kiinteää polkua
BePolku = Nz(DLookup("[VALUE]", "[SETTINGS]", "SETTING = 'DOCUMENTSBE'"), "W:\_DocumentDB\BackEnd\Documents_be.accdb")
If Len(Dir$(BePolku)) = 0 Then
    MsgBox "Documents-tietokantaa ei löytynyt:" & vbCrLf & BePolku, vbCritical, "Päivitä dokumenttitiedot"
    Exit Sub
End If

Set dbs = CurrentDb
' Back end avataan vain luku -tilassa, koska siitä ainoastaan luetaan
Set dbsDocuments = OpenDatabase(BePolku, False, True)

' Kysely usean BE-taulun yli; tuloksesta luodaan paikallinen taulu DocumentsDetail
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
"WHERE Documents.ProjectID=" & Me.RefID.Value
KopioiTulosTauluun dbsDocuments, dbs, SQLQry, "DocumentsDetail"

' Toinen kysely; tuloksesta luodaan paikallinen taulu DocumentsRevision
SQLQry = "SELECT " & _
"Documents.ProjectID, Revision.DocumentID, Revision.Revision, Revision.RevNote, Status.Status, " & _
"Status.Description , Revision.RevDate, Revision.RevAuthor, Revision.RevChecker, Revision.RevApprover, " & _
"Revision.RevID " & _
"FROM Status INNER JOIN (Documents RIGHT JOIN Revision ON Documents.DocumentID = Revision.DocumentID) " & _
"ON Status.StatusID = Revision.StatusID " & _
"WHERE Documents.ProjectID=" & Me.RefID.Value & ""
KopioiTulosTauluun dbsDocuments, dbs, SQLQry, "DocumentsRevision"

Debug.Print "PaivitaDocRev_Click: VALMIS onnistuneesti - DocumentsDetail- ja DocumentsRevision-taulut päivitetty"

Siivous:
    On Error Resume Next
    If Not dbsDocuments Is Nothing Then dbsDocuments.Close
    Set dbsDocuments = Nothing
    Set dbs = Nothing
    Exit Sub

ErrorHandler:
    Debug.Print "*** VIRHE aliohjelmassa PaivitaDocRev_Click: " & Err.Number & " - " & Err.Description
    Debug.Print "    Lähde: " & Err.Source
    MsgBox "Virhe dokumenttirevision päivityksessä: " & Err.Description, vbCritical
    Resume Siivous
End Sub

' Kopioi kyselyn tulokset paikalliseen tauluun. Vanha taulu korvataan, rivit lisätään
' yhdessä transaktiossa ja virhe välitetään kutsujalle.
Private Sub KopioiTulosTauluun(ByVal Lahde As DAO.Database, ByVal Kohde As DAO.Database, ByVal SQLQry As String, ByVal TaulunNimi As String)
    Dim rstLahde As DAO.Recordset
    Dim rstKohde As DAO.Recordset
    Dim Luontilause As String
    Dim Kentta As String
    Dim i As Long
    Dim Transaktio As Boolean
    Dim Virhenro As Long
    Dim Virhekuvaus As String

    On Error GoTo ErrorHandler

    Set rstLahde = Lahde.OpenRecordset(SQLQry, dbOpenSnapshot)

    If TableExists(TaulunNimi) Then
        Debug.Print "  Poistetaan aiempi taulu " & TaulunNimi & ""
        Kohde.Execute "DROP TABLE [" & TaulunNimi & "]", dbFailOnError
    End If

    ' ID-kentät tallennetaan Long-tyyppisinä, muut tekstinä
    For i = 0 To rstLahde.Fields.Count - 1
        Kentta = rstLahde.Fields(i).Name
        If i > 0 Then Luontilause = Luontilause & ", "
        If Kentta = "DocumentID" Or Kentta = "ProjectID" Then
            Luontilause = Luontilause & "[" & Kentta & "] LONG"
        Else
            Luontilause = Luontilause & "[" & Kentta & "] TEXT(250)"
        End If
    Next i

    Debug.Print "  Luodaan taulu " & TaulunNimi & ", kenttiä: " & rstLahde.Fields.Count & " fields"
    Kohde.Execute "CREATE TABLE [" & TaulunNimi & "] (" & Luontilause & ")", dbFailOnError

    Set rstKohde = Kohde.OpenRecordset(TaulunNimi, dbOpenDynaset)
    DBEngine.BeginTrans
    Transaktio = True
    Do While Not rstLahde.EOF
        rstKohde.AddNew
        For i = 0 To rstLahde.Fields.Count - 1
            ' Null-arvot jätetään tyhjiksi; CStr(Null) aiheuttaisi virheen
            If Not IsNull(rstLahde.Fields(i).Value) Then
                If rstKohde.Fields(i).Type = dbLong Then
                    rstKohde.Fields(i).Value = rstLahde.Fields(i).Value
                Else
                    rstKohde.Fields(i).Value = CStr(rstLahde.Fields(i).Value)
                End If
            End If
        Next i
        rstKohde.Update
        rstLahde.MoveNext
    Loop
    DBEngine.CommitTrans
    Transaktio = False

Siivous:
    On Error Resume Next
    If Not rstKohde Is Nothing Then rstKohde.Close
    If Not rstLahde Is Nothing Then rstLahde.Close
    Set rstKohde = Nothing
    Set rstLahde = Nothing
    On Error GoTo 0
    If Virhenro <> 0 Then Err.Raise Virhenro, "KopioiTulosTauluun", Virhekuvaus
    Exit Sub

ErrorHandler:
    Virhenro = Err.Number
    Virhekuvaus = Err.Description
    If Transaktio Then
        PeruTransaktio
        Transaktio = False
    End If
    Resume Siivous
End Sub

Private Sub PeruTransaktio()
    On Error Resume Next
    DBEngine.Rollback
End Sub

Public Function TableExists(sTable As String) As Boolean
    
    Dim DB As DAO.Database
    Dim tbl As DAO.TableDef
    Set DB = CurrentDb()
    
    TableExists = False
    
    For Each tbl In DB.TableDefs
        If tbl.Name = sTable Then
            TableExists = True
            Exit For
        End If
    Next tbl
    
End Function

Private Sub TeeKuvat_Click()
'Tämä avaa kuvan insertoi kuvaan blokin/blokit, vaihtaa attribuuttien tiedot ja tallentaa kuvan
On Error GoTo ErrorHandler
Debug.Print "TeeKuvat_Click: Kuvien luonti alkaa (kriittinen toiminto)"
Dim DB As Object ' DAO.Database
Dim Kuvat As Object ' DAO.Recordset
Dim DbBlokit As Object ' DAO.Recordset
Dim Tiedot As Object ' DAO.Recordset
Dim FSO As Object
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
Set FSO = CreateObject("Scripting.FileSystemObject")
Ajastin = Timer
LokiPituus = 0
OtsTaulukko = Me.TOtsTaulukko.Value
Debug.Print "  Title: " & OtsTaulukko

If Me.Translate.Value = True Then
    Debug.Print "  Käännös käytössä, kieli: " & Me.TLang.Value
    OtsTaulukko = OtsTaulukko & "_" & Me.TLang.Value
End If

BlockTaulukko = Me.TBlockTaulukko.Value
BlockTaulukkoTyhjennys = Me.TBlockTaulukko.Value

DoCmd.SetWarnings False 'MRu lisätty 26.04.2011
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
      LisaaLokiin "- Ajettiin kysely '" & Kysely & "'"
      Exit Do
    Else
      Kysely = Mid$(Me.TKyselyt.Value, OS, OS2 - OS)
      DB.Execute Kysely
      LisaaLokiin "- Ajettiin kysely '" & Kysely & "'"
      OS = OS2 + 2
    End If
  Loop
  If Me.CRef = True Then
    Debug.Print "  Luodaan taulu references"
    LuoReferenssit
  End If
  LisaaLokiin "- Päivitettiin " & BlockTaulukkoTyhjennys & "-taulukko"

ElseIf Vastaus = vbCancel Then
  Debug.Print "  Käyttäjä peruutti toiminnon"
  Set DB = Nothing
  Exit Sub
End If

If IsNull(Me.KuvaHakem.Value) Then
  Me.KuvaHakem.Value = Me.PohjaHakem.Value
  LisaaLokiin "-Kuva hakemistoa ei oltu määritelty, talletetaan oletuksena pohjahakemistoon.."
End If

If FSO.FolderExists(Me.PohjaHakem.Value) = False Then 'Hakemisto ei ollut laillinen annetaan virheilmoitus ja keskeytetään
  LisaaLokiin "Virhe: Hakemistoa ei löytynyt (" & Me.PohjaHakem.Value & ")"
  MsgBox "Hakemistoa: " & vbCrLf & Me.PohjaHakem.Value & vbCrLf & " ei löytynyt!" & vbCrLf & "Valitse hakemisto ja yritä uudelleen.", vbCritical, "Tee kuvat"

ElseIf FSO.FolderExists(Me.KuvaHakem.Value) = False Then 'Hakemisto ei ollut laillinen annetaan virheilmoitus ja keskeytetään
  LisaaLokiin "Virhe: Hakemistoa ei löytynyt (" & Me.KuvaHakem.Value & ")"
  MsgBox "Hakemistoa: " & vbCrLf & Me.KuvaHakem.Value & vbCrLf & " ei löytynyt!" & vbCrLf & "Valitse hakemisto ja yritä uudelleen.", vbCritical, "Tee kuvat"

Else
  
  Debug.Print "  Hakemistot tarkistettu: pohja=" & Me.PohjaHakem.Value & "; tuloste=" & Me.KuvaHakem.Value
  If Right$(Me.PohjaHakem.Value, 1) <> "\" Then Me.PohjaHakem.Value = Me.PohjaHakem.Value & "\"
  
  If Right$(Me.KuvaHakem.Value, 1) <> "\" Then Me.KuvaHakem.Value = Me.KuvaHakem.Value & "\"
  Set Kuvat = DB.OpenRecordset("SELECT * FROM " & OtsTaulukko)
  
  If Kuvat.EOF Then 'Taulukossa ei ollut tarvittavia tietoja, annetaan virhe ja keskeytetään
    LisaaLokiin "Virhe: Taulukosta (" & OtsTaulukko & ") ei löytynyt kuvia"
    MsgBox "Taulukosta/kyselystä """ & OtsTaulukko & """ ei löytynyt tietoja!" & vbCrLf & "Valitse taulukko ja yritä uudelleen.", vbCritical, "Tee kuvat"
  
  Else 'OK hakemisto löytyi ja taulukossa oli tietoja, ruvetaan käymään taulukon kuvia läpi
       
       'Otetaan yhteys AutoCADiin
  On Error Resume Next
  Set oAcad = GetObject(, "AutoCAD.Application") 'Koitetaan yhdistää AutoCADiin
  If Err <> 0 Then 'Käynnissä olevaa AutoCADiä ei löytynyt
    'MsgBox "Käynnissä olevaa AutoCADiä ei löytynyt!", vbCritical, "Virhe!"
    Set oAcad = CreateObject("AutoCAD.Application")
    'Exit Sub
  End If
  On Error GoTo ErrorHandler
    
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
          If VaihdaOtsikkotiedot(Kuvat, "", "", "") Then 'Jos Otsikkotiedot saatiin vaihdettua OK niin kannattaa mennään eteenpäin
            
            'Valitaan taulukosta kuvaan insertoitavat blokit
            sql = "SELECT Block, Blocknumber FROM " & BlockTaulukko & " WHERE FileID='" & Kuvat.Fields("FileID").Value & "' GROUP BY Block, Blocknumber HAVING (Blocknumber Is Not Null);"
            Set DbBlokit = DB.OpenRecordset(sql)
            
            Do While Not DbBlokit.EOF
              
              If FSO.FileExists(Me.PohjaHakem.Value & DbBlokit.Fields("Block").Value & ".dwg") Then 'Blokki löytyi
                sql = "SELECT * FROM " & BlockTaulukko & " WHERE FileID='" & Kuvat.Fields("FileID").Value & "' AND Block='" & DbBlokit.Fields("Block").Value & "'  AND Blocknumber='" & DbBlokit.Fields("Blocknumber").Value & "' ORDER BY AttrTag;"
                Set Tiedot = DB.OpenRecordset(sql)
                Paikka = CInt(DbBlokit.Fields("Blocknumber").Value)
                If Paikka < 1 Or Paikka > UBound(Paikat) Then
                  LisaaLokiin "Virhe: Blokin " & DbBlokit.Fields("Block").Value & " paikkanumeroa " & Paikka & " ei löytynyt paikkablokista (Kuva: " & Kuvat.Fields("FileID").Value & ")"
                  GoTo SeuraavaBlokki
                End If
                If Me.Translate.Value = True Then 'Tehdään eri kieliset kuvat
                    BNimi = Me.PohjaHakem.Value & DbBlokit.Fields("Block").Value & "_" & Me.TLang.Value & ".dwg" 'lisätään blokin nimeen kielen tunnus, esim. _RU
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
                  Tiedot.FindFirst "AttrTag = '" & Replace(Attribuutit(i).TagString, "'", "''") & "'" 'Etsitään tietokannasta attribuutin tagi
                  
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
              
SeuraavaBlokki:
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
        oDoc.PurgeAll

        oDoc.SetVariable "OSMODE", OSMODE 'Palautetaan OSNAP tila siihen joka se oli ennen sen poistamista pois pÃƒÂ¤ÃƒÂ¤ltÃƒÂ¤
        oDoc.SaveAs KuvaNimi 'Talletetaan kuva
        oDoc.Close
      Else 'Tiedostoa ei löytynyt
        LisaaLokiin "Virhe: määriteltyä tiedostoa ei löytynyt (" & Tiedosto & ")"
      End If
      Kuvat.MoveNext
    Loop
    Debug.Print "  Kaikki kuvat luotu onnistuneesti"
    MsgBox "Kuvat tehty", vbOKOnly, "Valmis"
  End If
End If
'oAcad.Quit

Debug.Print "TeeKuvat_Click: VALMIS onnistuneesti"
GoTo Loppu   ' Siirrytään siivouslohkoon — resurssit vapautetaan aina

ErrorHandler:
    Debug.Print "*** VIRHE aliohjelmassa TeeKuvat_Click: " & Err.Number & " - " & Err.Description
    Debug.Print "    Otsikkotaulu: " & OtsTaulukko & ", Blokkitaulu: " & BlockTaulukko
    Debug.Print "    Lähde: " & Err.Source & ", Rivi: " & Erl
    MsgBox "Virhe kuvien luonnissa: " & Err.Description, vbCritical
    SuljeAcadDokumentti True
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
' Muotoilee arvon SQL-merkkijonovakioon: Null muuttuu tyhjäksi ja heittomerkit kahdennetaan.
Private Function SqlTeksti(ByVal Arvo As Variant) As String
    If IsNull(Arvo) Then
        SqlTeksti = ""
    Else
        SqlTeksti = Replace(CStr(Arvo), "'", "''")
    End If
End Function

Private Sub LuoReferenssit()
'Luodaan referenssilinkit 2.3.2015 - VG
On Error GoTo ErrorHandler
Debug.Print "LuoReferenssit: Viittausten luonti alkaa"
Dim DB As Object ' DAO.Database
Dim Tiedot As Object ' DAO.Recordset
Dim sSQL As String
Dim RNro As Integer
Dim E_FileId As String
Dim E_Block As String
Dim E_Blocknumber As String
Dim Transaktio As Boolean
Dim Virhenro As Long
Dim Virhekuvaus As String

Set DB = CurrentDb
' Kaikki päivitykset samassa transaktiossa, jotta virhe ei jätä taulua puolivalmiiksi
DBEngine.BeginTrans
Transaktio = True
Set Tiedot = DB.OpenRecordset("SELECT * FROM Kuviin1_Ref_PAIVITA")

'Päivitetään ensin tuleviin attribuutteihin, mistä sivulta tagi tulee
Do While Not Tiedot.EOF
  sSQL = "UPDATE Kuviin1 SET Kuviin1.Teksti = '" & SqlTeksti(Tiedot("UusiTeksti")) & "' "
  sSQL = sSQL & "WHERE (((FileID)='" & SqlTeksti(Tiedot("FileID")) & "') "
  sSQL = sSQL & "AND ((Block)='" & SqlTeksti(Tiedot("Block")) & "') "
  sSQL = sSQL & "AND ((Blocknumber)='" & SqlTeksti(Tiedot("BlockNumber")) & "') "
  sSQL = sSQL & "AND ((AttrTag)='" & SqlTeksti(Tiedot("AttrTag")) & "'));"
  DB.Execute sSQL, dbFailOnError
  Tiedot.MoveNext
Loop
Tiedot.Close
Set Tiedot = DB.OpenRecordset("SELECT * FROM Kuviin1_Ref_LISAA")

'Lisätään sitten lähteviin attribuutteihin tieto siitä, mihin piiriin ja sivulle tieto menee
Do While Not Tiedot.EOF
 If E_FileId = Nz(Tiedot("FileId"), "") And E_Block = Nz(Tiedot("Block"), "") And E_Blocknumber = Nz(Tiedot("Blocknumber"), "") Then
   RNro = RNro + 1
 Else
   RNro = 1
 End If
 sSQL = "INSERT INTO Kuviin1 ( FileID, Block, Blocknumber, AttrTag, Teksti ) "
 sSQL = sSQL & "Values ('" & SqlTeksti(Tiedot("FileID")) & "', '" & SqlTeksti(Tiedot("Block")) & "','" & SqlTeksti(Tiedot("Blocknumber")) & "','FTO" & RNro & "','" & SqlTeksti(Tiedot("Teksti")) & "')"

 DB.Execute sSQL, dbFailOnError

 E_FileId = Nz(Tiedot("FileId"), "")
 E_Block = Nz(Tiedot("Block"), "")
 E_Blocknumber = Nz(Tiedot("Blocknumber"), "")

 Tiedot.MoveNext
Loop

Tiedot.Close
DBEngine.CommitTrans
Transaktio = False
Debug.Print "LuoReferenssit: VALMIS onnistuneesti"

Siivous:
    On Error Resume Next
    If Not Tiedot Is Nothing Then Tiedot.Close
    Set Tiedot = Nothing
    Set DB = Nothing
    On Error GoTo 0
    ' Virhe välitetään kutsujalle, jottei kuvia luoda puutteellisilla viittauksilla
    If Virhenro <> 0 Then Err.Raise Virhenro, "LuoReferenssit", Virhekuvaus
    Exit Sub

ErrorHandler:
    Virhenro = Err.Number
    Virhekuvaus = Err.Description
    Debug.Print "*** VIRHE aliohjelmassa LuoReferenssit: " & Virhenro & " - " & Virhekuvaus
    If Transaktio Then
        DBEngine.Rollback
        Transaktio = False
    End If
    Resume Siivous
End Sub
Private Function HaeIPoints(ByVal Kuva As String) As Boolean
On Error GoTo ErrorHandler
Debug.Print "HaeIPoints: Haetaan sijoituspisteet piirustuksesta: " & Kuva
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

FilterType(0) = 2             'Entityn nimi filtterinÃƒÂ¤
FilterData(0) = Paikkablokki  'Blokin nimi
Set Joukko = oDoc.ActiveSelectionSet
Joukko.Clear
Joukko.Select acSelectionSetAll, , , FilterType, FilterData
Debug.Print "  Löytyi " & Joukko.Count & " paikkablokkia"
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
        Paikat(Numero).Pisteet(0) = iPaikka(0)   'X:  Laitetaan paikan kordinaatit Paikat Arrayhyn kohtaan jonka attribuutin numero mÃƒÂ¤ÃƒÂ¤ritteli
        Paikat(Numero).Pisteet(1) = iPaikka(1)   'Y
        Paikat(Numero).Pisteet(2) = iPaikka(2)   'Z
      Next i
      Debug.Print "  Poimittiin " & UBound(Paikat) & " sijoituspistettä"
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
  Debug.Print "HaeIPoints: " & IIf(Virhe, "EPÄONNISTUI", "VALMIS onnistuneesti")
  Exit Function

ErrorHandler:
    Debug.Print "*** VIRHE aliohjelmassa HaeIPoints: " & Err.Number & " - " & Err.Description
    Debug.Print "    Piirustus: " & Kuva & ", Paikkablokki: " & Paikkablokki
    Debug.Print "    Lähde: " & Err.Source & ", Rivi: " & Erl
    HaeIPoints = False
End Function
Private Function VaihdaOtsikkotiedot(Taulu As DAO.Recordset, TallennusNimi As String, NextPage As String, CurrPage As String) As Boolean
    On Error GoTo ErrorHandler
    Debug.Print "VaihdaOtsikkotiedot: Päivitetään otsikkoblokin attribuutteja"
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

    FilterType(0) = 2             'Kohteen nimi suodattimena
    FilterData(0) = Titleblokki  'Blokin nimi
    Set Joukko = oDoc.ActiveSelectionSet
    Joukko.Clear
    Joukko.Select acSelectionSetAll, , , FilterType, FilterData
    Debug.Print "  Löytyi " & Joukko.Count & " otsikkoblokkia"
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
                ' Jos attribuutin teksti on kentän nimi, arvo haetaan kyseisestä kentästä
                If Taulu.Fields(j).Name = tagiValue Then
                  Attribuutit(i).TextString = IIf(IsNull(Taulu.Fields(tagiValue).Value), "", Taulu.Fields(tagiValue).Value)
                  Loytyi = True
                ' Muuten arvo haetaan attribuutin tunnisteen (tag) mukaisesta kentästä
                ElseIf Taulu.Fields(j).Name = tagi Then
                  Attribuutit(i).TextString = IIf(IsNull(Taulu.Fields(tagi).Value), "", Taulu.Fields(tagi).Value)
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
          LisaaLokiin "Virhe: Kuvasta " & Taulu.Fields("BaseDWG").Value & " ei löytynyt otsikkotietoja sisältävää blokkia (" & Titleblokki & ")"
          Virhe = True
        End If
      Next b
    End If
    Set Blokki = Nothing
    VaihdaOtsikkotiedot = Not Virhe
    Debug.Print "VaihdaOtsikkotiedot: " & IIf(Virhe, "EPÄONNISTUI", "VALMIS onnistuneesti")
    Exit Function

ErrorHandler:
    Debug.Print "*** VIRHE aliohjelmassa VaihdaOtsikkotiedot: " & Err.Number & " - " & Err.Description
    Debug.Print "    Otsikkoblokki: " & Titleblokki
    Debug.Print "    Lähde: " & Err.Source & ", Rivi: " & Erl
    VaihdaOtsikkotiedot = False
End Function
' Sulkee virhetilanteessa auki jääneen piirustuksen tallentamatta ja palauttaa OSNAP-tilan.
Private Sub SuljeAcadDokumentti(ByVal PalautaOsnap As Boolean)
    On Error Resume Next
    If Not oDoc Is Nothing Then
        If PalautaOsnap Then oDoc.SetVariable "OSMODE", OSMODE
        oDoc.Close False
    End If
    Set oDoc = Nothing
End Sub

Private Sub LisaaLokiin(Tieto As String)
  ' Kirjoitetaan loki-ruutuun ja Debug Windowiin
  Debug.Print Format(Now, "hh:mm:ss") & " [LisaaLokiin] " & Tieto
  If LokiPituus > 0 Then
    Me.Loki.Value = Me.Loki.Value & vbCrLf & Tieto
  Else
    Me.Loki.Value = Tieto
    LokiPituus = 1
  End If
  ' SelStart vaatii fokuksen ohitetaan virhe 2185, jos kontrolli ei ole aktiivinen (kosmeettinen vieritys)
  On Error Resume Next
  Me.Loki.SelStart = Len(Me.Loki.Value)
  On Error GoTo 0
  Me.Repaint
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
      .flags = &H8   ' OFN_NOCHANGEDIR: valinta ei muuta nykyistä työhakemistoa
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
Dim FSO As Object
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
Dim SivuTunnukset As Variant
ReDim Paikat(1 To 1)

On Error GoTo ErrorHandler
Set DB = CurrentDb
Set FSO = CreateObject("Scripting.FileSystemObject")
SivuTunnukset = Array("VI", "VII", "VIII", "IX", "X", "XI", "XII", "XIII", "XIV", "XV", "XVI", "XVII", "XVIII", "XIX", "XX")
Ajastin = Timer
LokiPituus = 0
OtsTaulukko = Me.TOtsTaulukko.Value



If IsNull(Me.KuvaHakem.Value) Then
  Me.KuvaHakem.Value = Me.PohjaHakem.Value
  LisaaLokiin "-Kuva hakemistoa ei oltu määritelty, talletetaan oletuksena pohjahakemistoon.."
End If
If FSO.FolderExists(Me.PohjaHakem.Value) = False Then 'Hakemisto ei ollut laillinen annetaan virheilmoitus ja keskeytetään
  LisaaLokiin "Virhe: Hakemistoa ei löytynyt (" & Me.PohjaHakem.Value & ")"
  MsgBox "Hakemistoa: " & vbCrLf & Me.PohjaHakem.Value & vbCrLf & " ei löytynyt!" & vbCrLf & "Valitse hakemisto ja yritä uudelleen.", vbCritical, "Tee kuvat"
ElseIf FSO.FolderExists(Me.KuvaHakem.Value) = False Then 'Hakemisto ei ollut laillinen annetaan virheilmoitus ja keskeytetään
  LisaaLokiin "Virhe: Hakemistoa ei löytynyt (" & Me.KuvaHakem.Value & ")"
  MsgBox "Hakemistoa: " & vbCrLf & Me.KuvaHakem.Value & vbCrLf & " ei löytynyt!" & vbCrLf & "Valitse hakemisto ja yritä uudelleen.", vbCritical, "Tee kuvat"
Else
  If Right$(Me.PohjaHakem.Value, 1) <> "\" Then Me.PohjaHakem.Value = Me.PohjaHakem.Value & "\"
  If Right$(Me.KuvaHakem.Value, 1) <> "\" Then Me.KuvaHakem.Value = Me.KuvaHakem.Value & "\"

' Käytetään jo käynnissä olevaa AutoCADia, jos sellainen on
On Error Resume Next
Set oAcad = GetObject(, "AutoCAD.Application")
If Err.Number <> 0 Then
    Err.Clear
    Set oAcad = CreateObject("AutoCAD.Application")
End If
On Error GoTo ErrorHandler

Tiedosto = Me.PohjaHakem.Value & Me.FrontBase.Value
If LCase$(Right$(Tiedosto, 4)) <> ".dwg" Then 'Varmistetaan että tiedostonimen pääte on .dwg
    Tiedosto = Tiedosto & ".dwg"
End If
    
If FSO.FileExists(Tiedosto) Then 'Tiedosto löytyi voidaan avata se AutoCADissä
    Set oDoc = oAcad.Documents.Open(Tiedosto) 'Avataan uusi dokumentti
    OSMODE = oDoc.GetVariable("OSMODE")
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
            
 ' Kun rivimäärä täyttyy, siirrytään seuraavaan sarakkeeseen
            
            If (k Mod Me.Rows.Value) = 0 Or k = 1 Then ' Rivinumerointi alkaa uudelleen uuden sarakkeen alussa
            
' Tallennetaan sivu, kun sarakkeet ovat täynnä.

                If LCase$(Right$(Kuvat.Fields("DWG_No"), 4)) = ".dwg" Then 'Poistetaan tiedostonimen pääte .dwg
                    TallennusNimi = Left$(Kuvat.Fields("DWG_No"), Len(Kuvat.Fields("DWG_No")) - 4)
                Else
                    TallennusNimi = Kuvat.Fields("DWG_NO")
                End If

                If Me.Columns.Value = m Then
                    n = n + 1
                    m = 0
                 
                    ' Revisiosivun tunnus: VI, VII, VIII jne.
                    If n <= UBound(SivuTunnukset) - 1 Then
                        TallennusNimi = TallennusNimi & "_" & SivuTunnukset(n)
                        CurrPage = SivuTunnukset(n)
                        NextPage = SivuTunnukset(n + 1)
                    Else
                        LisaaLokiin "Virhe: revisiosivuja on liikaa (" & n & "), sivun tunnusta ei voitu muodostaa"
                    End If
                 
                 
                 
                 Call VaihdaOtsikkotiedot(Kuvat, TallennusNimi & ".dwg", NextPage, CurrPage)
                 Valitallennus = TallennaUusiKuva(TallennusNimi, False, oAcad)
   
                
                
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
            LisaaLokiin "Virhe: Blokkia " & BNimi & " ei löytynyt"
        End If
        
        k = k + 1
        
    DbBlokit.MoveNext
    Loop

    ' Viimeinen sivu tallennetaan vasta silmukan jälkeen
    If VaihdaOtsikkotiedot(Kuvat, TallennusNimi & ".dwg", NextPage, CurrPage) Then
        Valitallennus = TallennaUusiKuva(TallennusNimi, True, oAcad)
    End If
    MsgBox "Revisiosivu(t) tehty", vbOKOnly, "Valmis"
    Else 'Tiedostoa ei löytynyt
        LisaaLokiin "Virhe: Määritelty tiedostoa ei löytynyt (" & Tiedosto & ")"
    End If

End If

' AutoCAD-istuntoa ei suljeta tässä, jotta käyttäjä voi jatkaa työskentelyä revisiosivujen luonnin jälkeen.
Loppu:
Set oAcad = Nothing
Set DB = Nothing
Set Kuvat = Nothing
Set FSO = Nothing
Set Blokki = Nothing
Set DbBlokit = Nothing
Exit Sub

ErrorHandler:
    Debug.Print "*** VIRHE aliohjelmassa TeeEtusivu_Click: " & Err.Number & " - " & Err.Description
    MsgBox "Virhe revisiosivujen luonnissa: " & Err.Description, vbCritical
    SuljeAcadDokumentti True
    Resume Loppu
End Sub

Private Function TallennaUusiKuva(TallNimi As String, Loppu As Boolean, oAcad As Object)
        
Dim FSO As Object
Dim KuvaNimi As String
Dim Tiedosto As String

Set FSO = CreateObject("Scripting.FileSystemObject")

        KuvaNimi = Me.KuvaHakem.Value & TallNimi
        
        If Right$(LCase$(KuvaNimi), 4) <> ".dwg" Then
        KuvaNimi = KuvaNimi & ".dwg"
        End If
        
        If FSO.FileExists(KuvaNimi) Then
          FSO.DeleteFile KuvaNimi
        End If
        
        oDoc.SetVariable "OSMODE", OSMODE 'Palautetaan OSNAP-tila aiempaan arvoonsa
        oDoc.SaveAs KuvaNimi 'Talletetaan kuva
        oDoc.Close
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