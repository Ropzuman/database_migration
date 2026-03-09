Option Compare Database   ' Käytä tietokantajärjestystä merkkijonovertailuissa
Option Explicit           ' Muuttujaesittely pakollinen
'==============================================================================
' Moduuli: DataToACAD
' Tarkoitus: Luo AutoCAD LISP-tiedostot tietokannassa olevista piirikaaviotiedoista
' Alkuperäinen: 1997-02-21 Fr 11:10 /tw
' Muokattu: 1997-03-19 We 15:46 /tw
' Muokattu: 1997-03-21 Fr 16:07 /tw
' Muokattu: 1997-07-14 Mo 14:29 /tw
' Päivitetty: 2025-11-11 — DAO-tyypit, virheenkäsittely, kommentit lisätty
' Päivitetty: 2026-03-09 — CrsRefLink muutettu välimuisti-hakuun (N+1-pullonkaula poistettu)
'==============================================================================

' --- Moduulitason välimuisti CrsRefLisps-haulle ---
' Ladataan taulukko kerran Dictionary-objektiin; jokainen haku on O(1) levyluvun sijaan
Private dictCrsRef As Object   ' Scripting.Dictionary
Private blnCrsRefLoaded As Boolean

'------------------------------------------------------------------------------
' Alirutiini: LoadCrsRefCache
' Tarkoitus: Lataa CrsRefLisps-taulukon kerran muistiin Dictionary-objektiin.
'            Kutsutaan laiskasti CrsRefLink-funktiosta ensimmäisellä hakukerralla.
'------------------------------------------------------------------------------
Private Sub LoadCrsRefCache()
    Dim tble As DAO.Recordset

    Set dictCrsRef = CreateObject("Scripting.Dictionary")
    ' dbOpenForwardOnly on nopein vaihtoehto pelkkään lukemiseen
    Set tble = CurrentDb.OpenRecordset("CrsRefLisps", dbOpenForwardOnly)

    Do Until tble.EOF
        ' Lisätään vain ensimmäinen osuma, jos duplikaatteja on
        If Not dictCrsRef.Exists(CStr(tble!CrsRefID)) Then
            dictCrsRef.Add CStr(tble!CrsRefID), CStr(tble!Lisp)
        End If
        tble.MoveNext
    Loop

    tble.Close
    Set tble = Nothing
    blnCrsRefLoaded = True
End Sub

'------------------------------------------------------------------------------
' Funktio: CrsRefLink
' Tarkoitus: Hakee LISP-koodin ristiviitetaulukosta välimuistin avulla.
'            Ensimmäisellä kutsulla lataa koko taulukon muistiin; sen jälkeen
'            jokainen haku on O(1)-nopea eikä avaa Recordsetia lainkaan.
' Parametrit:
'   tblnimi - Taulukonnimi (vain "CRSREF" käynnistää haun)
'   teksti  - Haettava ristiviite-ID
' Palautusarvo: LISP-koodi tai alkuperäinen teksti, jos ei löydy
'------------------------------------------------------------------------------
Function CrsRefLink(tblnimi As String, teksti As String) As String
On Error GoTo ErrorHandler

If tblnimi = "CRSREF" Then
    ' Ladataan välimuistiin vain ensimmäisellä kerralla
    If Not blnCrsRefLoaded Then LoadCrsRefCache

    ' Nopea O(1)-haku Dictionary-objektista levyn sijaan
    If dictCrsRef.Exists(teksti) Then
        CrsRefLink = dictCrsRef(teksti)
    Else
        CrsRefLink = teksti  ' Ei löydetty — palautetaan alkuperäinen teksti
    End If
Else
    ' Ei ristiviite — palautetaan alkuperäinen teksti
    CrsRefLink = teksti
End If

Exit Function

ErrorHandler:
    MsgBox "Error in CrsRefLink: " & Err.Description, vbCritical, "Cross-Reference Lookup Error"
    CrsRefLink = teksti  ' Virhetilanteessa palautetaan alkuperäinen teksti
End Function

'------------------------------------------------------------------------------
' Funktio: get_filename
' Tarkoitus: Poimii 8-merkkisen tiedostonimen taulukonnimestä
' Parametrit:
'   taulnimi - Taulukonnimi (voi sisältää tähti-erottimen)
' Palautusarvo: 8-merkkinen isolla kirjoitettu tiedostonimi
'------------------------------------------------------------------------------
Function get_filename(taulnimi As String) As String
On Error GoTo ErrorHandler

Dim ast As Integer

ast = InStr(taulnimi, "*")
If ast = 0 Then
    ' Ei tähteä — otetaan 8 ensimmäistä merkkiä
    get_filename = UCase$(Mid$(taulnimi, 1, 8))
Else
    ' Tähti löytyi — otetaan 8 ensimmäistä merkkiä ennen sitä
    get_filename = UCase$(Mid$(Mid$(taulnimi, 1, ast - 1), 1, 8))
End If

Exit Function

ErrorHandler:
    MsgBox "Error in get_filename: " & Err.Description, vbCritical, "Filename Extraction Error"
    get_filename = "ERROR"
End Function

'------------------------------------------------------------------------------
' Funktio: inch
' Tarkoitus: Korvaa lainausmerkit LISP-syntaksin \042-sekvenssillä
' Parametrit:
'   a - Merkkijono, joka sisältää lainausmerkkejä
' Palautusarvo: Merkkijono, jossa lainausmerkit korvattu \042-sekvenssillä
'------------------------------------------------------------------------------
Function inch(a As String) As String
On Error GoTo ErrorHandler

Dim L As String     ' Lainausmerkki
Dim E As String     ' Käsiteltävä merkkijono
Dim b As Integer    ' Lainausmerkin sijainti
Dim c As String     ' Merkkijono ennen lainausmerkkiä
Dim D As String     ' Merkkijono lainausmerkin jälkeen

L = Chr(34)  ' Lainausmerkki
E = a

Do
    b = InStr(1, E, L)
    If b = 0 Then
        ' Ei lisää lainausmerkkejä — palautetaan tulos
        inch = E
        Exit Function
    End If
    ' Jaetaan merkkijono lainausmerkin kohdasta
    c = Mid$(E, 1, b - 1)
    D = Mid$(E, b + 1, Len(a))
    ' Korvataan lainausmerkki LISP-erikoissekvenssillä
    E = c & "\042" & D
Loop

Exit Function

ErrorHandler:
    MsgBox "Error in inch: " & Err.Description, vbCritical, "LISP Quote Escaping Error"
    inch = a  ' Virhetilanteessa palautetaan alkuperäinen merkkijono
End Function

'------------------------------------------------------------------------------
' Funktio: makeFiles
' Tarkoitus: Pääorkestraattori — luo kaikki AutoCAD LISP-tiedostot tietokannasta
' Parametrit:
'   common - Konfiguraatiotaulun nimi
' Toiminta:
'   1. Lukee asetukset konfiguraatiotaulusta
'   2. Alustaa .txt-tulostiedostot
'   3. Luo silmukoimattomat listat
'   4. Luo silmukoihin perustuvat listat (jos käytössä)
'   5. Sulkee tiedostot asianmukaisesti
'------------------------------------------------------------------------------
Function makeFiles(common As String) As Integer
On Error GoTo ErrorHandler

Dim DB As DAO.Database
Dim cmmn As DAO.Recordset   ' Konfiguraatiotietue
Dim tbl As DAO.Recordset    ' Datatietue
Dim L As String             ' Lainausmerkki
Dim suod As Variant         ' Suodatusarvo
Dim direc As String         ' Tulostushakemiston polku

Set DB = CurrentDb
Set cmmn = DB.OpenRecordset(common, dbOpenDynaset)

L = Chr(34)  ' Lainausmerkki LISP-listoja varten

cmmn.MoveFirst
suod = cmmn.Fields("Filter")
direc = cmmn!AcadDirectory  ' Hakemisto, jonne LISP-tiedostot tallennetaan

' Jos vain skriptitiedosto — ohitetaan LISP-tiedostojen luonti
If cmmn!OnlyScript Then
    GoTo scrtest
End If

'--- Alustetaan kaikki tulostiedostot avaavalla sulkumerkillä ---
cmmn.MoveFirst
Do Until cmmn.EOF
    ' Alustetaan silmukoimattomat tiedostot
    If Not IsNull(cmmn!TablesOrQueriesNoLoop.Value) Then
        Open direc & get_filename(cmmn!TablesOrQueriesNoLoop.Value) & ".txt" For Output As #1
        Print #1, "("  ' LISP-listan avaava sulku
        Close #1
    End If
    ' Alustetaan silmukkaan perustuvat tiedostot
    If Not IsNull(cmmn!TablesOrQueries.Value) Then
        Open direc & get_filename(cmmn!TablesOrQueries.Value) & ".txt" For Output As #1
        Print #1, "("  ' LISP-listan avaava sulku
        Close #1
    End If
    cmmn.MoveNext
Loop

cmmn.MoveFirst

'--- Luodaan silmukoimattomat LISP-listat ---
' Yksinkertaiset listat ilman silmukka-ID-suodatusta
Do Until cmmn.EOF
    If Not IsNull(cmmn!TablesOrQueriesNoLoop.Value) Then
        MakeListNoLoopID cmmn!TablesOrQueriesNoLoop.Value, direc
    End If
    cmmn.MoveNext
Loop

cmmn.MoveFirst
' Jos ei silmukkatauluja — hypätään skriptinluontiin
If cmmn!NoLoopIDTables Then
    GoTo scrtest
End If

'--- Luodaan silmukkaan perustuvat LISP-listat ---
' Nämä listat suodatetaan silmukka-ID-sarakkeen mukaan
cmmn.MoveFirst
Do Until cmmn.EOF
    If Not IsNull(cmmn!TablesOrQueries.Value) Then
        MakeListWithLoopID cmmn!TablesOrQueries.Value, direc, cmmn!NoIDCount, suod, cmmn!LoopIDColumn
    End If
    cmmn.MoveNext
Loop

cmmn.MoveFirst

'--- Suljetaan tiedostot päättävällä sulkumerkillä ---
'--- Suljetaan tiedostot päättävällä sulkumerkillä ---
cmmn.MoveFirst
Do Until cmmn.EOF
    ' Suljetaan silmukoimattomat tiedostot
    If Not IsNull(cmmn!TablesOrQueriesNoLoop.Value) Then
        Open direc & get_filename(cmmn!TablesOrQueriesNoLoop.Value) & ".txt" For Append As #1
        Print #1, ")"  ' LISP-listan päättävä sulku
        Close #1
    End If
    ' Suljetaan silmukkaan perustuvat tiedostot
    If Not IsNull(cmmn!TablesOrQueries.Value) Then
        Open direc & get_filename(cmmn!TablesOrQueries.Value) & ".txt" For Append As #1
        Print #1, ")"  ' LISP-listan päättävä sulku
        Close #1
    End If
    cmmn.MoveNext
Loop

scrtest:
' Luodaan AutoCAD-skriptitiedosto eräajoa varten
cmmn.MoveFirst
MakeScript common, suod, cmmn!LoopIDColumn

' Siivotaan objektit
cmmn.Close
Set cmmn = Nothing
Set DB = Nothing

Exit Function

ErrorHandler:
    MsgBox "Error in makeFiles: " & Err.Description & vbCrLf & _
           "Error occurred while generating LISP files.", vbCritical, "File Generation Error"
    ' Siivotaan objektit virhetilanteessa
    On Error Resume Next
    Close #1  ' Suljetaan mahdollisesti auki oleva tiedostokahva
    If Not cmmn Is Nothing Then
        cmmn.Close
        Set cmmn = Nothing
    End If
    Set DB = Nothing
End Function

'------------------------------------------------------------------------------
' Proseduuri: MakeListNoLoopID
' Tarkoitus: Luo LISP-listat taulukoista/kyselyistä ilman silmukka-ID-suodatusta
' Parametrit:
'   tanimi - Taulu/kyselynnimi (voi sisältää * jokerina, esim. "CIRCUIT*")
'   Hakem  - Tulostushakemiston polku
'------------------------------------------------------------------------------
Sub MakeListNoLoopID(tanimi As String, Hakem As String)
On Error GoTo ErrorHandler

Dim DB As DAO.Database
Dim tble As DAO.Recordset
Dim L As String             ' Lainausmerkki
Dim aster As Integer        ' Tähtimerkin sijainti taulukonnimessä
Dim filenum As Integer      ' Tiedostokahvan numero
Dim i As Integer, ii As Integer  ' Silmukkalaskurit
Dim preref As String        ' LISP-muuttujien etuliiteviite

Set DB = CurrentDb

L = Chr(34)  ' Lainausmerkki LISP-listoja varten

aster = InStr(tanimi, "*")

'--- Käsitellään jokerilliset taulukotnimet (esim. "CIRCUIT*") ---
If aster <> 0 Then
  filenum = FreeFile
  Open Hakem & get_filename(tanimi) & ".txt" For Append As filenum

  ' Käydään läpi kaikki etuliitettä vastaavat taulukot
  For i = 0 To DB.TableDefs.Count - 1
      If Mid$(DB.TableDefs(i).Name, 1, aster - 1) = get_filename(tanimi) Then
        Set tble = DB.OpenRecordset(DB.TableDefs(i).Name, dbOpenDynaset)
        If Not tble.EOF Then tble.MoveFirst
        preref = get_filename(tanimi)
        
        ' Käsitellään jokainen tietue
        Do Until tble.EOF
            preref = get_filename(tanimi)
            ' Rakennetaan viittausprefiksi ID-kentät
            For ii = 0 To tble.Fields.Count - 1
                If Right$(tble.Fields(ii).Name, 2) = "ID" Then
                    preref = preref & "." & tble.Fields(ii).Value
                Else
                    Exit For
                End If
            Next
            ' Kirjoitetaan ei-tyhjät kenttäarvot LISP-tiedostoon
            For ii = 0 To tble.Fields.Count - 1
                If Not IsNull(tble.Fields(ii).Value) Then
                    Print #filenum, "( " & L & UCase$(preref) & "." & UCase$(tble.Fields(ii).Name);
                    Print #filenum, L & " " & L & inch(tble.Fields(ii).Value) & L & " )"
                End If
            Next
            tble.MoveNext
        Loop
        tble.Close
      End If
  Next
  Close filenum

'--- Käsitellään yksittäinen taulu/kysely ---
Else
  Set tble = DB.OpenRecordset(tanimi, dbOpenDynaset)
  If Not tble.EOF Then tble.MoveFirst

  filenum = FreeFile
  Open Hakem & get_filename(tanimi) & ".txt" For Append As filenum

  ' Käsitellään jokainen tietue
  Do Until tble.EOF
    preref = get_filename(tanimi)
    ' Rakennetaan viittausprefiksi ID-kentät
    For ii = 0 To tble.Fields.Count - 1
        If Right$(tble.Fields(ii).Name, 2) = "ID" Then
            preref = preref & "." & tble.Fields(ii).Value
        Else
            Exit For
        End If
    Next
    ' Kirjoitetaan ei-tyhjät arvot LISP-tiedostoon (ristiviitehaku mukana)
    For ii = 0 To tble.Fields.Count - 1
        If Not IsNull(tble.Fields(ii).Value) Then
            Print #filenum, "( " & L & UCase$(preref) & "." & UCase$(tble.Fields(ii).Name);
            Print #filenum, L & " " & L & inch(CrsRefLink(tanimi, tble.Fields(ii).Value)) & L & " )"
        End If
    Next
    tble.MoveNext
  Loop

  Close filenum
  tble.Close
End If

' Siivotaan objektit
Set tble = Nothing
Set DB = Nothing

Exit Sub

ErrorHandler:
    MsgBox "Virhe MakeListNoLoopID-rutiinissa: " & Err.Description & vbCrLf & _
           "Taulu/Kysely: " & tanimi, vbCritical, "LISP-generointi epäonnistui"
    ' Siivotaan objektit virhetilanteessa
    On Error Resume Next
    Close filenum
    If Not tble Is Nothing Then
        tble.Close
        Set tble = Nothing
    End If
    Set DB = Nothing
End Sub

Sub MakeListWithLoopID(tblnimipre As String, Hakem As String, idsyst As String, suoda As Variant, Looppid As Integer)
On Error GoTo ErrorHandler

Dim DB As DAO.Database
Dim tble As DAO.Recordset
Dim L As String
Dim aster As Integer
Dim filenum As Integer
Dim i As Integer
Dim ii As Integer
Dim iii As Integer
Dim preref As String

Set DB = CurrentDb
L = Chr(34)

aster = InStr(tblnimipre, "*")
If aster <> 0 Then

  filenum = FreeFile
  Open Hakem & get_filename(tblnimipre) & ".txt" For Append As filenum
  ' Käydään läpi kaikki taulukot
  For i = 0 To DB.TableDefs.Count - 1
      If Mid$(DB.TableDefs(i).Name, 1, aster - 1) = get_filename(tblnimipre) Then
        Set tble = DB.OpenRecordset(DB.TableDefs(i).Name, dbOpenDynaset)
        If Not tble.EOF Then tble.MoveFirst
        ' Käydään läpi kaikki tietueet
        Do Until tble.EOF
            If tble.Fields(0).Value = suoda Then
                preref = tble.Fields(Looppid).Value & "." & get_filename(tblnimipre)
                If idsyst = 0 Then
                    For ii = 1 To tble.Fields.Count - 1
                        If Not tble.Fields(ii).Name = Looppid Then
                            If Right$(tble.Fields(ii).Name, 2) = "ID" Then
                                preref = preref & "." & tble.Fields(ii).Value
                            Else
                                Exit For
                            End If
                        End If
                    Next
                Else
                    For ii = 1 To idsyst
                        If Not tble.Fields(ii).Name = Looppid Then
                            preref = preref & "." & tble.Fields(ii).Value
                        End If
                    Next
                End If
              
                For iii = 0 To tble.Fields.Count - 1
                    If Not IsNull(tble.Fields(iii).Value) Then
                        Print #filenum, "( " & L & UCase$(preref) & "." & UCase$(tble.Fields(iii).Name);
                        Print #filenum, L & " " & L & inch(tble.Fields(iii).Value) & L & " )"
                    End If
                Next
            End If
            tble.MoveNext
        Loop
         
      End If
  Next
  Close

Else

  Set tble = DB.OpenRecordset(tblnimipre, dbOpenDynaset)
  If Not tble.EOF Then tble.MoveFirst

  filenum = FreeFile
  Open Hakem & Mid$(tble.Name, 1, 8) & ".txt" For Append As filenum
  Do Until tble.EOF
    If tble.Fields(0).Value = suoda Then
        preref = tble.Fields(Looppid).Value & "." & get_filename(tblnimipre)
        If idsyst = 0 Then
            For ii = 1 To tble.Fields.Count - 1
                If Not tble.Fields(ii).Name = Looppid Then
                    If Right$(tble.Fields(ii).Name, 2) = "ID" Then
                        preref = preref & "." & tble.Fields(ii).Value
                    Else
                        Exit For
                    End If
                End If
            Next
        Else
            For ii = 1 To idsyst
                If Not tble.Fields(ii).Name = Looppid Then
                    preref = preref & "." & tble.Fields(ii).Value
                End If
            Next
        End If
      
        For ii = 0 To tble.Fields.Count - 1
            If Not IsNull(tble.Fields(ii).Value) Then
                Print #filenum, "( " & L & UCase$(preref) & "." & UCase$(tble.Fields(ii).Name);
                Print #filenum, L & " " & L & inch(tble.Fields(ii).Value) & L & " )"
            End If
        Next
    End If
    tble.MoveNext
  Loop
  
  Close filenum

End If

Exit Sub

ErrorHandler:
    MsgBox "Error in MakeListWithLoopID: " & Err.Description, vbCritical, "Loop ID List Error"
    On Error Resume Next
    If Not tble Is Nothing Then tble.Close
    Close filenum
End Sub

'------------------------------------------------------------------------------
' Funktio: MakeLocFiles
' Tarkoitus: Luo asennussijainnin tiedostot AutoCADille
'
' KOVAKOODATUT POLUT — projektikohtaiset:
'   P:\acaddata\projekti\agropm10\tyo\instloc.txt
'
' Huom: Polut ovat agropm10-projektille. Muokkaa tai siirrä konfiguraatiotauluun.
'------------------------------------------------------------------------------
Function MakeLocFiles()
On Error GoTo ErrorHandler

Dim DB As DAO.Database
Dim cmmn As DAO.Recordset
Dim tbl As DAO.Recordset
Dim Taulukko As DAO.TableDef
Dim Taul As DAO.Recordset
Dim tble As DAO.Recordset
Dim L As String
Dim i As Integer
Dim kentta1 As Variant
Dim kentta2 As Variant

Set DB = CurrentDb
Set tble = DB.OpenRecordset("Loops", dbOpenDynaset)

L = Chr(34)

' Alustetaan txt-tiedostot
        Open "p:\acaddata\projekti\agropm10\tyo\instloc.txt" For Output As #1
        Print #1, "(";
        Close
 
 
 For i = 0 To DB.TableDefs.Count - 1
  Set Taulukko = DB.TableDefs(i)
  If Left$(Taulukko.Name, 6) = "devTbl" Then 'valitaan taulukot
   If Right$(Taulukko.Name, 6) <> "Common" Then
    If Right$(Taulukko.Name, 12) <> "Positioner01" Then
     Set Taul = DB.OpenRecordset(DB.TableDefs(i).Name)
 
        If Not Taul.EOF Then Taul.MoveFirst
          Do Until Taul.EOF
            Open "p:\acaddata\projekti\agropm10\tyo\instloc.txt" For Append As #1
            Print #1, "(" & L;
            kentta1 = Taul.Fields(0).Value
            kentta2 = Taul.Fields(1).Value
            Print #1, (Taul.Fields(0).Value);
            Print #1, (Taul.Fields(1).Value);
            Print #1, L & " " & L;
            Print #1, (Taul.Fields(0).Value) & "-";

                tble.MoveFirst
                 Do Until tble.EOF
                  If Left$(Taul.Fields(2).Value, 2) = "ZS" Then Exit Do
                  If Left$(Taul.Fields(2).Value, 2) = "EV" Then Exit Do
                  If tble!AreaCode.Value = kentta1 And tble!LoopNo.Value = kentta2 Then
                  Print #1, tble!LoopFID.Value;
                  Exit Do
                  Else: tble.MoveNext
                  End If
                 Loop

        Print #1, (Taul.Fields(2).Value);
        If (Taul.Fields(3).Value) <> "-" Then Print #1, (Taul.Fields(3).Value);
        Print #1, "-" & (Taul.Fields(1).Value) & L & " " & L;
        Print #1, Taulukko.Name & "." & Taul!CounterID.Value;
        Print #1, L & ")"
        Close
     Taul.MoveNext
    Loop
   End If
  End If
 End If
Next

' Kirjoitetaan päättävä sulkumerkki tiedostoon
        Open "p:\acaddata\projekti\agropm10\tyo\instloc.txt" For Append As #1
        Print #1, ")"
        Close

Exit Function

ErrorHandler:
    MsgBox "Error in MakeLocFiles: " & Err.Description, vbCritical, "Location Files Error"
    On Error Resume Next
    Close #1
    If Not tble Is Nothing Then tble.Close
    If Not Taul Is Nothing Then Taul.Close
End Function

Sub MakeScript(common As String, suod As Variant, Looppid As Integer)
On Error GoTo ErrorHandler

'common = "COMMON"

Dim DB As DAO.Database
Dim cmmn As DAO.Recordset
Dim tblmain As DAO.Recordset
Dim L As String
Dim iii As Integer

Set DB = CurrentDb
Set cmmn = DB.OpenRecordset(common, dbOpenDynaset)

L = Chr(34)
cmmn.MoveFirst
Set tblmain = DB.OpenRecordset(cmmn.Fields(0).Value, dbOpenDynaset)

Open cmmn!AcadDirectory.Value & cmmn!ScriptFileName.Value For Output As #1

tblmain.MoveFirst
cmmn.MoveFirst

If Not IsNull(cmmn.Fields("ScriptInTheBegining").Value) Then Print #1, cmmn.Fields("ScriptInTheBegining").Value

Print #1, "(QMEM " & L & "W" & L & " 1 " & L & "CRSREF.TXT" & L & ")'nil"
Print #1, "(QMEM " & L & "W" & L & " 0 " & L & "QMEMLIST.TXT" & L & ")'nil"

Open cmmn!AcadDirectory.Value & "qmemlist.txt" For Output As #2
iii = 2
Print #2, "("
Do Until cmmn.EOF
    If Not IsNull(cmmn.Fields(0).Value) Then
      Print #2, "( " & L & get_filename(cmmn.Fields(0).Value) & L & " " & L & iii & L & " )"
      Print #1, "(QMEM " & L & "W" & L & " " & iii & " " & L & get_filename(cmmn.Fields(0).Value) & ".TXT" & L & ")'nil"
    End If
    If Not IsNull(cmmn.Fields(1).Value) Then
      Print #2, "( " & L & get_filename(cmmn.Fields(1).Value) & L & " " & L & iii + 1 & L & " )"
      Print #1, "(QMEM " & L & "W" & L & " " & iii + 1 & " " & L & get_filename(cmmn.Fields(1).Value) & ".TXT" & L & ")'nil"
    End If
    cmmn.MoveNext
    iii = iii + 2
Loop
Print #2, ")"
Close #2

tblmain.MoveFirst
cmmn.MoveFirst

Do Until tblmain.EOF
    If tblmain.Fields(0).Value = suod Then
        If Not IsNull(cmmn.Fields("ScriptBeforeLoop1").Value) Then Print #1, cmmn.Fields("ScriptBeforeLoop1").Value
        If cmmn.Fields!New.Value Then Print #1, "(New " & L & tblmain.Fields(cmmn.Fields("FileNameColumn").Value).Value & L & L & tblmain.Fields(cmmn.Fields("BaseDwgColumn").Value).Value & L & ")"
        If Not IsNull(cmmn.Fields("ScriptBeforeLoop2").Value) Then Print #1, cmmn.Fields("ScriptBeforeLoop2").Value
        Print #1, "(setq loop " & L & tblmain.Fields(Looppid).Value & L & ")"
        If Not IsNull(cmmn.Fields("ScriptAfterLoop").Value) Then Print #1, cmmn.Fields("ScriptAfterLoop").Value
        If cmmn.Fields("Save").Value Then Print #1, "(save " & L & tblmain.Fields(cmmn.Fields("FileNameColumn").Value).Value & L & ")"
    End If
    
    tblmain.MoveNext
Loop

cmmn.MoveFirst
tblmain.MoveLast
If Not IsNull(cmmn.Fields("ScriptInTheEnd").Value) Then Print #1, cmmn.Fields("ScriptInTheEnd").Value

Close

Exit Sub

ErrorHandler:
    MsgBox "Error in MakeScript: " & Err.Description, vbCritical, "Script Generation Error"
    On Error Resume Next
    Close #1
    Close #2
End Sub

