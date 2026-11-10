Option Compare Database   ' K‰ytt‰‰ tietokantaj‰rjestyst‰ merkkijonovertailuissa
Option Explicit           ' Muuttujaesittely pakollinen
'==============================================================================
' Moduuli: DataToACAD
' Tarkoitus: Luo AutoCAD LISP-tiedostot tietokannassa olevista piirikaaviotiedoista
' Alkuper‰inen: 1997-02-21 Fr 11:10 /tw
' Muokattu: 1997-03-19 We 15:46 /tw
' Muokattu: 1997-03-21 Fr 16:07 /tw
' Muokattu: 1997-07-14 Mo 14:29 /tw
' P‰ivitetty: 2025-11-11 - DAO-tyypit, virheenk‰sittely, kommentit lis‰tty
' P‰ivitetty: 2026-03-09 - CrsRefLink muutettu v‰limuisti-hakuun (N+1-pullonkaula poistettu)
' P‰ivitetty: 2026-11-10 - Tiedostokahvat FreeFile-kutsuilla, tiedostot suljetaan aina,
'             k‰ytt‰m‰tˆn MakeLocFiles (kovakoodatut agropm10-polut) poistettu,
'             virheilmoitukset suomeksi
'==============================================================================

' --- Moduulitason v‰limuisti CrsRefLisps-haulle ---
' Ladataan taulukko kerran Dictionary-objektiin; jokainen haku on O(1) levyluvun sijaan
Private dictCrsRef As Object   ' Scripting.Dictionary
Private blnCrsRefLoaded As Boolean

'------------------------------------------------------------------------------
' Alirutiini: NollaaCrsRefValimuisti
' Tarkoitus: Tyhjent‰‰ CrsRefLisps-v‰limuistin, jotta taulun muutokset tulevat
'            mukaan seuraavassa LISP-generoinnissa (v‰limuisti el‰‰ koko istunnon).
'------------------------------------------------------------------------------
Public Sub NollaaCrsRefValimuisti()
    Set dictCrsRef = Nothing
    blnCrsRefLoaded = False
End Sub

'------------------------------------------------------------------------------
' Alirutiini: LoadCrsRefCache
' Tarkoitus: Lataa CrsRefLisps-taulukon kerran muistiin Dictionary-objektiin.
'            Kutsutaan laiskasti CrsRefLink-funktiosta ensimm‰isell‰ hakukerralla.
'------------------------------------------------------------------------------
Private Sub LoadCrsRefCache()
    Dim tble As DAO.Recordset

    Set dictCrsRef = CreateObject("Scripting.Dictionary")
    ' dbOpenForwardOnly on nopein vaihtoehto pelkk‰‰n lukemiseen
    Set tble = CurrentDb.OpenRecordset("CrsRefLisps", dbOpenForwardOnly)

    Do Until tble.EOF
        ' Lis‰t‰‰n vain ensimm‰inen osuma, jos duplikaatteja on
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
' Tarkoitus: Hakee LISP-koodin ristiviitetaulukosta v‰limuistin avulla.
'            Ensimm‰isell‰ kutsulla lataa koko taulukon muistiin; sen j‰lkeen
'            jokainen haku on O(1)-nopea eik‰ avaa Recordsetia lainkaan.
' Parametrit:
'   tblnimi - Taulukonnimi (vain "CRSREF" k‰ynnist‰‰ haun)
'   teksti  - Haettava ristiviite-ID
' Palautusarvo: LISP-koodi tai alkuper‰inen teksti, jos ei lˆydy
'------------------------------------------------------------------------------
Function CrsRefLink(tblnimi As String, teksti As String) As String
On Error GoTo ErrorHandler

If tblnimi = "CRSREF" Then
    ' Ladataan v‰limuistiin vain ensimm‰isell‰ kerralla
    If Not blnCrsRefLoaded Then LoadCrsRefCache

    ' Nopea O(1)-haku Dictionary-objektista levyn sijaan
    If dictCrsRef.Exists(teksti) Then
        CrsRefLink = dictCrsRef(teksti)
    Else
        CrsRefLink = teksti  ' Ei lˆydetty - palautetaan alkuper‰inen teksti
    End If
Else
    ' Ei ristiviite - palautetaan alkuper‰inen teksti
    CrsRefLink = teksti
End If

Exit Function

ErrorHandler:
    MsgBox "Virhe CrsRefLink-funktiossa: " & Err.Description, vbCritical, "Ristiviitehaun virhe"
    CrsRefLink = teksti  ' Virhetilanteessa palautetaan alkuper‰inen teksti
End Function

'------------------------------------------------------------------------------
' Funktio: get_filename
' Tarkoitus: Poimii 8-merkkisen tiedostonimen taulukonnimest‰
' Parametrit:
'   taulnimi - Taulukonnimi (voi sis‰lt‰‰ t‰hti-erottimen)
' Palautusarvo: 8-merkkinen isolla kirjoitettu tiedostonimi
'------------------------------------------------------------------------------
Function get_filename(taulnimi As String) As String
On Error GoTo ErrorHandler

Dim ast As Integer

ast = InStr(taulnimi, "*")
If ast = 0 Then
    ' Ei t‰hte‰ - otetaan 8 ensimm‰ist‰ merkki‰
    get_filename = UCase$(Mid$(taulnimi, 1, 8))
Else
    ' T‰hti lˆytyi - otetaan 8 ensimm‰ist‰ merkki‰ ennen sit‰
    get_filename = UCase$(Mid$(Mid$(taulnimi, 1, ast - 1), 1, 8))
End If

Exit Function

ErrorHandler:
    MsgBox "Virhe get_filename-funktiossa: " & Err.Description, vbCritical, "Tiedostonimen poimintavirhe"
    get_filename = "ERROR"
End Function

'------------------------------------------------------------------------------
' Funktio: inch
' Tarkoitus: Korvaa lainausmerkit LISP-syntaksin \042-sekvenssill‰, jottei
'            merkkijono-arvo katkaise LISP-listan lainausmerkkej‰
' Parametrit:
'   a - Merkkijono, joka sis‰lt‰‰ lainausmerkkej‰
' Palautusarvo: Merkkijono, jossa lainausmerkit korvattu \042-sekvenssill‰
'------------------------------------------------------------------------------
Function inch(a As String) As String
    inch = Replace(a, Chr$(34), "\042")
End Function

'------------------------------------------------------------------------------
' Funktio: NormalisoiHakemisto
' Tarkoitus: Varmistaa, ett‰ hakemistopolku p‰‰ttyy kenoviivaan, jottei
'            tiedostonimi liity suoraan kansion nimeen
'------------------------------------------------------------------------------
Private Function NormalisoiHakemisto(ByVal Polku As String) As String
    If Len(Polku) > 0 And Right$(Polku, 1) <> "\" Then Polku = Polku & "\"
    NormalisoiHakemisto = Polku
End Function

'------------------------------------------------------------------------------
' Alirutiini: KirjoitaRivi
' Tarkoitus: Kirjoittaa yhden rivin tekstitiedostoon ja sulkee tiedoston heti
'            (oma FreeFile-kahva, ei tˆrm‰‰ muihin avoimiin tiedostoihin)
' Parametrit:
'   Polku - Tiedoston t‰ydellinen polku
'   Teksti - Kirjoitettava rivi
'   Uusi - True = luodaan tiedosto alusta, False = lis‰t‰‰n loppuun
'------------------------------------------------------------------------------
Private Sub KirjoitaRivi(ByVal Polku As String, ByVal Teksti As String, ByVal Uusi As Boolean)
    Dim Kahva As Integer  ' Vapaa tiedostokahva

    Kahva = FreeFile
    If Uusi Then
        Open Polku For Output As Kahva
    Else
        Open Polku For Append As Kahva
    End If
    Print #Kahva, Teksti
    Close Kahva
End Sub

'------------------------------------------------------------------------------
' Funktio: makeFiles
' Tarkoitus: P‰‰orkestraattori - luo kaikki AutoCAD LISP-tiedostot tietokannasta
' Parametrit:
'   common - Konfiguraatiotaulun nimi
' Toiminta:
'   1. Lukee asetukset konfiguraatiotaulusta
'   2. Alustaa .txt-tulostiedostot
'   3. Luo silmukoimattomat listat
'   4. Luo silmukoihin perustuvat listat (jos k‰ytˆss‰)
'   5. Sulkee listat p‰‰tt‰v‰ll‰ sulkumerkill‰ (myˆs silloin, kun silmukkalistoja ei ole)
'   6. Luo AutoCAD-skriptitiedoston
' Palautusarvo: 0 = onnistui, -1 = virhe
'------------------------------------------------------------------------------
Function makeFiles(common As String) As Integer
On Error GoTo ErrorHandler

Dim DB As DAO.Database
Dim cmmn As DAO.Recordset   ' Konfiguraatiotietue
Dim suod As Variant         ' Suodatusarvo
Dim direc As String         ' Tulostushakemiston polku

Set DB = CurrentDb
Set cmmn = DB.OpenRecordset(common, dbOpenDynaset)

' Ristiviitetaulun muutokset tulevat mukaan jokaisella ajokerralla
NollaaCrsRefValimuisti

cmmn.MoveFirst
suod = cmmn.Fields("Filter")
direc = NormalisoiHakemisto(cmmn!AcadDirectory)  ' Hakemisto, jonne LISP-tiedostot tallennetaan

' Jos vain skriptitiedosto - ohitetaan LISP-tiedostojen luonti
If Not cmmn!OnlyScript Then

    '--- Alustetaan kaikki tulostiedostot avaavalla sulkumerkill‰ ---
    cmmn.MoveFirst
    Do Until cmmn.EOF
        If Not IsNull(cmmn!TablesOrQueriesNoLoop.Value) Then
            KirjoitaRivi direc & get_filename(cmmn!TablesOrQueriesNoLoop.Value) & ".txt", "(", True
        End If
        If Not IsNull(cmmn!TablesOrQueries.Value) Then
            KirjoitaRivi direc & get_filename(cmmn!TablesOrQueries.Value) & ".txt", "(", True
        End If
        cmmn.MoveNext
    Loop

    '--- Luodaan silmukoimattomat LISP-listat ---
    ' Yksinkertaiset listat ilman silmukka-ID-suodatusta
    cmmn.MoveFirst
    Do Until cmmn.EOF
        If Not IsNull(cmmn!TablesOrQueriesNoLoop.Value) Then
            MakeListNoLoopID cmmn!TablesOrQueriesNoLoop.Value, direc
        End If
        cmmn.MoveNext
    Loop

    '--- Luodaan silmukkaan perustuvat LISP-listat ---
    ' N‰m‰ listat suodatetaan silmukka-ID:n perusteella (ei, jos silmukkatauluja ei ole)
    cmmn.MoveFirst
    If Not cmmn!NoLoopIDTables Then
        Do Until cmmn.EOF
            If Not IsNull(cmmn!TablesOrQueries.Value) Then
                MakeListWithLoopID cmmn!TablesOrQueries.Value, direc, CLng(Nz(cmmn!NoIDCount, 0)), suod, cmmn!LoopIDColumn
            End If
            cmmn.MoveNext
        Loop
    End If

    '--- Suljetaan tiedostot p‰‰tt‰v‰ll‰ sulkumerkill‰ ---
    ' Tehd‰‰n aina, jotta LISP-lista on sulkumerkeilt‰‰n validi silmukkalistojen puuttuessakin
    cmmn.MoveFirst
    Do Until cmmn.EOF
        If Not IsNull(cmmn!TablesOrQueriesNoLoop.Value) Then
            KirjoitaRivi direc & get_filename(cmmn!TablesOrQueriesNoLoop.Value) & ".txt", ")", False
        End If
        If Not IsNull(cmmn!TablesOrQueries.Value) Then
            KirjoitaRivi direc & get_filename(cmmn!TablesOrQueries.Value) & ".txt", ")", False
        End If
        cmmn.MoveNext
    Loop
End If

' Luodaan AutoCAD-skriptitiedosto er‰ajoa varten
cmmn.MoveFirst
MakeScript common, suod, cmmn!LoopIDColumn

makeFiles = 0

Cleanup:
    On Error Resume Next
    If Not cmmn Is Nothing Then cmmn.Close
    Set cmmn = Nothing
    Set DB = Nothing
    Exit Function

ErrorHandler:
    MsgBox "Virhe makeFiles-funktiossa: " & Err.Description & vbCrLf & _
           "LISP-tiedostojen luonti ep‰onnistui.", vbCritical, "Tiedostojen luontivirhe"
    makeFiles = -1
    Resume Cleanup
End Function

'------------------------------------------------------------------------------
' Proseduuri: MakeListNoLoopID
' Tarkoitus: Luo LISP-listat taulukoista/kyselyist‰ ilman silmukka-ID-suodatusta
' Parametrit:
'   tanimi - Taulu/kyselynnimi (voi sis‰lt‰‰ * jokerina, esim. "CIRCUIT*")
'   Hakem  - Tulostushakemiston polku
'------------------------------------------------------------------------------
Sub MakeListNoLoopID(tanimi As String, Hakem As String)
On Error GoTo ErrorHandler

Dim DB As DAO.Database
Dim tble As DAO.Recordset
Dim L As String             ' Lainausmerkki
Dim aster As Integer        ' T‰htimerkin sijainti taulukonnimess‰
Dim filenum As Integer      ' Tiedostokahvan numero (0 = ei avattu)
Dim i As Integer, ii As Integer  ' Silmukkalaskurit
Dim preref As String        ' LISP-muuttujien etuliiteviite

Set DB = CurrentDb

L = Chr$(34)  ' Lainausmerkki LISP-listoja varten

aster = InStr(tanimi, "*")

'--- K‰sitell‰‰n jokerilliset taulukonnimet (esim. "CIRCUIT*") ---
If aster <> 0 Then
  filenum = FreeFile
  Open Hakem & get_filename(tanimi) & ".txt" For Append As filenum

  ' K‰yd‰‰n l‰pi kaikki etuliitett‰ vastaavat taulukot
  For i = 0 To DB.TableDefs.Count - 1
      If Mid$(DB.TableDefs(i).Name, 1, aster - 1) = get_filename(tanimi) Then
        Set tble = DB.OpenRecordset(DB.TableDefs(i).Name, dbOpenDynaset)
        If Not tble.EOF Then tble.MoveFirst

        ' K‰sitell‰‰n jokainen tietue
        Do Until tble.EOF
            preref = get_filename(tanimi)
            ' Rakennetaan viittausprefiksi ID-kentist‰ (per‰kk‰iset ...ID-p‰‰tteiset kent‰t)
            For ii = 0 To tble.Fields.Count - 1
                If Right$(tble.Fields(ii).Name, 2) = "ID" Then
                    preref = preref & "." & tble.Fields(ii).Value
                Else
                    Exit For
                End If
            Next
            ' Kirjoitetaan ei-tyhj‰t kentt‰arvot LISP-tiedostoon
            For ii = 0 To tble.Fields.Count - 1
                If Not IsNull(tble.Fields(ii).Value) Then
                    Print #filenum, "( " & L & UCase$(preref) & "." & UCase$(tble.Fields(ii).Name);
                    Print #filenum, L & " " & L & inch(tble.Fields(ii).Value) & L & " )"
                End If
            Next
            tble.MoveNext
        Loop
        tble.Close
        Set tble = Nothing
      End If
  Next
  Close filenum
  filenum = 0

'--- K‰sitell‰‰n yksitt‰inen taulu/kysely ---
Else
  Set tble = DB.OpenRecordset(tanimi, dbOpenDynaset)
  If Not tble.EOF Then tble.MoveFirst

  filenum = FreeFile
  Open Hakem & get_filename(tanimi) & ".txt" For Append As filenum

  ' K‰sitell‰‰n jokainen tietue
  Do Until tble.EOF
    preref = get_filename(tanimi)
    ' Rakennetaan viittausprefiksi ID-kentist‰
    For ii = 0 To tble.Fields.Count - 1
        If Right$(tble.Fields(ii).Name, 2) = "ID" Then
            preref = preref & "." & tble.Fields(ii).Value
        Else
            Exit For
        End If
    Next
    ' Kirjoitetaan ei-tyhj‰t arvot LISP-tiedostoon (ristiviitehaku mukana)
    For ii = 0 To tble.Fields.Count - 1
        If Not IsNull(tble.Fields(ii).Value) Then
            Print #filenum, "( " & L & UCase$(preref) & "." & UCase$(tble.Fields(ii).Name);
            Print #filenum, L & " " & L & inch(CrsRefLink(tanimi, tble.Fields(ii).Value)) & L & " )"
        End If
    Next
    tble.MoveNext
  Loop

  Close filenum
  filenum = 0
  tble.Close
End If

' Siivotaan objektit
Set tble = Nothing
Set DB = Nothing

Exit Sub

ErrorHandler:
    MsgBox "Virhe MakeListNoLoopID-rutiinissa: " & Err.Description & vbCrLf & _
           "Taulu/Kysely: " & tanimi, vbCritical, "LISP-generointi ep‰onnistui"
    ' Siivotaan objektit virhetilanteessa
    On Error Resume Next
    If filenum <> 0 Then Close filenum
    If Not tble Is Nothing Then
        tble.Close
        Set tble = Nothing
    End If
    Set DB = Nothing
End Sub

'------------------------------------------------------------------------------
' Proseduuri: MakeListWithLoopID
' Tarkoitus: Luo LISP-listat taulukoista/kyselyist‰ vain valitun silmukan riveille
' Parametrit:
'   tblnimipre - Taulu/kyselynnimi (voi sis‰lt‰‰ * jokerina)
'   Hakem      - Tulostushakemiston polku
'   idsyst     - Viittausprefiksiin otettavien ID-kenttien m‰‰r‰ (0 = automaattinen)
'   suoda      - Suodatusarvo (verrataan taulun ensimm‰iseen kentt‰‰n)
'   Looppid    - Silmukka-ID-kent‰n indeksi
'------------------------------------------------------------------------------
Sub MakeListWithLoopID(tblnimipre As String, Hakem As String, idsyst As Long, suoda As Variant, Looppid As Integer)
On Error GoTo ErrorHandler

Dim DB As DAO.Database
Dim tble As DAO.Recordset
Dim L As String
Dim aster As Integer
Dim filenum As Integer      ' Tiedostokahva (0 = ei avattu)
Dim i As Integer
Dim ii As Integer
Dim iii As Integer
Dim preref As String

Set DB = CurrentDb
L = Chr$(34)

aster = InStr(tblnimipre, "*")
filenum = FreeFile
Open Hakem & get_filename(tblnimipre) & ".txt" For Append As filenum

If aster <> 0 Then

  ' K‰yd‰‰n l‰pi kaikki etuliitett‰ vastaavat taulukot
  For i = 0 To DB.TableDefs.Count - 1
      If Mid$(DB.TableDefs(i).Name, 1, aster - 1) = get_filename(tblnimipre) Then
        Set tble = DB.OpenRecordset(DB.TableDefs(i).Name, dbOpenDynaset)
        If Not tble.EOF Then tble.MoveFirst
        ' K‰yd‰‰n l‰pi kaikki tietueet
        Do Until tble.EOF
            If tble.Fields(0).Value = suoda Then
                preref = tble.Fields(Looppid).Value & "." & get_filename(tblnimipre)
                If idsyst = 0 Then
                    For ii = 1 To tble.Fields.Count - 1
                        If ii <> Looppid Then
                            If Right$(tble.Fields(ii).Name, 2) = "ID" Then
                                preref = preref & "." & tble.Fields(ii).Value
                            Else
                                Exit For
                            End If
                        End If
                    Next
                Else
                    For ii = 1 To idsyst
                        If ii <> Looppid Then
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
        tble.Close
        Set tble = Nothing
      End If
  Next

Else

  Set tble = DB.OpenRecordset(tblnimipre, dbOpenDynaset)
  If Not tble.EOF Then tble.MoveFirst

  Do Until tble.EOF
    If tble.Fields(0).Value = suoda Then
        preref = tble.Fields(Looppid).Value & "." & get_filename(tblnimipre)
        If idsyst = 0 Then
            For ii = 1 To tble.Fields.Count - 1
                If ii <> Looppid Then
                    If Right$(tble.Fields(ii).Name, 2) = "ID" Then
                        preref = preref & "." & tble.Fields(ii).Value
                    Else
                        Exit For
                    End If
                End If
            Next
        Else
            For ii = 1 To idsyst
                If ii <> Looppid Then
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
  tble.Close
  Set tble = Nothing

End If

Close filenum
filenum = 0
Set DB = Nothing

Exit Sub

ErrorHandler:
    MsgBox "Virhe MakeListWithLoopID-rutiinissa: " & Err.Description & vbCrLf & _
           "Taulu/Kysely: " & tblnimipre, vbCritical, "Silmukkalistan luontivirhe"
    On Error Resume Next
    If Not tble Is Nothing Then tble.Close
    Set tble = Nothing
    If filenum <> 0 Then Close filenum
    Set DB = Nothing
End Sub

'------------------------------------------------------------------------------
' Proseduuri: MakeScript
' Tarkoitus: Luo AutoCAD-skriptitiedoston ja qmemlist.txt-tiedoston er‰ajoa varten
' Parametrit:
'   common  - Konfiguraatiotaulun nimi
'   suod    - Suodatusarvo (verrataan p‰‰taulun ensimm‰iseen kentt‰‰n)
'   Looppid - Silmukka-ID-kent‰n indeksi p‰‰taulussa
'------------------------------------------------------------------------------
Sub MakeScript(common As String, suod As Variant, Looppid As Integer)
On Error GoTo ErrorHandler

Dim DB As DAO.Database
Dim cmmn As DAO.Recordset
Dim tblmain As DAO.Recordset
Dim L As String
Dim iii As Integer
Dim Hakem As String         ' Tulostushakemisto
Dim ScriptKahva As Integer  ' Skriptitiedoston kahva (0 = ei avattu)
Dim ListaKahva As Integer   ' qmemlist.txt-tiedoston kahva (0 = ei avattu)

Set DB = CurrentDb
Set cmmn = DB.OpenRecordset(common, dbOpenDynaset)

L = Chr$(34)
cmmn.MoveFirst
Hakem = NormalisoiHakemisto(cmmn!AcadDirectory.Value)
Set tblmain = DB.OpenRecordset(cmmn.Fields(0).Value, dbOpenDynaset)

ScriptKahva = FreeFile
Open Hakem & cmmn!ScriptFileName.Value For Output As ScriptKahva

cmmn.MoveFirst

If Not IsNull(cmmn.Fields("ScriptInTheBegining").Value) Then Print #ScriptKahva, cmmn.Fields("ScriptInTheBegining").Value

Print #ScriptKahva, "(QMEM " & L & "W" & L & " 1 " & L & "CRSREF.TXT" & L & ")'nil"
Print #ScriptKahva, "(QMEM " & L & "W" & L & " 0 " & L & "QMEMLIST.TXT" & L & ")'nil"

ListaKahva = FreeFile
Open Hakem & "qmemlist.txt" For Output As ListaKahva
iii = 2
Print #ListaKahva, "("
Do Until cmmn.EOF
    If Not IsNull(cmmn.Fields(0).Value) Then
      Print #ListaKahva, "( " & L & get_filename(cmmn.Fields(0).Value) & L & " " & L & iii & L & " )"
      Print #ScriptKahva, "(QMEM " & L & "W" & L & " " & iii & " " & L & get_filename(cmmn.Fields(0).Value) & ".TXT" & L & ")'nil"
    End If
    If Not IsNull(cmmn.Fields(1).Value) Then
      Print #ListaKahva, "( " & L & get_filename(cmmn.Fields(1).Value) & L & " " & L & iii + 1 & L & " )"
      Print #ScriptKahva, "(QMEM " & L & "W" & L & " " & iii + 1 & " " & L & get_filename(cmmn.Fields(1).Value) & ".TXT" & L & ")'nil"
    End If
    cmmn.MoveNext
    iii = iii + 2
Loop
Print #ListaKahva, ")"
Close ListaKahva
ListaKahva = 0

' Kirjoitetaan skriptirivit jokaiselle valitun silmukan tietueelle
cmmn.MoveFirst
If Not tblmain.EOF Then tblmain.MoveFirst
Do Until tblmain.EOF
    If tblmain.Fields(0).Value = suod Then
        If Not IsNull(cmmn.Fields("ScriptBeforeLoop1").Value) Then Print #ScriptKahva, cmmn.Fields("ScriptBeforeLoop1").Value
        If cmmn.Fields!New.Value Then Print #ScriptKahva, "(New " & L & tblmain.Fields(cmmn.Fields("FileNameColumn").Value).Value & L & L & tblmain.Fields(cmmn.Fields("BaseDwgColumn").Value).Value & L & ")"
        If Not IsNull(cmmn.Fields("ScriptBeforeLoop2").Value) Then Print #ScriptKahva, cmmn.Fields("ScriptBeforeLoop2").Value
        Print #ScriptKahva, "(setq loop " & L & tblmain.Fields(Looppid).Value & L & ")"
        If Not IsNull(cmmn.Fields("ScriptAfterLoop").Value) Then Print #ScriptKahva, cmmn.Fields("ScriptAfterLoop").Value
        If cmmn.Fields("Save").Value Then Print #ScriptKahva, "(save " & L & tblmain.Fields(cmmn.Fields("FileNameColumn").Value).Value & L & ")"
    End If

    tblmain.MoveNext
Loop

If Not IsNull(cmmn.Fields("ScriptInTheEnd").Value) Then Print #ScriptKahva, cmmn.Fields("ScriptInTheEnd").Value

Close ScriptKahva
ScriptKahva = 0

tblmain.Close
cmmn.Close
Set tblmain = Nothing
Set cmmn = Nothing
Set DB = Nothing

Exit Sub

ErrorHandler:
    MsgBox "Virhe MakeScript-rutiinissa: " & Err.Description, vbCritical, "Skriptin luontivirhe"
    On Error Resume Next
    If ScriptKahva <> 0 Then Close ScriptKahva
    If ListaKahva <> 0 Then Close ListaKahva
    If Not tblmain Is Nothing Then tblmain.Close
    If Not cmmn Is Nothing Then cmmn.Close
    Set tblmain = Nothing
    Set cmmn = Nothing
    Set DB = Nothing
End Sub
