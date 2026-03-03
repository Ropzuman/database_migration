Option Compare Database
Option Explicit

'================================================================================
' Lomake: FrmMUOKKAUS (Muokkaus/Luonti)
' Tarkoitus: Luo Word- ja HTML-toimintakuvaukset piireille ja laitteistolle
' Päivitetty: 2025-11-13 — VBA7/64-bit tuki lisätty
'             2026-03-03 — Me.-etuliite lomakeohjaimille; SelStart-suojaukset
'
' Description:
'   Main editing and generation form for creating function descriptions.
'   Generates both Word (.docx) and HTML documents from database templates.
'   Handles loops (PIIRIT) and equipment (MOOTTORIT) descriptions.
'
' Document Generation Process:
'   1. User selects template type (loops/motors) and output format (Word/HTML)
'   2. Form reads data from TIEDOT_PIIRIT or TIEDOT_MOOTTORIT tables
'   3. Text fields support special syntax:
'      - {LOOPTAG} inserts hyperlinked loop references
'      - {Equipment references} insert equipment tags
'      - Lines starting with "-" create indented lists
'      - Lines starting with "--" create double-indented lists
'   4. For Word: Opens template, replaces bookmarks, saves as new document
'   5. For HTML: Generates HTML file with CSS styling and hyperlinks
'
' Template Bookmarks (Word):
'   - _REV, _STAT, _AUTH, _DOCNO, _PL1, _PL2 (header fields)
'   - _LTAG, _DESCR, _FUNCT, _INTERL, _ALARMS, _CROSS, _REV1 (content fields)
'   - Alarm levels: _HH, _H, _LL, _L with _ALARMDESCR
'
' Dependencies:
'   - Word Automation (Microsoft Word Object Library)
'   - FileSystemObject (Microsoft Scripting Runtime)
'   - DAO.Recordset
'   - Forms: MOTORS subform, PIIRIT subform
'   - Tables: TIEDOT_PIIRIT, TIEDOT_MOOTTORIT, SELOSTUKSET
'   - Queries: LUOPIIRIT (creates new loop descriptions)
'   - Module: USysCheck (SniffUser function)
'   - Public variables: KohdeTextBox As TextBox, Kursori As Long
'   - Controls: LOOPTAG, CLoppuun, FKieli, LomakePohja, FTyyppi
'             Alarms, Function, Interlocks, Crossref, RevHistoria
'================================================================================

' Module-level variables
Dim POHJA As String            ' Template file path
Dim POLKU As String            ' Output directory path
Dim REV As String              ' Revision number
Dim STATUS As String           ' Document status
Dim AUTHOR As String           ' Document author
Dim DOCNO As String            ' Document number
Dim PROJECT_L1 As String       ' Project line 1
Dim PROJECT_L2 As String       ' Project line 2
Dim oWord As Word.Application  ' Word application object
Dim oTemplate As Word.Document ' Word document object
Dim Tiedot As DAO.Recordset    ' Data recordset
Dim Samaan As Boolean          ' Generate all into one file
Dim Eka As Boolean             ' First document flag
Dim Suomeksi As Boolean        ' Finnish language flag
Dim TOUT As TextStream         ' HTML output file stream

'================================================================================
' Tapahtumat: Tekstikenttien fokus ja kursoriseuranta
'================================================================================
Private Sub Alarms_GotFocus()
    On Error Resume Next
    Set KohdeTextBox = Me.Alarms
End Sub

Private Sub Alarms_KeyUp(KeyCode As Integer, Shift As Integer)
    On Error Resume Next
    Kursori = Me.Alarms.SelStart
End Sub

Private Sub Alarms_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Error Resume Next
    Kursori = Me.Alarms.SelStart
End Sub

'================================================================================
' Tapahtuma: Command13_Click
' Tarkoitus: Luo uudet piirikuvaukset SELOSTUKSET-tauluun
'================================================================================
Private Sub Command13_Click()
    On Error GoTo ErrorHandler
    
    If MsgBox("Luodaanko uudet piirit SELOSTUKSET taulukkoon?", vbOKCancel, "LUO PIIRIT") = vbOK Then
        CurrentDb.Execute "LUOPIIRIT"
        Me.Requery
    End If
    Exit Sub
    
ErrorHandler:
    MsgBox "Virhe piirikuvausten luonnissa: " & Err.Description, vbExclamation
End Sub

'================================================================================
' Tapahtuma: Command16_Click
' Tarkoitus: Luo yksittäinen toimintakuvaus
'================================================================================
Private Sub Command16_Click()
    On Error GoTo ErrorHandler
    Me.Refresh
    TeeSelostus False
    Exit Sub
    
ErrorHandler:
    MsgBox "Virhe kuvauksen luonnissa: " & Err.Description, vbExclamation
End Sub

'================================================================================
' Tapahtuma: Command17_Click
' Tarkoitus: Luo kaikki toimintakuvaukset
'================================================================================
Private Sub Command17_Click()
    On Error GoTo ErrorHandler
    Me.Refresh
    TeeSelostus True
    Exit Sub
    
ErrorHandler:
    MsgBox "Virhe kuvausten luonnissa: " & Err.Description, vbExclamation
End Sub

'================================================================================
' Aliohjelma: TeeSelostus
' Tarkoitus: Luo Word- tai HTML-toimintakuvauksia
' Parametrit:
'   Kaikki — True = kaikki kuvaukset, False = vain nykyinen tietue
'================================================================================
Private Sub TeeSelostus(Kaikki As Boolean)
    Dim Nimi As String
    Dim TTIEDOT As String
    Dim oPohja As Word.Document
    Dim FS As New FileSystemObject
    
    On Error GoTo ErrorHandler
   
    Suomeksi = False
'  If Me.FKieli.Value = 2 Then Suomeksi = True

    PoimiAsetukset
    Samaan = False
    Eka = True
    If Kaikki Then
'    If MsgBox("Tehdäänkä yhteen tiedostoon?", vbYesNo) = vbYes Then
'      Samaan = True
'    End If
    End If
    
    ' Varmistetaan, että tulostuspolku on määritelty ja hakemisto on olemassa
    If POLKU = "" Then
        MsgBox "OUTPUTPATH on tyhjä SETTINGS-taulussa. Tarkista taulun sisältö.", vbExclamation, "TeeSelostus"
        Exit Sub
    End If
    If Me.FTyyppi <> 1 Then  ' HTML-polku: tarkistetaan hakemiston olemassaolo
        If Not FS.FolderExists(POLKU) Then
            MsgBox "Tulostushakemistoa ei löydy:" & vbCrLf & POLKU & vbCrLf & _
                   "Luo hakemisto ensin tai korjaa SETTINGS.OUTPUTPATH.", vbExclamation, "TeeSelostus"
            Exit Sub
        End If
    End If
    
    ' Valitaan tietolähde lomakepohjan tyypin mukaan
    If Me.LomakePohja.Value = 1 Then
        TTIEDOT = "TIEDOT_PIIRIT"
    Else
        TTIEDOT = "TIEDOT_MOOTTORIT"
    End If
  
    If Me.FTyyppi = 1 Then 'WORD
        Set oWord = CreateObject("Word.Application") ' oTemplate.Application
        oWord.Visible = True
    Else
    
    End If
    
    If Kaikki Then
        Set Tiedot = CurrentDb.OpenRecordset(TTIEDOT)
    Else
        If Me.LomakePohja.Value = 1 Then
            Set Tiedot = CurrentDb.OpenRecordset("SELECT * FROM " & TTIEDOT & " WHERE Looptag='" & Me.LOOPTAG & "'")
        Else
            Set Tiedot = CurrentDb.OpenRecordset("SELECT * FROM " & TTIEDOT & " WHERE Looptag='" & Me.LOOPTAG & "'")
        End If
    End If
    
    If Samaan Then
        Set oTemplate = oWord.Documents.Open(POHJA)
        Nimi = POLKU & DOCNO & "_KAIKKI.docx"
        oTemplate.SaveAs Nimi
        Set oPohja = oWord.Documents.Open(POHJA, , True)
        oTemplate.Activate
        Do While Not Tiedot.EOF
            MuutaTiedot
            oPohja.Activate
            oWord.Selection.WholeStory
            oWord.Selection.Copy
            oTemplate.Activate
            oWord.Selection.EndKey wdStory
            oWord.Selection.InsertBreak wdPageBreak
            oWord.Selection.Paste
        Loop
    Else
        Do While Not Tiedot.EOF
            Dim sTagiEnnen As String
            sTagiEnnen = Tiedot.Fields("LoopTAG")
            If Me.FTyyppi = 1 Then 'Word versiot
                Set oTemplate = oWord.Documents.Open(POHJA)
                Nimi = POLKU & DOCNO & "_" & sTagiEnnen & ".docx"
                MuutaTiedot
                oTemplate.SaveAs Nimi
                If Kaikki Then oTemplate.Close False
                ' Turvasilmukka Word-polulle: varmistetaan eteneminen vaikka MuutaTiedot ei edistäisi
                Do While Not Tiedot.EOF
                    If Tiedot.Fields("LoopTAG") <> sTagiEnnen Then Exit Do
                    Tiedot.MoveNext
                Loop
            Else  'HTML versiot
                Set TOUT = FS.CreateTextFile(POLKU & sTagiEnnen & ".html")
                TeeHTML
                TOUT.Close
                ' Turvasilmukka: jos TeeHTML ei edennyt (virhe keskeyt\u00e4nyt MoveNext-kutsun),
                ' hyp\u00e4t\u00e4\u00e4n k\u00e4sin t\u00e4m\u00e4n piirin kaikki j\u00e4ljell\u00e4 olevat rivit
                Do While Not Tiedot.EOF
                    If Tiedot.Fields("LoopTAG") <> sTagiEnnen Then Exit Do
                    Tiedot.MoveNext
                Loop
            End If
        Loop
    End If
    
    MsgBox "Valmis"
    Exit Sub
    
ErrorHandler:
    MsgBox "Virhe kuvausten luonnissa: " & Err.Description, vbExclamation
    If Not Tiedot Is Nothing Then Tiedot.Close
    If Not oTemplate Is Nothing Then oTemplate.Close False
    If Not oWord Is Nothing Then oWord.Quit False
End Sub

'================================================================================
' Sub: TeeHTML (Generate HTML)
' Purpose: Generate HTML function description file
'
' Description:
'   Creates HTML file with loop/motor function description, including:
'   - Document header with title and CSS link
'   - Function, Interlocks, Alarms sections
'   - Cross-references to other loops
'   - Revision history
'   Special formatting: {LOOPTAG} becomes hyperlink, "-" lines become indented
'================================================================================
Private Sub TeeHTML()
    'Luodaan html muotoinen selostus. TOUT on tiedosto-objekti, joka on jo määritelty
    'Luodaan ensin header osuus
    Dim TEKSTI As String
    Dim Rivit As Variant
    Dim r As Integer
    Dim ADescr As String
    Dim Halyt As String
    Dim Tagi As String
    Dim Laitteet As String
    
    On Error GoTo ErrorHandler

    Tagi = Tiedot.Fields("LOOPTAG")

TOUT.WriteLine "<!DOCTYPE HTML PUBLIC ""-//W3C//DTD HTML 4.01 Transitional//EN"">"
TOUT.WriteLine "<html>"
TOUT.WriteLine "<head>"
TOUT.WriteLine "    <meta http-equiv=""Content-Type"" content=""text/html; charset=ISO-8859-1"">"
TOUT.WriteLine "    <link type=""text/css"" rel=""stylesheet"" href=""FuncDesc.css"">"
TOUT.WriteLine "    <title>" & Tiedot.Fields("LoopTAG") & " " & Tiedot.Fields("Description") & "</title>"
TOUT.WriteLine "</head>"
TOUT.WriteLine "<body>"
'OTSIKKO - Piirin positio ja nimitys
TOUT.WriteLine "    <h1 id=""loopname"">"
TOUT.WriteLine "        <a href=""" & Tiedot.Fields("LOOPTAG") & ".html"" target=""_parent"">" _
        & Tiedot.Fields("LOOPTAG") & "</a> " _
        & Tiedot.Fields("Description") '& "<br>"
TOUT.WriteLine "    </h1>"

'TOIMINTA
TOUT.WriteLine "    <h2 id=""function"">"
TOUT.WriteLine "        Toiminta"
TOUT.WriteLine "    </h2>"

TOUT.WriteLine "    <div class=""description"">"
TEKSTI = Nz(Tiedot.Fields("Function")) 'Poimitaan teksti muuttujaan
MuutaTeksti TEKSTI                     'Muuttaa tekstin ja kirjoittaa sen suoraan tiedostoon
TOUT.WriteLine "    </div>"


'LUKITUKSET
TOUT.WriteLine "    <h2 id=""interlocks"">"
TOUT.WriteLine "        Lukitukset"
TOUT.WriteLine "    </h2>"
TEKSTI = Nz(Tiedot.Fields("Interlocks"))
MuutaTeksti TEKSTI

'HÄLYTYKSET
TOUT.WriteLine "    <h2 id=""alarms"">"
TOUT.WriteLine "        Hälytykset"
TOUT.WriteLine "    </h2>"
TEKSTI = Nz(Tiedot.Fields("Alarms"))
MuutaTeksti TEKSTI
    
  If IsNull(Tiedot.Fields("AlarmDescrBL")) And IsNull(Tiedot.Fields("Alarms")) Then
'    TOUT.WriteLine "    <p>"
'    TOUT.WriteLine "Ei erityisiä hälytyksiä"
'    TOUT.WriteLine "    </p>"
  Else
    If IsNull(Tiedot.Fields("HH")) And IsNull(Tiedot.Fields("H")) And IsNull(Tiedot.Fields("LL")) And IsNull(Tiedot.Fields("L")) Then
      If IsNull(Tiedot.Fields("Alarms")) Then
'        TOUT.WriteLine "    <p>"
'        TOUT.WriteLine "Ei hälytyksiä"
'        TOUT.WriteLine "    </p>"
      End If
    Else
      ADescr = Tiedot.Fields("AlarmDescrBL")
      TOUT.WriteLine "<dl>"
        LisaaHalyHTML Tiedot.Fields("HH"), "HH: " & ADescr & " > "
        LisaaHalyHTML Tiedot.Fields("H"), "H: " & ADescr & " > "
        LisaaHalyHTML Tiedot.Fields("L"), "L: " & ADescr & " < "
        LisaaHalyHTML Tiedot.Fields("LL"), "LL: " & ADescr & " < "
      TOUT.WriteLine "</dl>"
    End If
  End If
'----------------------------------------


'TIEDOT MUIHIN PIIREIHIN
TOUT.WriteLine "    <h2 id=""crossref"">"
TOUT.WriteLine "        Tiedot muihin piireihin"
TOUT.WriteLine "    </h2>"
TEKSTI = Nz(Tiedot.Fields("Linkit")) 'Poimitaan teksti muuttujaan
MuutaTeksti TEKSTI                     'Muuttaa tekstin ja kirjoittaa sen suoraan tiedostoon


'MUUTOSHISTORIA
TOUT.WriteLine "    <h2 id=""rev"">"
TOUT.WriteLine "        Muutoshistoria"
TOUT.WriteLine "    </h2>"
TEKSTI = Nz(Tiedot.Fields("Muutokset")) 'Poimitaan teksti muuttujaan
MuutaTeksti TEKSTI                     'Muuttaa tekstin ja kirjoittaa sen suoraan tiedostoon

'Tiedoston loppu
TOUT.WriteLine "</body>"
TOUT.WriteLine "</html>"

'Muilla riveillä olisi piirin kaikki laitteet, mutta niitä ei tuloste näkyviin, joten ohitetaan ne
  Do While Not Tiedot.EOF
    If Tiedot.Fields("LoopTAG") <> Tagi Then Exit Do
    If Laitteet <> "" Then Laitteet = Laitteet & "<br>"
'    Laitteet = Laitteet & Tiedot.Fields("TAGNAME")
    Tiedot.MoveNext
  Loop
  
  Exit Sub
  
ErrorHandler:
    MsgBox "Virhe HTML-kuvauksen luonnissa: " & Err.Description, vbExclamation
End Sub

'================================================================================
' Aliohjelma: MuutaTeksti
' Tarkoitus: Muuttaa pelkän tekstin HTML-muotoiseksi
' Parametrit:
'   TEKSTI — Muunnettava teksti ({LOOPTAG}-linkit ja sisennykset tuettu)
'
' Kuvaus:
'   Muuntaa tekstiöntö HTML-muotoon:
'   - {LOOPTAG} muuttuu hyperlinkiksi
'   - "-"-alkuiset rivit sisennetään (dd dl:ssä)
'   - "--"-alkuiset rivit sisennetään kaksinkertaisesti
'   - Tyhjä rivi ennen sisennystä muuttuu otsikkoriviksi (dt dl:ssä)
'================================================================================
Private Sub MuutaTeksti(TEKSTI As String)
    Dim Rivit As Variant
    Dim Riveja As Integer
    Dim r As Integer
    Dim Sis As Boolean
    
    On Error GoTo ErrorHandler

    If TEKSTI <> "" Then
        TEKSTI = MuutaLinkit(TEKSTI)  'Muutetaan aaltosulkeissa olevat piiripositiot linkeiksi
        Rivit = Split(TEKSTI, vbCrLf) 'Splitataan rivit ja muutetaan html kappeliksi
        Riveja = UBound(Rivit)
        For r = 0 To Riveja
            If Left$(Rivit(r), 2) = "--" Then 'Ollaan kahdesti sisennetyllä rivillä tehdään rivistä sisennetty
      If Sis = False Then
          TOUT.WriteLine "      <dl>"
          Sis = True
      End If
      TOUT.WriteLine "        <dl>"
      TOUT.WriteLine "          <dd>"
      TOUT.WriteLine "             " & Trim$(Mid$(Rivit(r), 3))
      TOUT.WriteLine "          </dd>"
      TOUT.WriteLine "        </dl>"
      If r = Riveja Then ' Ollaan viimeisellä rivillä, lopetetaan sisennys
        TOUT.WriteLine "      </dl>"
      End If
     ElseIf Left$(Rivit(r), 1) = "-" Then 'Ollaan sisennetyllä rivillä tehdään rivistä sisennetty
      If Sis = False Then
          TOUT.WriteLine "      <dl>"
          Sis = True
      End If
      TOUT.WriteLine "        <dd>"
      TOUT.WriteLine "           " & Trim$(Mid$(Rivit(r), 2))
      TOUT.WriteLine "        </dd>"
      If r = Riveja Then ' Ollaan viimeisellä rivillä, lopetetaan sisennys
        TOUT.WriteLine "      </dl>"
      End If
    Else
      If r > 0 Then 'Ei ole ensimmäinen rivi
        If Sis Then 'On Aloitettu sisennys aikaisemmin, joten päätetään edellinen sisennys
          TOUT.WriteLine "      </dl>"
        End If
      End If
      If r < Riveja Then  'Ei olla viimeisellä rivillä
        If Left$(Rivit(r + 1), 1) = "-" Then 'Seuraava rivi on sisennetty, tehdään tästä otsikkorivi
          If Rivit(r) = "" Then
            TOUT.WriteLine "      <dl>"
            Sis = True
          Else
            TOUT.WriteLine "      <dl>"
            TOUT.WriteLine "        <dt>"
            TOUT.WriteLine "          " & Rivit(r)
            TOUT.WriteLine "        </dt>"
            Sis = True
          End If
        ElseIf (Rivit(r)) = "" Then
          Sis = False
        Else 'Seuraava rivi ei ole sisennetty, tehdään tästä rivistä normaali kappale
          TOUT.WriteLine "      <p>"
          TOUT.WriteLine "      " & Rivit(r)
          TOUT.WriteLine "      </p>"
          Sis = False
        End If
      ElseIf (Rivit(r)) = "" Then
          Sis = False
      Else
        If (Rivit(r)) = "" Then
            Sis = False
        Else
            'Viimeinen rivi, joka ei ole sisennetty, eikä seuraava rivi voi olla sisennetty, joten tehdään normaali kappale
            TOUT.WriteLine "      <p>"
            TOUT.WriteLine "      " & Rivit(r)
            TOUT.WriteLine "      </p>"
            Sis = False
        End If
      End If
      
     End If
  Next r
End If

  Exit Sub
  
ErrorHandler:
    MsgBox "Virhe tekstin HTML-muuntamisessa: " & Err.Description, vbExclamation
End Sub

'================================================================================
' Funktio: MuutaLinkit
' Tarkoitus: Muuntaa {LOOPTAG}-viittaukset HTML-hyperlinkeiksi
' Parameters:
'   Tieto - Text containing {LOOPTAG} references
' Returns:
'   Text with {LOOPTAG} converted to <a href="LOOPTAG.html">LOOPTAG</a>
'
' Description:
'   Searches for {POS Description} patterns and converts to hyperlinks
'   where POS is loop tag. Supports both {POS} and {POS Description} formats.
'================================================================================
Private Function MuutaLinkit(Tieto As String) As String
    Dim OS As Long, OS2 As Long, OS3 As Long
    Dim POS As String
    Dim DESC As String
    
    On Error GoTo ErrorHandler
    
    MuutaLinkit = Tieto
    Do
        OS = InStr(MuutaLinkit, "{")
        If OS = 0 Then
            Exit Do
        End If
        OS2 = InStr(OS, MuutaLinkit, " ")
        OS3 = InStr(MuutaLinkit, "}")
        If OS2 > OS3 Then OS2 = OS3
        POS = Mid$(MuutaLinkit, OS + 1, OS2 - OS - 1)
        If OS2 < OS3 Then
            DESC = " " & Mid$(MuutaLinkit, OS2 + 1, OS3 - OS2 - 1)
        End If
        MuutaLinkit = Left$(MuutaLinkit, OS - 1) & "<a href=""" & POS & ".html"">" & POS & DESC & "</a>" & Mid$(MuutaLinkit, OS3 + 1)
    Loop
    
    Exit Function
    
ErrorHandler:
    MuutaLinkit = Tieto  ' Return original on error
End Function

'================================================================================
' Aliohjelma: PoimiAsetukset
' Tarkoitus: Lataa luontiasetukset SETTINGS-taulusta
'
' Description:
'   Reads configuration from SETTINGS table including template paths,
'   output directory, and project information for document headers.
'================================================================================
Private Sub PoimiAsetukset()
    Dim Taul As DAO.Recordset
    
    On Error GoTo ErrorHandler
    
    Set Taul = CurrentDb.OpenRecordset("SETTINGS")
    
    ' Tarkistetaan, että SETTINGS-taulussa on edes yksi rivi
    If Taul.EOF Then
        Taul.Close
        Set Taul = Nothing
        MsgBox "SETTINGS-taulukko on tyhjä. Asetusten hakeminen epäonnistui.", vbExclamation, "PoimiAsetukset"
        Exit Sub
    End If
    Taul.MoveFirst
    
    If Suomeksi Then
        POHJA = Nz(Taul.Fields("TEMPLATE_FI"), "")
    Else
        POHJA = Nz(Taul.Fields("TEMPLATE"), "")
    End If
    POLKU = Nz(Taul.Fields("OUTPUTPATH"), "")
    PROJECT_L1 = Nz(Taul.Fields("PROJECT_LINE1"), "")
    PROJECT_L2 = Nz(Taul.Fields("PROJECT_LINE2"), "")
    AUTHOR = Nz(Taul.Fields("AUTHOR"), "")
    REV = Nz(Taul.Fields("REV"), "")
    STATUS = Nz(Taul.Fields("STATUS"), "")
    DOCNO = Nz(Taul.Fields("DOCNO"), "")
    If Right$(POLKU, 1) <> "\" Then POLKU = POLKU & "\"
    
    Taul.Close
    Set Taul = Nothing
    Exit Sub
  
ErrorHandler:
    MsgBox "Virhe asetusten latauksessa: " & Err.Description, vbExclamation
    If Not Taul Is Nothing Then Taul.Close
    Set Taul = Nothing
End Sub

'================================================================================
' Aliohjelma: MuutaTiedot
' Tarkoitus: Täyttää Word-mallin kirjanmerkit piirin/laitteen tiedoilla
'
' Description:
'   Replaces Word template bookmarks with data from current record.
'   Handles headers (_REV, _STAT, _AUTH, etc.) and content fields
'   (_LTAG, _DESCR, _FUNCT, etc.). Generates alarm lists and equipment lists.
'================================================================================
Private Sub MuutaTiedot()
    Dim Tagi As String
    Dim Laitteet  As String
    Dim Halyt As String
    Dim Controlmodes As String
    Dim ADescr As String
    Dim PID_D As Variant
    Dim PID_U As Variant
    
    On Error GoTo ErrorHandler
    
    Tagi = Tiedot.Fields("LOOPTAG")
    If IsNull(Tiedot.Fields("AlarmDescrBL")) Then
'    If Tiedot.Fields("VakioHaly") = "" Then
'      If Suomeksi Then
'        Halyt = "Ei erityisiä hälytyksiä"
'      Else
'        Halyt = "No special alarms"
'      End If
'    Else
'      Halyt = Tiedot.Fields("VakioHaly")
'    End If
  Else
    If IsNull(Tiedot.Fields("HH")) And IsNull(Tiedot.Fields("H")) And IsNull(Tiedot.Fields("LL")) And IsNull(Tiedot.Fields("L")) Then
      If Suomeksi Then
        Halyt = "Ei hälytyksiä"
      Else
        Halyt = "No alarms"
      End If
    Else
      If Suomeksi Then
        ADescr = Nz(Tiedot.Fields("AlarmDescrBL"), "")
      Else
        ADescr = Nz(Tiedot.Fields("AlarmDescrBL_FI"), "")
      End If
      LisaaHaly Tiedot.Fields("HH"), Halyt, "HH: " & ADescr & " > "
      LisaaHaly Tiedot.Fields("H"), Halyt, "H: " & ADescr & " > "
      LisaaHaly Tiedot.Fields("L"), Halyt, "L: " & ADescr & " < "
      LisaaHaly Tiedot.Fields("LL"), Halyt, "LL: " & ADescr & " < "
    End If
  End If

  
  ' Me.LomakePohja.Value tarkistettu ylempänä — tässä otetaan arvo moduulitasolta
  If Samaan = False Or Eka = True Then
    KorvaaOts "[REV]", REV
    KorvaaOts "[PROJECT_LINE1]", PROJECT_L1
    KorvaaOts "[PROJECT_LINE2]", PROJECT_L2
    KorvaaOts "[STATUS]", STATUS
    KorvaaOts "[DOCNO]", DOCNO
    KorvaaOts "[AUTHOR]", AUTHOR
    KorvaaOts "[DATE]", Day(Now()) & "." & Month(Now()) & "." & Year(Now())
    Eka = False
  End If
  Korvaa "[PIDwgNo]", Nz(Tiedot.Fields("PIDwgNo"))
  If Suomeksi Then
    Korvaa "[LOOPDESCR]", UCase$(Tiedot.Fields("DescriptionFI"))
    Korvaa "[LOOPTAG]", Tagi
    Korvaa "[PROCESS]", Nz(Tiedot.Fields("ProcessFI"))
    Korvaa "[FUNCTION]", Tiedot.Fields("FunctionFI")
  Else
    Korvaa "[LOOPDESCR]", UCase$(Tiedot.Fields("Description"))
    Korvaa "[LOOPTAG]", Tagi
    Korvaa "[PROCESS]", Nz(Tiedot.Fields("Process"))
    Korvaa "[FUNCTION]", Nz(Tiedot.Fields("Function"))
  End If
  Korvaa "[LINKS]", Nz(Tiedot.Fields("Linkit"))
  Korvaa "[CHANGES]", Nz(Tiedot.Fields("Muutokset"))
  
  If Me.LomakePohja.Value = 1 Then ' Kysymyksessä piiri
    PID_D = DLookup("FBlock", "CONTROL", "AreaCode = '" & Tiedot.Fields("AreaCode") & "' AND LoopNo='" & Tiedot.Fields("LoopNo") & "' AND FBlock='PID_D'")
    PID_U = DLookup("FBlock", "CONTROL", "AreaCode = '" & Tiedot.Fields("AreaCode") & "' AND LoopNo='" & Tiedot.Fields("LoopNo") & "' AND FBlock='PID_U'")
     
   If IsNull(PID_D) = False Or IsNull(PID_U) = False Then
     If Suomeksi Then
       Korvaa "[ACTION_T]", "Säätimen toiminta"
     Else
       Korvaa "[ACTION_T]", "Action of the controller"
     End If
   End If
   If IsNull(PID_U) = False Then
     If Suomeksi Then
       Korvaa "[ACTION]", "Suora"
     Else
       Korvaa "[ACTION]", "Direct"
     End If
   ElseIf IsNull(PID_D) = False Then
     If Suomeksi Then
       Korvaa "[ACTION]", "Käänteinen"
     Else
       Korvaa "[ACTION]", "Reverse"
     End If
   Else
     POISTA "[ACTION_T]"
     POISTA "[ACTION]"
   End If
   
   If IsNull(Tiedot.Fields("Setpoint")) Or Tiedot.Fields("Setpoint") = "na" Then
     POISTA "[SETPOINT]"
   Else
     If Suomeksi Then
       Korvaa "[SETPOINT]", "Säätimen tyypillinen asetusarvoalue on: " & Tiedot.Fields("SetPoint")
     Else
       Korvaa "[SETPOINT]", "Typical set point range of controller is: " & Tiedot.Fields("SetPoint")
     End If
   End If
  Else
     POISTA "[ACTION_T]"
     POISTA "[ACTION]"
  End If
  If IsNull(Tiedot.Fields("RANGE")) Then
  'If (Tiedot.Fields("RANGE") = "-") Then
    POISTA "[RANGE_T]"
    POISTA "[RANGE]"
  Else
    If Suomeksi Then
      Korvaa "[RANGE_T]", "Mittausalue"
    Else
      Korvaa "[RANGE_T]", "Measuring range"
    End If
    Korvaa "[RANGE]", Tiedot.Fields("RANGE")
  End If
  If IsNull(Tiedot.Fields("Other")) Then
    POISTA "[OTHER_T]"
    POISTA "[OTHER]"
  Else
    If Suomeksi Then
        Korvaa "[OTHER_T]", "Muut"
    Else
        Korvaa "[OTHER_T]", "Other"
    End If
    Korvaa "[OTHER]", Tiedot.Fields("Other")
  End If
  Controlmodes = Nz(Tiedot.Fields("ControlMode"), "")  ' Null kaataisi MuutaTiedot ilman Nz()
  Korvaa2 "[CONTROLMODES]", Controlmodes
  If Suomeksi Then
    Korvaa2 "[INTERLOCKS]", Nz(Tiedot.Fields("InterlocksFI"))
  Else
    Korvaa2 "[INTERLOCKS]", Nz(Tiedot.Fields("Interlocks"))
  End If
  
  
  Korvaa2 "[ALARMS]", Halyt
  
  ' Kerätään laitelistaus (TAGNAME) — Nz() välttää Invalid use of Null -virhen
  Do While Not Tiedot.EOF
    If Tiedot.Fields("LoopTAG") <> Tagi Then Exit Do
    If Laitteet <> "" Then Laitteet = Laitteet & vbCrLf
    If Suomeksi Then
      Laitteet = Laitteet & Nz(Tiedot.Fields("TAGNAME_FI"), "")
    Else
      Laitteet = Laitteet & Nz(Tiedot.Fields("TAGNAME"), "")
    End If
'    If Tiedot.Fields("suffix") <> "-" Then Laitteet = Laitteet & " (" & Tiedot.Fields("suffix") & ")"
'    Laitteet = Laitteet & Tiedot.Fields("maara")
'    If Tiedot.Fields("maara") > 1 Then
'      Laitteet = Laitteet & " pcs)"
'    Else
'      Laitteet = Laitteet & " pc)"
'    End If
    Tiedot.MoveNext
  Loop
  Korvaa2 "[EQUIPMENT]", Laitteet
  
  Exit Sub
  
ErrorHandler:
    MsgBox "Virhe Word-dokumentin tietojen muokkauksessa: " & Err.Description, vbExclamation
End Sub

'================================================================================
' Aliohjelma: KorvaaOts
' Tarkoitus: Korvaa teksti Word-dokumentin otsikossa
' Parameters:
'   VanhaTeksti - Old text to find
'   UusiTeksti - New text to replace with
'================================================================================
Private Sub KorvaaOts(VanhaTeksti As String, UusiTeksti As String)
    Dim rngStory As Variant
    
    On Error Resume Next  ' Header may not exist
    
    UusiTeksti = Replace(UusiTeksti, "{", "")
    UusiTeksti = Replace(UusiTeksti, "}", "")
    oWord.Activate
    
    oTemplate.StoryRanges(wdPrimaryHeaderStory).Select
    With oWord.Selection.Find
        .Replacement.ClearFormatting
        .Text = VanhaTeksti
        .Replacement.Text = UusiTeksti
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
        .Execute Replace:=wdReplaceAll
    End With
End Sub

'================================================================================
' Aliohjelma: Korvaa
' Tarkoitus: Korvaa teksti Word-dokumentin päätekstissä
' Parameters:
'   VanhaTeksti - Old text to find (bookmark)
'   UusiTeksti - New text to replace with
'
' Description:
'   Replaces bookmark text in document. Handles long text (>255 chars) separately
'   due to Word limitations. Strips {braces} from replacement text.
'================================================================================
Private Sub Korvaa(VanhaTeksti As String, UusiTeksti As String)
    Dim rngStory As Variant
    
    On Error Resume Next  ' Continue if bookmark not found
    
    UusiTeksti = Replace(UusiTeksti, "{", "")
    UusiTeksti = Replace(UusiTeksti, "}", "")
    oWord.Activate
    
    If Len(UusiTeksti) > 255 Then
    
      oTemplate.StoryRanges(wdMainTextStory).Select
      With oWord.Selection.Find
          .Replacement.ClearFormatting
          .Text = VanhaTeksti
          .Forward = True
          .Wrap = wdFindContinue
          .Format = False
          .MatchCase = False
          .MatchWholeWord = False
          .MatchWildcards = False
          .MatchSoundsLike = False
          .MatchAllWordForms = False
          .Execute
          If .Found Then
              oWord.Selection.TypeText ""
              oWord.Selection.TypeText UusiTeksti
          End If
      End With
    
   Else
      oTemplate.StoryRanges(wdMainTextStory).Select
      With oWord.Selection.Find
          .Replacement.ClearFormatting
          .Text = VanhaTeksti
          .Replacement.Text = UusiTeksti
          .Forward = True
          .Wrap = wdFindContinue
          .Format = False
          .MatchCase = False
          .MatchWholeWord = False
          .MatchWildcards = False
          .MatchSoundsLike = False
          .MatchAllWordForms = False
          .Execute Replace:=wdReplaceAll
      End With
  End If
End Sub

'================================================================================
' Aliohjelma: Korvaa2
' Tarkoitus: Korvaa teksti Word-dokumentissa säilyttäen kappaleiden vaihdot
' Parameters:
'   VanhaTeksti - Old text to find (bookmark)
'   UusiTeksti - New text with vbCrLf paragraph separators
'
' Description:
'   Replaces bookmark text and converts vbCrLf to Word paragraphs.
'   Strips {braces} from replacement text.
'================================================================================
Private Sub Korvaa2(VanhaTeksti As String, ByVal UusiTeksti As String)
    Dim rngStory As Variant
    Dim Tekstit As Variant
    Dim i As Integer
    
    On Error Resume Next  ' Continue if bookmark not found
    
    UusiTeksti = Replace(UusiTeksti, "{", "")
    UusiTeksti = Replace(UusiTeksti, "}", "")
    Tekstit = Split(UusiTeksti, vbCrLf)
    
    oWord.Activate
    oTemplate.StoryRanges(wdMainTextStory).Select
      With oWord.Selection.Find
          .Text = VanhaTeksti
          .Forward = True
          .Wrap = wdFindContinue
          .Format = False
          .MatchCase = False
          .MatchWholeWord = False
          .MatchWildcards = False
          .MatchSoundsLike = False
          .MatchAllWordForms = False
          .Execute
      End With
      If oWord.Selection.Find.Found Then
        
        If UBound(Tekstit) = -1 Then oWord.Selection.TypeText "-"
        For i = 0 To UBound(Tekstit)
            oWord.Selection.TypeText Tekstit(i)
            If i < UBound(Tekstit) Then oWord.Selection.TypeParagraph
        Next i
      End If

End Sub

'================================================================================
' Aliohjelma: POISTA
' Tarkoitus: Poistaa kirjanmerkin tekstin Word-dokumentista
' Parameters:
'   TEKSTI - Bookmark text to find and delete
'================================================================================
Private Sub POISTA(TEKSTI As String)
    Dim rngStory As Variant
    Dim i As Integer
    
    On Error Resume Next  ' Continue if bookmark not found
    
    oWord.Activate
    oTemplate.StoryRanges(wdMainTextStory).Select
    With oWord.Selection.Find
          .Text = TEKSTI
          .Forward = True
          .Wrap = wdFindContinue
          .Format = False
          .MatchCase = False
          .MatchWholeWord = False
          .MatchWildcards = False
          .MatchSoundsLike = False
          .MatchAllWordForms = False
          .Execute
      End With
      If oWord.Selection.Find.Found Then
        oWord.Selection.Delete , 2
      End If
End Sub

'================================================================================
' Aliohjelma: LisaaHaly
' Tarkoitus: Lisää hälytystasorivi hälytyksien tekstimerkkijonoon (Word-tulostus)
' Parameters:
'   Tieto - Alarm level value
'   Haly - Current alarm text
'   Alku - Alarm prefix text (e.g., "HH: > ")
'================================================================================
Private Sub LisaaHaly(Tieto As Variant, Haly As String, Alku As String)
    On Error Resume Next
    If IsNull(Tieto) Then Exit Sub
    If Haly <> "" Then Haly = Haly & vbCrLf
    Haly = Haly & Alku & Tieto & Tiedot.Fields("Unit")
End Sub

'================================================================================
' Aliohjelma: LisaaHalyHTML
' Tarkoitus: Kirjoittaa hälytystasorivi suoraan HTML-tulostiedostoon
' Parameters:
'   Tieto - Alarm level value
'   Alku - Alarm prefix text (e.g., "HH: > ")
'================================================================================
Private Sub LisaaHalyHTML(Tieto As Variant, Alku As String)
    On Error Resume Next
    If IsNull(Tieto) Then Exit Sub
    TOUT.WriteLine "<dd>"
    TOUT.WriteLine Alku & Tieto & Tiedot.Fields("Unit")
    TOUT.WriteLine "</dd>"
End Sub

'================================================================================
' Aliohjelma: LisaaTeksti
' Tarkoitus: Lisää teksti kursorin kohdalle kohdetekstikenttään
' Parametrit:
'   TEKSTI — Lisättävä teksti
' Huom: SelStart vaatii fokuksen — kietotaan virheensuojauksen sisään (Virhe 2185)
'================================================================================
Private Sub LisaaTeksti(TEKSTI As String)
    Dim Alku As String
    Dim Loppu As String
    
    On Error GoTo ErrorHandler
    
    If KohdeTextBox.Name = "Interlocks" Or KohdeTextBox.Name = "TFunction" Then
        Alku = Left$(Nz(KohdeTextBox.Value, ""), Kursori)
        Loppu = Mid$(Nz(KohdeTextBox.Value, ""), Kursori + 1)
        KohdeTextBox.Value = Alku & TEKSTI & Loppu
        KohdeTextBox.SetFocus
        ' SelStart ja SelLength vaativat fokuksen — estetään Virhe 2185
        On Error Resume Next
        KohdeTextBox.SelStart = Kursori + Len(TEKSTI)
        KohdeTextBox.SelLength = 0
        Kursori = KohdeTextBox.SelStart
        On Error GoTo ErrorHandler
    End If
    Exit Sub
    
ErrorHandler:
    ' Tekstin lisäys epäonnistui hiljaisesti
End Sub

'================================================================================
' Tapahtumat: Tekstiä lisäävät painikkeet
'================================================================================
Private Sub Command32_Click()
    On Error Resume Next
    LisaaTeksti " JA" & vbCrLf
End Sub

Private Sub Command33_Click()
    On Error Resume Next
    LisaaTeksti " TAI" & vbCrLf
End Sub

Private Sub Command34_Click()
    On Error Resume Next
    LisaaTeksti " säädin on vapautettu kun:" & vbCrLf & vbCrLf
End Sub

Private Sub Command35_Click()
    On Error Resume Next
    LisaaTeksti " venttiili on vapautettu kun:" & vbCrLf & vbCrLf
End Sub

Private Sub Command36_Click()
    On Error Resume Next
    LisaaTeksti " on auki rajalla"
End Sub

Private Sub Command37_Click()
    On Error Resume Next
    LisaaTeksti " on kiinni rajalla"
End Sub

Private Sub Command38_Click()
    On Error Resume Next
    LisaaTeksti " käy"
End Sub

Private Sub Command40_Click()
    On Error GoTo ErrorHandler
    
    If MsgBox("Luodaanko uudet moottorit SELOSTUKSET taulukkoon?", vbOKCancel, "LUO MOOTTORI") = vbOK Then
        CurrentDb.Execute "LUOMOOTTORIT"
        Me.Requery
    End If
    Exit Sub
    
ErrorHandler:
    MsgBox "Error creating motor descriptions: " & Err.Description, vbExclamation
End Sub

Private Sub Command60_Click()
    On Error Resume Next
    LisaaTeksti " Moottori on vapautettu kun:" & vbCrLf & vbCrLf
End Sub

Private Sub Command61_Click()
    On Error Resume Next
    LisaaTeksti " , päästohidastettu 5s"
End Sub

Private Sub Command62_Click()
    On Error Resume Next
    LisaaTeksti " , vetohidastettu 5s"
End Sub

Private Sub Command63_Click()
LisaaTeksti " on >" & vbCrLf
End Sub

Private Sub Command64_Click()
LisaaTeksti " on <" & vbCrLf
End Sub

Private Sub Command65_Click()
LisaaTeksti " °C"
End Sub

Private Sub Command66_Click()
LisaaTeksti " µS/cm"
End Sub

Private Sub Command67_Click()
LisaaTeksti " , päästöhidastettu 10s"
End Sub

Private Sub Command68_Click()
LisaaTeksti " , päästöhidastettu 30s"
End Sub

Private Sub Command69_Click()
LisaaTeksti " , päästöhidastettu 60s"
End Sub

Private Sub Command70_Click()
LisaaTeksti " , vetohidastettu 10s"
End Sub

Private Sub Command71_Click()
LisaaTeksti " , vetohidastettu 30s"
End Sub

Private Sub Command72_Click()
LisaaTeksti " , vetohidastettu 60s"
End Sub

Private Sub Command73_Click()
LisaaTeksti " , sends running enable signal"
End Sub

Private Sub Command75_Click()
LisaaTeksti " , on OK"
End Sub

Private Sub Command76_Click()
LisaaTeksti "Ei lukituksia" & vbCrLf & vbCrLf
End Sub

Private Sub Command81_Click()

End Sub

Private Sub Command99_Click()
If KohdeTextBox.Name = "TFunction" Then
    LisaaTeksti vbCrLf & "-Odotusaika: 0 s" & vbCrLf & _
                "-Valvonta-aika: 1 s" & vbCrLf & vbCrLf & _
                "-Toimenpiteet:" & vbCrLf
End If
End Sub

Private Sub Command100_Click()
If KohdeTextBox.Name = "TFunction" Then
    LisaaTeksti vbCrLf & "-Toimenpiteet:" & vbCrLf
End If
End Sub

Private Sub Command101_Click()
If KohdeTextBox.Name = "TFunction" Then
    LisaaTeksti vbCrLf & vbCrLf & "-Ehdot:" & vbCrLf & "-"
End If
End Sub

Private Sub Command102_Click()
If KohdeTextBox.Name = "TFunction" Then
    LisaaTeksti vbCrLf & "<h3>Käsiajolla</h3>" & vbCrLf
End If
End Sub

Private Sub Command104_Click()
If KohdeTextBox.Name = "TFunction" Then
    LisaaTeksti vbCrLf & vbCrLf & "<h4>Askel nro. " & Me.Text105.Value & ": " & Me.Text107.Value & "</h4>" & _
                vbCrLf & "-Odotusaika: 0 s" & vbCrLf & _
                "-Valvonta-aika: 1 s" & vbCrLf & vbCrLf & _
                "-Toimenpiteet:" & vbCrLf & _
                "-"

End If
End Sub

Private Sub Command103_Click()
If KohdeTextBox.Name = "TFunction" Then
    LisaaTeksti vbCrLf & "<h3>Automaatilla</h3>" & vbCrLf
End If
End Sub


Private Sub Form_AfterUpdate()
    On Error Resume Next
'  cUI.Requery
    'Me.Refresh
End Sub

Private Sub Form_Load()
    On Error Resume Next
'  CLoppuun.Value = False
End Sub

'================================================================================
' Tapahtumat: Tekstikenttien fokusseuranta ja kursorin paikka
'================================================================================
Private Sub INTERLOCKS_FROM__Click()
    On Error Resume Next
    LisaaTeksti "Lukitsevat tiedot:" & vbCrLf & vbCrLf
End Sub

Private Sub Interlocks_GotFocus()
    On Error Resume Next
    Set KohdeTextBox = Me.Interlocks
End Sub

Private Sub Interlocks_KeyUp(KeyCode As Integer, Shift As Integer)
    On Error Resume Next
    Kursori = Me.Interlocks.SelStart
End Sub

Private Sub Interlocks_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Error Resume Next
    Kursori = Me.Interlocks.SelStart
End Sub

Private Sub INTERLOCKS_TO__Click()
    On Error Resume Next
    LisaaTeksti "Liittyvät tiedot:" & vbCrLf & vbCrLf
End Sub

Private Sub Linkit_GotFocus()
    On Error Resume Next
    Set KohdeTextBox = Me.Linkit
End Sub

Private Sub Linkit_KeyUp(KeyCode As Integer, Shift As Integer)
    On Error Resume Next
    Kursori = Me.Linkit.SelStart
End Sub

Private Sub Linkit_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Error Resume Next
    Kursori = Me.Linkit.SelStart
End Sub

Private Sub LomakePohja_AfterUpdate()
    On Error GoTo ErrorHandler
    
    ' Vaihdetaan tietolähde valitun pohjan mukaan
    If Me.LomakePohja.Value = 1 Then
        Me.RecordSource = "MUOKKAUSPIIRIT"
        Me.Halyt.Visible = True
    Else
        Me.RecordSource = "MUOKKAUSMOOTTORIT"
        Me.Halyt.Visible = False
    End If
    Me.Requery
    Exit Sub
    
ErrorHandler:
    MsgBox "Virhe pohjan vaihtamisessa: " & Err.Description, vbExclamation
End Sub

Private Sub TFunction_GotFocus()
    On Error Resume Next
    Set KohdeTextBox = Me.TFunction
End Sub

Private Sub TFunction_KeyUp(KeyCode As Integer, Shift As Integer)
    On Error Resume Next
    Kursori = Me.TFunction.SelStart
End Sub

Private Sub TFunction_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Error Resume Next
    Kursori = Me.TFunction.SelStart
End Sub
