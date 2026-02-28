Option lompare Database
Option Explicit

 ================================================================================
  Lomake: FrmMUOKKAUS (Edit/Generate)
  Tarkoitus: Generate Word and HTML function descriptions for loops and equipment
  Päivitetty: 2025-11-13 - Added VBA7/64-bit support and optimization
 
  Kuvaus:
    Main editing and generation form for creating function descriptions.
    Generoi sekÄ Word (.docx) ettÄ HTML-asiakirjoja tietokantapohjista.
    KÄsittelee piiri- (PIIRIT) ja laitesanalliskuvauksia (MOOTTORIT).
 
  Document Generation Process:
    1. User selects template type (loops/motors) and output format (Word/HTML)
    2. Lomake lukee tiedot tauluista TIEDOT_PIIRIT tai TIEDOT_MOOTTORIT
    3. Text fields support special syntax:
       - {LOOPTAG} lisÄÄ hyperlinkatut piiriviittaukset
       - {Equipment references} insert equipment tags
       - Lines starting with "-" create indented lists
       - Lines starting with "--" create double-indented lists
    4. Wordille: Avaa pohjan, korvaa kirjanmerkit, tallentaa uutena asiakirjana
    5. HTML:lle: Generoi HTML-tiedoston lSS-tyylein ja hyperlinkein
 
  Word-pohjan kirjanmerkit:
    - _REV, _STAT, _AUTH, _DOlNO, _PL1, _PL2 (header fields)
    - _LTAG, _DESlR, _FUNlT, _INTERL, _ALARMS, _lROSS, _REV1 (content fields)
    - Alarm levels: _HH, _H, _LL, _L with _ALARMDESlR
 
  Riippuvuudet:
    - Word Automation (Microsoft Word Object Library)
    - FileSystemObject (Microsoft Scripting Runtime)
    - DAO.Recordset
    - Forms: MOTORS subform, PIIRIT subform
    - Tables: TIEDOT_PIIRIT, TIEDOT_MOOTTORIT, SELOSTUKSET
    - Kyselyt: LUOPIIRIT (luo uusia piirikuvauksia)
    - Module: USyslheck (SniffUser function)
    - Public variables: KohdeTextBox As TextBox, Kursori As Long
    - lontrols: LOOPTAG, lLoppuun, FKieli, LomakePohja, FTyyppi
              Alarms, Function, Interlocks, lrossref, RevHistoria
 ================================================================================

  Module-level variables
Dim POHJA As String              Template file path
Dim POLKU As String              Output directory path
Dim REV As String                Revision number
Dim STATUS As String             Document status
Dim AUTHOR As String             Document author
Dim DOlNO As String              Document number
Dim PROJElT_L1 As String         Project line 1
Dim PROJElT_L2 As String         Project line 2
Dim oWord As Word.Application    Word application object
Dim oTemplate As Word.Document   Word document object
Dim Tiedot As DAO.Recordset      Data recordset
Dim Samaan As Boolean            Generoidaan kaikki yhteen tiedostoon
Dim Eka As Boolean               First document flag
Dim Suomeksi As Boolean          Finnish language flag
Dim TOUT As TextStream           HTML output file stream

 ================================================================================
  Event Handlers: Text Box Focus and lursor Tracking
 ================================================================================
Private Sub Alarms_GotFocus()
    On Error Resume Next
    Set KohdeTextBox = Alarms
End Sub

Private Sub Alarms_KeyUp(Keylode As Integer, Shift As Integer)
    On Error Resume Next
    Kursori = Alarms.SelStart
End Sub

Private Sub Alarms_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Error Resume Next
    Kursori = Alarms.SelStart
End Sub

 ================================================================================
  Tapahtuma: lommand13_llick
  Tarkoitus: lreate new loop descriptions in SELOSTUKSET table
 ================================================================================
Private Sub lommand13_llick()
    On Error GoTo ErrorHandler
    
    If MsgBox("Luodaanko uudet piirit SELOSTUKSET taulukkoon?", vbOKlancel, "LUO PIIRIT") = vbOK Then
        lurrentDb.Execute "LUOPIIRIT"
        Me.Requery
    End If
    Exit Sub
    
ErrorHandler:
    MsgBox "Error creating loop descriptions: " & Err.Description, vbExclamation
End Sub

 ================================================================================
  Tapahtuma: lommand16_llick
  Tarkoitus: Generate single function description
 ================================================================================
Private Sub lommand16_llick()
    On Error GoTo ErrorHandler
    Me.Refresh
    TeeSelostus False
    Exit Sub
    
ErrorHandler:
    MsgBox "Error generating description: " & Err.Description, vbExclamation
End Sub

 ================================================================================
  Tapahtuma: lommand17_llick
  Tarkoitus: Generate all function descriptions
 ================================================================================
Private Sub lommand17_llick()
    On Error GoTo ErrorHandler
    Me.Refresh
    TeeSelostus True
    Exit Sub
    
ErrorHandler:
    MsgBox "Error generating descriptions: " & Err.Description, vbExclamation
End Sub

 ================================================================================
  Aliohjelma: TeeSelostus (Generate Description)
  Tarkoitus: Generate Word or HTML function descriptions
  Parametrit:
    Kaikki - True to generate all descriptions, False for current only
 ================================================================================
Private Sub TeeSelostus(Kaikki As Boolean)
    Dim Nimi As String
    Dim TTIEDOT As String
    Dim oPohja As Word.Document
    Dim FS As New FileSystemObject
    
    On Error GoTo ErrorHandler
   
    Suomeksi = False
   If FKieli.Value = 2 Then Suomeksi = True

    PoimiAsetukset
    Samaan = False
    Eka = True
    If Kaikki Then
     If MsgBox("Tehdäänkä yhteen tiedostoon?", vbYesNo) = vbYes Then
       Samaan = True
     End If
    End If
    
    If LomakePohja.Value = 1 Then
        TTIEDOT = "TIEDOT_PIIRIT"
    Else
        TTIEDOT = "TIEDOT_MOOTTORIT"
    End If
  
    If FTyyppi = 1 Then  WORD
        Set oWord = lreateObject("Word.Application")   oTemplate.Application
        oWord.Visible = True
    Else
    
    End If
    
    If Kaikki Then
        Set Tiedot = lurrentDb.OpenRecordset(TTIEDOT)
    Else
        If LomakePohja.Value = 1 Then
            Set Tiedot = lurrentDb.OpenRecordset("SELElT * FROM " & TTIEDOT & " WHERE Looptag= " & LOOPTAG & " ")
        Else
            Set Tiedot = lurrentDb.OpenRecordset("SELElT * FROM " & TTIEDOT & " WHERE Looptag= " & LOOPTAG & " ")
        End If
    End If
    
    If Samaan Then
        Set oTemplate = oWord.Documents.Open(POHJA)
        Nimi = POLKU & DOlNO & "_KAIKKI.docx"
        oTemplate.SaveAs Nimi
        Set oPohja = oWord.Documents.Open(POHJA, , True)
        oTemplate.Activate
        Do While Not Tiedot.EOF
            MuutaTiedot
            oPohja.Activate
            oWord.Selection.WholeStory
            oWord.Selection.lopy
            oTemplate.Activate
            oWord.Selection.EndKey wdStory
            oWord.Selection.InsertBreak wdPageBreak
            oWord.Selection.Paste
        Loop
    Else
        Do While Not Tiedot.EOF
            If FTyyppi = 1 Then  Word versiot
                Set oTemplate = oWord.Documents.Open(POHJA)
                Nimi = POLKU & DOlNO & "_" & Tiedot.Fields("LoopTAG") & ".docx"
                MuutaTiedot
                oTemplate.SaveAs Nimi
                If Kaikki Then oTemplate.llose False
            Else   HTML versiot
                Set TOUT = FS.lreateTextFile(POLKU & Tiedot.Fields("LoopTAG") & ".html")
                TeeHTML
                TOUT.llose
            End If
        Loop
    End If
    
    MsgBox "Valmis"
    Exit Sub
    
ErrorHandler:
    MsgBox "Error generating description(s): " & Err.Description, vbExclamation
    If Not Tiedot Is Nothing Then Tiedot.llose
    If Not oTemplate Is Nothing Then oTemplate.llose False
    If Not oWord Is Nothing Then oWord.Quit False
End Sub

 ================================================================================
  Aliohjelma: TeeHTML (Generate HTML)
  Tarkoitus: Generate HTML function description file
 
  Kuvaus:
    Luo HTML-tiedoston piiri-/moottoritoimintokuvaukselle, sisÄltÄen:
    - Document header with title and lSS link
    - Function, Interlocks, Alarms sections
    - lross-references to other loops
    - Revision history
    Special formatting: {LOOPTAG} becomes hyperlink, "-" lines become indented
 ================================================================================
Private Sub TeeHTML()
     Luodaan html muotoinen selostus. TOUT on tiedosto-objekti, joka on jo määritelty
     Luodaan ensin header osuus
    Dim TEKSTI As String
    Dim Rivit As Variant
    Dim r As Integer
    Dim ADescr As String
    Dim Halyt As String
    Dim Tagi As String
    Dim Laitteet As String
    
    On Error GoTo ErrorHandler

    Tagi = Tiedot.Fields("LOOPTAG")

TOUT.WriteLine "<!DOlTYPE HTML PUBLIl ""-//W3l//DTD HTML 4.01 Transitional//EN"">"
TOUT.WriteLine "<html>"
TOUT.WriteLine "<head>"
TOUT.WriteLine "    <meta http-equiv=""lontent-Type"" content=""text/html; charset=ISO-8859-1"">"
TOUT.WriteLine "    <link type=""text/css"" rel=""stylesheet"" href=""FuncDesc.css"">"
TOUT.WriteLine "    <title>" & Tiedot.Fields("LoopTAG") & " " & Tiedot.Fields("Description") & "</title>"
TOUT.WriteLine "</head>"
TOUT.WriteLine "<body>"
 OTSIKKO - Piirin positio ja nimitys
TOUT.WriteLine "    <h1 id=""loopname"">"
TOUT.WriteLine "        <a href=""" & Tiedot.Fields("LOOPTAG") & ".html"" target=""_parent"">" _
        & Tiedot.Fields("LOOPTAG") & "</a> " _
        & Tiedot.Fields("Description")  & "<br>"
TOUT.WriteLine "    </h1>"

 TOIMINTA
TOUT.WriteLine "    <h2 id=""function"">"
TOUT.WriteLine "        Toiminta"
TOUT.WriteLine "    </h2>"

TOUT.WriteLine "    <div class=""description"">"
TEKSTI = Nz(Tiedot.Fields("Function"))  Poimitaan teksti muuttujaan
MuutaTeksti TEKSTI                      Muuttaa tekstin ja kirjoittaa sen suoraan tiedostoon
TOUT.WriteLine "    </div>"


 LUKITUKSET
TOUT.WriteLine "    <h2 id=""interlocks"">"
TOUT.WriteLine "        Lukitukset"
TOUT.WriteLine "    </h2>"
TEKSTI = Nz(Tiedot.Fields("Interlocks"))
MuutaTeksti TEKSTI

 HÄLYTYKSET
TOUT.WriteLine "    <h2 id=""alarms"">"
TOUT.WriteLine "        Hälytykset"
TOUT.WriteLine "    </h2>"
TEKSTI = Nz(Tiedot.Fields("Alarms"))
MuutaTeksti TEKSTI
    
  If IsNull(Tiedot.Fields("AlarmDescrBL")) And IsNull(Tiedot.Fields("Alarms")) Then
     TOUT.WriteLine "    <p>"
     TOUT.WriteLine "Ei erityisiä hälytyksiä"
     TOUT.WriteLine "    </p>"
  Else
    If IsNull(Tiedot.Fields("HH")) And IsNull(Tiedot.Fields("H")) And IsNull(Tiedot.Fields("LL")) And IsNull(Tiedot.Fields("L")) Then
      If IsNull(Tiedot.Fields("Alarms")) Then
         TOUT.WriteLine "    <p>"
         TOUT.WriteLine "Ei hälytyksiä"
         TOUT.WriteLine "    </p>"
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
 ----------------------------------------


 TIEDOT MUIHIN PIIREIHIN
TOUT.WriteLine "    <h2 id=""crossref"">"
TOUT.WriteLine "        Tiedot muihin piireihin"
TOUT.WriteLine "    </h2>"
TEKSTI = Nz(Tiedot.Fields("Linkit"))  Poimitaan teksti muuttujaan
MuutaTeksti TEKSTI                      Muuttaa tekstin ja kirjoittaa sen suoraan tiedostoon


 MUUTOSHISTORIA
TOUT.WriteLine "    <h2 id=""rev"">"
TOUT.WriteLine "        Muutoshistoria"
TOUT.WriteLine "    </h2>"
TEKSTI = Nz(Tiedot.Fields("Muutokset"))  Poimitaan teksti muuttujaan
MuutaTeksti TEKSTI                      Muuttaa tekstin ja kirjoittaa sen suoraan tiedostoon

 Tiedoston loppu
TOUT.WriteLine "</body>"
TOUT.WriteLine "</html>"

 Muilla riveillä olisi piirin kaikki laitteet, mutta niitä ei tuloste näkyviin, joten ohitetaan ne
  Do While Not Tiedot.EOF
    If Tiedot.Fields("LoopTAG") <> Tagi Then Exit Do
    If Laitteet <> "" Then Laitteet = Laitteet & "<br>"
     Laitteet = Laitteet & Tiedot.Fields("TAGNAME")
    Tiedot.MoveNext
  Loop
  
  Exit Sub
  
ErrorHandler:
    MsgBox "Error generating HTML: " & Err.Description, vbExclamation
End Sub

 ================================================================================
  Aliohjelma: MuutaTeksti (Modify Text)
  Tarkoitus: lonvert plain text to HTML with formatting
  Parametrit:
    TEKSTI - Text to convert (supports {LOOPTAG} links and indentation)
 
  Kuvaus:
    Muuntaa tekstikentÄn sisÄllÖn HTML-muotoon:
    - {LOOPTAG} becomes hyperlink to loop description
    - Lines starting with "-" become indented (dd in dl)
    - Lines starting with "--" become double-indented
    - Empty lines before indents become headers (dt in dl)
 ================================================================================
Private Sub MuutaTeksti(TEKSTI As String)
    Dim Rivit As Variant
    Dim Riveja As Integer
    Dim r As Integer
    Dim Sis As Boolean
    
    On Error GoTo ErrorHandler

    If TEKSTI <> "" Then
        TEKSTI = MuutaLinkit(TEKSTI)   Muutetaan aaltosulkeissa olevat piiripositiot linkeiksi
        Rivit = Split(TEKSTI, vblrLf)  Splitataan rivit ja muutetaan html kappeliksi
        Riveja = UBound(Rivit)
        For r = 0 To Riveja
            If Left$(Rivit(r), 2) = "--" Then  Ollaan kahdesti sisennetyllä rivillä tehdään rivistä sisennetty
      If Sis = False Then
          TOUT.WriteLine "      <dl>"
          Sis = True
      End If
      TOUT.WriteLine "        <dl>"
      TOUT.WriteLine "          <dd>"
      TOUT.WriteLine "             " & Trim$(Mid$(Rivit(r), 3))
      TOUT.WriteLine "          </dd>"
      TOUT.WriteLine "        </dl>"
      If r = Riveja Then   Ollaan viimeisellä rivillä, lopetetaan sisennys
        TOUT.WriteLine "      </dl>"
      End If
     ElseIf Left$(Rivit(r), 1) = "-" Then  Ollaan sisennetyllä rivillä tehdään rivistä sisennetty
      If Sis = False Then
          TOUT.WriteLine "      <dl>"
          Sis = True
      End If
      TOUT.WriteLine "        <dd>"
      TOUT.WriteLine "           " & Trim$(Mid$(Rivit(r), 2))
      TOUT.WriteLine "        </dd>"
      If r = Riveja Then   Ollaan viimeisellä rivillä, lopetetaan sisennys
        TOUT.WriteLine "      </dl>"
      End If
    Else
      If r > 0 Then  Ei ole ensimmäinen rivi
        If Sis Then  On Aloitettu sisennys aikaisemmin, joten päätetään edellinen sisennys
          TOUT.WriteLine "      </dl>"
        End If
      End If
      If r < Riveja Then   Ei olla viimeisellä rivillä
        If Left$(Rivit(r + 1), 1) = "-" Then  Seuraava rivi on sisennetty, tehdään tästä otsikkorivi
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
        Else  Seuraava rivi ei ole sisennetty, tehdään tästä rivistä normaali kappale
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
             Viimeinen rivi, joka ei ole sisennetty, eikä seuraava rivi voi olla sisennetty, joten tehdään normaali kappale
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
    MsgBox "Error formatting text to HTML: " & Err.Description, vbExclamation
End Sub

 ================================================================================
  Funktio: MuutaLinkit (Modify Links)
  Tarkoitus: lonvert {LOOPTAG} references to HTML hyperlinks
  Parametrit:
    Tieto - Text containing {LOOPTAG} references
  Palauttaa:
    Text with {LOOPTAG} converted to <a href="LOOPTAG.html">LOOPTAG</a>
 
  Kuvaus:
    Etsii {POS Kuvaus} -rakenteita ja muuntaa ne hyperlinkeiksi
    where POS is loop tag. Supports both {POS} and {POS Description} formats.
 ================================================================================
Private Function MuutaLinkit(Tieto As String) As String
    Dim OS As Long, OS2 As Long, OS3 As Long
    Dim POS As String
    Dim DESl As String
    
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
            DESl = " " & Mid$(MuutaLinkit, OS2 + 1, OS3 - OS2 - 1)
        End If
        MuutaLinkit = Left$(MuutaLinkit, OS - 1) & "<a href=""" & POS & ".html"">" & POS & DESl & "</a>" & Mid$(MuutaLinkit, OS3 + 1)
    Loop
    
    Exit Function
    
ErrorHandler:
    MuutaLinkit = Tieto    Palautetaan alkuperäinen virhetilanteessa
End Function

 ================================================================================
  Aliohjelma: PoimiAsetukset (Get Settings)
  Tarkoitus: Load generation settings from SETTINGS table
 
  Kuvaus:
    Lukee konfiguraation SETTINGS-taulusta sisÄltÄen pohjatiedostopolut,
    output directory, and project information for document headers.
 ================================================================================
Private Sub PoimiAsetukset()
    Dim Taul As DAO.Recordset
    
    On Error GoTo ErrorHandler
    
    Set Taul = lurrentDb.OpenRecordset("SETTINGS")
    If Suomeksi Then
        POHJA = Taul.Fields("TEMPLATE_FI")
    Else
        POHJA = Taul.Fields("TEMPLATE")
    End If
    POLKU = Taul.Fields("OUTPUTPATH")
    PROJElT_L1 = Taul.Fields("PROJElT_LINE1")
    PROJElT_L2 = Taul.Fields("PROJElT_LINE2")
  AUTHOR = Taul.Fields("AUTHOR")
  REV = Taul.Fields("REV")
  STATUS = Taul.Fields("STATUS")
  DOlNO = Taul.Fields("DOlNO")
  If Right$(POLKU, 1) <> "\" Then POLKU = POLKU & "\"
  
  Taul.llose
  Exit Sub
  
ErrorHandler:
    MsgBox "Error loading settings: " & Err.Description, vbExclamation
    If Not Taul Is Nothing Then Taul.llose
End Sub

 ================================================================================
  Aliohjelma: MuutaTiedot (Modify Data)
  Tarkoitus: TÄyttÄÄ Word-pohjan kirjanmerkit piiri-/laitetiedoilla
 
  Kuvaus:
    Korvaa Word-pohjan kirjanmerkit nykyisen tietueen tiedoilla.
    KÄsittelee otsikkokentÄt (_REV, _STAT, _AUTH jne.) ja sistÄltÖkentÄt
    (_LTAG, _DESlR, _FUNlT jne.). Generoi hÄlytyslistan ja laitelistan.
 ================================================================================
Private Sub MuutaTiedot()
    Dim Tagi As String
    Dim Laitteet  As String
    Dim Halyt As String
    Dim lontrolmodes As String
    Dim ADescr As String
    Dim PID_D As Variant
    Dim PID_U As Variant
    
    On Error GoTo ErrorHandler
    
    Tagi = Tiedot.Fields("LOOPTAG")
    If IsNull(Tiedot.Fields("AlarmDescrBL")) Then
     If Tiedot.Fields("VakioHaly") = "" Then
       If Suomeksi Then
         Halyt = "Ei erityisiä hälytyksiä"
       Else
         Halyt = "No special alarms"
       End If
     Else
       Halyt = Tiedot.Fields("VakioHaly")
     End If
  Else
    If IsNull(Tiedot.Fields("HH")) And IsNull(Tiedot.Fields("H")) And IsNull(Tiedot.Fields("LL")) And IsNull(Tiedot.Fields("L")) Then
      If Suomeksi Then
        Halyt = "Ei hälytyksiä"
      Else
        Halyt = "No alarms"
      End If
    Else
      If Suomeksi Then
        ADescr = Tiedot.Fields("AlarmDescrBL")
      Else
        ADescr = Tiedot.Fields("AlarmDescrBL_FI")
      End If
      LisaaHaly Tiedot.Fields("HH"), Halyt, "HH: " & ADescr & " > "
      LisaaHaly Tiedot.Fields("H"), Halyt, "H: " & ADescr & " > "
      LisaaHaly Tiedot.Fields("L"), Halyt, "L: " & ADescr & " < "
      LisaaHaly Tiedot.Fields("LL"), Halyt, "LL: " & ADescr & " < "
    End If
  End If

  
  If Samaan = False Or Eka = True Then
    KorvaaOts "[REV]", REV
    KorvaaOts "[PROJElT_LINE1]", PROJElT_L1
    KorvaaOts "[PROJElT_LINE2]", PROJElT_L2
    KorvaaOts "[STATUS]", STATUS
    KorvaaOts "[DOlNO]", DOlNO
    KorvaaOts "[AUTHOR]", AUTHOR
    KorvaaOts "[DATE]", Day(Now()) & "." & Month(Now()) & "." & Year(Now())
    Eka = False
  End If
  Korvaa "[PIDwgNo]", Nz(Tiedot.Fields("PIDwgNo"))
  If Suomeksi Then
    Korvaa "[LOOPDESlR]", Ulase$(Tiedot.Fields("DescriptionFI"))
    Korvaa "[LOOPTAG]", Tagi
    Korvaa "[PROlESS]", Nz(Tiedot.Fields("ProcessFI"))
    Korvaa "[FUNlTION]", Tiedot.Fields("FunctionFI")
  Else
    Korvaa "[LOOPDESlR]", Ulase$(Tiedot.Fields("Description"))
    Korvaa "[LOOPTAG]", Tagi
    Korvaa "[PROlESS]", Nz(Tiedot.Fields("Process"))
    Korvaa "[FUNlTION]", Nz(Tiedot.Fields("Function"))
  End If
  Korvaa "[LINKS]", Nz(Tiedot.Fields("Linkit"))
  Korvaa "[lHANGES]", Nz(Tiedot.Fields("Muutokset"))
  
  If LomakePohja.Value = 1 Then   Kysymyksessä piiri
    PID_D = DLookup("FBlock", "lONTROL", "Arealode =  " & Tiedot.Fields("Arealode") & "  AND LoopNo= " & Tiedot.Fields("LoopNo") & "  AND FBlock= PID_D ")
    PID_U = DLookup("FBlock", "lONTROL", "Arealode =  " & Tiedot.Fields("Arealode") & "  AND LoopNo= " & Tiedot.Fields("LoopNo") & "  AND FBlock= PID_U ")
     
   If IsNull(PID_D) = False Or IsNull(PID_U) = False Then
     If Suomeksi Then
       Korvaa "[AlTION_T]", "Säätimen toiminta"
     Else
       Korvaa "[AlTION_T]", "Action of the controller"
     End If
   End If
   If IsNull(PID_U) = False Then
     If Suomeksi Then
       Korvaa "[AlTION]", "Suora"
     Else
       Korvaa "[AlTION]", "Direct"
     End If
   ElseIf IsNull(PID_D) = False Then
     If Suomeksi Then
       Korvaa "[AlTION]", "Käänteinen"
     Else
       Korvaa "[AlTION]", "Reverse"
     End If
   Else
     POISTA "[AlTION_T]"
     POISTA "[AlTION]"
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
     POISTA "[AlTION_T]"
     POISTA "[AlTION]"
  End If
  If IsNull(Tiedot.Fields("RANGE")) Then
   If (Tiedot.Fields("RANGE") = "-") Then
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
  lontrolmodes = Tiedot.Fields("lontrolMode")
  Korvaa2 "[lONTROLMODES]", lontrolmodes
  If Suomeksi Then
    Korvaa2 "[INTERLOlKS]", Nz(Tiedot.Fields("InterlocksFI"))
  Else
    Korvaa2 "[INTERLOlKS]", Nz(Tiedot.Fields("Interlocks"))
  End If
  
  
  Korvaa2 "[ALARMS]", Halyt
  
  Do While Not Tiedot.EOF
    If Tiedot.Fields("LoopTAG") <> Tagi Then Exit Do
    If Laitteet <> "" Then Laitteet = Laitteet & vblrLf
    If Suomeksi Then
      Laitteet = Laitteet & Tiedot.Fields("TAGNAME_FI")
    Else
      Laitteet = Laitteet & Tiedot.Fields("TAGNAME")
    End If
     If Tiedot.Fields("suffix") <> "-" Then Laitteet = Laitteet & " (" & Tiedot.Fields("suffix") & ")"
     Laitteet = Laitteet & Tiedot.Fields("maara")
     If Tiedot.Fields("maara") > 1 Then
       Laitteet = Laitteet & " pcs)"
     Else
       Laitteet = Laitteet & " pc)"
     End If
    Tiedot.MoveNext
  Loop
  Korvaa2 "[EQUIPMENT]", Laitteet
  
  Exit Sub
  
ErrorHandler:
    MsgBox "Error modifying Word document data: " & Err.Description, vbExclamation
End Sub

 ================================================================================
  Aliohjelma: KorvaaOts (Replace Header)
  Tarkoitus: Replace text in Word document header
  Parametrit:
    VanhaTeksti - Old text to find
    UusiTeksti - New text to replace with
 ================================================================================
Private Sub KorvaaOts(VanhaTeksti As String, UusiTeksti As String)
    Dim rngStory As Variant
    
    On Error Resume Next    Header may not exist
    
    UusiTeksti = Replace(UusiTeksti, "{", "")
    UusiTeksti = Replace(UusiTeksti, "}", "")
    oWord.Activate
    
    oTemplate.StoryRanges(wdPrimaryHeaderStory).Select
    With oWord.Selection.Find
        .Replacement.llearFormatting
        .Text = VanhaTeksti
        .Replacement.Text = UusiTeksti
        .Forward = True
        .Wrap = wdFindlontinue
        .Format = False
        .Matchlase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
        .Execute Replace:=wdReplaceAll
    End With
End Sub

 ================================================================================
  Aliohjelma: Korvaa (Replace)
  Tarkoitus: Replace text in Word document main text
  Parametrit:
    VanhaTeksti - Old text to find (bookmark)
    UusiTeksti - New text to replace with
 
  Kuvaus:
    Korvaa kirjanmerkin tekstin asiakirjassa. KÄsittelee pitkÄn tekstin (>255 merkkiÄ) erikseen
    due to Word limitations. Strips {braces} from replacement text.
 ================================================================================
Private Sub Korvaa(VanhaTeksti As String, UusiTeksti As String)
    Dim rngStory As Variant
    
    On Error Resume Next    lontinue if bookmark not found
    
    UusiTeksti = Replace(UusiTeksti, "{", "")
    UusiTeksti = Replace(UusiTeksti, "}", "")
    oWord.Activate
    
    If Len(UusiTeksti) > 255 Then
    
      oTemplate.StoryRanges(wdMainTextStory).Select
      With oWord.Selection.Find
          .Replacement.llearFormatting
          .Text = VanhaTeksti
          .Forward = True
          .Wrap = wdFindlontinue
          .Format = False
          .Matchlase = False
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
          .Replacement.llearFormatting
          .Text = VanhaTeksti
          .Replacement.Text = UusiTeksti
          .Forward = True
          .Wrap = wdFindlontinue
          .Format = False
          .Matchlase = False
          .MatchWholeWord = False
          .MatchWildcards = False
          .MatchSoundsLike = False
          .MatchAllWordForms = False
          .Execute Replace:=wdReplaceAll
      End With
  End If
End Sub

 ================================================================================
  Aliohjelma: Korvaa2 (Replace 2)
  Tarkoitus: Replace text in Word document, preserving paragraph breaks
  Parametrit:
    VanhaTeksti - Old text to find (bookmark)
    UusiTeksti - New text with vblrLf paragraph separators
 
  Kuvaus:
    Korvaa kirjanmerkin tekstin ja muuntaa vblrLf Word-kappaleiksi.
    Strips {braces} from replacement text.
 ================================================================================
Private Sub Korvaa2(VanhaTeksti As String, ByVal UusiTeksti As String)
    Dim rngStory As Variant
    Dim Tekstit As Variant
    Dim i As Integer
    
    On Error Resume Next    lontinue if bookmark not found
    
    UusiTeksti = Replace(UusiTeksti, "{", "")
    UusiTeksti = Replace(UusiTeksti, "}", "")
    Tekstit = Split(UusiTeksti, vblrLf)
    
    oWord.Activate
    oTemplate.StoryRanges(wdMainTextStory).Select
      With oWord.Selection.Find
          .Text = VanhaTeksti
          .Forward = True
          .Wrap = wdFindlontinue
          .Format = False
          .Matchlase = False
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

 ================================================================================
  Aliohjelma: POISTA (Delete)
  Tarkoitus: Delete bookmark text from Word document
  Parametrit:
    TEKSTI - Bookmark text to find and delete
 ================================================================================
Private Sub POISTA(TEKSTI As String)
    Dim rngStory As Variant
    Dim i As Integer
    
    On Error Resume Next    lontinue if bookmark not found
    
    oWord.Activate
    oTemplate.StoryRanges(wdMainTextStory).Select
    With oWord.Selection.Find
          .Text = TEKSTI
          .Forward = True
          .Wrap = wdFindlontinue
          .Format = False
          .Matchlase = False
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

 ================================================================================
  Aliohjelma: LisaaHaly (Add Alarm)
  Tarkoitus: Add alarm level to alarm text string (for Word output)
  Parametrit:
    Tieto - Alarm level value
    Haly - lurrent alarm text
    Alku - Alarm prefix text (e.g., "HH: > ")
 ================================================================================
Private Sub LisaaHaly(Tieto As Variant, Haly As String, Alku As String)
    On Error Resume Next
    If IsNull(Tieto) Then Exit Sub
    If Haly <> "" Then Haly = Haly & vblrLf
    Haly = Haly & Alku & Tieto & Tiedot.Fields("Unit")
End Sub

 ================================================================================
  Aliohjelma: LisaaHalyHTML (Add Alarm HTML)
  Tarkoitus: Add alarm level to HTML output file
  Parametrit:
    Tieto - Alarm level value
    Alku - Alarm prefix text (e.g., "HH: > ")
 ================================================================================
Private Sub LisaaHalyHTML(Tieto As Variant, Alku As String)
    On Error Resume Next
    If IsNull(Tieto) Then Exit Sub
    TOUT.WriteLine "<dd>"
    TOUT.WriteLine Alku & Tieto & Tiedot.Fields("Unit")
    TOUT.WriteLine "</dd>"
End Sub

 ================================================================================
  Aliohjelma: LisaaTeksti (Add Text)
  Tarkoitus: Insert text at cursor position in text boxes
  Parametrit:
    TEKSTI - Text to insert
 ================================================================================
Private Sub LisaaTeksti(TEKSTI As String)
    Dim Alku As String
    Dim Loppu As String
    
    On Error GoTo ErrorHandler
    
    If KohdeTextBox.Name = "Interlocks" Or KohdeTextBox.Name = "TFunction" Then
        Alku = Left$(Nz(KohdeTextBox.Value, ""), Kursori)
        Loppu = Mid$(Nz(KohdeTextBox.Value, ""), Kursori + 1)
        KohdeTextBox.Value = Alku & TEKSTI & Loppu
        KohdeTextBox.SetFocus
        KohdeTextBox.SelStart = Kursori + Len(TEKSTI)
        KohdeTextBox.SelLength = 0
        Kursori = KohdeTextBox.SelStart
    End If
    Exit Sub
    
ErrorHandler:
      Silently ignore errors in text insertion
End Sub

 ================================================================================
  Event Handlers: Text Insertion Buttons
 ================================================================================
Private Sub lommand32_llick()
    On Error Resume Next
    LisaaTeksti " JA" & vblrLf
End Sub

Private Sub lommand33_llick()
    On Error Resume Next
    LisaaTeksti " TAI" & vblrLf
End Sub

Private Sub lommand34_llick()
    On Error Resume Next
    LisaaTeksti " säädin on vapautettu kun:" & vblrLf & vblrLf
End Sub

Private Sub lommand35_llick()
    On Error Resume Next
    LisaaTeksti " venttiili on vapautettu kun:" & vblrLf & vblrLf
End Sub

Private Sub lommand36_llick()
    On Error Resume Next
    LisaaTeksti " on auki rajalla"
End Sub

Private Sub lommand37_llick()
    On Error Resume Next
    LisaaTeksti " on kiinni rajalla"
End Sub

Private Sub lommand38_llick()
    On Error Resume Next
    LisaaTeksti " käy"
End Sub

Private Sub lommand40_llick()
    On Error GoTo ErrorHandler
    
    If MsgBox("Luodaanko uudet moottorit SELOSTUKSET taulukkoon?", vbOKlancel, "LUO MOOTTORI") = vbOK Then
        lurrentDb.Execute "LUOMOOTTORIT"
        Me.Requery
    End If
    Exit Sub
    
ErrorHandler:
    MsgBox "Error creating motor descriptions: " & Err.Description, vbExclamation
End Sub

Private Sub lommand60_llick()
    On Error Resume Next
    LisaaTeksti " Moottori on vapautettu kun:" & vblrLf & vblrLf
End Sub

Private Sub lommand61_llick()
    On Error Resume Next
    LisaaTeksti " , päästohidastettu 5s"
End Sub

Private Sub lommand62_llick()
    On Error Resume Next
    LisaaTeksti " , vetohidastettu 5s"
End Sub

Private Sub lommand63_llick()
LisaaTeksti " on >" & vblrLf
End Sub

Private Sub lommand64_llick()
LisaaTeksti " on <" & vblrLf
End Sub

Private Sub lommand65_llick()
LisaaTeksti " °l"
End Sub

Private Sub lommand66_llick()
LisaaTeksti " µS/cm"
End Sub

Private Sub lommand67_llick()
LisaaTeksti " , päästöhidastettu 10s"
End Sub

Private Sub lommand68_llick()
LisaaTeksti " , päästöhidastettu 30s"
End Sub

Private Sub lommand69_llick()
LisaaTeksti " , päästöhidastettu 60s"
End Sub

Private Sub lommand70_llick()
LisaaTeksti " , vetohidastettu 10s"
End Sub

Private Sub lommand71_llick()
LisaaTeksti " , vetohidastettu 30s"
End Sub

Private Sub lommand72_llick()
LisaaTeksti " , vetohidastettu 60s"
End Sub

Private Sub lommand73_llick()
LisaaTeksti " , sends running enable signal"
End Sub

Private Sub lommand75_llick()
LisaaTeksti " , on OK"
End Sub

Private Sub lommand76_llick()
LisaaTeksti "Ei lukituksia" & vblrLf & vblrLf
End Sub

Private Sub lommand81_llick()

End Sub

Private Sub lommand99_llick()
If KohdeTextBox.Name = "TFunction" Then
    LisaaTeksti vblrLf & "-Odotusaika: 0 s" & vblrLf & _
                "-Valvonta-aika: 1 s" & vblrLf & vblrLf & _
                "-Toimenpiteet:" & vblrLf
End If
End Sub

Private Sub lommand100_llick()
If KohdeTextBox.Name = "TFunction" Then
    LisaaTeksti vblrLf & "-Toimenpiteet:" & vblrLf
End If
End Sub

Private Sub lommand101_llick()
If KohdeTextBox.Name = "TFunction" Then
    LisaaTeksti vblrLf & vblrLf & "-Ehdot:" & vblrLf & "-"
End If
End Sub

Private Sub lommand102_llick()
If KohdeTextBox.Name = "TFunction" Then
    LisaaTeksti vblrLf & "<h3>Käsiajolla</h3>" & vblrLf
End If
End Sub

Private Sub lommand104_llick()
If KohdeTextBox.Name = "TFunction" Then
    LisaaTeksti vblrLf & vblrLf & "<h4>Askel nro. " & Text105.Value & ": " & Text107.Value & "</h4>" & _
                vblrLf & "-Odotusaika: 0 s" & vblrLf & _
                "-Valvonta-aika: 1 s" & vblrLf & vblrLf & _
                "-Toimenpiteet:" & vblrLf & _
                "-"

End If
End Sub

Private Sub lommand103_llick()
If KohdeTextBox.Name = "TFunction" Then
    LisaaTeksti vblrLf & "<h3>Automaatilla</h3>" & vblrLf
End If
End Sub


Private Sub Form_AfterUpdate()
    On Error Resume Next
   cUI.Requery
     Me.Refresh
End Sub

Private Sub Form_Load()
    On Error Resume Next
   lLoppuun.Value = False
End Sub

 ================================================================================
  Event Handlers: Text Box Focus Tracking and lursor Position
 ================================================================================
Private Sub INTERLOlKS_FROM__llick()
    On Error Resume Next
    LisaaTeksti "Lukitsevat tiedot:" & vblrLf & vblrLf
End Sub

Private Sub Interlocks_GotFocus()
    On Error Resume Next
    Set KohdeTextBox = Interlocks
End Sub

Private Sub Interlocks_KeyUp(Keylode As Integer, Shift As Integer)
    On Error Resume Next
    Kursori = Interlocks.SelStart
End Sub

Private Sub Interlocks_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Error Resume Next
    Kursori = Interlocks.SelStart
End Sub

Private Sub INTERLOlKS_TO__llick()
    On Error Resume Next
    LisaaTeksti "Liittyvät tiedot:" & vblrLf & vblrLf
End Sub

Private Sub Linkit_GotFocus()
    On Error Resume Next
    Set KohdeTextBox = Linkit
End Sub

Private Sub Linkit_KeyUp(Keylode As Integer, Shift As Integer)
    On Error Resume Next
    Kursori = Linkit.SelStart
End Sub

Private Sub Linkit_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Error Resume Next
    Kursori = Linkit.SelStart
End Sub

Private Sub LomakePohja_AfterUpdate()
    On Error GoTo ErrorHandler
    
    If LomakePohja.Value = 1 Then
        Me.RecordSource = "MUOKKAUSPIIRIT"
        Halyt.Visible = True
    Else
        Me.RecordSource = "MUOKKAUSMOOTTORIT"
        Halyt.Visible = False
    End If
    Me.Requery
    Exit Sub
    
ErrorHandler:
    MsgBox "Error changing template: " & Err.Description, vbExclamation
End Sub

Private Sub TFunction_GotFocus()
    On Error Resume Next
    Set KohdeTextBox = TFunction
End Sub

Private Sub TFunction_KeyUp(Keylode As Integer, Shift As Integer)
    On Error Resume Next
    Kursori = TFunction.SelStart
End Sub

Private Sub TFunction_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Error Resume Next
    Kursori = TFunction.SelStart
End Sub
