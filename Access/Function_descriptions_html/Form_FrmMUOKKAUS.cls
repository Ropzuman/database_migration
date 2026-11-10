Option Compare Database
Option Explicit

'================================================================================
' Lomake: FrmMUOKKAUS (Muokkaus/Luonti)
' Tarkoitus: Luo Word- ja HTML-toimintakuvaukset piireille ja laitteistolle
' P‰ivitetty: 2025-11-13 ó VBA7/64-bit tuki lis‰tty
'             2026-03-03 ó Me.-etuliite lomakeohjaimille; SelStart-suojaukset
'             2026-11-10 - Katsaus: ryhm‰n ohitus virhetilanteessa, HTML-escape, Null-suojaukset, h‰lytyskuvauksen kielivalinta, tiedostonimen siivous, SQL-heittomerkit, resurssien sulku
'             2026-11-10 - Viimeistely: kuollut Samaan/Eka-haara ja kommentoitu koodi poistettu, k‰ytt‰m‰ttˆm‰t muuttujat poistettu, ByVal-parametrit, kommentit korjattu (paikkamerkit), p‰iv‰m‰‰r‰ Format-funktiolla
'
' Kuvaus:
'   P‰‰lomake toimintakuvausten muokkaukseen ja luontiin.
'   Luo sek‰ Word- (.docx) ett‰ HTML-tiedostoja tietokantapohjaisista malleista.
'   K‰sittelee piirit (PIIRIT) sek‰ laitteiston (MOOTTORIT).
'
' Luontiprosessi:
'   1. K‰ytt‰j‰ valitsee mallin tyypin (piirit/moottorit) ja tulostusmuodon (Word/HTML)
'   2. Lomake lukee tiedot TIEDOT_PIIRIT- tai TIEDOT_MOOTTORIT-taulusta
'   3. Tekstikent‰t tukevat erikoissyntaksia:
'      - {LOOPTAG} lis‰‰ hyperlinkatun piiriviittauksen
'      - {Laiteviittaukset} lis‰‰v‰t laitetagit
'      - "-"-alkuiset rivit sisennet‰‰n listaan
'      - "--"-alkuiset rivit sisennet‰‰n kaksinkertaisesti
'   4. Word-muoto: avaa mallin, korvaa paikkamerkit, tallentaa uutena dokumenttina
'   5. HTML-muoto: luo HTML-tiedoston CSS-tyyleill‰ ja hyperlinkill‰
'
' Word-mallin paikkamerkit (korvataan tekstihaulla):
'   - Otsikko: [REV], [STATUS], [AUTHOR], [DOCNO], [PROJECT_LINE1], [PROJECT_LINE2], [DATE]
'   - Sis‰ltˆ: [LOOPTAG], [LOOPDESCR], [PROCESS], [FUNCTION], [LINKS], [CHANGES],
'              [PIDwgNo], [CONTROLMODES], [INTERLOCKS], [ALARMS], [EQUIPMENT]
'   - Valinnaiset (poistetaan, jos tieto puuttuu): [ACTION_T], [ACTION], [SETPOINT],
'              [RANGE_T], [RANGE], [OTHER_T], [OTHER]
'
' Riippuvuudet:
'   - Word-automaatio (myˆh‰inen sidonta)
'   - FileSystemObject (myˆh‰inen sidonta)
'   - DAO.Recordset
'   - Lomakkeet: MOTORS subform, PIIRIT subform
'   - Taulut: TIEDOT_PIIRIT, TIEDOT_MOOTTORIT, SELOSTUKSET
'   - Kyselyt: LUOPIIRIT (luo uudet piirikuvaukset)
'   - Moduuli: USysCheck (SniffUser-funktio)
'   - Julkiset muuttujat: KohdeTextBox As TextBox, Kursori As Long
'   - Ohjaimet: LOOPTAG, CLoppuun, FKieli, LomakePohja, FTyyppi
'             Alarms, Function, Interlocks, Crossref, RevHistoria
'================================================================================

' Moduulitason muuttujat
Dim POHJA As String            ' Mallitiedoston polku
Dim POLKU As String            ' Tulostushakemiston polku
Dim REV As String              ' Revisionumero
Dim STATUS As String           ' Dokumentin tila
Dim AUTHOR As String           ' Dokumentin tekij‰
Dim DOCNO As String            ' Dokumenttinumero
Dim PROJECT_L1 As String       ' Projektin otsikkorivi 1
Dim PROJECT_L2 As String       ' Projektin otsikkorivi 2
Dim oWord As Object            ' Wordin sovellusolio (myˆh‰inen sidonta)
Dim oTemplate As Object        ' Wordin dokumenttiolio (myˆh‰inen sidonta)
Dim Tiedot As DAO.Recordset    ' Tietuejoukko (piirit tai moottorit)
Dim Suomeksi As Boolean        ' Kielilippu: True = suomi, False = englanti
Dim TOUT As Object             ' HTML-tulostiedoston virta (myˆh‰inen sidonta)

' Word-kirjaston vakiot pidet‰‰n paikallisina, jotta Word-viitett‰ ei tarvita k‰‰nnˆksess‰.
Private Const wdMainTextStory As Long = 1
Private Const wdPrimaryHeaderStory As Long = 7
Private Const wdFindContinue As Long = 1
Private Const wdReplaceAll As Long = 2

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
' Tarkoitus: Luo yksitt‰inen toimintakuvaus
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
'   Kaikki ó True = kaikki kuvaukset, False = vain nykyinen tietue
'================================================================================
Private Sub TeeSelostus(Kaikki As Boolean)
    Dim Nimi As String
    Dim TTIEDOT As String
    Dim FS As Object
    Dim sTagiEnnen As String
    
    On Error GoTo ErrorHandler
    
    Suomeksi = False
    ' Kieli on toistaiseksi kiinte‰ (englanti); suomenkielinen haara otetaan k‰yttˆˆn asettamalla Suomeksi = True

    Set FS = CreateObject("Scripting.FileSystemObject")

    PoimiAsetukset

    ' Varmistetaan, ett‰ tulostuspolku on m‰‰ritelty
    If POLKU = "" Then
        MsgBox "OUTPUTPATH on tyhj‰ SETTINGS-taulussa.", vbExclamation, "TeeSelostus"
        Exit Sub
    End If

    ' Tarkistetaan hakemiston olemassaolo
    If Not FS.FolderExists(POLKU) Then
        MsgBox "Tulostushakemistoa ei lˆydy:" & vbCrLf & POLKU, vbExclamation, "TeeSelostus"
        Exit Sub
    End If
    
    ' Valitaan tietol‰hde lomakepohjan tyypin mukaan
    If Me.LomakePohja.Value = 1 Then
        TTIEDOT = "TIEDOT_PIIRIT"
    Else
        TTIEDOT = "TIEDOT_MOOTTORIT"
    End If
  
    ' Alustetaan Word tarvittaessa
    If Me.FTyyppi = 1 Then
        Set oWord = CreateObject("Word.Application")
        oWord.Visible = True
    End If
    
    ' Avataan tietuejoukko (Recordset)
    If Kaikki Then
      Set Tiedot = CurrentDb.OpenRecordset("SELECT * FROM " & TTIEDOT & " ORDER BY LOOPTAG, AreaCode, LoopNo")
    Else
      Set Tiedot = CurrentDb.OpenRecordset("SELECT * FROM " & TTIEDOT & " WHERE Looptag='" & Replace(Nz(Me.LOOPTAG, ""), "'", "''") & "' ORDER BY AreaCode, LoopNo")
    End If
    
    ' K‰yd‰‰n tietueet l‰pi LoopTAG-ryhmitt‰in: yksi tiedosto per piiri
    Do While Not Tiedot.EOF
        sTagiEnnen = Nz(Tiedot.Fields("LoopTAG"), "")

        If sTagiEnnen = "" Then
            Tiedot.MoveNext ' Ohitetaan tyhj‰t tagit
        ElseIf Not RyhmassaOnSisaltoa(TTIEDOT, sTagiEnnen) Then
            ' Ryhm‰ss‰ on vain tunnistetiedot, joten tiedostoa ei luoda
            OhitaRyhma sTagiEnnen
        ElseIf Me.FTyyppi = 1 Then ' Word-versio
            Set oTemplate = oWord.Documents.Open(POHJA)
            Nimi = POLKU & TiedostoNimi(DOCNO & "_" & sTagiEnnen) & ".docx"

            ' MuutaTiedot siirt‰‰ kursorin ryhm‰n loppuun; OhitaRyhma varmistaa etenemisen virhetilanteessa
            MuutaTiedot
            OhitaRyhma sTagiEnnen

            oTemplate.SaveAs Nimi
            If Kaikki Then oTemplate.Close False
        Else ' HTML-versio
            Set TOUT = FS.CreateTextFile(POLKU & TiedostoNimi(sTagiEnnen) & ".html")
            TeeHTML
            OhitaRyhma sTagiEnnen
            TOUT.Close
            Set TOUT = Nothing
        End If
        DoEvents ' Pit‰‰ Windowsin vastaamassa pitk‰n ajon aikana (est‰‰ "Ei vastaa" -tilan)
    Loop

    Set FS = Nothing
    SuljeTiedot
    MsgBox "Valmis", vbInformation
    Exit Sub

ErrorHandler:
    MsgBox "Virhe TeeSelostus-ajossa: " & Err.Description, vbCritical
    If Not oWord Is Nothing Then oWord.Visible = True
    Set FS = Nothing
    SuljeTiedot
End Sub

'================================================================================
' Funktio: HtmlEsc
' Tarkoitus: Muuntaa tekstin HTML-turvalliseen muotoon (&, < ja >)
' Parametrit:
'   Teksti - Muunnettava arvo (Null palauttaa tyhj‰n merkkijonon)
'================================================================================
Private Function HtmlEsc(ByVal Teksti As Variant) As String
    If IsNull(Teksti) Then Exit Function
    HtmlEsc = CStr(Teksti)
    HtmlEsc = Replace(HtmlEsc, "&", "&amp;")
    HtmlEsc = Replace(HtmlEsc, "<", "&lt;")
    HtmlEsc = Replace(HtmlEsc, ">", "&gt;")
End Function

'================================================================================
' Funktio: TiedostoNimi
' Tarkoitus: Poistaa tiedostonimest‰ merkit, jotka eiv‰t ole sallittuja Windowsissa
'================================================================================
Private Function TiedostoNimi(ByVal Nimi As String) As String
    Dim Merkki As Variant
    TiedostoNimi = Nimi
    For Each Merkki In Array("\", "/", ":", "*", "?", """", "<", ">", "|")
        TiedostoNimi = Replace(TiedostoNimi, Merkki, "_")
    Next Merkki
End Function

'================================================================================
' Aliohjelma: OhitaRyhma
' Tarkoitus: Siirt‰‰ tietuekursorin annetun LoopTAG-ryhm‰n ohi
'
' Kuvaus:
'   Normaalisti MuutaTiedot ja TeeHTML siirt‰v‰t kursorin ryhm‰n loppuun itse.
'   Virhetilanteessa kursori j‰‰ paikalleen, jolloin p‰‰silmukka j‰isi
'   k‰sittelem‰‰n samaa tietuetta loputtomasti. T‰m‰ ohittaa ryhm‰n.
'================================================================================
Private Sub OhitaRyhma(Tagi As String)
    Do While Not Tiedot.EOF
        If Nz(Tiedot.Fields("LoopTAG"), "") <> Tagi Then Exit Do
        Tiedot.MoveNext
    Loop
End Sub

'================================================================================
' Aliohjelma: SuljeTiedot
' Tarkoitus: Sulkee tietuejoukon ja HTML-tiedostovirran sek‰ vapauttaa Word-viittaukset
'            (Word j‰‰ auki, jotta k‰ytt‰j‰ voi tarkastella dokumentteja)
'================================================================================
Private Sub SuljeTiedot()
    On Error Resume Next
    If Not Tiedot Is Nothing Then Tiedot.Close
    Set Tiedot = Nothing
    If Not TOUT Is Nothing Then TOUT.Close
    Set TOUT = Nothing
    Set oTemplate = Nothing
    Set oWord = Nothing
End Sub

'================================================================================
' Funktio: HtmlEscTeksti
' Tarkoitus: HTML-escape kuvaustekstille niin, ett‰ painikkeilla lis‰tyt
'            v‰liotsikot (<h3>, <h4>) s‰ilyv‰t otsikoina
'================================================================================
Private Function HtmlEscTeksti(ByVal Teksti As Variant) As String
    Dim Tunniste As Variant
    HtmlEscTeksti = HtmlEsc(Teksti)
    For Each Tunniste In Array("h3", "h4")
        HtmlEscTeksti = Replace(HtmlEscTeksti, "&lt;" & Tunniste & "&gt;", "<" & Tunniste & ">", , , vbTextCompare)
        HtmlEscTeksti = Replace(HtmlEscTeksti, "&lt;/" & Tunniste & "&gt;", "</" & Tunniste & ">", , , vbTextCompare)
    Next Tunniste
End Function

'================================================================================
' Funktio: TFunctionValittu
' Tarkoitus: Palauttaa True, jos kohdekentt‰n‰ on Toiminta-kentt‰ (TFunction)
'            Suojaa virheelt‰ 91, kun kentt‰‰ ei ole viel‰ valittu
'================================================================================
Private Function TFunctionValittu() As Boolean
    If KohdeTextBox Is Nothing Then Exit Function
    TFunctionValittu = (KohdeTextBox.Name = "TFunction")
End Function

'================================================================================
' Funktio: RyhmassaOnSisaltoa
' Tarkoitus: Tarkistaa, sis‰lt‰‰kˆ nykyinen LoopTAG-ryhm‰ muutakin kuin tunnistetiedot
' Parametrit:
'   TTIEDOT ó Tietol‰hteen nimi
'   Tagi    ó Tarkistettava LoopTAG
' Palauttaa:
'   True, jos ryhm‰ss‰ on jotakin tulostettavaa sis‰ltˆ‰
'================================================================================
Private Function RyhmassaOnSisaltoa(TTIEDOT As String, Tagi As String) As Boolean
  Dim Tarkistus As DAO.Recordset
  Dim Hakuehto As String

  On Error GoTo ErrorHandler

  Hakuehto = "SELECT * FROM " & TTIEDOT & " WHERE LoopTAG='" & Replace(Tagi, "'", "''") & "'"
  Set Tarkistus = CurrentDb.OpenRecordset(Hakuehto)

  Do While Not Tarkistus.EOF
    If Len(Trim$(Nz(Tarkistus.Fields("Function"), ""))) > 0 Then
      RyhmassaOnSisaltoa = True
      Exit Do
    End If

    Tarkistus.MoveNext
  Loop

  Tarkistus.Close
  Set Tarkistus = Nothing
  Exit Function

ErrorHandler:
  If Not Tarkistus Is Nothing Then Tarkistus.Close
  Set Tarkistus = Nothing
  RyhmassaOnSisaltoa = True
End Function

'================================================================================
' Aliohjelma: TeeHTML
' Tarkoitus: Luo HTML-muotoinen toimintakuvaus
'
' Kuvaus:
'   Luo HTML-tiedoston piirin tai laitteen toimintakuvauksesta:
'   - Dokumentin otsikko-osa, otsikko ja CSS-linkki
'   - Toiminta-, lukitukset- ja h‰lytykset-osiot
'   - Ristiviittaukset muihin piireihin
'   - Muutoshistoria
'   Erikoismuotoilu: {LOOPTAG} muuttuu hyperlinkiksi, "-"-rivit sisennet‰‰n
'================================================================================
Private Sub TeeHTML()
    ' TOUT-tiedostovirta on avattu kutsujassa (TeeSelostus)
    Dim TEKSTI As String
    Dim ADescr As String
    Dim Tagi As String
    
    On Error GoTo ErrorHandler

    Tagi = Tiedot.Fields("LOOPTAG")

TOUT.WriteLine "<!DOCTYPE html>"
TOUT.WriteLine "<html lang=""fi"">"
TOUT.WriteLine "<head>"
TOUT.WriteLine "    <meta charset=""windows-1252"">"
TOUT.WriteLine "    <meta name=""viewport"" content=""width=device-width, initial-scale=1.0"">"
TOUT.WriteLine "    <link rel=""stylesheet"" href=""FuncDesc.css"">"
TOUT.WriteLine "    <title>" & HtmlEsc(Tiedot.Fields("LoopTAG")) & " " & HtmlEsc(Tiedot.Fields("Description")) & "</title>"
TOUT.WriteLine "</head>"
TOUT.WriteLine "<body>"

'OTSIKKO - Piirin positio ja nimitys
TOUT.WriteLine "    <h1 id=""loopname"">"
TOUT.WriteLine "        <a href=""" & HtmlEsc(TiedostoNimi(Tagi)) & ".html"" target=""_parent"">" _
        & HtmlEsc(Tiedot.Fields("LOOPTAG")) & "</a> " _
        & HtmlEsc(Tiedot.Fields("Description")) '& "<br>"
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

'HƒLYTYKSET
TOUT.WriteLine "    <h2 id=""alarms"">"
TOUT.WriteLine "        H‰lytykset"
TOUT.WriteLine "    </h2>"
TEKSTI = Nz(Tiedot.Fields("Alarms"))
MuutaTeksti TEKSTI
    
' H‰lytysrajat kirjoitetaan vain, jos h‰lytyskuvaus tai h‰lytysteksti on annettu
' ja v‰hint‰‰n yksi raja-arvo (HH/H/L/LL) on asetettu
If Not (IsNull(Tiedot.Fields("AlarmDescrBL")) And IsNull(Tiedot.Fields("Alarms"))) Then
    If Not (IsNull(Tiedot.Fields("HH")) And IsNull(Tiedot.Fields("H")) _
            And IsNull(Tiedot.Fields("LL")) And IsNull(Tiedot.Fields("L"))) Then
        ADescr = Nz(Tiedot.Fields("AlarmDescrBL"), "")
        TOUT.WriteLine "<dl>"
        LisaaHalyHTML Tiedot.Fields("HH"), "HH: " & ADescr & " > "
        LisaaHalyHTML Tiedot.Fields("H"), "H: " & ADescr & " > "
        LisaaHalyHTML Tiedot.Fields("L"), "L: " & ADescr & " < "
        LisaaHalyHTML Tiedot.Fields("LL"), "LL: " & ADescr & " < "
        TOUT.WriteLine "</dl>"
    End If
End If


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

' Muilla riveill‰ on piirin laitteet, joita ei tulosteta; siirret‰‰n kursori ryhm‰n ohi
OhitaRyhma Tagi
  
  Exit Sub
  
ErrorHandler:
    MsgBox "Virhe HTML-kuvauksen luonnissa: " & Err.Description, vbExclamation
End Sub

'================================================================================
' Aliohjelma: MuutaTeksti
' Tarkoitus: Muuttaa pelk‰n tekstin HTML-muotoiseksi
' Parametrit:
'   TEKSTI ó Muunnettava teksti ({LOOPTAG}-linkit ja sisennykset tuettu)
'
' Kuvaus:
'   Muuntaa tekstiˆntˆ HTML-muotoon:
'   - {LOOPTAG} muuttuu hyperlinkiksi
'   - "-"-alkuiset rivit sisennet‰‰n (dd dl:ss‰)
'   - "--"-alkuiset rivit sisennet‰‰n kaksinkertaisesti
'   - Tyhj‰ rivi ennen sisennyst‰ muuttuu otsikkoriviksi (dt dl:ss‰)
'================================================================================
Private Sub MuutaTeksti(TEKSTI As String)
    Dim Rivit As Variant
    Dim Riveja As Integer
    Dim r As Integer
    Dim Sis As Boolean
    
    On Error GoTo ErrorHandler

    If TEKSTI <> "" Then
        TEKSTI = MuutaLinkit(HtmlEscTeksti(TEKSTI))  'Muutetaan aaltosulkeissa olevat piiripositiot linkeiksi
        Rivit = Split(TEKSTI, vbCrLf) 'Splitataan rivit ja muutetaan html kappeliksi
        Riveja = UBound(Rivit)
        For r = 0 To Riveja
            If Left$(Rivit(r), 2) = "--" Then 'Ollaan kahdesti sisennetyll‰ rivill‰ tehd‰‰n rivist‰ sisennetty
      If Sis = False Then
          TOUT.WriteLine "      <dl>"
          Sis = True
      End If
      TOUT.WriteLine "        <dl>"
      TOUT.WriteLine "          <dd>"
      TOUT.WriteLine "             " & Trim$(Mid$(Rivit(r), 3))
      TOUT.WriteLine "          </dd>"
      TOUT.WriteLine "        </dl>"
      If r = Riveja Then ' Ollaan viimeisell‰ rivill‰, lopetetaan sisennys
        TOUT.WriteLine "      </dl>"
      End If
     ElseIf Left$(Rivit(r), 1) = "-" Then 'Ollaan sisennetyll‰ rivill‰ tehd‰‰n rivist‰ sisennetty
      If Sis = False Then
          TOUT.WriteLine "      <dl>"
          Sis = True
      End If
      TOUT.WriteLine "        <dd>"
      TOUT.WriteLine "           " & Trim$(Mid$(Rivit(r), 2))
      TOUT.WriteLine "        </dd>"
      If r = Riveja Then ' Ollaan viimeisell‰ rivill‰, lopetetaan sisennys
        TOUT.WriteLine "      </dl>"
      End If
    Else
      If r > 0 Then 'Ei ole ensimm‰inen rivi
        If Sis Then 'On Aloitettu sisennys aikaisemmin, joten p‰‰tet‰‰n edellinen sisennys
          TOUT.WriteLine "      </dl>"
        End If
      End If
      If r < Riveja Then  'Ei olla viimeisell‰ rivill‰
        If Left$(Rivit(r + 1), 1) = "-" Then 'Seuraava rivi on sisennetty, tehd‰‰n t‰st‰ otsikkorivi
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
        Else 'Seuraava rivi ei ole sisennetty, tehd‰‰n t‰st‰ rivist‰ normaali kappale
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
            'Viimeinen rivi, joka ei ole sisennetty, eik‰ seuraava rivi voi olla sisennetty, joten tehd‰‰n normaali kappale
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
' Parametrit:
'   Tieto ó {LOOPTAG}-viittauksia sis‰lt‰v‰ teksti
' Palauttaa:
'   Teksti, jossa {LOOPTAG} on muutettu muotoon <a href="LOOPTAG.html">LOOPTAG</a>
'
' Kuvaus:
'   Etsii {POS Kuvaus}-muotoiset rakenteet ja muuntaa ne hyperlinkeiksi,
'   miss‰ POS on piirin tagi. Tukee sek‰ {POS}- ett‰ {POS Kuvaus}-muotoja.
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
      DESC = ""
        OS2 = InStr(OS, MuutaLinkit, " ")
        OS3 = InStr(OS, MuutaLinkit, "}")
      If OS3 = 0 Then Exit Do
      If OS2 = 0 Or OS2 > OS3 Then OS2 = OS3
        POS = Mid$(MuutaLinkit, OS + 1, OS2 - OS - 1)
        If OS2 < OS3 Then
            DESC = " " & Mid$(MuutaLinkit, OS2 + 1, OS3 - OS2 - 1)
        End If
        MuutaLinkit = Left$(MuutaLinkit, OS - 1) & "<a href=""" & POS & ".html"">" & POS & DESC & "</a>" & Mid$(MuutaLinkit, OS3 + 1)
    Loop
    
    Exit Function
    
ErrorHandler:
    MuutaLinkit = Tieto  ' Palautetaan alkuper‰inen virhetilanteessa
End Function

'================================================================================
' Aliohjelma: PoimiAsetukset
' Tarkoitus: Lataa luontiasetukset SETTINGS-taulusta
'
' Kuvaus:
'   Lukee SETTINGS-taulusta konfiguraation: mallipolut, hakemistot ja
'   projektin otsikkotiedot dokumenttien yl‰tunnistetta varten.
'================================================================================
Private Sub PoimiAsetukset()
    Dim Taul As DAO.Recordset
    
    On Error GoTo ErrorHandler
    
    Set Taul = CurrentDb.OpenRecordset("SETTINGS")
    
    ' Tarkistetaan, ett‰ SETTINGS-taulussa on edes yksi rivi
    If Taul.EOF Then
        Taul.Close
        Set Taul = Nothing
        MsgBox "SETTINGS-taulukko on tyhj‰. Asetusten hakeminen ep‰onnistui.", vbExclamation, "PoimiAsetukset"
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
' Tarkoitus: T‰ytt‰‰ Word-mallin paikkamerkit piirin/laitteen tiedoilla
'
' Kuvaus:
'   Korvaa Word-mallin paikkamerkit nykyisen tietueen tiedoilla.
'   K‰sittelee otsikkotiedot ([REV], [STATUS], [AUTHOR] jne.) ja sis‰ltˆkent‰t
'   ([LOOPTAG], [FUNCTION], [ALARMS] jne.). Rakentaa h‰lytys- ja laitelistan.
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
    ' H‰lytysteksti muodostetaan vain, jos h‰lytyskuvaus on annettu
    If Not IsNull(Tiedot.Fields("AlarmDescrBL")) Then
    If IsNull(Tiedot.Fields("HH")) And IsNull(Tiedot.Fields("H")) And IsNull(Tiedot.Fields("LL")) And IsNull(Tiedot.Fields("L")) Then
      If Suomeksi Then
        Halyt = "Ei h‰lytyksi‰"
      Else
        Halyt = "No alarms"
      End If
    Else
      If Suomeksi Then
        ADescr = Nz(Tiedot.Fields("AlarmDescrBL_FI"), "")
      Else
        ADescr = Nz(Tiedot.Fields("AlarmDescrBL"), "")
      End If
      LisaaHaly Tiedot.Fields("HH"), Halyt, "HH: " & ADescr & " > "
      LisaaHaly Tiedot.Fields("H"), Halyt, "H: " & ADescr & " > "
      LisaaHaly Tiedot.Fields("L"), Halyt, "L: " & ADescr & " < "
      LisaaHaly Tiedot.Fields("LL"), Halyt, "LL: " & ADescr & " < "
    End If
  End If

  
  ' Otsikkotiedot korvataan mallin yl‰tunnisteeseen
  KorvaaOts "[REV]", REV
  KorvaaOts "[PROJECT_LINE1]", PROJECT_L1
  KorvaaOts "[PROJECT_LINE2]", PROJECT_L2
  KorvaaOts "[STATUS]", STATUS
  KorvaaOts "[DOCNO]", DOCNO
  KorvaaOts "[AUTHOR]", AUTHOR
  KorvaaOts "[DATE]", Format$(Date, "d.m.yyyy")
  Korvaa "[PIDwgNo]", Nz(Tiedot.Fields("PIDwgNo"))
  If Suomeksi Then
    Korvaa "[LOOPDESCR]", UCase$(Nz(Tiedot.Fields("DescriptionFI"), ""))
    Korvaa "[LOOPTAG]", Tagi
    Korvaa "[PROCESS]", Nz(Tiedot.Fields("ProcessFI"))
    Korvaa "[FUNCTION]", Nz(Tiedot.Fields("FunctionFI"), "")
  Else
    Korvaa "[LOOPDESCR]", UCase$(Nz(Tiedot.Fields("Description"), ""))
    Korvaa "[LOOPTAG]", Tagi
    Korvaa "[PROCESS]", Nz(Tiedot.Fields("Process"))
    Korvaa "[FUNCTION]", Nz(Tiedot.Fields("Function"))
  End If
  Korvaa "[LINKS]", Nz(Tiedot.Fields("Linkit"))
  Korvaa "[CHANGES]", Nz(Tiedot.Fields("Muutokset"))
  
  If Me.LomakePohja.Value = 1 Then ' Kysymyksess‰ piiri
    PID_D = Null
    PID_U = Null
    On Error Resume Next  ' CONTROL-taulu voi puuttua linkityksist‰, jolloin s‰‰timen toimintaa ei m‰‰ritet‰
    PID_D = DLookup("FBlock", "CONTROL", "AreaCode = '" & Tiedot.Fields("AreaCode") & "' AND LoopNo='" & Tiedot.Fields("LoopNo") & "' AND FBlock='PID_D'")
    PID_U = DLookup("FBlock", "CONTROL", "AreaCode = '" & Tiedot.Fields("AreaCode") & "' AND LoopNo='" & Tiedot.Fields("LoopNo") & "' AND FBlock='PID_U'")
    On Error GoTo ErrorHandler
     
   If IsNull(PID_D) = False Or IsNull(PID_U) = False Then
     If Suomeksi Then
       Korvaa "[ACTION_T]", "S‰‰timen toiminta"
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
       Korvaa "[ACTION]", "K‰‰nteinen"
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
       Korvaa "[SETPOINT]", "S‰‰timen tyypillinen asetusarvoalue on: " & Tiedot.Fields("SetPoint")
     Else
       Korvaa "[SETPOINT]", "Typical set point range of controller is: " & Tiedot.Fields("SetPoint")
     End If
   End If
  Else
     POISTA "[ACTION_T]"
     POISTA "[ACTION]"
  End If
  If IsNull(Tiedot.Fields("RANGE")) Then
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
  
  ' Ker‰t‰‰n piirin laitteet (TAGNAME); Nz() est‰‰ Null-arvon aiheuttaman virheen
  Do While Not Tiedot.EOF
    If Tiedot.Fields("LoopTAG") <> Tagi Then Exit Do
    If Laitteet <> "" Then Laitteet = Laitteet & vbCrLf
    If Suomeksi Then
      Laitteet = Laitteet & Nz(Tiedot.Fields("TAGNAME_FI"), "")
    Else
      Laitteet = Laitteet & Nz(Tiedot.Fields("TAGNAME"), "")
    End If
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
' Parametrit:
'   VanhaTeksti ó Korvattava teksti
'   UusiTeksti  ó Uusi teksti
'================================================================================
Private Sub KorvaaOts(ByVal VanhaTeksti As String, ByVal UusiTeksti As String)
    
    On Error Resume Next  ' Otsikko saattaa puuttua ó jatketaan hiljaisesti
    
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
' Tarkoitus: Korvaa teksti Word-dokumentin p‰‰tekstiss‰
' Parametrit:
'   VanhaTeksti ó Korvattava teksti (paikkamerkki)
'   UusiTeksti  ó Uusi teksti
'
' Kuvaus:
'   Korvaa paikkamerkin tekstin dokumentissa. Pitk‰ teksti (>255 merkki‰)
'   k‰sitell‰‰n erikseen Word-rajoituksen vuoksi. Poistaa {aaltosulkeet}.
'================================================================================
Private Sub Korvaa(ByVal VanhaTeksti As String, ByVal UusiTeksti As String)
    
    On Error Resume Next  ' Paikkamerkki saattaa puuttua ó jatketaan hiljaisesti
    
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
' Tarkoitus: Korvaa teksti Word-dokumentissa s‰ilytt‰en kappaleiden vaihdot
' Parametrit:
'   VanhaTeksti ó Korvattava teksti (paikkamerkki)
'   UusiTeksti  ó Uusi teksti, kappaleet erotettu vbCrLf-merkill‰
'
' Kuvaus:
'   Korvaa paikkamerkin tekstin ja muuntaa vbCrLf-erottimet Word-kappaleiksi.
'   Poistaa {aaltosulkeet} tekstist‰.
'================================================================================
Private Sub Korvaa2(VanhaTeksti As String, ByVal UusiTeksti As String)
    Dim Tekstit As Variant
    Dim i As Integer
    
    On Error Resume Next  ' Paikkamerkki saattaa puuttua ó jatketaan hiljaisesti
    
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
' Tarkoitus: Poistaa paikkamerkin tekstin Word-dokumentista
' Parametrit:
'   TEKSTI ó Poistettava paikkamerkki
'================================================================================
Private Sub POISTA(ByVal TEKSTI As String)
    
    On Error Resume Next  ' Paikkamerkki saattaa puuttua ó jatketaan hiljaisesti
    
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
' Tarkoitus: Lis‰‰ h‰lytystasorivi h‰lytyksien tekstimerkkijonoon (Word-tulostus)
' Parametrit:
'   Tieto ó H‰lytystason arvo
'   Haly  ó Nykyinen h‰lytysmerkijono
'   Alku  ó H‰lytyksen etuliiteteksti (esim. "HH: > ")
'================================================================================
Private Sub LisaaHaly(Tieto As Variant, Haly As String, Alku As String)
    On Error Resume Next
    If IsNull(Tieto) Then Exit Sub
    If Haly <> "" Then Haly = Haly & vbCrLf
    Haly = Haly & Alku & Tieto & Tiedot.Fields("Unit")
End Sub

'================================================================================
' Aliohjelma: LisaaHalyHTML
' Tarkoitus: Kirjoittaa h‰lytystasorivi suoraan HTML-tulostiedostoon
' Parametrit:
'   Tieto ó H‰lytystason arvo
'   Alku  ó H‰lytyksen etuliiteteksti (esim. "HH: > ")
'================================================================================
Private Sub LisaaHalyHTML(Tieto As Variant, Alku As String)
    On Error Resume Next
    If IsNull(Tieto) Then Exit Sub
    TOUT.WriteLine "<dd>"
    TOUT.WriteLine HtmlEsc(Alku & Tieto & Tiedot.Fields("Unit"))
    TOUT.WriteLine "</dd>"
End Sub

'================================================================================
' Aliohjelma: LisaaTeksti
' Tarkoitus: Lis‰‰ teksti kursorin kohdalle kohdetekstikentt‰‰n
' Parametrit:
'   TEKSTI ó Lis‰tt‰v‰ teksti
' Huom: SelStart vaatii fokuksen ó kietotaan virheensuojauksen sis‰‰n (Virhe 2185)
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
        ' SelStart ja SelLength vaativat fokuksen ó estet‰‰n Virhe 2185
        On Error Resume Next
        KohdeTextBox.SelStart = Kursori + Len(TEKSTI)
        KohdeTextBox.SelLength = 0
        Kursori = KohdeTextBox.SelStart
        On Error GoTo ErrorHandler
    End If
    Exit Sub
    
ErrorHandler:
    ' Tekstin lis‰ys ep‰onnistui hiljaisesti
End Sub

'================================================================================
' Tapahtumat: Teksti‰ lis‰‰v‰t painikkeet
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
    LisaaTeksti " s‰‰din on vapautettu kun:" & vbCrLf & vbCrLf
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
    LisaaTeksti " k‰y"
End Sub

Private Sub Command40_Click()
    On Error GoTo ErrorHandler
    
    If MsgBox("Luodaanko uudet moottorit SELOSTUKSET taulukkoon?", vbOKCancel, "LUO MOOTTORI") = vbOK Then
        CurrentDb.Execute "LUOMOOTTORIT"
        Me.Requery
    End If
    Exit Sub
    
ErrorHandler:
    MsgBox "Virhe moottorikuvausten luonnissa: " & Err.Description, vbExclamation
End Sub

Private Sub Command60_Click()
    On Error Resume Next
    LisaaTeksti " Moottori on vapautettu kun:" & vbCrLf & vbCrLf
End Sub

Private Sub Command61_Click()
    On Error Resume Next
    LisaaTeksti " , p‰‰stohidastettu 5s"
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
LisaaTeksti " ∞C"
End Sub

Private Sub Command66_Click()
LisaaTeksti " µS/cm"
End Sub

Private Sub Command67_Click()
LisaaTeksti " , p‰‰stˆhidastettu 10s"
End Sub

Private Sub Command68_Click()
LisaaTeksti " , p‰‰stˆhidastettu 30s"
End Sub

Private Sub Command69_Click()
LisaaTeksti " , p‰‰stˆhidastettu 60s"
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

Private Sub Command99_Click()
If TFunctionValittu Then
    LisaaTeksti vbCrLf & "-Odotusaika: 0 s" & vbCrLf & _
                "-Valvonta-aika: 1 s" & vbCrLf & vbCrLf & _
                "-Toimenpiteet:" & vbCrLf
End If
End Sub

Private Sub Command100_Click()
If TFunctionValittu Then
    LisaaTeksti vbCrLf & "-Toimenpiteet:" & vbCrLf
End If
End Sub

Private Sub Command101_Click()
If TFunctionValittu Then
    LisaaTeksti vbCrLf & vbCrLf & "-Ehdot:" & vbCrLf & "-"
End If
End Sub

Private Sub Command102_Click()
If TFunctionValittu Then
    LisaaTeksti vbCrLf & "<h3>K‰siajolla</h3>" & vbCrLf
End If
End Sub

Private Sub Command104_Click()
If TFunctionValittu Then
    LisaaTeksti vbCrLf & vbCrLf & "<h4>Askel nro. " & Me.Text105.Value & ": " & Me.Text107.Value & "</h4>" & _
                vbCrLf & "-Odotusaika: 0 s" & vbCrLf & _
                "-Valvonta-aika: 1 s" & vbCrLf & vbCrLf & _
                "-Toimenpiteet:" & vbCrLf & _
                "-"

End If
End Sub

Private Sub Command103_Click()
If TFunctionValittu Then
    LisaaTeksti vbCrLf & "<h3>Automaatilla</h3>" & vbCrLf
End If
End Sub


Private Sub Form_Load()
    ' Lomakkeen alustus ó ei toimenpiteit‰
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
    LisaaTeksti "Liittyv‰t tiedot:" & vbCrLf & vbCrLf
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
    
    ' Vaihdetaan tietol‰hde valitun pohjan mukaan
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
