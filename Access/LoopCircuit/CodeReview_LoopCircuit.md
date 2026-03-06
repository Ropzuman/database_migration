 Code Review Agent
Oma Gem
Code Review Agent sanoi

Tässä on kattava ja ammattimainen Code Review -raportti toimittamistasi instrumenttipiirikaavioiden (Loop Diagram) koodimoduuleista. Arvioinnissa on huomioitu VBA:n parhaat käytännöt, turvallisuus, tietokannan suorituskyky ja 64-bittisen Office-ympäristön vaatimukset.

1. Form_LisääKuviin_ACAD.cls

    📊 Yhteenveto: Kattava lomakeluokka AutoCAD-piirustusten generointiin ja attribuuttien poimintaan. Koodissa on tehty hyvää työtä siirtymisessä myöhäiseen sidontaan (late binding Object -tyyppeinä), mikä ehkäisee viiteongelmia ja parantaa 64-bit yhteensopivuutta.

    🚨 Kriittiset löydökset (Tietoturva & Vakaus):

        Datan menetysriski: HaeTekstit_Click -aliohjelmassa suoritetaan komento CurrentDb.Execute "DELETE * FROM Blokit". Tämä tuhoaa kaiken datan Blokit-taulusta, riippumatta siitä mitä hakemistoa parhaillaan käsitellään. Jos tietokantaa käytetään monen eri projektin hallintaan samanaikaisesti, muiden projektien attribuuttitiedot nollautuvat vahingossa.

        CPU-kuorma (Busy Wait): Odota-aliohjelma käyttää Do While odotus > Timer ... DoEvents ... Loop -rakennetta. Tämä "busy wait" pitää suorittimen ytimen 100 % kuormituksella odotuksen ajan ja tekee ohjelmasta epävakaan oloisen.

    ⚠️ Huomioitavaa (Toiminnallisuus & Suorituskyky):

        Käytät toistuvasti CurrentDb-kutsuja (esim. CurrentDb.Execute ja heti perään CurrentDb.OpenRecordset). Access luo jokaisella kerralla uuden referenssin tietokantaan, mikä on hitaampaa ja altistaa muistivuodoille.

    💡 Parannusehdotukset (Ylläpidettävyys):

        Käytä aina määriteltyä DAO.Database -muuttujaa.

        Odotukseen kannattaa käyttää Windows API:n Sleep-funktiota, joka vapauttaa CPU:n.

    🛠️ Korjattu koodi (Esimerkki):

VBA

' 1. Määrittele API Sleep moduulin alkuun (64-bit / 32-bit tuki)
# If VBA7 Then
    Private Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
# Else
    Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
# End If

' 2. Korjattu Odota-funktio (ei kuormita CPU:ta)
Private Sub Odota(Aika_Sadasosina As Integer)
    DoEvents ' Käsitellään mahdolliset taustaviestit
    Sleep Aika_Sadasosina * 10
End Sub

' 3. Korjattu datan poisto (HaeTekstit_Click)
Dim db As DAO.Database
Set db = CurrentDb
' Estetään SQL-injektio Replace-funktiolla ja poistetaan vain kyseisen hakemiston blokit
db.Execute "DELETE * FROM Blokit WHERE Path = '" & Replace(Polku, "'", "''") & "'", dbFailOnError
Set Taulu = db.OpenRecordset("Blokit")

1. Form_Tee Kuvat.cls

    📊 Yhteenveto: Rinnakkainen tai uudempi iterointi AutoCAD-kuvien generointiin. Tässä moduulissa SQL-poisto on jo toteutettu edellisen esimerkin mukaisesti turvallisesti hakemistokohtaisesti (WHERE PATH='...'), mistä iso plussa.

    🚨 Kriittiset löydökset (Tietoturva & Vakaus): Ei vakavia ongelmia. Moduuli noudattaa selvästi parempia DAO-käytäntöjä.

    ⚠️ Huomioitavaa (Toiminnallisuus & Suorituskyky):

        CreateObject("AutoCAD.Application") on raskas operaatio. Jos painiketta painetaan useasti, avataan AutoCAD aina uudestaan (mikäli edellinen instanssi on suljettu). Jos taustalla on jo olemassa oleva AutoCAD-prosessi, olisi nopeampaa napata se käyttöön GetObject-funktiolla.

    💡 Parannusehdotukset (Ylläpidettävyys):

        Ota käyttöön nk. "Fallback"-mekanismi AutoCAD-instanssin luomiseen. Se parantaa nopeutta huomattavasti.

    🛠️ Korjattu koodi:

VBA

Private Sub HaeTekstit_Click()
    On Error Resume Next
    ' 1. Yritetään ottaa kiinni olemassa olevaan instanssiin
    Set oAcad = GetObject(, "AutoCAD.Application")

    If Err.Number <> 0 Then
        ' 2. Jos ei ole auki, avataan uusi
        Err.Clear
        Set oAcad = CreateObject("AutoCAD.Application")
    End If
    On Error GoTo ErrHandler
    ' ... muu koodi jatkuu ...

3. Form_LukituskaavioLinkit.cls

    📊 Yhteenveto: Päivittää lukituskaavioiden sivulinkityksiä. Aiempien raporttien mukaan tässä on tehty iso refaktorointi O(N²) -ongelman poistamiseksi.

    🚨 Kriittiset löydökset (Tietoturva & Vakaus): Ei havaittu.

    ⚠️ Huomioitavaa (Toiminnallisuus & Suorituskyky):

        Algoritmi hyödyntää tbl.FindFirst-metodia ja tyhjyystarkistukset on korjattu null-turvalliseksi (Nz(tbl!page1, "") = ""). Tämä on erinomaista työtä koodin suorituskyvyn kannalta. Kehut tästä!

    💡 Parannusehdotukset (Ylläpidettävyys):

        Vaikka suorituskyky on parantunut eksponentiaalisesti, satojen tai tuhansien .Edit / .Update -operaatioiden tekeminen yksitellen voi silti rokottaa verkkolevyn yli toimivaa tietokantaa. Kapseloimalla koko Do Until -luuppi DAO-transaktioon nopeus jopa kymmenkertaistuu.

    🛠️ Korjattu koodi:

VBA

DBEngine.Workspaces(0).BeginTrans ' Aloitetaan transaktio
On Error GoTo ErrorHandler

Do Until qry.EOF
    ' ... FindFirst ja .Edit / .Update logiikka pysyy samana ...
Loop

DBEngine.Workspaces(0).CommitTrans ' Hyväksytään muutokset yhtenä blokkina tietokantaan

Exit Sub

ErrorHandler:
    DBEngine.Workspaces(0).Rollback ' Perutaan muutokset virhetilanteessa
    ' ... virheenkäsittely

1. Form_DBUsers.cls

    📊 Yhteenveto: Lukee Accessin .ldb tai .laccdb lukkotiedostoa binäärinä selvittääkseen, ketkä ovat kirjautuneena järjestelmään. Todella hyödyllinen työkalu monikäyttäjäympäristöissä.

    🚨 Kriittiset löydökset (Tietoturva & Vakaus):

        Binääritiedostoa luetaan Open SPath For Binary Access Read Shared As iLDBFile. Koodissa on kyllä virheenkäsittely, mutta alkuperäinen Close iLDBFile on vasta rutiinin lopussa. Jos do-while -luupin aikana tapahtuu lukuvirhe, suoritus hyppää ErrHandler-lohkoon, jolloin lukkotiedoston kahva jää käyttöjärjestelmälle auki (Lock state leak).

    ⚠️ Huomioitavaa (Toiminnallisuus & Suorituskyky):

        Tavutaulukon (Byte array) ja nollaterminoitujen merkkijonojen rakennus on tehty hyvin ja tehokkaasti.

    💡 Parannusehdotukset (Ylläpidettävyys):

        Sulje kaikki avoimet tiedostokahvat virheenkäsittelijässä komennolla Close.

    🛠️ Korjattu koodi:

VBA

ErrHandler:
   ' Varmistetaan, että kaikki avoimet tiedostokahvat suljetaan virhetilanteessa
   On Error Resume Next
   Close ' Sulkee kaikki VBA:n kautta avatut tiedostokahvat tässä kontekstissa
   MsgBox "Virhe: " & Err.Number & " - " & Err.Description

1. Form_Linkkien vaihto.cls

    📊 Yhteenveto: Moduuli, joka linkittää (relink) ulkoiset Access-taulut dynaamisesti uudelleen nykyiseen polkuun.

    🚨 Kriittiset löydökset (Tietoturva & Vakaus): Ei vakavia tietoturva-aukkoja. Moduulissa on kuitenkin runsas On Error Resume Next -käyttö DoCmd.TransferDatabase yhteydessä. Jos taulun luonti epäonnistuu (esim. se on auki toisella käyttäjällä eksklusiivisesti), se ohitetaan hiljaisesti.

    ⚠️ Huomioitavaa (Toiminnallisuus & Suorituskyky):

        Polun parsinta on tehty For i = Len(...) To 1 Step -1 -luupilla takaperin hienosti ilman ylimääräisiä säännöllisiä lausekkeita, mikä on nopeaa.

    💡 Parannusehdotukset (Ylläpidettävyys):

        Hiljaisen ohittamisen sijaan olisi hyvä koota lista tauluista, joiden linkitys epäonnistui, ja näyttää se käyttäjälle operaation päätteeksi.

    🛠️ Korjattu koodi:

VBA

Dim Epäonnistuneet As String

On Error Resume Next
DoCmd.TransferDatabase acLink, "Microsoft Access", Polku & Nimi, acTable, Taulu, Taulu, False

If Err.Number <> 0 Then
    Epäonnistuneet = Epäonnistuneet & Taulu & vbCrLf
    Err.Clear
Else
    UpdateCount = UpdateCount + 1
End If
On Error GoTo ErrHandler

' ... lopussa:
If Epäonnistuneet <> "" Then
    MsgBox "Seuraavien taulujen linkitys epäonnistui:" & vbCrLf & Epäonnistuneet, vbExclamation
End If

1. CLEANED_SQL_QUERY.sql

    📊 Yhteenveto: Siivottu tietokantakysely, jota hyödynnetään OLE DB -yhteyksissä (esim. Excelin Power Queryn kautta).

    🚨 Kriittiset löydökset (Tietoturva & Vakaus):

        Ei löydöksiä. Ongelmallinen, pelkästään Accessin sisäinen Nz() -funktio on havaittu ja raportoitu dokumentaatiossa korvattavaksi. Tämä ratkaisee tunnetun "Undefined function in expression" -OLE DB -bugin.

    ⚠️ Huomioitavaa (Toiminnallisuus & Suorituskyky):

        SQL-kyselyssä käytetään raskaasti merkkijonojen manipulointia (Left([Rev],InStr([Rev],"/")-1) AS RevNum). Jos aineisto kasvaa suureksi, nämä hidastavat kyselyä, koska niitä ei voida indeksoida.

    💡 Parannusehdotukset (Ylläpidettävyys):

        Harkitse revision ja sen aputietojen tallentamista omiin sarakkeisiinsa kantaan (RevNum, RevDate) sen sijaan, että ne parsitaan lennosta monimutkaisista stringeistä. Se nopeuttaa tiedonhakua eksponentiaalisesti, kun datamäärä ylittää 10 000 riviä.

2. For ACAD Utility.bas

    📊 Yhteenveto: API-määrityksiä AutoCAD-integraatiolle.

    ⚠️ Huomioitavaa: Hiiren sijaintia (GetCursorPos) hakevan API-funktion deklarointi (vrt. 64-bit PtrSafe) on tehty oikeaoppisesti #If VBA7 -ehtolauseella (kuten MIGRATION_SUMMARY-dokumettikin toteaa). Tyypit on sijoitettu moduulitasolle (ratkaistu bugi), joten tämä moduuli on teknisesti todella vakaa 64-bit ympäristöihin.

Review Agentin Loppusanat: Koodikanta on käynyt läpi hienoja 64-bittiseen ympäristöön siirtymiseen liittyviä päivityksiä. Eniten arvoa saat korjaamalla tietokantakyselyiden transaktionhallinnan ja poistamalla CPU:ta kuormittavat odotusluupit (busy waits). Näillä muutoksilla AutoCAD-automaatio toimii kevyesti ja vakaasti myös raskaassa tuotantokäytössä. Työ on erittäin ammattimaista!
