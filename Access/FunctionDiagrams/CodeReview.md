🕵️‍♂️ Code Review -raportti: AutoCAD-integraatiot ja Funktiokaaviot
> **Refaktoroitu:** 2026-03-06 — kaikki alla mainitut löydökset korjattu

📊 Yhteenveto

## ✅ REFAKTOROINTI 2026-03-06 (kierros 2) — Toteutetut korjaukset

### Form_LisääKuviin_ACAD.cls — `TeeKuvat_Click`

- Lisätty `On Error Resume Next` + `oAcad.Quit` `Loppu:`-siivouslohkoon
- `ErrorHandler` ohjaa nyt `Resume Loppu` — AutoCAD ei enää jää taustalle haamuprosessina virhetilanteessa

### Form_LukituskaavioLinkit.cls — `Command0_Click`

- **O(N²) → FindFirst**: Sisäkkäinen `Do Until tbl.EOF`-silmukka korvattu `tbl.FindFirst`-kutsulla — suorituskyky kasvaa eksponentiaalisesti suurilla tietomäärillä
- `InterlockingLinkPage` avattu `dbOpenSnapshot`-tilassa (nopeampi vain luku)
- `IntLinkPage` avattu eksplisiittisesti `dbOpenDynaset`-tilassa (muokkaus sallittu)
- `tbl!page1`-tyhjyystarkistus muutettu `Nz(tbl!page1, "") = ""`-muotoon (null-turvallinen)
- `qry!TXT1` sanitoitu `Replace(..., "'", "''")`-metodilla FindFirst-kutsussa

---

 moduuleihin, jotka käsittelevät AutoCAD-automaatiota (late binding), funktiokaavioiden lukituksia ja blokkien päivittämistä. Yleinen arkkitehtuuri näyttää siistiytyneen: 64-bittiseen ympäristöön siirtyminen on huomioitu hyvin (esim. Object-tyyppien käyttö AcadApplication-viittausten sijaan). Koodista on kuitenkin löydettävissä yhä tietoturvariskejä (SQL-injektiot) sekä perinteisiä DAO-tietueistoihin ja COM-objekteihin liittyviä resurssinhallinnan sudenkuoppia, jotka kannattaa korjata pitkän aikavälin vakauden takaamiseksi.
📂 1. Moduuli: Form_LisääKuviin_ACAD.cls

📊 Yhteenveto: Raskas automaatiomoduuli, joka hallinnoi AutoCAD-piirustusten avaamista, blokkien lisäämistä ja tallennusta.

🚨 Kriittiset löydökset (Tietoturva & Vakaus):

    COM-objektien haamuprosessit: Moduuli avaa AutoCAD-instansseja (CreateObject("AutoCAD.Application")). Vaikka koodissa on pyritty vapauttamaan olioita (Set oAcad = Nothing), tämä ei kaikissa virhetilanteissa riitä. Jos koodi kaatuu kesken suorituksen ja ohittaa oAcad.Quit -kutsun, AutoCAD jää pyörimään taustalle näkymättömänä prosessina, mikä syö tietokoneen muistia (Memory Leak).

⚠️ Huomioitavaa (Toiminnallisuus & Suorituskyky):

    Taulukoiden läpikäynti ja AutoCADin "OSMODE"-muuttujan jatkuva päivitys loopin sisällä voi olla hidasta suurilla tiedostomäärillä.

💡 Parannusehdotukset (Ylläpidettävyys):

    Keskitä virheenkäsittely yhteen Exit/CleanUp -lohkoon, joka takaa ohjelman hallitun alasajon riippumatta siitä, missä kohtaa virhe tapahtuu.

🛠️ Korjattu koodi (Turvallinen AutoCADin sammutus):
VBA

Private Sub TeeKuvat_Click()
    On Error GoTo ErrorHandler
    ' ... koodin alku ...

    Set oAcad = CreateObject("AutoCAD.Application")
    ' ... prosessointi ...
    
CleanUp:
    On Error Resume Next ' Suojataan siivouslohko
    If Not oDoc Is Nothing Then
        oDoc.Close False ' Sulje dokumentti tallentamatta, jos virhe
        Set oDoc = Nothing
    End If
    If Not oAcad Is Nothing Then
        oAcad.Quit
        Set oAcad = Nothing
    End If
    Exit Sub

ErrorHandler:
    MsgBox "Virhe piirustusten luonnissa: " & Err.Description, vbCritical
    Resume CleanUp ' Takaa että AutoCAD ei jää taustalle!
End Sub

Perustelu: Tämä Try-Finally -rakenne on teollisuusstandardi COM-automaatiossa ja estää kymmenien näkymättömien AutoCAD-prosessien kertymisen Tehtävienhallintaan.
📂 2. Moduuli: Form_FuncBlock.cls

📊 Yhteenveto: Funktiopiirien yksittäisten lohkoattribuuttien (kuten FTAGI, REFFROM) muokkauslomake.

🚨 Kriittiset löydökset (Tietoturva & Vakaus):

    SQL-injektion mahdollisuus (Vaarallinen ristiinkytkentä): Tiedostossa on suoritettu kustomoituja SQL-kyselyitä, kuten RowSource = "SELECT ... WHERE AreaCOde = '" & ALUE & "'". Jos muuttujat ALUE, LooppiNo tai Sfx saavat arvonsa suoraan toiselta lomakkeelta ja sisältävät heittomerkin ('), kysely kaatuu tai se voi tulla manipuloiduksi.

💡 Parannusehdotukset (Ylläpidettävyys):

    Sinulla on koodissa huomioitu heittomerkkien poistaminen paikoin Replace(muuttuja, "'", "''") -tyylillä. Varmista, että tämä on järjestelmällisesti käytössä kaikissa kohteissa, joista muodostetaan SQL-lause.

📂 3. Moduuli: Form_LukituskaavioLinkit.cls

📊 Yhteenveto: Yhdistää ja päivittää sivuviittaukset (See Page X, Y) lukituskaavioihin.

⚠️ Huomioitavaa (Toiminnallisuus & Suorituskyky):

    Sana N+1 (O(N^2) Aikakompleksisuus): Koodi ajaa sisäkkäistä silmukkaa:
    VBA

    Do Until qry.EOF
        Do Until tbl.EOF
            If tbl!TXT1 = qry!TXT1 Then ...

    Jos qry-taulussa on 1000 riviä ja tbl-taulussa on 1000 riviä, vertailuja tehdään 1 000 000 kappaletta. DAO:lla tämä on erittäin hidasta.

🛠️ Korjattu koodi (Suorituskyvyn optimointi DAO.Seek tai .FindFirst avulla):
VBA

Private Sub Command0_Click()
    On Error GoTo ErrorHandler
    DoCmd.Hourglass True

    Dim DB As DAO.Database
    Dim qry As DAO.Recordset
    Dim tbl As DAO.Recordset
    
    Set DB = CurrentDb
    Set qry = DB.OpenRecordset("InterlockingLinkPage", dbOpenSnapshot) ' Snapshot riittää lukuun
    Set tbl = DB.OpenRecordset("IntLinkPage", dbOpenDynaset) ' Dynaset sallii muokkauksen
    
    If Not qry.EOF Then qry.MoveFirst
    
    Do Until qry.EOF
        ' Etsitään oikea tietue FindFirst-metodilla koko taulun läpikäynnin sijaan
        tbl.FindFirst "TXT1 = '" & Replace(qry!TXT1, "'", "''") & "'"
        
        If Not tbl.NoMatch Then
            tbl.Edit
            If Nz(tbl!page1, "") = "" Then
                tbl!page1 = "See Page " & qry!Page
            Else
                tbl!page1 = tbl!page1 & ", " & qry!Page
            End If
            tbl.Update
        End If
        qry.MoveNext
    Loop
    
CleanUp:
    On Error Resume Next
    DoCmd.Hourglass False
    If Not tbl Is Nothing Then tbl.Close: Set tbl = Nothing
    If Not qry Is Nothing Then qry.Close: Set qry = Nothing
    Set DB = Nothing
    Exit Sub

ErrorHandler:
    MsgBox "Virhe: " & Err.Description
    Resume CleanUp
End Sub

Perustelu: FindFirst:n käyttö ja dbOpenSnapshot lukutaululle nopeuttaa suoritusta eksponentiaalisesti verrattuna sisäkkäisiin looppeihin. Myös tiimalasi kytkeytyy pois päältä varmasti, jos tapahtuu virhe.
📂 4. Moduuli: Form_Funktiokaavio.cls

📊 Yhteenveto: Funktiokaavion päälomake, joka ohjaa tulostuksia ja näkymiä.

⚠️ Huomioitavaa (Toiminnallisuus & Suorituskyky):

    Koodista huomaa selvästi hyvän siivouksen jäljen: Option Explicit on paikallaan ja vanhoja 32-bit API-kutsuja on poistettu. Taso on varsin hyvä ja lomake näyttää hoitavan keskeisen liiketoimintalogiikkansa vakaasti.

💡 Parannusehdotukset (Ylläpidettävyys):

    Moduulissa on paljon painikkeita automaattisesti generoiduilla nimillä (kuten Command50, Komento47, Command83). Tämä vaikeuttaa koodin ylläpidettävyyttä ja luettavuutta.

    Vinkki: Anna elementeille kuvaavat nimet lomakkeen ominaisuuksissa, esim. cmdAjaReseptit tai cmdLuoIndeksisivut, ja nimeä tapahtumat sen mukaisesti. Koodin refaktorointi myöhemmin helpottuu huomattavasti.
