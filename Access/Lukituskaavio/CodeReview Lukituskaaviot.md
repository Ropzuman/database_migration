1. Form_Funktiokaavio.cls

📊 Yhteenveto: Tämä moduuli hallitsee funktiokaavioiden revisioita ja tietokantapäivityksiä. Koodi on toiminnallista, mutta sisältää kriittisiä tietoturvariskejä SQL-kyselyiden muodostamisessa sekä puutteita virheidenkäsittelyssä globaalien asetusten (SetWarnings) osana.

🚨 Kriittiset löydökset (Tietoturva & Vakaus):

    SQL-injektioalttius: Command50_Click ja Command83_Click -aliohjelmissa SQL-lauseet muodostetaan yhdistämällä merkkijonoja suoraan käyttöliittymän kentistä (esim. Me!Muokkaus0.VALUE). Tämä on vakava tietoturvariski.

    DoCmd.SetWarnings -riski: Koodi kytkee varoitukset pois päältä (DoCmd.SetWarnings False). Jos koodi kaatuu ennen kuin ne kytketään takaisin, Access ei enää varoita käyttäjää esimerkiksi tietojen poistamisesta muissakaan toiminnoissa.

⚠️ Huomioitavaa (Toiminnallisuus & Suorituskyky):

    Kova koodatut päivämääräformaatit: revpaiva = paiva & "." & kk & "." & vuosi voi aiheuttaa ongelmia, jos tietokannan kieliasetukset odottavat ISO-standardia tai eri erotinta.

💡 Parannusehdotukset:

    Käytä parametrisoituja kyselyitä tai Replace(str, "'", "''") -funktiota syötteiden puhdistamiseen.

    Siirrä DoCmd.SetWarnings True myös ErrorHandler-lohkoon varmuuden vuoksi.

🛠️ Korjattu koodiesimerkki (SQL-suojaus):
VBA

' Alkuperäinen haavoittuva koodi:
' CurrentDb.Execute "UPDATE Control SET Sel = '" & KAYTTAJA() & "' WHERE AreaCode = """ & Me!Muokkaus0.VALUE & """ ..."

' Korjattu ja turvallisempi tapa (VBA-ympäristössä):
Dim strSQL As String
strSQL = "UPDATE Control SET Sel = '" & Replace(KAYTTAJA(), "'", "''") & "' " & _
         "WHERE AreaCode = '" & Replace(Me!Muokkaus0.Value, "'", "''") & "' " &_
         "And LoopNo = '" & Replace(Me!Muokkaus4.Value, "'", "''") & "'"
CurrentDb.Execute strSQL, dbFailOnError

Perustelu: dbFailOnError varmistaa, että virheet tulevat esiin ilman SetWarnings-kikkailua, ja Replace estää heittomerkkipohjaiset SQL-injektiot.
2. Form_Interlocking.cls

📊 Yhteenveto: Moduuli vastaa lukituskaavioiden (Interlocking) attribuuttien lukemisesta ja kirjoittamisesta AutoCAD-blokkeihin. Se käyttää Late Binding -tekniikkaa, mikä on hyvä ratkaisu yhteensopivuuden (32-bit/64-bit) kannalta.

🚨 Kriittiset löydökset (Tietoturva & Vakaus):

    Objektien hallinta: LueAttribuutit-funktiossa ja muissa CAD-rutiineissa AutoCAD-objekteja (kuten oDOC, oACAD) käytetään globaalisti, mutta niiden tilan tarkistus (is Nothing) on puutteellista ennen kutsuja.

    Resurssien vapautus: Command151_Click-aliohjelmassa avataan useita DWG-tiedostoja loopissa. Jos virhe tapahtuu loopin sisällä, tiedostot saattavat jäädä auki muistiin, mikä johtaa lopulta resurssien loppumiseen.

⚠️ Huomioitavaa (Toiminnallisuus & Suorituskyky):

    DLookup loopin sisällä: Command151_Click käyttää DLookup-funktiota loopin sisällä tai toistuvasti. Tämä on erittäin hidasta suurilla tietomäärillä.

💡 Parannusehdotukset:

    Lataa asetukset kertaalleen muuttujiin tai Recordsetiin ennen looppia suorituskyvyn parantamiseksi.

    Lisää oDOC.Close False virheenkäsittelijään.

3. APIKoodit.bas

📊 Yhteenveto: Windows API -kutsuja sisältävä kirjastomoduuli. Moduuli on hyvin päivitetty tukemaan 64-bittistä VBA7-ympäristöä PtrSafe-avainsanoilla ja LongPtr-tyypeillä.

🚨 Kriittiset löydökset (Tietoturva & Vakaus):

    Puskurin ylivuotoriski: wu_GetUserName-kutsussa käytetään nSize-muuttujaa. On varmistettava, että puskuri (lpBuffer) on alustettu riittävän suureksi (kuten Space$(255)) ennen kutsua, jotta API ei kirjoita muistialueen yli.

💡 Parannusehdotukset:

    Moduulissa on kommentti "KORJATTU: GetUserNameA kirjoittaa DWORD:n...". Tämä on erinomaista huomiointia API-tyyppien eroista.

4. Form_Aloitus.cls & Form_FromTo.cls

📊 Yhteenveto: Nämä moduulit hoitavat tietokantojen linkitystä ja "From-To" -yhteyksien laskentaa.

🚨 Kriittiset löydökset:

    Suorituskyky (N+1 ongelma): Form_FromTo.cls sisältää sisäkkäisiä looppeja (Do Until qry.EOF -> Do Until qry2.EOF), jotka käyvät läpi recordsettejä. Jos molemmissa on 1000 tietuetta, operaatioita tehdään 1 000 000.

        Ratkaisu: Tämä tulisi hoitaa yhdellä SQL-kyselyllä (JOIN) recordsettien sijaan.

⚠️ Huomioitavaa:

    Kiinteät polut: APIKoodit.bas ja Form_Aloitus.cls sisältävät kiinteitä polkuja kuten "K:\PROJECTS\". Nämä tulisi siirtää asetustauluun.

🛠️ Yleinen korjausehdotus (Markdown formaatissa IDE:lle)
Markdown

# Code Review Improvements - Phase 1

## SQL Injection Prevention

In all modules (especially `Form_Funktiokaavio`), replace direct string concatenation in SQL:

- **Risk:** High (OWASP A03:2021)
- **Fix:** Use a helper function for sanitization.

## Resource Management (AutoCAD)

Ensure CAD documents are closed in Error Handlers:

```vba
ErrorHandler:
    If Not oDOC Is Nothing Then oDOC.Close False
    MsgBox Err.Description

Performance Optimization

Replace nested Recordset loops in Form_FromTo with SQL JOIN operations to avoid O(n²) complexity.

1. Koodit.bas (Apuohjelmamoduuli)

📊 Yhteenveto: Tämä on kriittinen moduuli, joka sisältää sovelluksen keskeisimmät AutoCAD-ohjauslogiikat ja tietokantayhteyksien hallinnan (linkitykset). Koodi on päivitetty käyttämään Late Bindingia, mikä on erinomainen valinta 64-bittisen Officen ja eri AutoCAD-versioiden yhteensopivuuden kannalta.

🚨 Kriittiset löydökset (Tietoturva & Vakaus):

    SQL-injektio tietokantatoiminnoissa: KillLinks-aliohjelmassa suoritetaan CurrentDb.Execute ("DROP TABLE " & T.Name). Vaikka tässä käydään läpi olemassa olevia tauluja, on hyvä käytäntö suojata taulunimet hakasulkeilla siltä varalta, että nimeämisessä on käytetty erikoismerkkejä tai varattuja sanoja.

    AutoCAD-instanssin hallinta: AvaaBlock-funktiossa luotetaan oACAD-objektin olemassaoloon. Jos AutoCAD on suljettu välissä, koodi kaatuu.

⚠️ Huomioitavaa (Toiminnallisuus):

    AppActivate-riski: AppActivate "AutoCAD" voi epäonnistua, jos useita AutoCAD-ikkunoita on auki tai jos ikkunan otsikko ei täsmää tarkasti. Tämä aiheuttaa usein "Run-time error 5" -virheen.

💡 Parannusehdotukset:

    Käytä On Error Resume Next -rakennetta tarkistamaan, onko oACAD edelleen hengissä ennen sen käyttöä, ja alusta se tarvittaessa uudelleen.

2. Form_Linkkien vaihto.cls

📊 Yhteenveto: Moduuli huolehtii MS Access -taulujen dynaamisesta uudelleenlinkityksestä. Se on elintärkeä sovelluksen siirrettävyyden kannalta.

🚨 Kriittiset löydökset (Vakaus):

    Kova koodattu logiikka: If Taul.Fields("Name") <> Taulu Then. Muuttujaa Taulu ei alusteta ennen looppia, jolloin ensimmäinen iteraatio riippuu oletusarvosta.

    Virheiden ohittaminen: On Error Resume Next -lauseen käyttö DoCmd.TransferDatabase-kutsun ympärillä on vaarallista, jos linkitys epäonnistuu esimerkiksi lukitun tiedoston vuoksi; käyttäjä ei saa tästä ilmoitusta.

⚠️ Huomioitavaa:

    Koodi käy läpi MSysObjects-taulun. Tämä vaatii luku-oikeudet järjestelmätauluihin, mikä voi olla estetty joissakin tiukoissa IT-ympäristöissä.

3. TOACAD-alilomakkeet (Loops, Motors, Sekvens)

📊 Yhteenveto: Nämä lomakkeet toimivat liittymänä tietokannan ja AutoCAD-lomakkeen välillä. Niiden tehtävänä on päivittää emolomakkeen (Form_Interlocking) kenttiä valitun tietueen perusteella.

⚠️ Huomioitavaa (Toiminnallisuus & Ylläpidettävyys):

    Riippuvuus emolomakkeesta: Me.Parent.TyhjennaKentat ja Form_Interlocking.OldText luovat tiukan kytkennän (tight coupling). Jos alilomaketta käytetään toisessa yhteydessä, koodi kaatuu.

    Toisto (DRY-periaate): Kaikissa neljässä alilomakkeessa on lähes identtinen VaihdaTiedot-metodi.

💡 Parannusehdotukset:

    Logiikka kannattaisi siirtää yhteen paikkaan emolomakkeelle, jota alilomakkeet kutsuvat, tai käyttää Interface-tyyppistä ratkaisua.

🛠️ Korjattu koodi (Esimerkkejä)
Turvallisempi KillLinks (Koodit.bas):
VBA

' Perustelu: Hakasulkeet [] estävät virheet, jos taulun nimessä on välilyöntejä tai se on varattu sana.
Public Sub KillLinks()
    Dim T As DAO.TableDef
    On Error GoTo ErrorHandler
    
    ' Käytetään takaperin looppia tai kerätään nimet, jotta kokoelman muuttaminen ei sotke iterointia
    Dim i As Integer
    For i = CurrentDb.TableDefs.Count - 1 To 0 Step -1
        Set T = CurrentDb.TableDefs(i)
        If Len(T.Connect) > 0 And Left$(T.Name, 1) <> "~" Then
            CurrentDb.Execute "DROP TABLE [" & T.Name & "]", dbFailOnError
        End If
    Next i
    Exit Sub
ErrorHandler:
    MsgBox "Linkkien poisto epäonnistui: " & Err.Description, vbExclamation
End Sub

Parannettu AutoCAD-aktivointi:
VBA

' Perustelu: AppActivate on epävarma. Käytetään virheenkäsittelyä.
On Error Resume Next
AppActivate oACAD.Caption
If Err.Number <> 0 Then
    ' Jos otsikolla ei löydy, kokeillaan yleisnimeä
    Err.Clear
    AppActivate "AutoCAD"
End If
On Error GoTo ErrorHandler
