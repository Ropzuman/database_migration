# Code Review Raportti

## 📊 Yhteenveto

[cite_start]Koodi on rakenteeltaan järjestelmällistä ja siinä on selvästi otettu huomioon suorituskyky, mistä on osoituksena näytönpäivitysten hallinta (ns. Fast Mode) [cite: 93] [cite_start]sekä taulukoiden käsittely muistissa jatkuvan levy-/soluoperaatioiden sijaan[cite: 110, 244]. Koodista löytyy kuitenkin merkittäviä arkkitehtuurisia päällekkäisyyksiä ja tietoturvaan liittyviä heikkouksia, jotka keskittyvät erityisesti tietokantayhteyksien ja kyselyiden validointiin.

## 🚨 Kriittiset löydökset (Tietoturva & Vakaus)

* [cite_start]**Heikko SQL-injektiosuojaus (Security):** Koodin `OnTurvallinenSQL`-funktio pyrkii estämään tuhoisat kyselyt käyttämällä *denylist*-menetelmää [etsimällä avainsanoja kuten "DROP", "DELETE", "UPDATE"](cite: 100). OWASP-standardien mukaisesti tällainen suodatus on erittäin herkkä virheille (CWE-89). Menetelmä voidaan ohittaa erilaisilla SQL-syntaksin säännöillä ja varioinneilla, jättäen tietokannan yhä alttiiksi virheelliselle manipuloinnille.
* [cite_start]**Puuttuva Read-Only -pakotus (Security/Vakaus):** Koska ohjelman on tarkoitus vain hakea tietoa kytkentälistoja varten, DAO-tietokantayhteys tulisi tietoturvan nimissä avata vain luku -tilassa [Read-Only](cite: 106). Tekstipohjaisen filtteröinnin sijaan on paljon turvallisempaa antaa tietokantamoottorin itsensä estää kantaa muuttavat kyselyt suoraan arkkitehtuurin tasolla.
* [cite_start]**Käyttöliittymän lukittumisriski (Vakaus):** Makro laittaa Excelin "Fast Mode" -tilaan maksimoidakseen suorituskyvyn[cite: 93]. [cite_start]Mikäli rutiini kaatuu ajonaikaiseen virheeseen, jota `ErrorHandler` ei saavuta tai joka keskeyttää suorituksen yllättäen, `EndFastMode` saattaa jäädä ajamatta[cite: 94]. Tällöin käyttäjän Excel jää täysin jumiin (automaattilaskenta ja näytönpäivitys pois päältä).

## ⚠️ Huomioitavaa (Toiminnallisuus & Suorituskyky)

* [cite_start]**DAO- ja ADODB-kirjastojen sekakäyttö (Suorituskyky & Ylläpidettävyys):** Makrossa avataan samanaikaisesti yhteydet tietokantaan käyttämällä kahta eri rajapintaa [DAO ja ADODB](cite: 106, 33). [cite_start]Vaikka koodin kommentin mukaan DAO:a tarvitaan tallennetuille kyselyille[cite: 103], myös ADODB tukee natiivisti tallennettujen Access-kyselyiden suorittamista. Kahden päällekkäisen rajapinnan käyttö kuluttaa turhaan resursseja ja tekee ylläpidosta raskasta.
* [cite_start]**Rivinvaihtojen pilkkominen (Toiminnallisuus):** Metadatan haussa revision osat erotetaan taulukoksi `Split(DIRev, Chr(10))` -kutsulla. Koska datassa voi olla Windowsin normaaleja CR+LF (`Chr(13)` + `Chr(10)`) rivinvaihtoja, merkkijonojen perään jää todennäköisesti näkymättömiä Carriage Return -merkkejä (`Chr(13)`), jotka voivat aiheuttaa vertailuvirheitä ja asetteluongelmia Excelissä.
* [cite_start]**Recordsetin manuaalinen purkaminen (Suorituskyky):** Fix3-kohdassa tehty koodi, joka siirtää DAO:n tulosrivin solutasolta matriisiin (`dataArr`) on jo iso parannus. Se on kuitenkin yhä huomattavasti hitaampi ja monimutkaisempi ratkaisu verrattuna Excelin VBA:n sisäänrakennettuun `.CopyFromRecordset`-menetelmään.

## 💡 Parannusehdotukset (Ylläpidettävyys)

* [cite_start]**Globaalien muuttujien siivous (Clean Code):** `Module1.bas` -tiedoston alussa on julistettu yli 30 erillistä `Public`-muuttujaa[cite: 90]. Valtava määrä globaaleja tilamuuttujia aiheuttaa nopeasti vaikeasti jäljitettäviä "spagettikoodin" piirteitä. Harkitse vahvasti näiden kapseloimista User Defined Type (UDT) -tietueeseen (esim. `Type DocumentInfo`) tai omaan Luokkamoduuliin (Class Module).
* [cite_start]**Kovakoodatut soluviittaukset:** Käyttöliittymädatan haku, kuten `Sheets("Main").Cells(8 + Valinta, 3).Value`[cite: 105], on ns. *brittle*-koodia eli se menee helposti rikki, jos joku lisää Excel-välilehdelle yhdenkin ylimääräisen visuaalisen välirivin. Käytä Excelin nimettyjä alueita (Named Ranges) turvataksesi kestävyyden.
* **Error Resume Next -käytäntö:** Moduuleissa käytetään laajalti `On Error Resume Next` -lausekkeita ohittamaan tarkistuksia. Sitä kannattaa käyttää harkiten, sillä se voi piilottaa taustalla tapahtuvat kriittiset virheet.

## 🛠️ Korjattu koodi

### 1. Luotettava tietoturva: DAO-yhteyden avaaminen Read-Only -tilassa

[cite_start]Voit poistaa monimutkaisen ja vuotavan `OnTurvallinenSQL` -funktion kokonaan, kun pakotat tietokantayhteyden tapahtumaan puhtaasti lukutilassa[cite: 100, 106]. Access estää tällöin kaikki injektiot, jotka pyrkivät muuttamaan kantaa.

```vba
' Muutos tiedostossa Module1.bas
' Alkuperäinen implementaatio: Set dbDAO = CreateObject("DAO.DBEngine.120").OpenDatabase(Kanta)

' KORJATTU TAPA:
' Avataan yhteys vain luku (ReadOnly:=True) tilassa. 
' Tietokantamoottori itse heittää virheen turvallisesti, jos joku yrittää ajaa DROP/DELETE/UPDATE -kyselyitä.
On Error Resume Next
Set dbDAO = CreateObject("DAO.DBEngine.120").OpenDatabase(Name:=Kanta, Options:=False, ReadOnly:=True)
If Err.Number <> 0 Then
    MsgBox "Tietokannan avaaminen epäonnistui.", vbCritical
    GoTo SafeExit
End If
On Error GoTo ErrorHandler

2. Datan viennin massiivinen optimointi (CopyFromRecordset)

Voit korvata Module1:ssä olevan Do-While -silmukan natiivilla ja erittäin nopealla Excel-kutsulla. Tämä nopeuttaa makroa etenkin isoilla kyselytuloksilla.

' KORJATTU: C++ -optimoitu datan siirto suoraan Recordsetistä työkirjaan.
If Not rsDAO.EOF Then
    ' Otsikot kuten ennenkin
    colData = 1
    For Each fldDAO In rsDAO.Fields
        ws.Cells(1, colData).Value = fldDAO.Name
        colData = colData + 1
    Next fldDAO
    
    ' SIIVOTTU KOODI: Koko aiempi dataArr-taulukon luonti ja Do-While-looppi voidaan 
    ' korvata yhdellä suorituskykyisemmällä rivillä:
    ws.Range("A2").CopyFromRecordset rsDAO
End If

3. Rivinvaihtojen turvallinen pilkkominen

Muokkaa Module2.bas -tiedoston merkkijonon käsittelyä, jotta mahdolliset "Carriage Return" -jäänteet (Chr(13)) puhdistuvat.

' Muutos tiedostossa Module2.bas (Case "rev")
Case "rev"
    DIRev = CStr(valArr(1, i) & "")
    Erase DIRevArr
    
    If Len(DIRev) > 0 Then
        ' KORJATTU: Siivotaan piilevä Chr(13) (Carriage Return) pois ennen merkkijonon jakamista
        DIRev = Replace(DIRev, vbCr, "") 
        DIRevArr = Split(DIRev, vbLf)
    Else
        ReDim DIRevArr(0): DIRevArr(0) = ""
    End If
