# Vastaanottajavalintojen käyttöönotto

1. Ota varmuuskopio tietokannasta ja varmista, ettei sitä käytetä asennuksen aikana.
2. Luo `USysDistrSelection` **samaan tietokantaan kuin `USysRecipients`**, jotta kaikki käyttäjät näkevät taulun. Suorita [taulun luontilause](USysDistrSelection.sql) Accessin SQL-näkymässä. Jos sovelluksessa on erillinen käyttöliittymä- ja taustakanta, liitä uusi taulu käyttöliittymään. Jokaisella käyttäjällä pitää olla tauluun luku-, lisäys- ja poisto-oikeus.
3. Päivitä molemmat lomakemoduulit: [Form_USysNewDistribution.cls](Form_USysNewDistribution.cls) ja [Form_USysNewRecipient.cls](Form_USysNewRecipient.cls). Aseta `USysNewDistribution`-lomakkeen **On Unload** -tapahtumaksi **[Event Procedure]**, jotta istunnon valinnat siivotaan myös peruutettaessa. Tarkista VBA-editorissa **Debug → Compile**.
4. Tarkista `USysNewDistribution`-lomakkeen `Lista`-ohjaimen asetukset: sidottu sarake 1, sarakkeiden määrä 9 ja sarakeleveydet `0cm;0cm;2cm;2,524cm;3cm;4cm;4cm`. Koodi asettaa rivilähteen jokaisella avauskerralla ja säilyttää alkuperäisen sarakejärjestyksen. Uusi vastaanottaja valitaan listalta käsin.
5. Testaa kahdella samanaikaisella Access-istunnolla saman vastaanottajan valinta eri kopiomäärillä. Peruuta toinen ja tallenna toinen; varmista, että jäljelle jääneen istunnon valinta ja tallennettu kopiomäärä säilyvät. Testaa myös uuden vastaanottajan lisääminen, tallennusvirhe ja lomakkeen sulkeminen.

Jokaisella lomakkeen avauksella on oma satunnainen istuntotunniste. Valinnat poistetaan tallennuksen yhteydessä tai lomakkeen sulkeutuessa. Jos Access tai verkkoyhteys katkeaa, tauluun voi jäädä rivejä. `CreatedAt` kertoo valinnan paikallisen luontihetken; automaattista ikään perustuvaa siivousta ei tehdä, koska pitkäkestoinen aktiivinen istunto voisi menettää valintansa. Poista vanhoja rivejä käsin vain, kun olet varmistanut, etteivät kyseiset istunnot ole aktiivisia.

Vanhat `USysRecipients`-kentät `To Distribution` ja `SelectedBy` säilyvät, mutta tämä valintatoiminto ei enää käytä niitä. Muiden lomakkeiden, kyselyjen tai raporttien mahdolliset riippuvuudet näistä kentistä pitää tarkistaa varsinaisessa tietokannassa ennen käyttöönottoa.
