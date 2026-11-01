# Module 1.bas muutosloki

- **Tiedosto:** `Module 1.bas` (Excel\Moduulit\AcadDATA)
- **Päivämäärä:** 2026-11-01

## Kriittiset muutokset

- Poistettu `BlockArray`-muuttujan kaksoismääritys `TuoDATA`:ssa (käännösvirhe "Duplicate declaration").
- `Filtterit`-parametri on nyt käytössä: `RefNumerointi` tuo vain `REFERENCE`-blokit (aiemmin parametria ei käytetty).
- Excelin asetukset (`ScreenUpdating`, `EnableEvents`, `Calculation`) talletetaan ennen virheenkäsittelyn alkua ja palautetaan aina `Cleanup`-osassa. Pakotettu `CalculateFull` poistettu.
- `Numerointi` ja `RefNumerointi` palauttavat käyttäjän alkuperäiset asetukset (ei pakotettua `Automatic`/`True`).
- `RefNumerointi`: kirjain pysähtyy Z:hen ja yli 26 viitettä antaa virheilmoituksen.
- `VieDATA` kirjoitettu uudelleen:
  - tyhjä Excel-solu ei koskaan tyhjennä AutoCAD-attribuuttia; tarkoituksellinen tyhjennys merkinnällä `<TYHJÄ>` (`TYHJENNYS_MERKKI`),
  - koko taulukko luetaan muistiin yhdellä kutsulla (`Value2`), virhearvot käsitellään tyhjinä, Excelin soluihin ei kirjoiteta,
  - saman nimiset tagit tunnistetaan (vain ensimmäinen päivitetään),
  - jokainen kirjoitettu arvo luetaan takaisin ja verrataan; yhteenveto lopuksi,
  - otsikkosarakkeet luetaan myös yli sarakkeen IV,
  - puuttuva kahva (poistettu entiteetti) ohitetaan keskeyttämättä vientiä.
- Piirustusten tallennus: edellinen piirustus tallennetaan aina vaihdettaessa (`AvaaDoc`) ja suljetaan vain, jos makro avasi sen. Virhetilanteessa mitään ei tallenneta eikä suljeta.
- `PoistaBlokit`: tallennus vain, jos jotain poistettiin; tuki usean alueen valinnalle; tarkistus, että valinta on solualue.
- Taulukkotuonti (`TuoDATA`) palautettu: täsmälleen oikean kokoinen kirjoitustaulukko (blokit + tekstit), kohdealueen tyhjyystarkistus (`CountA`), tekstimuodot asetetaan ennen kirjoitusta, ensimmäisen ja viimeisen kahvan tarkistus kirjoituksen jälkeen. Varatila: `KAYTA_TAULUKKOKIRJOITUSTA = False` kirjoittaa solu kerrallaan.
- Kaikki `Cells`/`Rows`/`Columns`/`Range`-viittaukset osoittavat eksplisiittisesti `DATA`-taulukkoon (`TIEDLISTA` listalle) aktiivisen taulukon sijaan.
- Ei `Declare`-lauseita, `Nz()`-kutsuja eikï¿½ Jet-ajureita; AutoCAD pysyy myï¿½hï¿½isessï¿½ sidonnassa (ei sidottu AutoCAD-versioon). Toimii 32- ja 64-bittisessï¿½ Excelissï¿½.

- Toinen katselmointi (2026-11-01):
  - `LueSolu`: General-muotoisesta solusta luetaan täysi tarkkuus (`.Text` katkoi desimaaleja).
  - `TuoDATA`: virhetilanteessa listalta avattu piirustus suljetaan tallentamatta.
  - `EXCELHAKU`-valintajoukot poistetaan lopusta alkuun (indeksien siirtyminen).
  - Virheilmoitukset tarkennettu: aiemmin käsitellyt piirustukset on jo tallennettu.

## Siivous ja optimointi

- Poistettu käyttämättömät muuttujat (`Jakso`, `BlockNimi`, `k` VieDATA:ssa, `textStartRow`, `j` RefNumeroinnissa) sekä "Dim poistettu" -kommentit.
- Poistettu vanhat korjausmerkinnï¿½t (BUG/Issue-viittaukset) ja toistuvat "Muutettu Integer -> Long" -kommentit.
-- `Dir()` korvattu `TiedostoNimi`-funktiolla; lisätty apufunktiot `LueSolu` ja `TallennaDoc`.
- `OtsS` ei enï¿½ï¿½ silmukoi loputtomasti, vaan antaa virheen, jos sarakkeet loppuvat.
-- Tuplailmoitus "ei löytynyt tietoja" poistettu.
-- Kommentit ja vaihe-ilmoitukset käännetty suomeksi, puuttuvat ä/ö-merkit korjattu, kysymysmerkitiksi muuttuneet nuolet korvattu `->`-merkinnällä.
- Tiedosto tallennettu yhtenäisenä Windows-1252-koodauksena (CRLF) – aiempi sekakoodaus korjattu.
- `Lisaa` ja `Yhd` jätetty paikoilleen, koska ne voivat olla taulukkokaavojen käytössä.

## Huomioitavaa

- `RefNumerointi` lukee sivunumeron solusta D17 ja kirjoittaa seuraavan solusta D18 ï¿½ tarkista, onko tï¿½mï¿½ tarkoituksellista.
- Koodia ei ole kï¿½ï¿½nnetty eikï¿½ ajettu: tuo moduuli uudelleen VBA-editoriin, aja Debug ? Compile ja testaa testikuvalla (tee varmuuskopio).
