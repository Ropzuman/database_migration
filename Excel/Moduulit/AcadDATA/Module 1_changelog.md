# Module 1.bas muutosloki

- **Tiedosto:** `Module 1.bas` (Excel\Moduulit\AcadDATA)
- **P‰iv‰m‰‰r‰:** 2026-11-01

## Kriittiset muutokset

- Poistettu `BlockArray`-muuttujan kaksoism‰‰ritys `TuoDATA`:ssa (k‰‰nnˆsvirhe "Duplicate declaration").
- `Filtterit`-parametri on nyt k‰ytˆss‰: `RefNumerointi` tuo vain `REFERENCE`-blokit (aiemmin parametria ei k‰ytetty).
- Excelin asetukset (`ScreenUpdating`, `EnableEvents`, `Calculation`) talletetaan ennen virheenk‰sittelyn alkua ja palautetaan aina `Cleanup`-osassa. Pakotettu `CalculateFull` poistettu.
- `Numerointi` ja `RefNumerointi` palauttavat k‰ytt‰j‰n alkuper‰iset asetukset (ei pakotettua `Automatic`/`True`).
- `RefNumerointi`: kirjain pys‰htyy Z:hen ja yli 26 viitett‰ antaa virheilmoituksen.
- `VieDATA` kirjoitettu uudelleen:
  - tyhj‰ Excel-solu ei koskaan tyhjenn‰ AutoCAD-attribuuttia; tarkoituksellinen tyhjennys merkinn‰ll‰ `<TYHJƒ>` (`TYHJENNYS_MERKKI`),
  - koko taulukko luetaan muistiin yhdell‰ kutsulla (`Value2`), virhearvot k‰sitell‰‰n tyhjin‰, Excelin soluihin ei kirjoiteta,
  - saman nimiset tagit tunnistetaan (vain ensimm‰inen p‰ivitet‰‰n),
  - jokainen kirjoitettu arvo luetaan takaisin ja verrataan; yhteenveto lopuksi,
  - otsikkosarakkeet luetaan myˆs yli sarakkeen IV,
  - puuttuva kahva (poistettu entiteetti) ohitetaan keskeytt‰m‰tt‰ vienti‰.
- Piirustusten tallennus: edellinen piirustus tallennetaan aina vaihdettaessa (`AvaaDoc`) ja suljetaan vain, jos makro avasi sen. Virhetilanteessa mit‰‰n ei tallenneta eik‰ suljeta.
- `PoistaBlokit`: tallennus vain, jos jotain poistettiin; tuki usean alueen valinnalle; tarkistus, ett‰ valinta on solualue.
- Taulukkotuonti (`TuoDATA`) palautettu: t‰sm‰lleen oikean kokoinen kirjoitustaulukko (blokit + tekstit), kohdealueen tyhjyystarkistus (`CountA`), tekstimuodot asetetaan ennen kirjoitusta, ensimm‰isen ja viimeisen kahvan tarkistus kirjoituksen j‰lkeen. Varatila: `KAYTA_TAULUKKOKIRJOITUSTA = False` kirjoittaa solu kerrallaan.
- Kaikki `Cells`/`Rows`/`Columns`/`Range`-viittaukset osoittavat eksplisiittisesti `DATA`-taulukkoon (`TIEDLISTA` listalle) aktiivisen taulukon sijaan.
- Ei `Declare`-lauseita, `Nz()`-kutsuja eik‰ Jet-ajureita; AutoCAD pysyy myˆh‰isess‰ sidonnassa (ei sidottu AutoCAD-versioon). Toimii 32- ja 64-bittisess‰ Exceliss‰.

## Siivous ja optimointi

- Poistettu k‰ytt‰m‰ttˆm‰t muuttujat (`Jakso`, `BlockNimi`, `k` VieDATA:ssa, `textStartRow`, `j` RefNumeroinnissa) sek‰ "Dim poistettu" -kommentit.
- Poistettu vanhat korjausmerkinn‰t (BUG/Issue-viittaukset) ja toistuvat "Muutettu Integer -> Long" -kommentit.
- `Dir()` korvattu `TiedostoNimi`-funktiolla; lis√§tty apufunktiot `LueSolu` ja `TallennaDoc`.
- `OtsS` ei en‰‰ silmukoi loputtomasti, vaan antaa virheen, jos sarakkeet loppuvat.
- Tuplailmoitus "ei l√∂ytynyt tietoja" poistettu.
- Kommentit ja vaihe-ilmoitukset k√§√§nnetty suomeksi, puuttuvat √§/√∂-merkit korjattu, kysymysmerkeiksi muuttuneet nuolet korvattu `->`-merkinn√§ll√§.
- Tiedosto tallennettu yhten√§isen√§ Windows-1252-koodauksena (CRLF) ‚Äì aiempi sekakoodaus korjattu.
- `Lisaa` ja `Yhd` j√§tetty paikoilleen, koska ne voivat olla taulukkokaavojen k√§yt√∂ss√§.

## Huomioitavaa

- `RefNumerointi` lukee sivunumeron solusta D17 ja kirjoittaa seuraavan solusta D18 ñ tarkista, onko t‰m‰ tarkoituksellista.
- Koodia ei ole k‰‰nnetty eik‰ ajettu: tuo moduuli uudelleen VBA-editoriin, aja Debug ? Compile ja testaa testikuvalla (tee varmuuskopio).
