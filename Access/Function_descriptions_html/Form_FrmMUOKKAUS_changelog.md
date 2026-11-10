# Muutosloki

- **Tiedosto:** `Form_FrmMUOKKAUS.cls`
- **P‰iv‰m‰‰r‰:** 2026-11-10

## Kriittiset muutokset

- `TeeSelostus`: uusi `OhitaRyhma`-apurutiini siirt‰‰ tietuekursorin LoopTAG-ryhm‰n ohi myˆs virhetilanteessa. Aiemmin sama tietue j‰i k‰sittelyyn ja virheilmoitus toistui loputtomasti. Samalla poistui EOF-virhe (3021) viimeisen ryhm‰n ohituksessa.
- HTML-tuloste: uusi `HtmlEsc` suojaa `&`, `<` ja `>` merkit otsikossa, teksteiss‰ ja h‰lytysriveill‰. `HtmlEscTeksti` s‰ilytt‰‰ painikkeilla lis‰tyt `<h3>`- ja `<h4>`-v‰liotsikot.
- `meta charset` on `windows-1252`, koska FileSystemObject kirjoittaa ANSI-tiedostoja.
- Null-suojaukset: `DescriptionFI`, `FunctionFI` ja `AlarmDescrBL`.
- H‰lytyskuvauksen kielivalinta korjattu (`AlarmDescrBL_FI` suomeksi).
- `CONTROL`-taulun `DLookup`-haut suojattu, joten taulun puuttuminen linkityksist‰ ei kaada Word-luontia.
- `LOOPTAG`-hakuehdon heittomerkit kaksinkerrotaan (SQL).
- Tiedostonimet siistit‰‰n `TiedostoNimi`-funktiolla.
- `Command99`ñ`Command104`: uusi `TFunctionValittu` est‰‰ virheen 91, kun kohdekentt‰‰ ei ole valittu.

## Siivous ja optimointi

- Poistettu kuollut `Samaan`/`Eka`-haara, `oPohja`, vakiot `wdStory` ja `wdPageBreak` sek‰ kaikki kommentoitu vanha koodi (`VakioHaly`, `suffix`/`maara`, `TAGNAME`-rivit).
- Poistettu k‰ytt‰m‰ttˆm‰t muuttujat (`Rivit`, `r`, `Halyt`, `Laitteet` TeeHTML:st‰; `rngStory`, `i` Word-apurutiineista).
- H‰lytyslohkojen sis‰kk‰iset tyhj‰t `If`-haarat yksinkertaistettu.
- `SuljeTiedot` sulkee recordsetin ja HTML-virran sek‰ vapauttaa `oTemplate`- ja `oWord`-viittaukset; `Set FS = Nothing` lis‰tty. Word j‰‰ tarkoituksella auki.
- `Korvaa`, `KorvaaOts` ja `POISTA`: parametrit `ByVal` (ei muuta kutsujan muuttujia).
- P‰iv‰m‰‰r‰ muodostetaan `Format$(Date, "d.m.yyyy")`.
- Kommentit: englanninkieliset k‰‰nnetty, "kirjanmerkki" ? "paikkamerkki" (korvaus tehd‰‰n tekstihaulla), otsikon paikkamerkkiluettelo vastaa koodia.
- J‰tetty ennalleen: tyhj‰ `Form_Load` (lomakkeen tapahtumaominaisuus voi viitata siihen) ja `Suomeksi`-lippu.
