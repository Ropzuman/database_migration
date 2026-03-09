# Muutosloki – Listojen kyselyt

**Tiedostot:** `Module1.bas`, `Module2.bas`, `Module3.bas`  
**Päivämäärä:** 2026-03-06  
**Tekijä:** GitHub Copilot (Senior System Architect – 64-bit Migration Agent)  
**Peruste:** Code Review -raportti (`Logs/CodeReviewReport.md`)

---

## Kriittiset muutokset

### Tietoturva – SQL-injektiosuoja (CWE-89) — `Module1.bas`

- Lisätty `OnTurvallinenSQL()`-apufunktio, joka estää vaarallisten DML/DDL-komentojen
  (`DROP`, `DELETE`, `UPDATE`, `INSERT`, `ALTER`, `EXEC`) suorittamisen Excel-soluista
  luetuille dynaamisille SQL-kyselyille.
- Tarkistus ajetaan ennen jokaista `HaeData`-funktion tietokantakutsua (`sSQL(1)` ja `sSQL(2)`).
- Sallii vain `SELECT`-kyselyt ja tallennettujen kyselyiden nimet.

### Vakaus – Try-Finally -virheenkäsittelymalli — `Module1.bas`

- `HaeData`-funktio refaktoroitu käyttämään VBA:n Try-Finally -mallia (`SafeExit`-hyppymerkki).
- `BeginFastMode` tallentaa nyt alkuperäiset Excel-asetukset (`ScreenUpdating`, `Calculation`,
  `EnableEvents`, `DisplayAlerts`, `DisplayStatusBar`) ennen niiden muuttamista.
- `EndFastMode` palauttaa aina kaikki asetukset alkuperäiseen tilaan – myös virheen sattuessa.
- `ErrorHandler` ohjaa aina `SafeExit`-lohkoon, jolloin Excel ei jää jäätymistilaan.
- Resurssien siivous (`rsDAO`, `dbDAO`, `rs`, `conn`) keskitetty `SafeExit`-lohkoon.

### DRY – ADODB-yhteyden avaaminen eristetty apufunktioon — `Module1.bas`

- Lisätty `LuoADODBYhteys(kantaPolku)`-apufunktio, joka kokeilee ACE OLEDB -versioita
  prioriteettijärjestyksessä `16.0 → 15.0 → 12.0`.
- Kolminkertainen toisteinen yhteysyritysrakenne `HaeData`:ssa korvattu yhdellä kutsulla.
- Epäonnistuneet yritykset siivotaan (`Set conn = Nothing`) ennen seuraavaa kokeilua.
- Palauttaa `Nothing` jos kaikki versiot epäonnistuvat (kutsuva koodi käsittelee virheen).

---

## Toiminnallisuuskorjaukset

### `HaeDocTiedot` – Otsikkorivi luetaan yhtenä 2D-taulukkona — `Module2.bas`

- Sarakekohtainen soluiterointilisilmukka korvattu: otsikkorivi ja datarivi luetaan nyt
  yhdellä `Range.Value`-kutsulla `hdrArr`- ja `valArr`-taulukoihin.
- Vähennetty COM-kutsujen määrää merkittävästi suurilla taulukoilla.
- Lisätty `lastCol = wsDB2.Cells(1, wsDB2.Columns.Count).End(xlToLeft).Column`
  viimeisen sarakkeen löytämiseen `MAX_EXCEL_COLUMNS`-iteraation sijaan.

### `DIRev`-splitaus – Null-tarkistus ennen `Split()` — `Module2.bas`

- Lisätty `Len(DIRev) > 0` -tarkistus ennen `Split(DIRev, Chr(10))` -kutsua.
- Estää `Type Mismatch` -virheen (`Error 13`) tilanteessa, jossa solun arvo on tyhjä.

### Duplikaattifunktiot poistettu — `Module2.bas`

- Poistettu toisteiset funktiot, jotka olivat jo olemassa `Module1.bas`:ssä. (27.2.2026)

### `Linking` – Monilinkki-varoitus ja virheenkäsittely — `Module3.bas`

- Lisätty silmukka, joka laskee LINKING-nimisten sheetien määrän ennen näkyvyysmuutosta.
- Jos sheetejä on useampi kuin yksi, näytetään `MsgBox`-varoitus käyttäjälle. (27.2.2026)
- Lisätty `ErrorHandler` estämään hiljaiset virheet sheetin näkyvyyden vaihdossa.

---

## Siivous ja optimointi

- `BeginFastMode` / `EndFastMode` eriytetty omiksi `Private Sub` -prosedureiksi
  selkeyden ja uudelleenkäytettävyyden parantamiseksi.
- Kaikki tietokantamuuttujat (`dbDAO`, `rsDAO`, `conn`, `rs`) alustetaan `Nothing`:ksi
  proseduurin alussa, jotta `SafeExit`-siivous toimii myös haun epäonnistuessa.
- `Option Explicit` pakotettu kaikissa kolmessa moduulissa.
- `MAX_EXCEL_COLUMNS = 16384` -vakio lisätty `Module1.bas`:iin ja `Module2.bas`:iin
  suojaamaan silmukoita ikuiselta iteroinnilta.
- Debug-tulosteet (`Debug.Print`) lisätty avainkohtiin diagnostiikkaa varten.
