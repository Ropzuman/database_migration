# Muutosloki – AcadDATA

**Tiedostot:** `Koodit.bas`, `DATA.bas`, `AcadHelpers.bas`  
**Päivämäärä:** 2026-03-06  
**Tekijä:** GitHub Copilot (Senior System Architect – 64-bit Migration Agent)  
**Peruste:** HTML Code Review -raportti (`acad_code_review.html`) sekä kriittinen bugikorjaus (`VIEDATA_BUG_FIX.md`)

---

## Kriittiset muutokset (Bugit)

### BUG 1 [HIGH] – Teksti/blokki-kirjoitusristiriita `TuoDATA`:ssa — `Koodit.bas`

- TEXT/MTEXT-entiteetit kirjoittivat aiemmin suoraan `Cells(Rivi, ...)`-soluihin silmukan aikana,
  minkä jälkeen lohkopuskurin huuhtelu (`outRng.Value = buf`) ylikirjoitti ne hiljaisesti.
- **Korjaus:** Lisätty erillinen `textBuf()`-puskuri teksti-entiteeteille. Lohkopuskuri
  huuhdellaan ensin; teksti-entiteetit kirjoitetaan omille riveilleen lohkojen jälkeen.
  Värikoodi (`ColorIndex = 8`) säilytetty teksti-riveille.

### BUG 2 [HIGH] – `VieDATA` ei tallentanut jo auki olevia piirustuksia — `Koodit.bas`

- Tallennus oli aiemmin ehdollinen `If Not OliAuki Then`, jolloin jo avoinna olleiden
  piirustusten attribuuttimuutokset jäivät tallentamatta levylle.
- **Korjaus:** `SaveAs` kutsutaan aina riippumatta `OliAuki`-lipusta.
  Sulkeminen (`Close False`) on edelleen ehdollinen.

### BUG 3 [HIGH] – `PoistaBlokit` tuotti väärän rivitäsmäyksen — `Koodit.bas`

- `InStr(Kaydyt, "|" & Rivi.Row & "|")` -logiikka osui virheellisesti rivinumeroihin,
  joiden numero sisälsi tarkasteltavan numeron osamerkkijonona (esim. rivi 1 osui myös
  riveihin 10, 21 jne.).
- **Korjaus:** `Kaydyt As String` korvattu `Scripting.Dictionary`-rakenteella.
  Tarjoaa O(1) eksaktia kokonaislukutäsmäystä.

### BUG 4 [MED] – `ReDim Preserve Poista()` jokaisella kierroksella = O(n²) — `Koodit.bas`

- `ReDim Preserve` kutsuttiin jokaisella hylätyllä entiteetillä, mikä kopioi koko taulukon
  joka kierroksella.
- **Korjaus:** Taulukko esiallokoitu täyteen valintakokoon (`Joukko.Count`) ennen silmukkaa.
  Yksi `ReDim Preserve` silmukan jälkeen oikeaan lopulliseen kokoon.

### BUG 5 [MED] – `ScreenUpdating` ei palautunut virhetilanteessa — `DATA.bas`

- `Application.ScreenUpdating = False` jäi pysyväksi COM-virhetilanteissa.
- **Korjaus:** Lisätty `prevScreen`-muuttuja tallentamaan alkutila. Kaikki palautuskohdat
  (`Cleanup`, `ErrHandler`, varhaiset `Exit Sub` -haarat) käyttävät `Application.ScreenUpdating = prevScreen`.

### BUG 6 [HIGH] – `VieDATA` tyhjensi blokkiattribuutit (index-pohjainen logiikka) — `Koodit.bas`

- Alkuperäinen `VieDATA` käytti indeksipohjaista logiikkaa (`BlockArray(j).TextString = Cells(i, 8+j)`),
  joka ei vastannut `TuoDATA`:n TAG-pohjaista allokointia. Attribuutit ohitettiin tai tyhjennettiin.
- **Korjaus (27.2.2026):** Uusi TAG-pohjainen päivityslogiikka – symmetrinen `TuoDATA`:n kanssa.
  Attribuutit haetaan `TagString`-nimellä, ei järjestysnumerolla.

---

## 64-bittinen yhteensopivuus — `Koodit.bas`, `DATA.bas`

- Kaikki `Integer`-tyypit muutettu `Long`-tyypiksi 64-bittistä yhteensopivuutta varten
  (`i`, `j`, `Ver`, kierrosmuuttujat).
- Varhainen sidonta (`AcadApplication`, `AcadDocument` jne.) korvattu myöhäisellä
  sidonnalla (`As Object`) – ei vaadi AutoCAD-tyyppikirjaston viittausta.
- AutoCAD-vakiot (`acModelSpace`, `acMax`, `acSelectionSetAll` jne.) määritelty
  manuaalisesti niiden numeroarvoilla, koska myöhäinen sidonta ei tarjoa niitä automaattisesti.
- `acModelSpace As Integer` → `As Long` (`Koodit.bas`) yhtenäistetty `DATA.bas`:n kanssa.

---

## Suorituskykyoptimoint

### PERF 3 – `Cells.EntireColumn.AutoFit` → `UsedRange.Columns.AutoFit` — `Koodit.bas`

- Koko taulukon AutoFit pakotti täyden layout-laskennan. Rajattu käytettyyn alueeseen:
  `If Not DATA.UsedRange Is Nothing Then DATA.UsedRange.Columns.AutoFit`

### PERF 4 – `Numerointi`-kirjoitussilmukka suorituskykysuojauksella — `Koodit.bas`

- Lisätty `ScreenUpdating = False` ja `Calculation = xlCalculationManual` ympäröimään
  rivittäinen kirjoitussilmukka. Lisätty `On Error GoTo NumCleanup` varmistaen palautus
  myös virhetilanteessa.

---

## Tyyli ja laatu

### STYLE 1 – Loop-sisäiset `Dim`-lauseet proseduurin alkuun — `Koodit.bas`

- `entHandle`, `entType`, `isBlock`, `isText`, `isMText`, `tmp`, `effName`, `ip`, `tagName`,
  `colIdx` ym. siirretty `TuoDATA`-proseduurin `Dim`-lohkoon VBA:n käytännön mukaisesti.

### STYLE 2 – `BuildTypeFilter` otettu käyttöön — `Koodit.bas`

- `BuildTypeFilter`-apufunktio oli määritelty mutta kutsumaton; `TuoDATA` toteutti saman
  logiikan inline-koodina. Inline-lohko korvattu `BuildTypeFilter`-kutsulla.
- Päivitetty käyttämään pilkkueroteltu-merkkijono-lähestymistapaa `<or>`-ryhmityksen sijaan
  (luotettavampi AutoCAD 2019 myöhäisessä sidonnassa).

### STYLE 4 – `acModelSpace As Integer` → `As Long` — `Koodit.bas`

- Yhtenäistetty `DATA.bas`:n kanssa. `Long` on standardi 64-bittisessä ympäristössä.

### STYLE 5 – `Set BlockArray = Nothing` → `Erase BlockArray` — `Koodit.bas`

- `BlockArray` on `Variant`-muuttuja, ei `Object`. `Set ... = Nothing` Variant-muuttujalle
  aiheuttaisi runtime-virheen 91. Korvattu `Erase BlockArray`.

---

## Siivous ja optimointi

- Lisätty `AcadHelpers.bas` – sisältää `SafeZoomScaled()`-apufunktion, joka käsittelee
  `ZoomScaled`-kutsun vakioarvojen epäyhtenäisyyden eri AutoCAD-versioiden välillä.
- `Option Explicit` pakotettu kaikissa kolmessa moduulissa.
- `DEBUG_TRACE As Boolean` -vakio lisätty `Koodit.bas`:iin vianmääritystulosteiden
  hallintaan (`Trace`-apuproseduurin kautta).
- Kaikki COM-objektit (`oACAD`, `oDOC`, blokkiviittaukset) vapautetaan `Set ... = Nothing`
  virheen- ja siivouslohkoissa.

---

## Ei muutettu (tunnetut rajoitukset)

| Löydös | Syy |
|--------|-----|
| **PERF 1** – `buf()` sarakkeen `ReDim Preserve` | VBA sallii `Preserve` vain viimeiselle dimensiolle. Rows-dimensio on ensimmäinen; latentin bugin korjaus vaatii buffer-arkkitehtuurin uudelleensuunnittelua. |
| **PERF 2** – `OtsS`-lineaarihaku | `TagCol`-sanakirja välimuistittaa tuloksen ensimmäisen haun jälkeen → O(1). Ei muutosta tarvittu. |
| **STYLE 3** – `GoTo ContinueEntities` | Refaktorointi erilliseksi `ProcessEntity`-apufunktioksi vaatisi laajemman rakennemuutoksen. Jätetään tulevaan iteraatioon. |
