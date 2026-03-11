# AcadDATA – Code Review Korjaukset

**Tiedostot:** `Excel/Moduulit/AcadDATA/Koodit.bas`, `DATA.bas`, `AcadHelpers.bas`  
**Päivämäärä:** 2026-03-04  
**Peruste:** HTML code review -raportti (`acad_code_review.html`) – 5 bugia, 4 suorituskykyongelmaa, 5 tyyliseikkaa

---

## Kriittiset muutokset (Bugit)

### BUG 1 ✅ [HIGH] – Teksti/blokki-kirjoitusristiriita (`TuoDATA`) — `Koodit.bas`

**Ongelma:** TEXT/MTEXT-entiteetit kirjoittivat suoraan `Cells(Rivi, ...)`:lle silmukan aikana, mutta lohkodatan `buf`-puskurin huuhtelu (`outRng.Value = buf`) alkoi samasta `DocStartRow`-pisteestä. Lohkohuuhtelu ylikirjoitti aiemmin kirjoitetut teksti-rivit hiljaisesti.

**Korjaus:**

- Lisätty erillinen `textBuf()`, `textBufRows`, `textBufCap`-puskuri teksti-entiteeteille.
- Teksti-entiteetit kerätään `textBuf`-taulukkoon silmukan aikana (ei suoraan soluihin).
- Lohkopuskuri huuhdellaan ensin (`outRng.Value = buf`).
- Teksti-entiteetit huuhdellaan omille riveilleen lohkojen jälkeen solukohtaisella silmukalla (värikoodi `ColorIndex=8` säilytetty).
- Lisätty uudet `Dim`-lauseet proseduurin alkuun: `textBuf, textBufRows, textBufCap, tIdx, textStartRow, textX, textY, textStr`.

---

### BUG 2 ✅ [HIGH] – VieDATA ei tallenna jo auki olevia piirustuksia — `Koodit.bas`

**Ongelma:** Viimeisen käsitellyn piirustuksen tallennus oli ehdollinen `If Not OliAuki Then`. Jos piirustus oli jo auki (OliAuki = True), tehdyt attributtimuutokset jäivät tallentamatta levylle.

**Korjaus:** `SaveAs` kutsutaan aina riippumatta `OliAuki`-lipusta. Sulkeminen (`Close False`) on edelleen ehdollinen – jo auki olleita piirustuksia ei suljeta:

```vba
oDOC.SaveAs oDOC.FullName, Ver           ' Tallennetaan aina
If Not OliAuki Then oDOC.Close False     ' Suljetaan vain, jos macro avasi
```

---

### BUG 3 ✅ [HIGH] – `PoistaBlokit` Kaydyt-merkkijono tuotti väärän täsmäyksen — `Koodit.bas`

**Ongelma:** `InStr(Kaydyt, "|" & Rivi.Row & "|")` osui riville 1, mutta myös `"|10|"`, `"|21|"`, kaikille riveille joiden numero sisältää `"1"` merkkijonona. Rivejä ohitettiin virheellisesti.

**Korjaus:** Korvattu `Kaydyt As String` + `InStr`-logiikka `Scripting.Dictionary`-rakenteella. Dictionary tarjoaa O(1) eksaktin täsmäyksen kokonaisluvulle:

```vba
Dim Kaydyt As Object
Set Kaydyt = CreateObject("Scripting.Dictionary")
...
If Not Kaydyt.Exists(Rivi.Row) Then
    Kaydyt.Add RiviNo, True
```

---

### BUG 4 ✅ [MED] – `ReDim Preserve Poista()` per kierros = O(n²) — `Koodit.bas`

**Ongelma:** `ReDim Preserve Poista(L)` kutsuttiin jokaisella hylätyllä entiteetillä. Jokainen kutsu kopioi koko taulukon.

**Korjaus:** Taulukko esiallokoitu ennen silmukkaa täyteen valintakokoon (`Joukko.Count`). Silmukan jälkeen siivottu todelliseen kokoon yhdellä `ReDim Preserve`:llä.

---

### BUG 5 ✅ [MED] – `ScreenUpdating` ei palaudu virhetilanteessa – `DATA.bas`

**Ongelma:** Virhetilanteessa (esim. COM-kutsussa) `Application.ScreenUpdating = False` jäi pysyväksi.

**Korjaus:** Lisätty `prevScreen`-muuttuja tallentamaan alkutila; kaikki palautuskohdat (`Cleanup`, `ErrHandler`, varhaiset `Exit Sub` -haarat) käyttävät nyt `Application.ScreenUpdating = prevScreen`.

---

## Suorituskykyoptimoint

### PERF 3 ✅ [MED] – `Cells.EntireColumn.AutoFit` → `UsedRange.Columns.AutoFit` — `Koodit.bas`

Koko taulukon AutoFit pakottaa täyden layout-laskennan jokaiselle solulle. Rajattu käytettyyn alueeseen:

```vba
If Not DATA.UsedRange Is Nothing Then DATA.UsedRange.Columns.AutoFit
```

### PERF 4 ✅ [LOW] – `Numerointi` kirjoitussilmukka ilman suorituskykysuojausta — `Koodit.bas`

Lisätty `ScreenUpdating = False` + `Calculation = xlCalculationManual` ympäröimään rivittäinen kirjoitussilmukka. Lisätty `On Error GoTo NumCleanup` -merkki varmistaen palautus myös virhetilanteessa.

---

## Tyyli ja laatu

### STYLE 1 ✅ [LOW] – Loop-sisäiset `Dim`-lauseet siirretty proseduurin alkuun — `Koodit.bas`

VBA nostaa kaikki `Dim`-lauseet proseduuritasolle käännösaikana. `entHandle`, `entType`, `isBlock`/`isText`/`isMText`, `tmp`, `effName`, `ip`, `ipT`, `ipM`, `tagName`, `colIdx` siirretty `TuoDATA`-proseduurin `Dim`-lohkoon. Loop-kohtaan lisätty nollauskommentit (`entType = "": isBlock = False` jne.).

### STYLE 2 ✅ [LOW] – `BuildTypeFilter` kutsumaton kuollut koodi → otettu käyttöön — `Koodit.bas`

`BuildTypeFilter`-apufunktio oli määritelty mutta kutsumaton; `TuoDATA` toteutti saman logiikan inline-koodina. Inline-lohko korvattu `BuildTypeFilter`-kutsulla. `BuildTypeFilter` päivitetty käyttämään pilkkueroteltu-merkkijono-lähestymistapaa `<or>`-ryhmityksen sijaan (luotettavampi AutoCAD 2019 myöhäisessä sidonnassa).

### STYLE 4 ✅ [LOW] – `acModelSpace As Integer` → `As Long` — `Koodit.bas`

Yhtenäistetty `DATA.bas`:ssa käytetyn `Long`-tyypin kanssa. 64-bittisessä ympäristössä `Long` on standardi kokonaisluvuille.

### STYLE 5 ✅ [LOW] – `Set BlockArray = Nothing` → `Erase BlockArray` — `Koodit.bas`

`BlockArray` on `Variant`-muuttuja (sisältää `GetAttributes`-taulukon), ei `Object`. `Set ... = Nothing` Variant-muuttujalle aiheuttaisi runtime-virheen 91 `On Error Resume Next` -haarauksen ulkopuolella. Korvattu `Erase BlockArray`.

---

## Ei muutettu (tunnetut rajoitukset)

| Löydös | Syy |
|--------|-----|
| **PERF 1** – `buf()` sarakkeen ReDim Preserve | VBA sallii `ReDim Preserve` viimeiselle dimensiolle – `(rows, cols)` mallissa sarake on viimeinen, joten tämä toimii. Esiskannaus blokkimäärityksistä olisi parempi mutta vaatisi suuremman refaktoroinnin. |
| **PERF 2** – `OtsS`-lineaarihaku | `TagCol`-sanakirja välimuistittaa tuloksen ensimmäisen haun jälkeen. O(1) jokaiselle seuraavalle. Muutosta ei tarvita. |
| **STYLE 3** – `GoTo ContinueEntities` | Suurempi refaktorointi (erillinen `ProcessEntity`-apufunktio). Jätetään tulevaan iteraatioon. |
| **BUG 1 / PERF 1** – `ReDim Preserve buf` rividimenision muutos | Koodissa `rowCap` on ensimmäinen dimensio – VBA ei salli `Preserve` ensimmäiselle dimensiolle. Tämä on latentin runtime-bugin alku. Korjataan erillisessä iteraatiossa (vaatii buffer-arkkitehtuurin muutosta). |

---

## Vahvuudet (ei muutoksia tarvittu)

- ✅ `Scripting.Dictionary` tagi→sarake-välimuisti `TuoDATA`:ssa – hyvä O(1)-optimoint.
- ✅ TAG-pohjainen attribuuttipäivitys `VieDATA`:ssa (27.2.2026 korjaus) – oikea lähestymistapa.
- ✅ Massakirjoituspuskuri + yksittäinen `Range.Value = buf` -kirjoitus – tehokas tuontistrategia.
- ✅ `SafeZoomScaled` `AcadHelpers.bas`:ssa – versiovarianssien eristäminen omaan moduuliinsa.
- ✅ `StepMsg`-diagnostiikkajäljet virhetilanteissa – käyttökelpoinen paikannus.
