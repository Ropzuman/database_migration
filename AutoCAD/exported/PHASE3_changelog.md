# AutoCAD VBA — Phase 3 Viimeistely

**Tiedosto:** AutoCAD/exported (kaikki alikansiot)  
**Päivämäärä:** 2026-03-09  
**Tekijä:** GitHub Copilot — 64-bit migraatioagentti

---

## Kriittiset muutokset

- **`Jonotulostus_64bit/FileDialogs.cls`** — `ShowSave`-funktiossa `Len(udtStruct)` korjattu → `LenB(udtStruct)`. Kriittinen 64-bit bugi: `Len()` palauttaa merkkimäärän, mutta API-kutsuissa tarvitaan tavumäärä `LenB()`.

- **`Arkistotulostus/General.bas`** — Koodaus korjattu: `U+FFFD`-korvausmerkit (`EF BF BD`) muutettu oikeiksi `ä`-merkeiksi (`C3 A4`). Vaurioitunut kommentti: `'Näytetään tulostusformi` nyt oikein.

- **`Explode/Koodi.bas`** — Koodaus korjattu: `U+FFFD` → `ä`. Korjattu kommentti: `'Varmistetaan että ollaan poistuttu komennosta`.

- **`KuvienSelaus/General.bas`** — Koodaus korjattu: `U+FFFD` → `ä`. Korjattu kommentti: `'Näytetään Formi`.

- **`LoopInst/Koodit.bas`** — Koodaus korjattu: 5 riviä, kaikki `U+FFFD` → `ä`. Korjatut kommentit: `'Käydään läpi kaikki attribuutit`, `'siirrettyä blokkia`, `'Tarkistetaan että blokkia yleensä siirettiin`, `'Piirretään viiva`, `'yhtäkään viiva oli piirretty`.

- **`MultiPlot/General.bas`** — Koodaus korjattu: 8 riviä. Lähes kaikki `U+FFFD` → `ä`, poikkeuksena `sisältö`-sana (toinen merkki = `ö`). Korjatut kommentit kattavat listan käsittelyn koko logiikan.

- **`MultiPlot_TW/General.bas`** — Koodaus korjattu: `U+FFFD` → `ä`. Korjattu kommentti: `'Näytetään tulostusformi`.

- **`VBExec/General.bas`** — Koodaus korjattu: `U+FFFD` → `ä`. Korjattu kommentti: `'Näytetään tulostusformi`.

---

## Siivous ja optimointi

- **`Jonotulostus_64bit/General.bas`** — Lisätty `Option Explicit`. Poistettu käyttämättömät muuttujat `Dim i As Integer` ja `Dim Nimi As String`. Korjattu `Formi.show` → `Formi.Show` (oikea Camel Case). Korjattu Windows-1252-enkoodattu kommentti oikeaksi UTF-8:ksi.

---

## Huomiot

- `MultiPlot_OLD/General.bas` ja `MultiPlot_OLD2/General.bas` ovat vanhentuneita (`_OLD`-kansio), ne jätettiin käsittelemättä suunnitellusti.
- Koodauskorjaukset säilyttävät UTF-8 ilman BOM -enkoodauksen, joka on AutoCAD VBE:n eksporttistandardi.
- Kaikkien tiedostojen `Declare`-lausekkeet käyttävät jo `PtrSafe`-attribuuttia — ei muutoksia tarvittu.
