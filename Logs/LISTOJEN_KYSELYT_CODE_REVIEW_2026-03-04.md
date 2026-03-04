# Code Review — Listojen kyselyt (Module1, Module2, Module3)

**Tiedostot:** `Excel/Moduulit/Listojen kyselyt/Module1.bas`, `Module2.bas`, `Module3.bas`
**Päivämäärä:** 04.03.2026
**Review-työkalu:** Claude Code (automaattinen analyysi)
**Toteuttaja:** GitHub Copilot

---

## Kriittiset muutokset

### 1. Checkout — RMAX-rivimerkkiluku (`Module1.bas`)
- **Ongelma:** `Mid(Arvo, 4, 1)` luki välilyönnin `"£1: "` -muodon 4. merkistä → `CInt(" ") = 0`, jolloin `RMAX` jäi nollaksi ja moniriviset `£1/2/3`-merkit eivät toimineet oikein.
- **Korjaus:** `Mid(Arvo, 4, 1)` → `Mid(Arvo, 2, 1)` (lukee rivinumeron `£`-merkin välittömästi jälkeisestä positiosta).

### 2. VaihdaInfo — `Split()`-indeksin ulkopuolinen käyttö (`Module2.bas`)
- **Ongelma:** `Split(DIRevArr(r), "/")(1..4)` kaatui ajonaikaisella virheellä 9 jos revisioentry-merkkijono sisälsi vähemmän kuin odotettuja `/`-osia (esim. puuttuva approver tai desc).
- **Korjaus:** Kaikissa `designer`/`checker`/`approver`/`desc` -lohkoissa: `revParts = Split(...)` + `If UBound(revParts) >= N` ennen indeksointia.

### 3. GenPrintout — `ActiveWorkbook` epäluotettava (`Module1.bas`)
- **Ongelma:** `Set destWB = ActiveWorkbook` Copy-operaation jälkeen voi viitata väärään työkirjaan jos jokin lisäosa tai tapahtumakäsittelijä aktivoi toisen työkirjan `.Copy`-kutsun jälkeen.
- **Korjaus:** `Set destWB = Workbooks(Workbooks.Count)` — viittaa aina juuri lisättyyn työkirjaan.

### 4. HaeData — ADODB-yhteysobjektin nollaus provider-fallbackissa (`Module1.bas`)
- **Ongelma:** Kun 16.0-provider epäonnistui, sama `conn`-objekti yritettiin avata uudelleen ilman nollausta. Osittain avattu `ADODB.Connection` voi jäädä virhetilaan ja nostaa virheen 3709 seuraavalla `.Open`-kutsulla.
- **Korjaus:** `Set conn = Nothing` + `Set conn = CreateObject("ADODB.Connection")` jokaisen epäonnistuneen yrityksen jälkeen ennen seuraavaa provider-kokeilua.

---

## Suorituskykyparannukset

### 5. VaihdaInfo — `For Each cmt` indeksisilmukan sijaan (`Module2.bas`)
- **Ongelma:** `For i = 1 To .Comments.Count` + `.Comments(i)` käytti COM-kokoelman indeksointia, joka on O(n) per askel (traversointi alusta).
- **Korjaus:** `For Each cmt In .Comments` — käyttää enumeraattoria, O(1) per iteraatio.

### 6. TeeLinkingKommentit — `On Error Resume Next` silmukan ulkopuolelle (`Module2.bas`)
- **Ongelma:** `On Error Resume Next` / `On Error GoTo 0` asetettiin uudelleen joka kaavasolun kohdalla → virheenkäsittelytilan vaihtelua jokaisen iteraation yhteydessä.
- **Korjaus:** `On Error Resume Next` kerran ennen silmukkaa, `On Error GoTo 0` kerran silmukan jälkeen.

---

## Tyyli- ja laaturikorjaukset

### 7. VaihdaInfo — `ws Is Nothing` -antipattern (`Module2.bas`)
- **Ongelma:** `Set ws = Sheets(SheetName)` ilman virhesuojausta, jonka jälkeen `If ws Is Nothing` — `Sheets()` ei koskaan palauta `Nothing`, se nostaa Error 9 jos sheetti puuttuu, joten haara oli aina saavuttamaton.
- **Korjaus:** `On Error Resume Next` ennen `Set ws = ...`, `On Error GoTo ErrHandler` välittömästi Set-lauseen jälkeen.

### 8. GenPrintout — `Dim` For-silmukan sisällä → Sub-alkuun (`Module1.bas`)
- **Ongelma:** `Dim tCopy As Double: tCopy = Timer` ym. deklaroitu `For`-silmukan sisällä, mikä antaa harhaanjohtavan käsityksen muuttujan elinajasta — VBA nostaa kaikki `Dim`-lauseet kääntöaikana Sub-alkuun.
- **Korjaus:** `Dim tCopy As Double, tShade As Double, tLink As Double` siirretty Sub-muuttujamääritysten joukkoon.

---

## Tunnistetut vahvuudet (ei muutoksia)

- **`BeginFastMode`/`EndFastMode`** — Kaikki viisi Application-tilalippua tallennetaan ja palautetaan symmetrisesti, myös virheenkäsittelijöissä. Ammattimainen defensiivinen ohjelmointi.
- **DAO/ADODB provider-fallback** — ACE 16.0 → 15.0 → 12.0 ja DAO 120 → 36 -ketju kattaa eri M365-asennusversiot hyvin.
- **`VaihdaLinkit` Comments-kokoelmalla** — Erillisten solujen iteraation sijaan Comments-kokoelman käyttö on tunnistettu ja kommentoitu oikein (30–50% nopeampi).
- **Kontekstuaaliset virhekäsittelijät** — `GenPrintoutError` ja `CheckoutError` kääntävät VBA-virhekoodit (91, 1004, 9, 13) suomenkielisiksi selityksiksi.
- **Suorituskykyprofilointi** — `perfCopy`/`perfLink`/`perfShade`-ajoitusmittarit Immediate Windowiin. Oikea optimointistrategia: mittaa ensin.
