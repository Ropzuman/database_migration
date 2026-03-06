# Muutosloki — FunctionDiagrams

**Tiedosto:** `Access/FunctionDiagrams/` (kaikki moduulit)  
**Päivämäärä:** 2026-03-06  
**Projekti:** 64-bittinen siirtymä — VBA7 / Microsoft 365, tietoturva- ja vakauskorjaukset

---

## Kriittiset muutokset (tietoturva & vakaus)

### Form_FuncBlock.cls

- **SQL-injektiosuojaus — `NEWVALUE_Enter`:** `Atag`-kenttä sanitoitu `sAtag = Replace(Atag, "'", "''")` ennen RowSource-SQL-lauseen rakentamista
- **SQL-injektiosuojaus — `SelectValue_Enter`:** `Atag`, `AreaCode`, `LoopNo` ja `Suffix` sanitoitu `Replace(..., "'", "''")` kaikissa SQL-lausehaaroissa — heittomerkki attribuutin nimessä ei enää kaada kyselyä

### Form_Funktiokaavio.cls

- **SQL-injektiosuojaus — `Command50_Click`:** `KAYTTAJA()`-palaute (yksöislainaus) sekä `Muokkaus0`, `Muokkaus4` ja `Muokkaus3` -kenttäarvot (kaksoislainaus) sanitoitu paikallisiin muuttujiin ennen UPDATE-lauseen rakentamista
- **SQL-injektiosuojaus — `Komento67_Click`:** `RProcess.Value` ja `RCode.Value` sanitoitu muuttujiin `sProcess`/`sCode`; `Me.AreaCode` ja `Me.LoopNo` sanitoitu `sAreaCode2`/`sLoopNo2` INSERT- ja UPDATE-lauseita varten
- **Resurssivuoto poistettu — `Komento67_Click` (varhainen poistumapolku):** `Taul.Close` ja `Set Taul = Nothing` lisätty ennen varhaisinta `Exit Sub` — tietueisto ei enää jää avoimeksi kun resepti löytyy jo kannasta
- **RecordCount → EOF — `Komento67_Click`:** `Taul.RecordCount > 0` korvattu `Not Taul.EOF` — luotettavampi olemassaolotarkistus ilman `MoveLast`-kutsua

### Form_LisääKuviin_ACAD.cls

- **COM-haamuprosessin esto — `TeeKuvat_Click`:** `Exit Sub` vaihdettu `GoTo Loppu` — siivouslohko suoritetaan aina myös onnistumisreitillä; `ErrorHandler` ohjaa `Resume Loppu` → AutoCAD ei jää Tehtävienhallintaan näkymättömäksi prosessiksi myöskään virhetilanteessa
- **COM-siivouslohko — `TeeKuvat_Click` `Loppu:`:** Lisätty `On Error Resume Next` + `If Not oAcad Is Nothing Then oAcad.Quit` — kaikkien resurssien vapautus taattu
- **Resurssivuoto poistettu — `PaivitaDocRev_Click`:** `rstSnapshot.Close / Set = Nothing`, `dbsDocuments.Close / Set = Nothing` ja `Set dbs = Nothing` lisätty sekä normaali- että virhepolkuun — ulkoinen Documents-kantayhteys suljetaan aina
- **SQL-injektion esto INSERT-silmukoissa — `PaivitaDocRev_Click`:** `DocumentsDetail`- ja `DocumentsRevision`-INSERT-lauseiden kenttäarvot sanitoitu `Replace(CStr(...), "'", "''")` — heittomerkkiä sisältävät dokumenttinimet (esim. "O'Brien") eivät enää kaada lausetta

### Form_LukituskaavioLinkit.cls

- **Virheenkäsittely lisätty — `Command0_Click`:** `On Error GoTo ErrorHandler` ja täydellinen `ErrorHandler`-lohko — virhetilanteessa `tbl`/`qry` suljetaan, tiimalasi poistetaan ja käyttäjälle näytetään virheilmoitus

---

## Siivous ja optimointi

### Form_LukituskaavioLinkit.cls

- **O(N²) → O(N) suorituskykyjkorjaus — `Command0_Click`:** Sisäkkäinen `Do Until tbl.EOF` -silmukka (1&nbsp;000&nbsp;×&nbsp;1&nbsp;000 = 1&nbsp;000&nbsp;000 vertailua) korvattu `tbl.FindFirst`-kutsulla (1&nbsp;000 hakua) — suorituskyky kasvaa eksponentiaalisesti suurilla tietomäärillä
- **Avaustyyppi eksplisiittiseksi:** `InterlockingLinkPage` avattu `dbOpenSnapshot`-tilassa (ei muokkausta → nopeampi); `IntLinkPage` avattu `dbOpenDynaset`-tilassa (muokkaus vaaditaan)
- **Null-turvallinen tyhjyystarkistus:** `tbl!page1 = ""` korvattu `Nz(tbl!page1, "") = ""` — Null-arvo ei enää aiheuta Type Mismatch -virhettä
- **SQL-sanitointi FindFirst-kutsussa:** `qry!TXT1` sanitoitu `Replace(..., "'", "''")` — heittomerkki TXT1-kentässä ei kaada FindFirst-hakua
