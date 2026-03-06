# Muutosloki — LoopCircuit

**Tiedosto:** `Access/LoopCircuit/` (kaikki moduulit)
**Päivämäärä:** 03.03.2026

---

## 2026-03-07 — Code Review -korjaukset (suorituskyky & vakaus)

### Form_Tee Kuvat.cls

- **Busy wait poistettu — `Odota`:** `Do While odotus > Timer / DoEvents / Loop` -rakenne korvattu `Sleep`-kutsulla (`kernel32.dll`). CPU:n ytimet eivät enää pyöri täysillä odotuksen ajan.
- **Sleep API -deklaraatio lisätty:** `#If VBA7 / PtrSafe` -lohko moduulin alkuun — 64-bit yhteensopiva.
- **GetObject-fallback — `HaeTekstit_Click`:** `CreateObject("AutoCAD.Application")` edeltää nyt `GetObject`-yritys. Jos AutoCAD on jo auki, liitytään siihen — ei avata turhaa uutta instanssia.
- **`oDoc = Nothing` Cleanup-lohkoon — `TeeKuvat_Click`:** Moduulitason `oDoc`-objekti nollataan siivouslohkossa `oAcad`-siivauksen jälkeen — estää haamuviittauksen muistissa virhetilanteessakin.

### Form_Linkkien vaihto.cls

- **Epäonnistuneet linkitykset raportoitu:** Lisätty `Epäonnistuneet As String` -muuttuja keräämään taulut joiden `DoCmd.TransferDatabase` epäonnistui. Käyttäjälle näytetään lista epäonnistuneista operaation päätteeksi — hiljainen ohittaminen poistettu.

### General.bas

- **Transaktio nimenomaisen työtilan kautta — `SniffUser`:** `DBEngine.BeginTrans/CommitTrans/Rollback` korvattu `DBEngine.Workspaces(0).BeginTrans/CommitTrans/Rollback` — transaktio koskee vain kyseistä yhteyttä, ei kaikkia avoimia tietokantayhteyksiä instanssissa. `ws`-muuttuja deklaroitu `DAO.Workspace`-tyyppisenä.

### Module1.bas

- **Hadouken-sisennys litistetty — `CustomMessage`:** Sisäkkäinen `Do While continueLoop / If / Else / If / Else` -rakenne korvattu `Do...Loop` + fail-fast-tyylillä. `continueLoop`-apumuuttuja poistettu tarpeettomana. Koodi tasaisempaa ja helpommin luettavaa.

---

## Kriittiset muutokset (64-bit ja API)

### General.bas

- `api_GetUserName` / `api_GetComputerName`: `nSize As LongPtr` → `ByRef nSize As Long`
  — Win32 `GetUserNameA`/`GetComputerNameA` kirjoittaa 4-tavuisen DWORD:n, ei 8-tavuista LongPtr:ää. Virheellinen tyyppi aiheutti Type Mismatch ajonaikaisesti.
- `SniffUser()`: Poistettu tarpeeton `#If VBA7 Then / BufferSize_Ptr As LongPtr / #End If` -kiertotie — yksi `BufferSize As Long` -muuttuja riittää korjatun deklaraation jälkeen.

### USysCheck.bas

- `wu_GetUserName`: `nSize As LongPtr) As LongPtr` → `ByRef nSize As Long) As Long`
  — Sama DWORD-virhe kuin `General.bas`:ssa. Palautusarvo myös korjattu: `GetUserNameA` palauttaa BOOL (32-bit), ei osoitetta.

---

## Siivous ja optimointi (Phase 3)

### Form_Tee Kuvat.cls

- Lisätty `Me.`-etuliite kaikkiin lomakekontrolliviitteisiin (n. 18 viitettä): `Me.PohjaHakem`, `Me.KuvaHakem`, `Me.TOtsTaulukko`, `Me.TBurst`, `Me.TBlockTaulukko`, `Me.TKyselyt`, `Me.TPaikkaBlokki`, `Me.TTitleBlokki`, `Me.Loki`.
  — Pakollinen `Option Explicit` -yhteensopivuuden vuoksi; puuttuvat etuliitteet aiheuttavat `Variable not defined` -virheen.
- Poistettu kuolleet muuttujat `HaeValitutTekstit_Click`-funktiosta:
  - `Dim DWGName As String` — asetettu, mutta ei käytetty missään
  - `Dim Nimi As String` — ei koskaan käytetty tässä funktiossa
  - `Dim Tied As String` — ei koskaan käytetty tässä funktiossa
  - `Dim Polku As String` + kaksi käyttöriviä — asetettu, mutta ei käytetty myöhemmin

### Kaikki tiedostot

- Kaikki englanninkieliset kommentit käännetty suomeksi: `For ACAD Utility.bas`, `USysCheck.bas`, `General.bas`, `Form_DBUsers.cls`, `Form_Linkkien vaihto.cls`, `Form_Tee Kuvat.cls`, `Module1.bas`.
- Kommenteissa varmistettu oikeat Ä- ja Ö-merkit kautta linjan.

---

## Ei muutoksia tarvinnut

| Tiedosto | Syy |
|---|---|
| `For ACAD Utility.bas` | `PtrSafe` + `POINTAPI` jo oikein ennen migraatiota |
| `Form_DBUsers.cls` | `Me.`-etuliite jo käytössä, DAO-rakenne kunnossa |
| `Form_Linkkien vaihto.cls` | Ei lomakekontrolliviittauksia, logiikka kunnossa |
| `Module1.bas` | Ei API-kutsuja, tyyppiturvallisuus kunnossa |
