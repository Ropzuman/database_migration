# Muutosloki — Access/DOCUMENTS

**Tiedosto:** `Access/DOCUMENTS/` (koko kansio)
**Päivämäärä:** 2026-03-03
**Haara:** `main`
**Commitit:** `5fc6eaf` → `b81bb72`

---

## Kriittiset muutokset

### 64-bit API -korjaukset
- `GlobalVBAs.vba` ja `ForDocuments.vba`: `GetUserNameA`- ja `GetComputerNameA`-kutsuissa `nSize` muutettu `LongPtr` → `ByRef Long` — API kirjoittaa LPDWORD-osoittimeen (32-bit DWORD, 4 tavua), ei 64-bittiseen LongPtr:iin. Aiheutti Type Mismatch -ajovirheen 64-bittisessä Office-ympäristössä.
- `ForDocuments.vba`: `lngStringLength` muutettu `LongPtr` → `Long` samasta syystä.

### DAO-transaktio-ongelmat
- Kaikki `DB.BeginTrans` / `DB.CommitTrans` / `DB.Rollback` -kutsut muutettu `DBEngine.*` -muotoon — DAO:ssa transaktiot kuuluvat `DBEngine`-tasolle, ei `Database`-objektille. Aiheutti "function or interface marked as restricted" -käännösvirheen.
- Korjattu tiedostoissa: `Form_DISTRIBUTION.cls`, `Form_USysAddToDistr.cls`, `Form_USysEditDistribution.cls`, `Form_USysNewRecipient.cls`.

### DAO:lle sopimaton `.State`-tarkistus
- Poistettu kaikki `If taulu.State = 1 Then taulu.Close` -rakenteet — `.State`-ominaisuus kuuluu ADO:lle, ei DAO:lle. Aiheutti "method or data member not found" -käännösvirheen.
- Korjattu tiedostoissa: `Form_USysAddedDistr.cls`, `Form_USysOpenFile.cls`, `Form_DISTRIBUTION.cls`, `Form_USysAddToDistr.cls`, `Form_USysEditDistribution.cls`, `Form_USysNewRecipient.cls`, `Form_USysNewDistribution.cls`.

### Me.-etuliite (Option Explicit -yhteensopivuus)
- Lisätty `Me.`-etuliite kaikkiin lomakekontrolliviittauksiin kaikissa 20 lomake- ja raporttitiedostossa. Ilman etuliitettä `Option Explicit` tulkitsee kontrollin nimen määrittelemättömäksi muuttujaksi → "Variable not defined" / "method or data member not found".
- Erityistapaukset: `Form_SETTINGS.cls`-funktiossa `PoimiPolku(Kentta As Control)` parametri `Kentta` on tarkoituksellisesti ilman `Me.` (se on `Control`-tyyppinen funktioparametri, ei lomakekontrolli).

### Muut ajonaikaiset virheet
- `Form_USysRevText.cls`: `MsgBox`-merkkijonossa `\"` → `""` (JSON-eskejppaus jäi koodiin, aiheutti syntaksivirheen VBA:ssa).
- `Form_USysOpenFile.cls`, rivi `Tied.RowSource` → `Me.Tied.RowSource` (jäi edelliseltä kierrokselta).
- `Form_USysEditDistribution.cls`: `Lista.Requery` → `Me.Lista.Requery`.
- `Form_USysReserve.cls`: `TPaperSize`, `TDiscipline`, `TNumber`, `TClientNo`, `ClientNo` → kaikki `Me.`-etuliitteellä.

---

## Siivous ja optimointi

### Kuollut koodi poistettu
- `GlobalVBAs.vba`: Kaikki `Debug.Print`-rivit poistettu parsintafunktioista (`HaeTekija`, `HaeRevisioija`, `HaeRevisioijaPvm`, `EkaRevRivi`, `HaeRevisio`, `HaeViimPaiva`, `HaePaiva`) — nämä olivat kehityksenaikaisia lokitulostuksia, jotka jäivät koodiin.
- `Form_USysRevText.cls`: Käyttämätön `Dim Rivit As String` -muuttuja poistettu `Form_Resize`-aliohjelmasta.
- Tuplarivit poistettu parsintafunktioiden alusta.

### Resurssien sulkeminen
- `GlobalVBAs.vba` / `SetStartup`: Lisätty puuttuvat `taulu.Close` ja `DB.Close` ennen `Set x = Nothing` -kutsuja. Ilman sulkemista DAO-tietueet voivat jäädä auki tietokantaan.

### Kommentit suomeksi
- `GlobalVBAs.vba`: Kaikki englanninkieliset `'''`-JSDoc-kommentit ja `'`-inline-kommentit muutettu suomeksi revisioparsintafunktioissa.
- `GlobalVBAs.vba`: Englanninkielinen funktiolistablokkikommentti (`' Functions for parsing revisions`) käännetty suomeksi.
- Kaikki 24 tiedostoa: Lisätty suomalainen moduuliotsikko (LOMAKE/MODUULI/RAPORTTI, SOVELLUS, KUVAUS, PÄIVITETTY).

### Virheenkäsittely
- `Form_USysOpenFile.cls`: `Command52_Click` — kommentti "Hakemistoa ei ole" suomennettu.
- `Form_USysShowCommon.cls`: Otsikon kirjoitusvirheet korjattu (`Yleiskkäyttöinen` → `Yleiskäyttöinen`, `muokkauksest` → `muokkaus`; kuvaus täsmennetty).
- `Form_USysOpenFile.cls`: Vanha Windows NT -polku `C:\WINNT\EXPLORER.EXE` → `C:\Windows\explorer.exe`.

---

## Tiedostot joihin ei tehty muutoksia

| Tiedosto | Syy |
|---|---|
| `Form_USysDISTRIB.cls` | Kaikki koodi kommentoitu pois alkuperäisessä |
| `Report_USYSTRANSMITTALFP.cls` | Tyhjä tiedosto |
