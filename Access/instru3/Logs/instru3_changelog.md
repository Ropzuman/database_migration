# Muutosloki – Access/instru3/

**Tiedosto:** `Access/instru3/` (kansio)
**Päivämäärä:** 2026-03-03
**Haara:** `instru3`

---

## Kriittiset muutokset (64-bit-yhteensopivuus)

- **USysCheck.bas** – `api_GetUserName` ja `api_GetComputerName`: `nSize As LongPtr` → `ByRef nSize As Long`. `GetUserNameA`/`GetComputerNameA` kirjoittavat vain 32-bittisen DWORD:n; `LongPtr` (8 tavua) aiheutti Type Mismatch -ajonaikaisen virheen 64-bittisellä Officella.

---

## Muuttuja- ja viittauskorjaukset

- **Form_DBUsers.cls** – Lisätty puuttuva `Option Explicit`. `NetworkName.Value` → `Me.NetworkName.Value`.
- **Form_CopyLoops.cls** – `Me.`-etuliite lisätty kaikkiin lomakekontrolliviittauksiin: `TTiedot`, `Kanta`, `Loopit` (11 kohtaa).
- **Form_SizingOut.cls** – `Me.`-etuliite lisätty kontrolleihin `Hakem`, `Taulukko`, `FileName` (4 kohtaa).

---

## Siivous ja optimointi

- **Form_DBUsers.cls** – Poistettu kuolleet muuttujat `iLOF` (ei koskaan luettu) ja `iStart` (inkrementoitu mutta ei koskaan käytetty).
- **Form_SizingOut.cls** – Poistettu kuollut kommentoitu koodi `'Hakem.VALUE = "d:\tilap\"`. Korjattu kirjoitusvirhe: `'Oletusky sely` → `'Oletuskysely`.
- **Kaikki tiedostot** – Kommentit käännetty suomeksi oikeilla Ä/Ö-kirjaimilla. Kaksikieliset kommentit (suomi–englanti) siistitty suomeksi.
- **Kaikki tiedostot** – Moduuliotsikot päivitetty `Päivitetty`-kenttään 2026-03-03-merkintä.
- **Form_CopyLoops.cls** – MsgBox-viestit suomeksi: `"Select database first!"` → `"Valitse ensin tietokanta!"`, `"Ready!"` → `"Valmis!"`, virheviestit suomeksi.
- **USysCheck.bas** – Tuntematon käyttäjä/kone: `"Unknown"` → `"Tuntematon"`.

---

## 2026-03-06 – Code Review -korjaukset (tietoturva & vakaus)

### Kriittiset muutokset

- **Form_DBUsers.cls** – Command Injection -haavoittuvuus korjattu: `net send` → `msg.exe` (Windows Vista+ yhteensopiva). Käyttäjäsyötteet (`Viesti`, `NetworkName`) sanitoidaan: poistetaan `"`, `&`, `|` ennen Shell-kutsuun välittämistä. Häkättyyn InputBox-muotoiluun (7× vbCrLf) siistiminen.
- **Form_SizingOut.cls** – RFC 4180 -standardin mukainen CSV-lainausmerkkien tuplataminen lisätty: `Replace(Nz(Arvo, ""), """", """""")`. Formula Injection -esto lisätty Excelin varalle: `=`, `@`, `+`, `-` -alkuiset arvot saavat eteen heittomerkin `'`. Käsittelyjärjestys korjattu: formula injection -tarkistus ennen RFC 4180 -muunnosta. Aiempi virheellinen kolmoislainaus (`""""""""`) korvattu oikealla (`""""""`).
- **Form_Linkkien vaihto.cls** – Vakava tiedonhäviöriski poistettu: `DROP TABLE` + `DoCmd.TransferDatabase` -ketju korvattu turvallisella `tdf.Connect` + `tdf.RefreshLink` -päivitysmekanismilla. Linkitetyn taulun relaatiot säilyvät ehjinä.
- **Form_CopyLoops.cls** – `Form_Unload`-lukkiutumisriski korjattu: alilomakkeen `Me.Loopit.Form.RecordSource = ""` tyhjennetään ja `DoEvents`-kutsu suoritetaan ennen `LOOPLINK`-taulun poistoa, jotta Access ehtii vapauttaa tiedostolukot.

### API- ja yhteensopivuuskorjaukset

- **general.bas** – 64-bittisessä Officessa epävakaa `GetOpenFileNameA` (comdlg32.dll) poistettu kokonaan. Korvattu Office-natiivilla `Application.FileDialog(3)` -pohjaisella `HaeTiedostoNimi()`-funktiolla (toimii luotettavasti 32/64-bit).
- **USysCheck.bas** – ANSI-versiot `GetUserNameA` ja `GetComputerNameA` korvattu Unicode-versioilla `GetUserNameW` ja `GetComputerNameW`. Estää skandinaavisten merkkien (Ä, Ö) korruptoitumisen koodisivumuunnosten yhteydessä. Muutos tehty sekä `#If VBA7` että `#Else`-haaroihin.

### Code review -korjaukset (2026-11-10)

- **Form_CopyLoops.cls** – `HaeLoopit_Click` tekee koko tuonnin yhdessä transaktiossa (rollback virheessä). `On Error Resume Next` rajattu pelkkään kenttäkopiointiin, `Update`-virheet keskeyttävät tuonnin ja loopit lisätään laiteriviluetteloon vasta onnistuneen tallennuksen jälkeen. SQL-heittomerkit tuplataan. Recordsetit suljetaan, `devTable` avataan kerran, tyhjä alilomake ei enää kaada `MoveFirst`-kutsua. Vanhentuneet `GetOpenFileName`/`comdlg32`-maininnat poistettu kommenteista.
- **general.bas** – `EtsiLoop`: SQL-heittomerkkien suojaus. `UdNoteToRev`: tarkistus puuttuvalle `|`-merkille (ei enää negatiivista `Left$`-pituutta). Kommentit suomeksi.
- **USysCheck.bas** – `GetUserNameW`/`GetComputerNameW` saavat puskurin `LongPtr`-osoittimena (`StrPtr`), jotta W-API kirjoittaa UTF-16-merkit oikein. Kuollut 32-bittinen `#Else`-haara poistettu.
- **Form_DBUsers.cls** – `msg`-kohde rajattu sallittuihin merkkeihin ja viestistä poistetaan lisää Shell-erikoismerkkejä (`^<>%!`, rivinvaihdot). `LACCDB`-polku muodostetaan `InStrRev`-haulla. Virheviestit suomeksi.
- **Form_SizingOut.cls** – Taulun/kyselyn nimi validoidaan ennen `SELECT`-lausetta. Kaikki kentät (myös ensimmäinen) käsitellään samalla tavalla. Desimaalipilkku muutetaan pisteeksi vain luvuille, ja formula-injection-suoja ei enää riko negatiivisia lukuja. Hakemisto lainataan `EXPLORER`-kutsussa ja CSV-tiedosto suljetaan myös virhetilanteessa.
- **Form_Linkkien vaihto.cls** – Käyttämättömät muuttujat poistettu, `Polku2`/`Nimi` nollataan joka kierroksella, kommentit ja viestit suomeksi.
- **Sivunumerointi.bas** – `CInt(Alue)` korvattu `CInt(Val(Alue))`-muunnoksella ja lisätty dokumentaatio.

### Viimeistely (2026-11-10)

- **Form_CopyLoops.cls** – Englanninkieliset `Debug.Print`-tulosteet poistettu, tiedostovalintaikkunan otsikko ja suodatin suomennettu, vanha `LahdeDB`-yhteys suljetaan ennen uuden avaamista (ei vuoda yhteyksiä uudelleenvalinnassa). Kommentit tarkistettu.
- Muissa instru3-tiedostoissa ei löytynyt jäljellä olevia englanninkielisiä kommentteja, käyttämättömiä muuttujia tai kuollutta koodia; objektiviittaukset vapautetaan (`Set ... = Nothing`) normaali- ja virhepolussa.

### Siivous

- **Access/instru3/CodeReview_instru3** – Tiedostonimi muutettu: nimetty uudelleen `CodeReview_instru3.md`-tiedostoksi.
