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
