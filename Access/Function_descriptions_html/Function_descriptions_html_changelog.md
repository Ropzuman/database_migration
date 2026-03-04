# Muutosloki — Function_descriptions_html

**Tiedosto:** `Access/Function_descriptions_html/` (kaikki moduulit)  
**Päivämäärä:** 2026-03-03  
**Projekti:** 64-bittinen siirtymä — VBA7 / Microsoft 365

---

## Kriittiset muutokset (64-bit & ajuri)

### USysCheck.bas

- `Declare`-lauseet muutettu `PtrSafe`-muotoon (`#If VBA7 Then` -estolohkolla)
- `GetUserNameA` ja `GetComputerNameA`: `nSize`-parametri säilytetty `Long`-tyyppisenä (ei `LongPtr`) — välttää Type Mismatch -virheen 64-bittisessä Officessa
- DAO-tietueiston sulkeminen korjattu: `.Close` kutsutaan ennen `Set x = Nothing`
- ErrorHandler päivitetty: `If Not x Is Nothing Then x.Close` -malli (DAO:lla ei `.State`-ominaisuutta)

### Form_FrmASETUKSET.cls

- `Me.`-etuliite lisätty kaikille lomakeohjaimille (`TINSTRU`, `TMAINEQ`, `TFUNCDIAG`) — välttää `Variable not defined` -virheet `Option Explicit` -tilassa

### Form_FrmMUOKKAUS.cls

- `Me.`-etuliite lisätty kaikille lomakeohjaimille läpi koko lomakkeen
- `SelStart`/`SelLength`-asetukset suojattu `On Error Resume Next` / `On Error GoTo ErrorHandler` -lohkoilla (Error 2185: ohjain vaatii fokuksen)
- Word-automaatio (`CreateObject("Word.Application")`) — 64-bittisen Officen yhteensopivuus varmistettu

### Form_MOTORS subform.cls / Form_PIIRIT subform.cls

- `SelStart`-kirjoitukset kietottu virheensuojauksen sisään (Error 2185 -esto)
- `Me.`-etuliite lisätty ohjaimen viittauksiin

### Form_DBUsers.cls

- `.LACCDB`-lukitustiedoston lukeminen: `DBEngine.Workspaces(0).Databases(0).Name` — ei kutsuta `.Close` välillisesti (`dbCurrent`-viittausta ei käytetä)
- Turha, orvokoodi poistettu

---

## Siivous ja optimointi (Phase 3 — 2026-03-03)

### Form_FrmMUOKKAUS.cls

- **Kuollut koodi poistettu:**
  - `Command81_Click` (tyhjä tapahtumakäsittelijä) poistettu
  - `Form_AfterUpdate` (kommentoitu `cUI.Requery`) poistettu
  - `Form_Load` yksinkertaistettu — kommentoitu `CLoppuun.Value`-rivi poistettu
- **Englanninkieliset kommentit käännetty suomeksi:**
  - Moduuliotsikko: `Description:`, `Document Generation Process:`, `Template Bookmarks:`, `Dependencies:` → suomeksi
  - `TeeHTML`: `Sub:`, `Purpose:`, `Description:` → suomeksi
  - `MuutaLinkit`: `Parameters:`, `Returns:`, `Description:` → suomeksi
  - `PoimiAsetukset`, `MuutaTiedot`, `KorvaaOts`, `Korvaa`, `Korvaa2`, `POISTA`: englanninkieliset lohkot → suomeksi
  - `LisaaHaly`, `LisaaHalyHTML`: `Parameters:` → suomeksi
  - Inline-kommentit: `'Header may not exist'`, `'Continue if bookmark not found'`, `'Return original on error'` → suomeksi
- **Muuttujien kommentit suomeksi:** `POHJA`, `POLKU`, `REV`, `STATUS`, `AUTHOR`, `DOCNO`
- **Virheilmoitus korjattu:** `"Error creating motor descriptions"` → `"Virhe moottorikuvausten luonnissa"`
- **Unicode-escape-sekvenssit korjattu:** Kirjaimelliset `\u00e4`-merkkijonot korvattu oikeilla ä-kirjaimilla kahdessa kommenttirivissä

### GeneralCodes.bas

- Kommentit tarkastettu — kaikki jo suomeksi, ei muutoksia tarvittu

### KAANNOS.bas

- Kommentit tarkastettu — kaikki jo suomeksi, ei muutoksia tarvittu

### USysCheck.bas

- Kommentit tarkastettu — oikeat suomenkieliset ä/ö-merkit varmistettu

---

## Tiedostojen tila siirtymän jälkeen

| Tiedosto | 64-bit OK | Kommentit FI | Kuollut koodi poistettu |
|---|---|---|---|
| USysCheck.bas | ✅ | ✅ | ✅ |
| GeneralCodes.bas | ✅ | ✅ | ✅ |
| KAANNOS.bas | ✅ | ✅ | ✅ |
| Form_DBUsers.cls | ✅ | ✅ | ✅ |
| Form_FrmASETUKSET.cls | ✅ | ✅ | ✅ |
| Form_FrmMUOKKAUS.cls | ✅ | ✅ | ✅ |
| Form_MOTORS subform.cls | ✅ | ✅ | ✅ |
| Form_PIIRIT subform.cls | ✅ | ✅ | ✅ |
