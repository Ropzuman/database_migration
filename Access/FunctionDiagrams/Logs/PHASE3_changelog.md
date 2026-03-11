# Access/FunctionDiagrams & Function_descriptions_html — Phase 3 Viimeistely

**Tiedosto:** Access/FunctionDiagrams, Access/Function_descriptions_html  
**Päivämäärä:** 2026-03-09  
**Tekijä:** GitHub Copilot — 64-bit migraatioagentti

---

## Kriittiset muutokset

_Ei kriittisiä 32-bit/driver-ongelmia löytynyt._

---

## Siivous ja optimointi

### FunctionDiagrams

- **`Form_Funktiokaavio.cls`** — Poistettu 8 riviä vanhentunutta kommentoitua koodia, joka viittasi poistuneeseen `api_GetUserName`-API-kutsuun. Korvaus (`Me!RevBy`) on jo käytössä.

- **`Form_LisääKuviin_ACAD.cls`** — Poistettu tyhjä tapahtumankäsittelijä `Private Sub TTitleBlokki_BeforeUpdate(Cancel As Integer)`.

- **`Form_FuncBlock.cls`**, **`Form_Linkkien vaihto.cls`**, **`Form_Sub_RECIPES.cls`**, **`USysCheck.bas`**: analysoitu — DAO-cleanup (`Set x = Nothing`) löytyy oikein, ei muutoksia tarvittu.

### Function_descriptions_html

- **`GeneralCodes.bas`** — Lisätty defensiivinen `oTaulu`-cleanup `MuutaRev`-funktion alkuun: jos julkinen `oTaulu`-tietue on jo auki, se suljetaan ja vapautetaan ennen uuden avaamista. Tämä estää muistivuodon kun funktiota kutsutaan useita kertoja.

---

## Huomiot

- `Form_FUNC.cls`: `FBlock_DblClick` on tahallaan tyhjä ja dokumentoitu — varattuna tulevaa toiminnallisuutta varten. Ei poistettu.
- `Form_FrmASETUKSET.cls`, `Form_FrmMUOKKAUS.cls`, `Form_MOTORS subform.cls`, `Form_PIIRIT subform.cls`: analysoitu — ei merkittäviä ongelmia.
- `KAANNOS.bas`, `USysCheck.bas` (Function_descriptions_html): analysoitu — ei muutoksia tarvittu.
