# Muutosloki — Lukituskaavio

**Tiedosto:** `Access/Lukituskaavio/` (kaikki moduulit)
**Päivämäärä:** 2025-07-11
**Haara:** `Lukituskaavio`

---

## Kriittiset muutokset

### APIKoodit.bas

- `wu_GetUserName`: `nSize As LongPtr` → `ByRef nSize As Long` — Win32 DWORD on 32-bittinen, ei LongPtr
- `GetOpenFileName`: paluuarvo `As Long` → `As LongPtr` — 64-bit osoitinyhteensopivuus
- `OPENFILENAME`-rakenne: `hwndOwner`, `hInstance`, `lCustData`, `lpfnHook` kääritty `#If VBA7` ehtolausekkeeseen
- `ValitseTiedosto`: `lReturn`-muuttuja kääritty `#If VBA7 Then / Dim lReturn As LongPtr / #Else / Long / #End If`
- Englanninkielinen kommentti `'Return value` → `'Palautusarvo`

### Form_Aloitus.cls

- Lisätty `Option Explicit`
- `Me.`-etuliite lisätty kaikkiin `Polku`-, `LMAINEQ`- ja `LINSTRU`-viittauksiin

### Form_FromTo.cls

- Lisätty `Option Explicit`
- Lisätty puuttuvat muuttujaesittelyt: `Dim L As String`, `Dim edel As String`
- `AcadBlockReference` → `Object` (myöhäinen sidonta 64-bit-yhteensopivuuden vuoksi)
- `AcadSelectionSet` → `Object` (myöhäinen sidonta)
- `Me.`-etuliite lisätty: `TPolku`, `Lista`, `FROMTOBLOCKS` — kaikki esiintymät
- AutoCAD-kutsujen merkkijonot suomennettu (`"Set start point"` jne.)

### Form_Funktiokaavio.cls

- Poistettu käytöstä poistunut `Private Declare api_GetUserName` (ei kutsuja jäljellä)
- Lisätty `Option Explicit`
- Korjattu yhteen kirjoitetut rivit (`DoCmd.SetWarnings TrueDebug.Print...` jm.)
- `Viesti = MsgBox(...)` → `MsgBox "...", vbOKOnly` (määrittelemätön muuttuja poistettu)
- `LTITLE.Caption` → `Me.LTITLE.Caption` (Komento57\_Click, Komento58\_Click)
- `RProcess.VALUE`, `RCode.VALUE`, `RInfo.VALUE` → `Me.RProcess.Value` jne. (Komento67\_Click)
- `RECIPES.Form.ID` → `Me.RECIPES.Form.ID` (Command83\_Click, kaikki esiintymät)

### Form_Interlocking.cls

- Poistettu kaikki `Debug.Print`-lauseet (ml. `If BlockNimi <> "" Then Debug.Print ...`)
- `AcadObject`, `AcadBlockReference`, `AcadLine`, `AcadSelectionSet` → `Object`

### Form_LineForm.cls

- `Tyyppi.VALUE` → `Me.Tyyppi.Value` (COK\_Click)
- `Sijainti.VALUE` → `Me.Sijainti.Value` (COK\_Click)
- AutoCAD GetPoint -kehotteet suomennettu

---

## Siivous ja optimointi

- `Debug.Print` poistettu kaikista tuotantotiedostoista
- Kaikki suomalaiset kommentit tarkistettu; ä- ja ö-kirjaimet korjattu
- Vanha kommentoitu koodi (`api_GetUserName`-kutsu) poistettu Form_Funktiokaavio.cls:stä
- Tyhjiä `If`-lohkoja tunnistettu Command98:ssa (jätetty, koska kyseessä on toiminnallinen DebugPoint)
