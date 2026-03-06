# Muutosloki — Function_descriptions_html

**Tiedosto:** `Access/Function_descriptions_html/` (moduulikokoelma)
**Päivämäärä:** 2026-03-06
**Haara:** FunctionDescriptions

---

## Kriittiset muutokset

### KAANNOS.bas — SQL-injektiosuojaus DLookup-kutsuille

- **Ongelma:** `DLookup`-funktioiden ehtolausekkeet rakennettiin suoraan merkkijonoyhdistelyillä (`& Osat(0) &` jne.) ilman puhdistusta. Heittomerkki (`'`) parametreissä kaatoi SQL-moottorin syntaksivirheeseen.
- **Korjaus:** Lisätty `Dim sPar0 As String, sPar1 As String, sPar2 As String` ja kaikki `Osat`-arvot sanitoidaan `Replace(Nz(Osat(n), ""), "'", "''")` -kutsulla ennen kyselyä. Käytetty sekä MAINEQ- että Loops-hauissa.

### Form_FrmASETUKSET.cls — CurrentDb välimuistitettu Linkkaa-aliohjelmassa

- **Ongelma:** `CurrentDb` kutsuttiin viisi kertaa `Linkkaa`-aliohjelmassa, jokainen kutsu loi uuden `DAO.Database`-instanssin. Tämä on suorituskykysyöppö ja voi aiheuttaa ongelmia `.Refresh`-kutsun kanssa.
- **Korjaus:** Lisätty `Dim db As DAO.Database`, alustettu kerran `Set db = CurrentDb`. Kaikki `CurrentDb.`-viittaukset korvattu `db.`-viittauksilla. `Set tdf = Nothing` ja `Set db = Nothing` lisätty sekä normaaliin Exit Sub -polkuun että ErrorHandler-lohkoon.

### USysCheck.bas — Virheellinen DB.Close poistettu CurrentDb-viittaukselta

- **Ongelma:** `DB.Close` kutsuttiin `DB = CurrentDb` -viittauksen päälle, mikä on virheellinen tapa. `CurrentDb`-viittauksen sulkeminen `.Close`-kutsulla voi epävakaistaa Access-istunnon.
- **Korjaus:** Poistettu `DB.Close` sekä normaalilta polulta että ErrorHandler-lohkosta. Jätetty vain `Set DB = Nothing`. Lisätty selittävä kommentti.

---

## Toiminnalliset parannukset

### Form_MOTORS subform.cls — TurvallinenKursori-rajaus

- **Ongelma:** Globaali `Kursori`-muuttuja käytettiin suoraan `Left$`/`Mid$`-kutsuissa ilman tarkistusta. Jos `Kursori > Len(teksti)`, käytös voi olla odottamaton.
- **Korjaus:** Lisätty `Dim KohdeTeksti As String` ja `Dim TurvallinenKursori As Long`. `KohdeTeksti = Nz(KohdeTextBox.Value, "")` kerran (ei toistuvia `.Value`-viittauksia). `TurvallinenKursori` rajataan `[0, Len(KohdeTeksti)]`-välille ennen `Left$`/`Mid$`-kutsuja. `SelStart` käyttää myös `TurvallinenKursori`-arvoa. Turha `If IsNull(...)`-haara poistettu.

### Form_PIIRIT subform.cls — TurvallinenKursori-rajaus

- **Ongelma:** Sama kuin MOTORS: `Kursori` käytetty suoraan ilman rajaustarkoitusta.
- **Korjaus:** Sama ratkaisu kuin MOTORS: `KohdeTeksti`-välimuistitus, `TurvallinenKursori`-rajaus ja `SelStart`-päivitys.

---

## Siivous ja optimointi

- Kaikki muutokset dokumentoitu `' Päivitetty:`-headeriin ao. moduuleissa.
- `CodeReview_functiondescriptions.md` päivitetty `Refaktoroitu: 2026-03-06` -otsikolla ja yhteenvetotaululla.
- Tilapäiset skriptitiedostot (`_fix_motors_piirit.py`, `_fix_motors_piirit.ps1`) poistetaan Automations-kansiosta siivousvaiheessa.
