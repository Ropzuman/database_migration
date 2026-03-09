# Access/MAINEQ — Phase 3 Viimeistely

**Tiedosto:** Access/MAINEQ (kaikki lomakkeet ja raportit)  
**Päivämäärä:** 2026-03-09  
**Tekijä:** GitHub Copilot — 64-bit migraatioagentti

---

## Kriittiset muutokset

_Ei kriittisiä 32-bit/driver-ongelmia löytynyt. Kaikki `Declare`-lausekkeet käyttävät `PtrSafe`. Ei `Jet.OLEDB`-käyttöä._

---

## Siivous ja optimointi

- **`Form_Revisiointi.cls`** — `PaivitaIDRev`-funktiossa lisätty `Taul.Close` + `Set Taul = Nothing` ensin recordsetin sulkemiseen ennen uudelleenkäyttöä (aiemmin `OpenRecordset` kutsuttiin kahdesti ilman sulkemista). `Tee_Click`-tapahtumassa lisätty `Vert.Close` + `Set Vert = Nothing` ennen `"Valmis"`-ilmoitusta.

- **`Form_KuvienGenerointi.cls`** — `GenRev_Click`-funktiossa lisätty `Set Revisiot = Nothing` silmukan sisälle ennen `Kuvat.MoveNext`-kutsua. Lisätty `Set DB = Nothing` siivousosioihin kolmessa aliohjelmassa (`GenRev_Click`, `Teetaulukko_Click`, `GenKuvat_Click`). Poistettu duplikaatti `Set Kuvat = Nothing`.

- **`Form_EQUIPMENT.cls`** — Lisätty `Taulukko.Close` ennen `Set Taulukko = Nothing` `EqType_AfterUpdate`-tapahtumassa. DAO-recordset suljetaan nyt oikeaoppisesti ennen vapautusta.

- **`Form_EQUIPMENT_FI.cls`** — Sama korjaus kuin `Form_EQUIPMENT.cls`: lisätty `Taulukko.Close` ennen `Set Taulukko = Nothing`.

- **`Report_PÄÄLAITTEET_BAAN.cls`** — Lisätty `On Error GoTo ErrorHandler` -rakenne koko raportin generoinnille. Lisätty `Tiedot.Close` ennen `Set Tiedot = Nothing`. ErrorHandler varmistaa resurssien vapautuksen virhetilanteessa.

- **`Form__qryMotorData_subform.cls`** — Poistettu tyhjä tapahtumankäsittelijä `Private Sub ID_Click()`.

---

## Huomiot

- `Report_PÄÄLAITTEET_BAAN.cls`: `xlApp.Quit` on tahallaan jätetty pois — Excel jää näkyviin käyttäjälle interaktiiviseen käyttöön. Tämä on suunniteltu toimintamalli.
- `Report_MOOTTORIT.cls`, `Report_PÄÄLAITTEET.cls`, `Form_MoottTilaus.cls`: analysoitu — `xlApp.Quit` löytyy oikein, ei muutoksia tarvittu.
- `Form_USysRevText.cls`, `Form_UsysRevTextDrive.cls`, `Form_Motors_Subform.cls` yms. alilomaket: ei merkittäviä ongelmia.
- `GeneralCodes.bas`: ei ongelmia (ei sekoitettava `Function_descriptions_html`-kansion tiedostoon).
