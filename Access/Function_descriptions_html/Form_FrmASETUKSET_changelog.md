# Muutosloki

- **Tiedosto:** `Form_FrmASETUKSET.cls`
- **Päivämäärä:** 2026-11-10

## Kriittiset muutokset

- `Linkkaa` laskee epäonnistuneet linkitykset (`LinkkausVirheet`), ja `Command2_Click` ilmoittaa "Taulukot linkattu" vain, jos kaikki onnistuivat.
- Lähdetietokannan olemassaolo tarkistetaan (`Dir$`) ennen vanhan linkin poistoa, jotta toimiva linkki ei katoa.
- `DoCmd.Close acForm, "FrmASETUKSET"` (puuttuva objektityyppi).

## Siivous ja optimointi

- Poistettu rivin lopun turha välilyönti ja korjattu kommentin kirjoitusvirhe ("oletuspoluilla").
- Jätetty ennalleen: kommentoitu `CONTROL`-linkitys (toimintakaaviot, tarkoituksella pois käytöstä; `Form_FrmMUOKKAUS` suojattu taulun puuttumiselta) ja `SniffUser`-kutsu lomakkeen latauksessa.
