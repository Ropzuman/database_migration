# Code Review Raportti – Analyysi ulkoista auditointia vasten

## 📊 Yhteenveto

Koodikanta on kehittynyt valtavasti, ja suuret arkkitehtuuriset muutokset (kuten UDT-rakenteet metadatassa, `.CopyFromRecordset` -optimointi ja tietokantayhteyden `ReadOnly`-pakotus) on otettu onnistuneesti käyttöön. Liitteenä oleva Word-muotoinen katselmointiraportti nostaa kuitenkin esiin joukon tärkeitä reunaehtoja (edge cases) ja piileviä bugeja, joita ei ole vielä täysin korjattu koodissa. Näistä merkittävimmät liittyvät tyyppimuunnoksiin, ajastimen (Timer) käyttäytymiseen ja globaalin tilan hallintaan.

## 🚨 Kriittiset löydökset (Tietoturva & Vakaus)

* **Integer Overflow -riski (Vakaus):** Raportissa mainitaan "CInt → CLng for TRow/TCol in VaihdaLinkit". Excelin maksimirivimäärä on yli miljoona, mutta VBA:n `Integer` (`CInt`) kaatuu (Overflow error), jos arvo ylittää 32 767. Koska koodi käsittelee rivejä, sarakkeita ja iteraatioita, kaikki rivi- ja sarakeindeksit on poikkeuksetta pakotettava `Long`-tyyppisiksi.
* **Ajastimen (Timer) "Midnight Wraparound" -bugi (Vakaus):** VBA:n `Timer`-funktio palauttaa sekunnit keskiyöstä lähtien. Jos makron suoritus alkaa juuri ennen keskiyötä (esim. 23:59:59) ja päättyy sen jälkeen, laskukaava `Timer - StartTime` tuottaa negatiivisen luvun, mikä voi rikkoa suorituskykyraportoinnin ja lokituksen.
* **CheckOK-tilan nollaus (Toiminnallisuus):** Globaali muuttuja `CheckOK` jää muistiin edellisestä ajosta, mikäli sitä ei nollata työkirjan sulkeutuessa tai avautuessa. Tämä mahdollistaa skenaarion, jossa käyttäjä ajaa uuden, virheellisen datan läpi, mutta vanha `CheckOK = True` sallii prosessin jatkumisen.
* **SaveAs-formaatin varmistaminen (Tietoturva/Vakaus):** Raportti varoittaa "SaveAs format check for macro-enabled workbooks" -riskistä. Jos uusi työkirja tallennetaan `.xlsm` -päätteellä mutta tiedostoformaatiksi (FileFormat) on ohjelmoitu normaali `.xlsx`, tiedosto korruptoituu ja makrot menetetään.

## ⚠️ Huomioitavaa (Toiminnallisuus & Suorituskyky)

* **Vakioiden päällekkäisyys (DRY-periaate):** Kuten raportissa todetaan ("Centralise MAX_EXCEL_COLUMNS constant"), vakio `Private Const MAX_EXCEL_COLUMNS As Long = 16384` on julistettu sekä `Module1.bas` että `Module2.bas` -tiedostoissa. Se tulisi julistaa kerran `Public Const` -määrityksellä yhdessä paikassa, jotta ylläpidettävyys säilyy.
* **FreezePanes-kutsun järjestys:** FreezePanes (ruutujen kiinnitys) -operaatiot vaativat toimiakseen oikein sen, että ikkuna on aktiivinen ja päivitettävissä. Raportin mukaan nämä kutsut pitää varmistaa suoritettavaksi vasta `EndFastMode` -kutsun jälkeen (kun `ScreenUpdating = True`).
* **Vanhentunut DAO.DBEngine.36 (Ylläpidettävyys):** Auditointiraportti kehottaa poistamaan `DAO.DBEngine.36` fallback-koodin. Nykyaikaisissa 64-bittisissä Office-ympäristöissä tämä vanha Access 97/2000 -moottori on tarpeeton ja herkkä kaatumaan. `.120` (ACE OLEDB) riittää täysin.

## 💡 Parannusehdotukset (Ylläpidettävyys)

* **Tyhjä tietokanta (RMAX = 0):** Jos tietokantakysely ei palauta rivejä (`RecordCount = 0`), RMAX-arvoksi saattaa jäädä 0, mikä kaataa myöhemmät for-silmukat. Koodiin tulee lisätä eksplisiittinen virheilmoitus ja poistuminen (Exit Sub), jos RMAX on 0 `Checkout`-vaiheen jälkeen.

## 🛠️ Korjattu koodi

### 1. Integer-tyyppien päivittäminen Long-tyyppiin (Estää kaatumiset)

Kaikki rivi-/sarakekäsittelyt tulee refaktoroida käyttämään suuria kokonaislukuja.

```vba
' KORJATTU: CInt korvattu CLng-funktiolla ja muuttujat vaihdettu Long-tyyppiin.
' Esimerkki oletetusta VaihdaLinkit-funktiosta:
Dim TRow As Long, TCol As Long
' VÄÄRIN: TRow = CInt(Split(Osoite, ",")(0))
' OIKEIN:
TRow = CLng(Split(Osoite, ",")(0))
TCol = CLng(Split(Osoite, ",")(1))

2. Timer-bugin korjaus (Keskiyön ylitys)

Voit ratkaista ajastinongelman käyttämällä VBA:n Date ja Timer -funktioiden yhdistelmää tai yksinkertaisemmin modulo-operaattoria tai IF-lauseketta.
VBA

' KORJATTU: Turvallinen keston laskenta (estää negatiiviset arvot keskiyön ylittyessä)
Dim StartTime As Double
Dim EndTime As Double
Dim TotalTime As Double

StartTime = Timer
' ... [Makron suoritus] ...
EndTime = Timer

If EndTime < StartTime Then
    ' Keskiyö ylittyi makron aikana (Timer nollaantui)
    TotalTime = (86400 - StartTime) + EndTime
Else
    TotalTime = EndTime - StartTime
End If

Debug.Print "Makron suoritusaika: " & Format(TotalTime, "0.00") & " sekuntia."

3. Vakioiden keskittäminen

Poista Private Const MAX_EXCEL_COLUMNS As Long = 16384 Module2:sta ja muuta Module1:ssä oleva vakio julkiseksi:
VBA

' Module1.bas (Ylätunniste)
Public Const MAX_EXCEL_COLUMNS As Long = 16384

4. Turvallinen SaveAs-rutiini (GenPrintout)

Jos koodi tallentaa työkirjan, varmista oikea FileFormat tunnisteen perusteella:
VBA

' KORJATTU: Excelin tallennusformaatin varmistus
Dim SaveFileName As String
SaveFileName = DIPath & DIFile

If InStr(1, SaveFileName, ".xlsm", vbTextCompare) > 0 Then
    ' 52 = xlOpenXMLWorkbookMacroEnabled
    ActiveWorkbook.SaveAs Filename:=SaveFileName, FileFormat:=52
Else
    ' 51 = xlOpenXMLWorkbook
    ActiveWorkbook.SaveAs Filename:=SaveFileName, FileFormat:=51
End If
