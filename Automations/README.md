# VBA AUTOMATION SCRIPTS - KÄYTTÖOHJE

**Versio:** 2.0 (64-bit)  
**Päivitetty:** 27.2.2026  
**Ympäristö:** Microsoft Office 365 (64-bit)

---

## 📖 YLEISKATSAUS

Nämä PowerShell-skriptit automatisoivat VBA-moduulien päivittämisen Access-tietokantoihin ja Excel-työkirjoihin. Skriptit on optimoitu **64-bittiselle Office 365** -ympäristölle.

### Mitä skriptit tekevät?

#### **Access_automaatio.ps1**

- Avaa Access-tietokannan (.accdb)
- Päivittää VBA-moduulit (.bas) ja luokkamoduulit (.cls)
- Päivittää lomakkeiden (Form_*) VBA-koodin
- Poistaa automaattisesti tiedostojen header-metatiedot
- Tallentaa ja sulkee tietokannan turvallisesti

#### **Excel_automaatio.ps1**

- Käy läpi kaikki .xlsm-työkirjat hakemistossa
- Päivittää VBA-moduulit (.bas)
- Poistaa automaattisesti tiedostojen header-metatiedot
- Tallentaa työkirjat väliaikaisesti ja korvaa alkuperäiset

---

## ⚙️ JÄRJESTELMÄVAATIMUKSET

### Pakollinen

- ✅ **Windows 10/11** (64-bit)
- ✅ **PowerShell 5.1+** (64-bit)
- ✅ **Microsoft Office 365** (64-bit)
  - Access (Access_automaatio.ps1)
  - Excel (Excel_automaatio.ps1)

### Tarkista Office-versio

1. Avaa Word/Excel/Access
2. Tiedosto → Tili → Tietoja Wordista/Excelistä
3. Ylhäällä pitää näkyä **"Microsoft Office 365 (64-bit)"**

### Tarkista PowerShell-versio

Avaa PowerShell ja aja:

```powershell
[System.IntPtr]::Size
```

- Pitää palauttaa **8** (= 64-bit)
- Jos palauttaa **4** (= 32-bit), käynnistä **"Windows PowerShell"** (ei x86-versiota!)

---

## 🚀 PIKAOHJE

### Access-tietokannan päivitys

1. **Avaa PowerShell Administratorina**

2. **Siirry Automations-hakemistoon**

   ```powershell
   cd c:\database_migration\Automations
   ```

3. **Salli skriptin suoritus** (tarvitsee tehdä vain kerran)

   ```powershell
   Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass
   ```

4. **Aja skripti**

   ```powershell
   .\Access_automaatio.ps1
   ```

5. **Vastaa kysymyksiin:**
   - **Access file path:** Polku .accdb-tiedostoon (esim. `L:\PROJDATA\MAINEQ.accdb`)
   - **Component files folder:** Hakemisto, jossa .bas/.cls-tiedostot (esim. `c:\database_migration\Access\MAINEQ\`)

### Excel-työkirjojen päivitys

1. **Avaa PowerShell Administratorina**

2. **Siirry Automations-hakemistoon**

   ```powershell
   cd c:\database_migration\Automations
   ```

3. **Salli skriptin suoritus**

   ```powershell
   Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass
   ```

4. **Aja skripti**

   ```powershell
   .\Excel_automaatio.ps1
   ```

5. **Vastaa kysymyksiin:**
   - **Excel files folder:** Hakemisto, jossa .xlsm-tiedostot
   - **Module files folder:** Hakemisto, jossa .bas-tiedostot

---

## 🔧 OLETUSARVOJEN ASETTAMINEN

Jos käytät skriptejä toistuvasti samoilla poluilla, voit asettaa oletusarvot:

### Access_automaatio.ps1

Avaa tiedosto ja muokkaa rivejä 42-43:

```powershell
$DefaultAccessFilePath = 'L:\PROJDATA\MAINEQ.accdb'
$DefaultComponentPath = 'c:\database_migration\Access\MAINEQ\'
```

### Excel_automaatio.ps1

Avaa tiedosto ja muokkaa rivejä 56-57:

```powershell
$DefaultExcelFilesPath = 'c:\projektit\tools\'
$DefaultModulePath = 'c:\database_migration\Excel\Moduulit\'
```

Kun oletusarvot on asetettu, voit vain painaa `Enter` kysymysten kohdalla.

---

## 📋 KOMPONENTTIEN MÄÄRITTELY

### Access: Mitä moduuleja päivitetään?

Avaa `Access_automaatio.ps1` ja muokkaa riviä 66-73:

```powershell
$componentNames = @(
    "Module1",
    "General",
    "For ACAD Utility",
    "USysCheck",
    "Form_DBUsers", 
    "Form_Linkkien vaihto",
    "Form_Tee Kuvat"
)
```

- ✅ Lisää tai poista moduulien nimiä listalta
- ⚠️ **TÄRKEÄÄ:** Älä lisää tiedostopäätteitä (.bas/.cls)
- ⚠️ Lomakkeet (`Form_*`) päivitetään vain, jos ne ovat jo olemassa kannassa

### Excel: Mitä moduuleja päivitetään?

Avaa `Excel_automaatio.ps1` ja muokkaa riviä 79:

```powershell
$moduleNames = @("Module1", "Module2", "Module3")
```

---

## 🔒 TURVALLISUUSASETUKSET

Skriptit tarvitsevat pääsyn VBA-projekteihin. Jos saat virheen **"VBA Project is null"**, tee seuraavat asetukset:

### Access

1. Avaa Access
2. Tiedosto → Asetukset → Luottamuskeskus → Luottamuskeskuksen asetukset
3. **Makrojen asetukset:**
   - ✅ Valitse: "Luota VBA-projektin objektimallin käyttöön"
4. **Luotetut sijainnit:**
   - Lisää uusi sijainti: Tietokannan polku (esim. `L:\PROJDATA\`)
   - Lisää uusi sijainti: Moduulien polku (esim. `c:\database_migration\Access\`)
   - ✅ Valitse: "Tämän sijainnin alihakemistot ovat myös luotettuja"

### Excel

1. Avaa Excel
2. Tiedosto → Asetukset → Luottamuskeskus → Luottamuskeskuksen asetukset
3. **Makrojen asetukset:**
   - ✅ Valitse: "Luota VBA-projektin objektimallin käyttöön"
4. **Luotetut sijainnit:**
   - Lisää työkirjojen ja moduulien polut

### Windows: Poista tiedoston esto (jos verkosta ladattu)

1. Etsi .accdb/.xlsm-tiedosto
2. Napsauta hiiren kakkospainikkeella → Ominaisuudet
3. Jos näet alhaalla **"Turvallisuus: Tämä tiedosto on peräisin toisesta tietokoneesta..."**
4. ✅ Valitse **"Salli"** tai **"Unblock"**
5. OK → Käynnistä skripti uudelleen

---

## 📊 SKRIPTIN TULOSTE

### Onnistunut suoritus (Access)

```
09:15:23 [OK] Ajetaan 64-bittisessä PowerShellissä.
09:15:24 [ALUSTUS] Luodaan Access COM-objekti...
09:15:25 [OK] Access-objekti luotu.
Access file path:
Lisää polku Access-tiedostoon (.accdb) (paina Enter käyttääksesi oletusta): L:\PROJDATA\MAINEQ.accdb
09:15:35 [TIETOKANTA] Käsitellään: L:\PROJDATA\MAINEQ.accdb
09:15:35    ✓ Poistettiin Vain luku -attribuutti.
09:15:36    [AVAUS] Avataan tietokanta...
09:15:38    ✓ Tietokanta avattu onnistuneesti.
09:15:38    ✓ Access-varoitukset poistettu käytöstä.
09:15:38    [VBE] Haetaan VBA-projekti...
09:15:39    ✓ VBA-projekti avattu onnistuneesti.
09:15:39    [KOMPONENTIT] Aloitetaan päivitys...
09:15:39       [KÄSITTELY] Module1
09:15:40          ✓ Komponentti löytyi, päivitetään sisältö...
09:15:41          ✓ VALMIS: Module1 (45 → 48 riviä)
09:15:41       [KÄSITTELY] General
09:15:42          ✓ VALMIS: General (120 → 122 riviä)
...
09:16:15    [KOMPONENTIT] Kaikki komponentit käsitelty.
09:16:15    [TALLENNUS] Tallennetaan muutokset...
09:16:17    ✓ Tietokanta tallennettiin.
09:16:17    ✓ Tiedosto L:\PROJDATA\MAINEQ.accdb päivitetty onnistuneesti!
09:16:18 [CLEANUP] Siivotaan ja suljetaan Access-prosessi...
09:16:19    ✓ VBA Project vapautettu.
09:16:19    ✓ Database-objekti vapautettu.
09:16:19    ✓ Access.Quit() suoritettu.
09:16:20    ✓ Access COM-objekti vapautettu.
09:16:20 [OK] Siivous valmis.
```

### Onnistunut suoritus (Excel)

```
14:22:10 [OK] Ajetaan 64-bittisessä PowerShellissä.
14:22:11 [ALUSTUS] Luodaan Excel COM-objekti...
14:22:12 [OK] Excel-objekti luotu.
14:22:15 [TYÖKIRJAT] Haetaan .xlsm-tiedostot kohteesta: c:\projektit\tools\
14:22:15 [OK] Löytyi 3 työkirjaa käsiteltäväksi.

14:22:15 [TYÖKIRJA 1/3] KÄSITELLÄÄN: c:\projektit\tools\Kysely1.xlsm
14:22:15    ✓ Poistettiin Vain luku -attribuutti tiedostojärjestelmästä.
14:22:16    [AVAUS] Avataan työkirja...
14:22:18    ✓ Työkirja avattu onnistuneesti.
14:22:18    ✓ VBA-projekti avattu.
14:22:18    [MODUULIT] Aloitetaan päivitys...
14:22:18       [KÄSITTELY] Module1
14:22:19          ✓ Moduuli löytyi, päivitetään sisältö...
14:22:20          ✓ VALMIS: Module1 (85 → 88 riviä)
14:22:20    [MODUULIT] Kaikki moduulit käsitelty.
14:22:20    [TALLENNUS] Tallennetaan väliaikaiseen tiedostoon...
14:22:22    ✓ Väliaikainen tallennus onnistui.
14:22:22    ✓ Työkirja suljettu.
14:22:22    [KORVAUS] Poistetaan alkuperäinen tiedosto...
14:22:22    ✓ Alkuperäinen poistettu.
14:22:22    [KORVAUS] Nimetään uusi tiedosto alkuperäiseksi...
14:22:23    ✓ Tiedosto c:\projektit\tools\Kysely1.xlsm päivitetty onnistuneesti!
...
14:23:45 [VALMIS] Kaikki työkirjat käsitelty!
14:23:46 [CLEANUP] Siivotaan ja suljetaan Excel-prosessi...
14:23:47    ✓ Excel.Quit() suoritettu.
14:23:47    ✓ Excel COM-objekti vapautettu.
14:23:47 [OK] Siivous valmis.
```

---

## ⚠️ YLEISIMMÄT VIRHEET

### 1. "Tämä skripti on suoritettava 64-bittisessä PowerShellissä"

**Syy:** Avattiin PowerShell (x86) vahingossa  
**Ratkaisu:**

- Sulje nykyinen PowerShell
- Avaa: `Windows PowerShell` (ilman x86-merkintää)
- Tarkista: `[System.IntPtr]::Size` pitää palauttaa **8**

### 2. "VBA Project is null. Check Trust Center settings."

**Syy:** VBA-projektin objektimalliin ei ole pääsyä  
**Ratkaisu:** Katso kohta **"Turvallisuusasetukset"** yllä

### 3. "Access-tiedostoa ei löydy tai polku on hakemisto"

**Syy:** Väärä polku tai kirjoitusvirhe  
**Ratkaisu:**

- Tarkista polku (käytä Tab-key automaattitäydennykseen)
- Windows Explorerissa: Shift + Oikea nappi → "Kopioi polku nimellä"
- Liitä polku PowerShelliin (oikea nappi)

### 4. "Lomake X ei ole olemassa kannassa. Lomakkeita ei voi luoda automaattisesti."

**Syy:** Yritetään päivittää lomaketta, jota ei ole kannassa  
**Ratkaisu:**

- Poista lomake `$componentNames`-listalta (Access_automaatio.ps1, rivi 66-73)
- TAI luo lomake ensin manuaalisesti Accessissa

### 5. "Tiedostoa ei voitu avata 3 yrityksen jälkeen"

**Syy:** Tiedosto on lukittu (toinen käyttäjä, OneDrive-synkronointi)  
**Ratkaisu:**

- Sulje Access/Excel
- Tarkista Task Managerista, onko vanhoja Access.exe/Excel.exe-prosesseja (lopeta ne)
- Jos OneDrive: Odota synkronoinnin valmistumista
- Yritä uudelleen

### 6. "Script execution is disabled"

**Syy:** PowerShellin ExecutionPolicy on rajoitettu  
**Ratkaisu:**

```powershell
Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass
```

- Tai avaa PowerShell Administratorina ja aja:

```powershell
Set-ExecutionPolicy -Scope CurrentUser -ExecutionPolicy RemoteSigned
```

---

## 🗂️ TIEDOSTORAKENNE

```
Automations/
├── Access_automaatio.ps1          # Access-tietokantojen päivitys
├── Excel_automaatio.ps1           # Excel-työkirjojen päivitys
├── README.md                      # Tämä ohje
├── REFACTORING_DOCUMENTATION.md  # Tekninen dokumentaatio
└── access_vba_exports/            # VBA-moduulien varmistuskopiot
    ├── MAINEQ_OLD/
    ├── PIPE_OLD/
    └── ...
```

---

## 📞 TUKI JA VIANMÄÄRITYS

### Debug-moodi

Jos skripti kaatuu, tarkista:

1. **Aikaleima** viimeisestä onnistuneesta vaiheesta
2. **Virheviesti** (punainen teksti)
3. **Virhetyyppi** (jos tulostuu)
4. **Stack Trace** (jos tulostuu)

### Lokitiedostot

Skriptit eivät luo automaattisia lokitiedostoja, mutta voit ohjata tulosteen tiedostoon:

```powershell
.\Access_automaatio.ps1 *> loki.txt
```

### Lisätietoja

- Tekninen dokumentaatio: `REFACTORING_DOCUMENTATION.md`
- Projektiloki: `c:\database_migration\Logs\AUTOMATIONS_LOG.md`

---

## 🔄 VERSIOHISTORIA

### Versio 2.0 (27.2.2026) - 64-bit Migration

- ✅ Täydellinen refaktorointi 64-bit Office 365:lle
- ✅ Korjattu kriittinen VBA Project -bugi (Access)
- ✅ Lisätty finally-lohko (ei zombie-prosesseja)
- ✅ Timestamp-logging
- ✅ Parannettu header-parsaus
- ✅ Lomakkeiden turvallinen käsittely
- ✅ FileFormat-korjaus (Excel SaveAs)

### Versio 1.x (2025) - Legacy

- ⚠️ Suunniteltu 32-bitille
- ❌ Access-skripti ei toiminut (VBA Project -bugi)
- ❌ Zombie-prosessiongelmia

---

## 📄 LISENSSI JA TEKIJÄNOIKEUDET

**Projekti:** 64-bit Legacy VBA Migration  
**Kehittäjä:** Senior System Architect (Legacy Systems)  
**Asiakasorganisaatio:** [Yrityksen nimi]  
**Luotu:** 2025-2026

**Käyttöoikeus:** Sisäinen käyttö projektissa.

---

**Päivitetty:** 27.2.2026  
**Versio:** 2.0  
**Status:** ✅ PRODUCTION READY
