# VBA AUTOMATION SCRIPTS - KÄYTTÖOHJE

**Versio:** 2.3 (64-bit AutoCAD Migration)  
**Päivitetty:** 27.2.2026  
**Ympäristö:** Microsoft Office 365 (64-bit) + AutoCAD 2019 (64-bit)

---

## 📖 YLEISKATSAUS

Nämä PowerShell-skriptit automatisoivat VBA-moduulien päivittämisen Access-tietokantoihin, Excel-työkirjoihin ja AutoCAD DVB-projektien purkamisen. Skriptit on optimoitu **64-bittiselle Office 365** -ympäristölle ja **AutoCAD 2019** -integrointiin.

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

#### **AutoCAD_DVB_Import_Run.ps1** ⭐ UUSI

- Importoi korjatut VBA-tiedostot takaisin AutoCAD DVB-projekteihin
- Käy läpi kaikki `exported/<projektinimi>/` -kansiot automaattisesti
- Korvaa komponentit korjatuilla .bas/.cls/.frm-tiedostoilla
- Tallentaa korjatut projektit kansioon `AutoCAD/migrated/`
- Käyttää AutoCAD COM + RunMacro + SaveDVB -ketjua

#### **_scan_64bit.ps1** + **_scan_frm.ps1**

- Skannaavat kaikki exportoidut VBA-tiedostot 64-bitti-ongelmien varalta
- Tunnistavat: `Declare` ilman `PtrSafe`, `As Long` handle-muuttujissa
- `_scan_64bit.ps1` käsittelee .bas ja .cls tiedostot
- `_scan_frm.ps1` käsittelee .frm lomake-tiedostot

---

## ⚙️ JÄRJESTELMÄVAATIMUKSET

### Pakollinen

- ✅ **Windows 10/11** (64-bit)
- ✅ **PowerShell 5.1+** (64-bit)
- ✅ **Microsoft Office 365** (64-bit)
  - Access (Access_automaatio.ps1)
  - Excel (Excel_automaatio.ps1)
- ✅ **AutoCAD 2019** (64-bit) tai uudempi (AutoCAD_DVB_Export.ps1)

### Tarkista Office-versio

1. Avaa Word/Excel/Access
2. Tiedosto → Tili → Tietoja Wordista/Excelistä
3. Ylhäällä pitää näkyä **"Microsoft Office 365 (64-bit)"**

### Tarkista AutoCAD-versio

1. Avaa AutoCAD
2. Kirjoita komentoriviin: `ABOUT`
3. Tarkista: **"AutoCAD 2019"** tai uudempi + **"64-bit"**

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
   - **VAIHE 1 - Component files folder:** Hakemisto, jossa .bas/.cls-tiedostot (esim. `c:\database_migration\Access\MAINEQ\`)
   - **VAIHE 2 - Access file path:** Polku .accdb-tiedostoon (esim. `L:\PROJDATA\MAINEQ.accdb`)

   **HUOM:** Skripti skannaa automaattisesti kaikki .bas ja .cls -tiedostot lähdehakemistosta - ei tarvitse määritellä moduulilistaa!

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
   - **VAIHE 1 - Module files folder:** Hakemisto, jossa .bas-tiedostot
   - **VAIHE 2 - Excel files folder:** Hakemisto, jossa .xlsm-tiedostot

   **HUOM:** Skripti skannaa automaattisesti kaikki .bas-tiedostot lähdehakemistosta - ei tarvitse määritellä moduulilistaa!

### AutoCAD DVB-projektien 64-bitti-migraatio

#### Vaihe 1: Exportointi (kertaluonteinen — jo tehty)

VBA-koodi on jo exportoitu kansioon `AutoCAD/exported/` (43 projektia, 112 komponenttia).  
Export tehtiin manuaalisesti AutoCADin COM-rajapinnan kautta 27.2.2026.

#### Vaihe 2: Skannaus ja korjaus

1. **Aja 64-bitti-skanneri:**

   ```powershell
   powershell.exe -ExecutionPolicy Bypass -File ./_scan_64bit.ps1
   powershell.exe -ExecutionPolicy Bypass -File ./_scan_frm.ps1
   ```

   Tulostaa kaikki `DECLARE_NO_PTRSAFE` ja `LONG_HANDLE` -ongelmat tiedostoittain.

2. **Korjaa löydetyt ongelmat** manuaalisesti tai editorilla `AutoCAD/exported/`-kansiossa.

#### Vaihe 3: Import takaisin DVB-tiedostoihin

1. **Varmista AutoCAD on käynnissä** (AutoCAD 2019+)

2. **Siirry Automations-hakemistoon**

   ```powershell
   cd c:\database_migration\Automations
   ```

3. **Aja import-skripti WinPS 5.1:llä** (ei pwsh/PS7)

   ```powershell
   powershell.exe -ExecutionPolicy Bypass -File AutoCAD_DVB_Import_Run.ps1
   ```

   **HUOM:** Skripti käyttää `powershell.exe` (Windows PowerShell 5.1) — ei toimi `pwsh`:lla (PS7), koska `Marshal.GetActiveObject()` puuttuu PS7:sta.

4. **Korjatut DVB-tiedostot** tallennetaan kansioon `AutoCAD/migrated/`.

---

## 🔧 OLETUSARVOJEN ASETTAMINEN

Jos käytät skriptejä toistuvasti samoilla poluilla, voit asettaa oletusarvot:

### Access_automaatio.ps1

Avaa tiedosto ja muokkaa rivejä 43-44:

```powershell
$DefaultComponentPath = 'c:\database_migration\Access\MAINEQ\' # Päivitettyjen komponenttian polku
$DefaultAccessFilePath = 'L:\PROJDATA\MAINEQ.accdb' # Access-tiedoston polku
```

### Excel_automaatio.ps1

Avaa tiedosto ja muokkaa rivejä 46-47:

```powershell
$DefaultModulePath = 'c:\database_migration\Excel\Moduulit\' # Päivitettyjen moduulien polku
$DefaultExcelFilesPath = 'c:\projektit\tools\' # Excel-työkirjojen polku
```

### AutoCAD_DVB_Import_Run.ps1

Avaa tiedosto ja muokkaa riviä 5:

```powershell
$dvbSource  = 'C:\database_migration\AutoCAD'       # Alkuperäiset DVB-tiedostot
$exportRoot = 'C:\database_migration\AutoCAD\exported'  # Korjatut VBA-lähdekoodit
$migrRoot   = 'C:\database_migration\AutoCAD\migrated'  # Import-tulos
```

---

## 📋 AUTOMAATTINEN MODUULIHAKU (UUSI!)

### Moduulit skannataan automaattisesti

**Skriptit eivät enää vaadi kiinteäkoodattuja moduulilistoja!**

Molemmat skriptit **skannaavat automaattisesti** kaikki .bas ja .cls -tiedostot lähdehakemistosta:

#### Access_automaatio.ps1

```powershell
# Skannaa kaikki .bas ja .cls -tiedostot:
$componentFiles = Get-ChildItem -Path $componentPath -Filter "*.bas", "*.cls"
$componentNames = $componentFiles | ForEach-Object { $_.BaseName }
```

Tulostaa esim:

```
09:15:35 [OK] Löytyi 7 komponenttia:
  - Module1
  - General
  - For ACAD Utility
  - USysCheck
  - Form_DBUsers
  - Form_Linkkien vaihto
  - Form_Tee Kuvat
```

#### Excel_automaatio.ps1

```powershell
# Skannaa kaikki .bas -tiedostot:
$moduleFiles = Get-ChildItem -Path $modulePath -Filter "*.bas"
$moduleNames = $moduleFiles | ForEach-Object { $_.BaseName }
```

Tulostaa esim:

```
14:22:15 [OK] Löytyi 3 moduulia:
  - Module1
  - Module2
  - Module3
```

### Edut

- ✅ **Ei manuaalista konfigurointia** - lisää vain .bas/.cls-tiedosto kansioon
- ✅ **Ei virheitä kirjoitusvirheiden vuoksi** - tiedostonimet määrittävät moduulit
- ✅ **Helpompi ylläpito** - sama skripti toimii kaikille projekteille
- ⚠️ Lomakkeet (`Form_*`) päivitetään vain, jos ne ovat jo olemassa kannassa

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

- Poista lomakkeen .cls-tiedosto komponenttihakemistosta
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
├── AutoCAD_DVB_Import_Run.ps1     # AutoCAD DVB-projektien import (64-bit korjattu)
├── _scan_64bit.ps1                # 64-bitti-skanneri (.bas/.cls)
├── _scan_frm.ps1                  # 64-bitti-skanneri (.frm-lomakkeet)
├── _fix_encoding.ps1              # UTF-8 BOM -lisäys PS5.1-yhteensopivuuteen
├── README.md                      # Tämä ohje
├── REFACTORING_DOCUMENTATION.md  # Tekninen dokumentaatio
└── access_vba_exports/            # VBA-moduulien varmistuskopiot
    ├── MAINEQ_OLD/
    ├── PIPE_OLD/
    └── ...
```

```
AutoCAD/
├── *.dvb                          # Alkuperäiset binaariset DVB-projektit
├── exported/                      # Exportatut VBA-tiedostot (64-bit korjattu)
│   ├── Arkistotulostus/
│   │   ├── General.bas            # PtrSafe + LongPtr korjattu
│   │   └── Formi.frm
│   ├── MultiPlot/
│   │   └── General.bas            # PtrSafe + LongPtr + BrowseInfo korjattu
│   ├── ...                        # 43 projektia yhteensä
│   └── _64BIT_ANALYSIS.txt        # Skannausraportti
└── migrated/                      # Import-skriptin tulostama kansio
    └── *.dvb                      # 64-bit korjatut DVB-tiedostot
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

### Versio 2.3 (27.2.2026) - AutoCAD DVB 64-bit Migration

- ✅ **AutoCAD_DVB_Import_Run.ps1** - DVB-projektien import korjatuista .bas/.cls/.frm-tiedostoista
- ✅ **_scan_64bit.ps1** - Automaattinen 64-bitti-skanneri .bas/.cls-tiedostoille
- ✅ **_scan_frm.ps1** - Automaattinen 64-bitti-skanneri .frm-lomakkeille
- ✅ **64-bitti-korjaukset** - 78 ongelmaa korjattu 16 tiedostossa + 2 .frm-tiedostossa
- ✅ **PtrSafe-muunnos** - Kaikki Declare-lauseet päivitetty `PtrSafe`-avainsanalla
- ✅ **LongPtr-muunnos** - Handle-muuttujat muutettu `As LongPtr` -tyypeiksi
- ✅ **BrowseInfo-tyyppirakenne** - `SHBrowseForFolder`-kutsuihin liittyvät tyypit korjattu
- ✅ **RunMacro+SaveDVB** - Import käyttää RunMacro-tekniikkaa DVB-tallennukseen
- 🗑️ **Poistettu:** AutoCAD_DVB_Export.ps1, AutoCAD_DVB_Test.ps1 ja muut tilapäisskriptit

### Versio 2.2 (27.2.2026) - AutoCAD DVB Export

- ✅ Exportoitu 43/43 DVB-projektia (112 komponenttia) kansioon `AutoCAD/exported/`
- ✅ AutoCAD COM-integraatio (`GetActiveObject` + VBE-rajapinta)
- ✅ PowerShell 5.1 / UTF-8 BOM -yhteensopivuuskorjaukset

### Versio 2.1 (27.2.2026) - Automation & UX Improvements

- ✅ **Automaattinen moduulihaku** - Ei enää kiinteäkoodattuja listoja
- ✅ **Yhtenäinen käyttöliittymä** - Molemmat skriptit kysyvät: 1) Moduulit, 2) Kohteet
- ✅ **Dynaaminen skannaus** - .bas/.cls-tiedostot haetaan automaattisesti
- ✅ **Parempi käyttökokemus** - Värikoodatut vaiheet ja selkeät tulosteviestit

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
**Versio:** 2.3  
**Status:** ✅ PRODUCTION READY (AutoCAD DVB 64-bit migraatio valmis)
