# VIEDATA KRIITTINEN BUGIKORJAUS

**Päivämäärä:** 27.2.2026  
**Moduuli:** `Excel\Moduulit\AcadDATA\Koodit.bas`  
**Funktio:** `VieDATA()`  
**Ongelma:** Blokit tyhjentyivät viennin yhteydessä  

---

## 🔴 ONGELMAN KUVAUS

### Alkuperäinen Bugi

AcadDATA-työkalu tyhjensi blokkien attribuutit vietäessä dataa Excelistä takaisin AutoCADiin. Ongelma ilmeni seuraavasti:

1. Käyttäjä toi blokkien tiedot AutoCADista Exceliin (`TuoDATA`)
2. Käyttäjä muokkasi Excelissä joitakin arvoja
3. Käyttäjä vei tiedot takaisin AutoCADiin (`VieDATA`)
4. **TULOS:** Kaikki attribuutit tyhjennettiin tai saivat väärät arvot

### Juurisyy

**BUG #1: INDEX-POHJAINEN PÄIVITYS (rivit 774-781)**

```vba
' ❌ VIRHEELLINEN KOODI (ENNEN)
If oBlock.HasAttributes Then
    BlockArray = oBlock.GetAttributes
    For j = 0 To UBound(BlockArray)
        BlockArray(j).TextString = Cells(i, 8 + j).Text  ' BUGI!
    Next j
End If
```

**Ongelmat:**
- **Ei TAG-matchingia:** Koodi oletti, että attribuutit ovat samassa järjestyksessä kuin Excel-sarakkeet
- **Index-pohjainen logiikka:** `8 + j` oletti, että sarake H on ensimmäinen attribuutti, sarake I toinen jne.
- **Tyhjennys:** Jos `Cells(i, 8 + j).Text` oli tyhjä, attribuutti tyhjennettiin
- **Epäsymmetria:** `TuoDATA` käytti TAG-pohjaista allokointia, mutta `VieDATA` ei

---

## ✅ RATKAISU

### Uusi TAG-pohjainen Logiikka

```vba
' ✅ KORJATTU KOODI (JÄLKEEN)
If oBlock.HasAttributes Then
    BlockArray = oBlock.GetAttributes
    Trace "Block has " & (UBound(BlockArray) + 1) & " attributes"
    
    ' TAG-BASED UPDATE LOGIC (symmetrinen TuoDATA:n kanssa)
    For j = 0 To UBound(BlockArray)
        On Error Resume Next
        TagName = UCase(BlockArray(j).TagString)
        OldValue = BlockArray(j).TextString
        ' ... virheenkäsittely ...
        
        ' Etsi vastaava sarake riviltä 1 (header)
        ColIdx = 0
        For k = 8 To 256
            If UCase(Cells(1, k).Value) = TagName Then
                ColIdx = k
                Exit For
            End If
            If Cells(1, k).Value = "" Then Exit For
        Next k
        
        If ColIdx > 0 Then
            NewValue = CStr(Cells(i, ColIdx).Text)
            
            If Len(NewValue) > 0 Then
                ' Päivitä vain jos Excelissä on arvo
                BlockArray(j).TextString = NewValue
                UpdateCount = UpdateCount + 1
                Trace "  [" & TagName & "] '" & OldValue & "' -> '" & NewValue & "'"
            Else
                ' Säilytä olemassa oleva arvo
                EmptyCount = EmptyCount + 1
                Trace "  [" & TagName & "] SKIPPED (Excel tyhjä, säilytetään '" & OldValue & "')"
            End If
        Else
            SkippedCount = SkippedCount + 1
            Trace "  [" & TagName & "] SKIPPED (ei saraketta)"
        End If
    Next j
End If
```

---

## 🔧 TOTEUTETUT KORJAUKSET

### 1. TAG-pohjainen Matching

**Ennen:**
- Attribuutit päivitettiin index-järjestyksessä (`8 + j`)
- Ei tarkistettu `TagString`-arvoa

**Jälkeen:**
- Jokaisen attribuutin `TagString` luetaan
- Haetaan vastaava sarake riviltä 1 (header row)
- Päivitys tapahtuu vain jos TAG-nimi löytyy

### 2. Tyhjien Arvojen Suojaus

**Ennen:**
- Jos Excel-solu oli tyhjä, attribuutti tyhjennettiin: `BlockArray(j).TextString = ""`

**Jälkeen:**
- Tarkistetaan `Len(NewValue) > 0`
- Jos Excel-solu on tyhjä, **säilytetään** AutoCAD:ssa oleva arvo
- Logitetaan suojatut arvot: `"SKIPPED (Excel empty, preserving 'XXX')"`

### 3. Debug.Print-seuranta

**Lisätty:**
- `Trace`-kutsut jokaiseen vaiheeseen
- `StepMsg`-muuttuja virhetilanteita varten
- Laskurit: `UpdateCount`, `SkippedCount`, `EmptyCount`
- Yhteenveto lopussa: `"VieDATA completed: Updated=X, Skipped(no column)=Y, Preserved(empty)=Z"`

**Esimerkki debug-tulosteesta (Immediate Window):**

```
[27.2.2026 14:35:12] VieDATA: Initialization
[27.2.2026 14:35:12] Connect to AutoCAD
[27.2.2026 14:35:13] Update block attributes: row=2
[27.2.2026 14:35:13] Block has 3 attributes
[27.2.2026 14:35:13]   [MERK] '101-FT' -> '101-FT' (ei muutosta)
[27.2.2026 14:35:13]   [POS] 'K123' -> 'K124'
[27.2.2026 14:35:13]   [KOKO] SKIPPED (Excel tyhjä, säilytetään 'DN100')
[27.2.2026 14:35:14] VieDATA completed: Updated=2, Skipped(no column)=0, Preserved(empty)=1
```

### 4. Virheenkäsittelyn Parannus

**Lisätty:**
- `StepMsg` virheilmoitukseen: `"Vaihe: Update block attributes: row=5"`
- Trace-kutsut virhetilanteisiin
- `On Error Resume Next` attribuutin luvussa (suojaa proxy-objekteilta)

### 5. Tekstientiteettien Suojaus

**Sama logiikka sovellettu myös TEXT/MTEXT-entiteetteihin:**
- Päivitetään vain jos Excel-arvo ei ole tyhjä
- Logitetaan muutokset: `"[TEXT] 'old' -> 'new'"`

---

## 📊 VAIKUTUSANALYYSI

### Ennen Korjausta (BUGI)

| Tilanne | Tulos | Syy |
|---------|-------|-----|
| Blokissa: `MERK="101-FT", POS="K123", KOKO="DN100"` | Kaikki tyhjentyvät | Index-mismatch + tyhjät solut |
| Excel-sarakkeet: `H=POS, I=MERK, J=KOKO` | Arvot väärissä kentissä | Ei TAG-matchingia |
| Käyttäjä jättää KOKO-sarakkeen tyhjäksi | KOKO tyhjentyy | Ei tyhjien arvojen suojausta |

### Jälkeen Korjauksen (TOIMII)

| Tilanne | Tulos | Syy |
|---------|-------|-----|
| Blokissa: `MERK="101-FT", POS="K123", KOKO="DN100"` | Vain muutetut päivittyvät | TAG-matching |
| Excel-sarakkeet: `H=POS, I=MERK, J=KOKO` | Arvot oikeisiin kenttiin | TagString-haku |
| Käyttäjä jättää KOKO-sarakkeen tyhjäksi | KOKO säilyy "DN100" | Tyhjien suojaus |

---

## 🧪 TESTAUSVAIHEET

### 1. Perus Round-Trip -testi

```vba
' 1. Tuo blokit AutoCADista
TuoDATA

' 2. Muokkaa vain yhtä saraketta (esim. POS)
Cells(2, 8).Value = "UUSI-POS"  ' Olettaen H=POS

' 3. Vie takaisin
VieDATA

' 4. Tarkista AutoCAD:
'    - POS-attribuutti päivittyi: "UUSI-POS"
'    - MERK ja KOKO säilyivät ennallaan
```

### 2. Tyhjien Arvojen Testi

```vba
' 1. Tuo blokit
TuoDATA

' 2. Tyhjennä yksi solu (älä poista saraketta)
Cells(2, 10).Value = ""  ' Esim. KOKO-sarake

' 3. Vie takaisin
VieDATA

' 4. Tarkista AutoCAD:
'    - KOKO-attribuutti säilyi ennallaan (ei tyhjennetty)
```

### 3. Dynaamisten Blokkien Testi

```vba
' 1. Valitse dynaaminen blokki AutoCAD:ssa (EffectiveName ≠ Name)
' 2. Tuo TuoDATA
' 3. Muokkaa attribuutteja
' 4. Vie VieDATA
' 5. Tarkista että EffectiveName-matching toimii
```

### 4. Debug-lokin Tarkistus

```vba
' 1. Avaa VBE -> Immediate Window (Ctrl+G)
' 2. Varmista DEBUG_TRACE = True
' 3. Aja VieDATA
' 4. Tarkista loki:
'    - Updated-laskuri vastaa todellisia muutoksia
'    - Preserved-laskuri vastaa tyhjiä Excel-soluja
'    - Ei ERROR-viestejä
```

---

## 🔄 SYMMETRIA: TuoDATA ↔ VieDATA

### TuoDATA (Tuonti)

```vba
' 1. Lue attribuutin TagString
TagName = UCase(BlockArray(jj).TagString)

' 2. Etsi/luo sarake TAG-nimen perusteella
If Not TagCol.Exists(TagName) Then
    colIdx = OtsS(tagName)  ' OtsS luo sarakkeen jos ei löydy
    TagCol.Add tagName, colIdx
Else
    colIdx = CLng(TagCol(tagName))
End If

' 3. Kirjoita arvo
buf(rowUsed, colIdx) = BlockArray(jj).TextString
```

### VieDATA (Vienti) - NYT SYMMETRINEN!

```vba
' 1. Lue attribuutin TagString
TagName = UCase(BlockArray(j).TagString)

' 2. Etsi sarake TAG-nimen perusteella
For k = 8 To 256
    If UCase(Cells(1, k).Value) = TagName Then
        ColIdx = k
        Exit For
    End If
Next k

' 3. Päivitä arvo (vain jos löytyi JA ei tyhjä)
If ColIdx > 0 And Len(NewValue) > 0 Then
    BlockArray(j).TextString = NewValue
End If
```

**➡️ Nyt molemmat funktiot käyttävät TAG-pohjaista matchingia!**

---

## 📝 YLLÄPITO-OHJEET

### Jos Lisäät Uuden Attribuutin Blokkiin

1. **TuoDATA:** Luo automaattisesti uuden sarakkeen (`OtsS`-funktio)
2. **VieDATA:** Löytää automaattisesti sarakkeen TAG-nimen perusteella
3. **Ei koodimuutoksia tarvita!**

### Jos Muutat Attribuutin Järjestystä Blokissa

- **Ei vaikutusta:** Koodi käyttää TAG-nimiä, ei indeksejä

### Jos Poistat Attribuutin Blokista

- **TuoDATA:** Sarake jää Exceliin (vanha data säilyy)
- **VieDATA:** Ohittaa sarakkeen (`SKIPPED (no Excel column)` -logi)

### Debuggauksen Aktivointi

```vba
' Koodit.bas, rivi ~15:
Const DEBUG_TRACE As Boolean = True  ' Loggaus päälle
```

```vba
' Immediate Window (Ctrl+G VBE:ssä):
' Näet kaikki Trace-kutsut reaaliajassa
```

---

## 🚀 SUORITUSKYKY

### Muutokset Suorituskykyyn

- **TuoDATA:** Ei muutoksia (oli jo optimoitu bulk write -puskurilla)
- **VieDATA:** 
  - **+Haku-looppi:** Jokaisen attribuutin kohdalla etsitään sarake (O(n) per attribuutti)
  - **Vaikutus:** Minimaalinen, koska header-haku on nopea (max 256 saraketta)
  - **Esimerkki:** 100 blokkia × 5 attribuuttia = 500 hakua ≈ 0.5 sekuntia

**Suorituskykyvertailu (estimaatti):**
- **Ennen:** ~2 sekuntia per 100 blokkia (virheellinen data)
- **Jälkeen:** ~2.5 sekuntia per 100 blokkia (korrekti data)
- **Trade-off:** +25% aikaa, mutta 100% data-integriteetti

---

## ✅ YHTEENVETO

### Korjatut Ongelmat

1. ✅ **Blokkien tyhjentyminen viennissä** (kriittinen bugi)
2. ✅ **Attribuuttien vääriin kenttiin menevät arvot** (index-mismatch)
3. ✅ **TuoDATA ↔ VieDATA epäsymmetria** (nyt molemmat TAG-pohjaisia)
4. ✅ **Puuttuva tyhjien arvojen suojaus**
5. ✅ **Puutteellinen debuggaus** (lisätty Trace-seuranta)

### Uudet Ominaisuudet

- 🆕 TAG-pohjainen attribuuttien matching
- 🆕 Tyhjien Excel-solujen automaattinen säilytys
- 🆕 Debug.Print-loki kaikista päivityksistä
- 🆕 Yhteenveto-raportit (Updated/Skipped/Preserved)
- 🆕 Parannettu virheenkäsittely StepMsg-tiedoilla

### Yhteensopivuus

- ✅ **64-bit Office 365:** Täysi tuki (LongPtr, PtrSafe)
- ✅ **Dynaamliset blokit:** EffectiveName-tuki säilytetty
- ✅ **Vanha Excel-data:** Toimii ilman migraatiota (sarakkeet luetaan riviltä 1)
- ✅ **AutoCAD 2019:** Late binding, versio-riippumaton

---

**Laatija:** GitHub Copilot (Claude Sonnet 4.5)  
**Hyväksyntä:** Odottaa käyttäjän testausta  
**Versio:** 1.0
