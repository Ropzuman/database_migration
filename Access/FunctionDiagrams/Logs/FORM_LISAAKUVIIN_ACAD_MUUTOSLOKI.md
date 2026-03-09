# Form_LisääKuviin_ACAD.cls — Muutosloki

**Tiedosto:** `Access/FunctionDiagrams/Form_LisääKuviin_ACAD.cls`
**Päivitetty:** 2026-03-03
**Syy:** 64-bit M365 -migraatio + Option Explicit -korjaukset

---

## 1. AcadInsertPiste UDT + Paikat()-taulukko (moduulitaso)

**Ongelma:** `Paikat()`-taulukkoa käytettiin kolmessa aliohjelmassa (`HaeIPoints`, `TeeKuvat_Click`, `TeeEtusivu_Click`) ilman moduulitason julistusta — aiheuttaa `Compile Error: Variable not defined` Option Explicit -tilassa.

**Korjaus:** Lisätty moduulin alkuun `oDoc`-julistuksen jälkeen:

```vba
Private Type AcadInsertPiste
    Pisteet(0 To 2) As Double
End Type
Private Paikat() As AcadInsertPiste
```

---

## 2. Me.-etuliite kaikille lomakekontrolleille

**Ongelma:** ~20 lomakekontrolliviittausta ilman `Me.`-etuliitettä. Access tulkitsee ne määrittelemättömiksi muuttujiksi Option Explicit -tilassa.

**Korjatut kontrollit ja aliohjelmit:**

| Kontrolli | Aliohjelma/funktio |
|---|---|
| `Loki` | `Form_Load`, `LisaaLokiin` |
| `PohjaHakem` | `HaeTekstit_Click`, `HaeValitutTekstit_Click`, `TeeKuvat_Click`, `TeeEtusivu_Click`, `TallennaUusiKuva`, `ValHakem_Click` |
| `KuvaHakem` | `TeeKuvat_Click`, `TeeEtusivu_Click`, `TallennaUusiKuva`, `ValHakem2_Click` |
| `TOtsTaulukko` | `TeeKuvat_Click`, `TeeEtusivu_Click` |
| `Translate` | `TeeKuvat_Click` |
| `TLang` | `TeeKuvat_Click` |
| `TBlockTaulukko` | `TeeKuvat_Click` |
| `TKyselyt` | `TeeKuvat_Click` |
| `CRef` | `TeeKuvat_Click` |
| `FrontBase` | `TeeEtusivu_Click`, `TallennaUusiKuva` |
| `Rows` | `TeeEtusivu_Click` |
| `Columns` | `TeeEtusivu_Click` |
| `XCoord` | `TeeEtusivu_Click` |
| `YCoord` | `TeeEtusivu_Click` |
| `RefID` | `TeeEtusivu_Click` |
| `TPaikkaBlokki` | `HaeIPoints` |
| `TTitleBlokki` | `VaihdaOtsikkotiedot` |

---

## 3. Väärät kontrollin nimet (runtime "method or data member not found")

**Ongelma:** Koodissa käytettiin lomakkeelta puuttuvia kontrollin nimiä — Access heittää ajonaikaisesti `Error 2465: Method or data member not found`.

| Väärä nimi (koodissa) | Oikea nimi (lomakkeella) | Sijainti |
|---|---|---|
| `TFrontOtsTaulukko` | `TOtsTaulukko` | `TeeEtusivu_Click` rivi ~972 |
| `RefDocumenID` | `RefID` | `TeeEtusivu_Click` rivi ~1012 |

---

## 4. Irrallinen NextPage = "X" kommentoidussa lohkossa

**Ongelma:** `NextPage = "X"` oli jäänyt koodiin aktiivisena rivinä osittain kommentoidun lohkon sisällä — ylikirjoittaa oikean arvon ElseIf-ketjusta.

**Korjaus:** Rivi kommentoitu selityksen kera:

```vba
'NextPage = "X" ' Poistettu: alla oleva If/ElseIf lohko asettaa oikean arvon
```

---

## 5. LisaaLokiin — virhe 2185 (SelStart ilman fokusta)

**Ongelma:** `Me.Loki.SelStart` heittää `Error 2185: You can't reference a property or method for a control unless the control has the focus` kun `LisaaLokiin` kutsutaan pitkän operaation (esim. `TeeKuvat_Click`) aikana ilman fokusta `Loki`-kontrollin.

**Korjaus:** `SelStart`-rivi suojattu `On Error Resume Next` / `On Error GoTo 0` -parilla — vieritys on kosmeettinen eikä keskeytä operaatiota:

```vba
On Error Resume Next
Me.Loki.SelStart = Len(Me.Loki.Value)
On Error GoTo 0
```

---

## Lomakkeen vahvistetut kontrollit

Seuraavat kontrollin nimet on vahvistettu olemassa oleviksi lomakkeella:

`TOtsTaulukko`, `RefID`, `PohjaHakem`, `KuvaHakem`, `FrontBase`, `Rows`, `Columns`, `XCoord`, `YCoord`, `Translate`, `TLang`, `TBlockTaulukko`, `TKyselyt`, `CRef`, `TTitleBlokki`, `TPaikkaBlokki`, `Loki`

---

## Tiedostorakenne (moduulitason julistukset)

```vba
Private Const acSelectionSetAll As Integer = 5
Private Const acSelectionSetPrevious As Integer = 4
Dim Ajastin As Long
Dim min As Integer, sec As Integer
Dim OSMODE As Integer
Dim LokiPituus As Long
Dim oAcad As Object
Public oDoc As Object

Private Type AcadInsertPiste
    Pisteet(0 To 2) As Double
End Type
Private Paikat() As AcadInsertPiste
```
