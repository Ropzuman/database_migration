# Muutosloki

## Tiedosto

Form_LisääKuviin_ACAD.cls

## Päivämäärä

2026-03-13

## Kriittiset muutokset

- `Sleep`-kutsu päivitetty 64-bit-yhteensopivaksi `PtrSafe`-Declareksi.
- Poistettu compile-riski, jossa ei-PtrSafe-Declare saattoi kaataa M365 64-bit-käännöksen.

## Siivous ja optimointi

- Muutos rajattu API-deklarointiin; lomakkeen liiketoimintalogiikka jätetty koskematta.
- Yhteensopivuusparannus tehty minimaalisella muutospinnalla.
