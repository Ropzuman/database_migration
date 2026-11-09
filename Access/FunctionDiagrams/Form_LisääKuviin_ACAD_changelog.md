# Muutosloki

## Tiedosto

Form_Lis‰‰Kuviin_ACAD.cls

## P‰iv‰m‰‰r‰

2026-03-13

## Kriittiset muutokset

- `Sleep`-kutsu p‰ivitetty 64-bit-yhteensopivaksi `PtrSafe`-Declareksi.
- Poistettu compile-riski, jossa ei-PtrSafe-Declare saattoi kaataa M365 64-bit-k‰‰nnˆksen.

## Siivous ja optimointi

- Muutos rajattu API-deklarointiin; lomakkeen liiketoimintalogiikka j‰tetty koskematta.
- Yhteensopivuusparannus tehty minimaalisella muutospinnalla.
