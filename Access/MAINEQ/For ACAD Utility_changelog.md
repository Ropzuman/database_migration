# Muutosloki

## Tiedosto

For ACAD Utility.bas

## Päivämäärä

2026-03-13

## Kriittiset muutokset

- `GetCursorPos`-Declare päivitetty `PtrSafe`-muotoon 64-bit-yhteensopivuuden varmistamiseksi.
- Poistettu compile-riski 64-bit VBA-ympäristössä.

## Siivous ja optimointi

- Muutos kohdistettu vain API-deklarointiin.
- Muu ACAD-apulogiikka jätetty muuttumattomaksi.
