# Muutosloki

## Tiedosto

Koodit.bas

## Päivämäärä

2026-03-13

## Kriittiset muutokset

- `api_GetUserName` ja `api_GetComputerName` -declarit korjattu `PtrSafe`-muotoon.
- Poistettu 64-bit-käännöstä estävät legacy-declarit.

## Siivous ja optimointi

- Muutokset pidetty minimissä, vain API-yhteensopivuuteen.
- Putki- ja AutoCAD-liiketoimintalogiikka säilytetty ennallaan.
