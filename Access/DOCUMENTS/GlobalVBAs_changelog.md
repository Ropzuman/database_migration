# Muutosloki

## Tiedosto

GlobalVBAs.vba

## Päivämäärä

2026-03-13

## Kriittiset muutokset

- API-Declare-lauseet yhtenäistetty 64-bit-yhteensopiviksi `PtrSafe`-muotoon.
- `GetUserNameA` ja `GetComputerNameA` säilytetty `nSize As Long`-tyyppisenä DWORD-yhteensopivuuden varmistamiseksi.
- 64-bit-käännöstä estävä ei-PtrSafe-haara poistettu käytännön compile-riskinä.

## Siivous ja optimointi

- Käännösriskit poistettu ilman toiminnallisen liiketoimintalogiikan muutosta.
- Moduulin rakenne säilytetty ennallaan, jotta käyttöönotto pysyy turvallisena.
