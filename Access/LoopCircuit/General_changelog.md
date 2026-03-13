# Muutosloki

## Tiedosto

General.bas

## Päivämäärä

2026-03-13

## Kriittiset muutokset

- `api_GetUserName` ja `api_GetComputerName` -declarit päivitetty `PtrSafe`-muotoon.
- 64-bit-käännösvarmuus parannettu poistamalla ei-yhteensopivat Declare-rivit.

## Siivous ja optimointi

- Muutokset kohdistettu vain rajapintadeklaraatioihin.
- Käyttäjäseurannan liiketoimintalogiikka jätetty muuttamatta.
