# Muutosloki

## Tiedosto

USysCheck.bas

## Päivämäärä

2026-03-13

## Kriittiset muutokset

- `api_GetUserName` ja `api_GetComputerName` -declarit yhtenäistetty `PtrSafe`-muotoon.
- 64-bit Office -käännöksen kannalta ongelmalliset legacy-declarit korjattu.

## Siivous ja optimointi

- Muutokset rajattu API-rajapintaan.
- Käyttäjälokituksen toiminnallinen logiikka säilytetty.
