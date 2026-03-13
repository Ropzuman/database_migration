# Muutosloki

## Tiedosto

USysCheck.bas

## Päivämäärä

2026-03-13

## Kriittiset muutokset

- `api_GetUserName` ja `api_GetComputerName` -declarit päivitetty `PtrSafe`-muotoon.
- 64-bit-käännöstä rikkova ei-PtrSafe-Declare poistettu käytöstä.

## Siivous ja optimointi

- Muutos toteutettu ilman tietokantakirjauslogiikan muutoksia.
- Moduulin käyttäytyminen säilytetty ennallaan, mutta käännösvakaus parannettu.
