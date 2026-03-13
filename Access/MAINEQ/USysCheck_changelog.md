# Muutosloki

## Tiedosto

USysCheck.bas

## Päivämäärä

2026-03-13

## Kriittiset muutokset

- `wu_GetUserName` ja `GetOpenFileName` -declarit päivitetty `PtrSafe`-muotoon.
- OPENFILENAME-rakenteen osoitinkentät (`lCustData`, `lpfnHook`) varmistettu 64-bit-yhteensopiviksi.

## Siivous ja optimointi

- Moduulin compile-vakaus parannettu ilman toiminnallisia sivuvaikutuksia.
- Koodi pidetty tarkoituksella minimimuutoksilla ylläpidettävyyden vuoksi.
